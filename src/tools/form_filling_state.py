"""
FormFillingState — 表单填充状态机

核心职责：
1. 跟踪每个字段的填充状态（empty / filled / confirmed）
2. 批量写入（预填/知识文件提取）
3. 智能分批获取未填字段（优先关键字段）
4. 序列化/反序列化（跨消息持久化）
5. 与生成引擎对接（直接输出 field_id→value 映射）
"""

import json
import time
from typing import Optional


class FormFillingState:
    """表单填充状态机"""

    # 字段状态
    EMPTY = "empty"        # 未填
    FILLED = "filled"      # 已填但未确认
    CONFIRMED = "confirmed"  # 已确认

    # 关键字段关键词（优先收集）
    PRIORITY_KEYWORDS = [
        "课程名称", "教师", "班级", "学院", "开课", "学期",
        "课程性质", "考核方式", "考试", "命题",
    ]

    def __init__(self, template_name: str = "", template_path: str = "",
                 analysis_result: Optional[dict] = None, session_id: str = ""):
        self.session_id = session_id
        self.template_name = template_name
        self.template_path = template_path
        self.created_at = time.time()
        self.updated_at = time.time()

        # 字段清单（来自 analysis_result）
        self._fields = {}        # {field_id: {label, raw_label, fill_mode, status, value, confidence, source}}
        self._field_order = []   # 保留字段顺序
        self._analysis_result = None

        if analysis_result:
            self.init_from_analysis(analysis_result)

    # ── 初始化 ──

    def init_from_analysis(self, analysis_result: dict):
        """从 analyze_template 结果初始化字段状态"""
        self._analysis_result = analysis_result
        self._fields = {}
        self._field_order = []

        for f in analysis_result.get("label_fields", []):
            fid = f.get("field_id", "")
            self._fields[fid] = {
                "label": f.get("raw_label", f.get("label", "")),
                "raw_label": f.get("raw_label", ""),
                "fill_mode": f.get("fill_mode", "set"),
                "table_idx": f.get("table_idx", -1),
                "col_idx": f.get("col_idx", -1),
                "status": self.EMPTY,
                "value": None,
                "confidence": 0.0,
                "source": None,
                "is_group": f.get("fill_mode") == "group",
                "sub_fields": f.get("sub_field_ids", []),
            }
            self._field_order.append(fid)

    # ── 批量操作 ──

    def bulk_fill(self, values: dict, confidence: float = 0.8, source: str = "user"):
        """批量填入字段值

        Args:
            values: {label或field_id: value}
            confidence: 置信度
            source: 来源（user/prefill/rule）
        Returns:
            (matched_list, unmatched_list) 元组
        """
        matched = []
        unmatched = []
        for key, value in values.items():
            if not value:
                continue
            fid = self._resolve_field_id(key)
            if fid and self._fields[fid]["status"] != self.CONFIRMED:
                self._fields[fid]["status"] = self.FILLED if confidence < 1.0 else self.CONFIRMED
                self._fields[fid]["value"] = str(value)
                self._fields[fid]["confidence"] = confidence
                self._fields[fid]["source"] = source
                matched.append({"field_id": fid, "label": self._fields[fid]["label"], "value": str(value)})
            elif not fid:
                unmatched.append(key)
        self.updated_at = time.time()
        return matched, unmatched

    def confirm_fields(self, field_ids: Optional[list] = None):
        """确认字段值（将 FILLED → CONFIRMED）

        Args:
            field_ids: 要确认的字段ID列表，None=确认所有FILLED字段
        """
        for fid in (field_ids or self._fields.keys()):
            if fid in self._fields and self._fields[fid]["status"] == self.FILLED:
                self._fields[fid]["status"] = self.CONFIRMED
        self.updated_at = time.time()

    def set_field(self, field_id: str, value: str, status: str = "confirmed",
                  confidence: float = 1.0, source: str = "user"):
        """设置单个字段值"""
        if field_id in self._fields:
            self._fields[field_id]["value"] = str(value)
            self._fields[field_id]["status"] = status
            self._fields[field_id]["confidence"] = confidence
            self._fields[field_id]["source"] = source
            self.updated_at = time.time()
            return True
        return False

    # ── 查询 ──

    def get_missing(self, exclude_low_priority: bool = False) -> list:
        """获取未填字段列表

        Returns:
            [{field_id, label, fill_mode, is_group, sub_fields}]
        """
        missing = []
        for fid in self._field_order:
            f = self._fields[fid]
            if f["status"] == self.EMPTY:
                # 行组字段：检查子字段是否全空
                if f["is_group"]:
                    sub_empty = [sf for sf in f["sub_fields"]
                                 if sf in self._fields and self._fields[sf]["status"] == self.EMPTY]
                    if sub_empty:
                        missing.append({
                            "field_id": fid,
                            "label": f["label"],
                            "fill_mode": f["fill_mode"],
                            "is_group": True,
                            "sub_fields": sub_empty,
                        })
                else:
                    missing.append({
                        "field_id": fid,
                        "label": f["label"],
                        "fill_mode": f["fill_mode"],
                        "is_group": False,
                        "sub_fields": [],
                    })
        return missing

    def get_next_batch(self, n: int = 5) -> list:
        """智能分批：优先返回关键字段

        Args:
            n: 每批返回的字段数
        Returns:
            [{field_id, label, fill_mode, is_group}]
        """
        missing = self.get_missing()
        if not missing:
            return []

        # 分两堆：关键字段 + 普通字段
        priority = []
        normal = []
        for f in missing:
            is_priority = any(kw in f["label"] for kw in self.PRIORITY_KEYWORDS)
            if is_priority:
                priority.append(f)
            else:
                normal.append(f)

        # 优先关键字段，不足则补普通字段
        batch = priority[:n]
        if len(batch) < n:
            batch.extend(normal[:n - len(batch)])
        return batch

    def get_needs_review(self) -> list:
        """获取需要审核的低置信度字段"""
        return [
            {"field_id": fid, "label": f["label"], "value": f["value"],
             "confidence": f["confidence"], "source": f["source"]}
            for fid, f in self._fields.items()
            if f["status"] == self.FILLED
        ]

    def get_field_values(self, status_filter: Optional[str] = None) -> dict:
        """获取字段值映射

        Args:
            status_filter: 只返回指定状态的字段，None=返回所有非空
        Returns:
            {field_id: value}
        """
        result = {}
        for fid, f in self._fields.items():
            if f["value"] is not None:
                if status_filter is None or f["status"] == status_filter:
                    result[fid] = f["value"]
        return result

    def get_label_value_map(self) -> dict:
        """获取 label→value 映射（用于生成引擎）"""
        result = {}
        for fid, f in self._fields.items():
            if f["value"] is not None:
                result[f["raw_label"] or f["label"]] = f["value"]
        return result

    def is_complete(self) -> bool:
        """所有非行组字段是否已填"""
        for fid, f in self._fields.items():
            if f["status"] == self.EMPTY and not f["is_group"]:
                return False
        return True

    def get_progress(self) -> dict:
        """获取填写进度统计"""
        total = len([f for f in self._fields.values() if not f["is_group"]])
        filled = len([f for f in self._fields.values()
                       if f["status"] in (self.FILLED, self.CONFIRMED) and not f["is_group"]])
        confirmed = len([f for f in self._fields.values()
                          if f["status"] == self.CONFIRMED and not f["is_group"]])
        needs_review = len([f for f in self._fields.values()
                            if f["status"] == self.FILLED and not f["is_group"]])
        return {
            "total": total,
            "filled": filled,
            "confirmed": confirmed,
            "needs_review": needs_review,
            "empty": total - filled,
            "fill_rate": round(filled / total * 100, 1) if total > 0 else 0,
        }

    def get_summary(self) -> dict:
        """获取摘要信息（兼容 @tool 函数调用）"""
        progress = self.get_progress()
        return {
            "total_fields": progress["total"],
            "filled_count": progress["filled"],
            "empty_count": progress["empty"],
            "confirmed_count": progress["confirmed"],
            "fill_rate": progress["fill_rate"],
            "template_name": self.template_name,
            "template_path": self.template_path,
        }

    def get_missing_important(self) -> list:
        """获取未填的关键字段列表"""
        missing = self.get_missing()
        result = []
        for f in missing:
            is_priority = any(kw in f["label"] for kw in self.PRIORITY_KEYWORDS)
            if is_priority:
                result.append({
                    "field_id": f["field_id"],
                    "label": f["label"],
                    "fill_mode": f["fill_mode"],
                })
        return result

    def get_filled_values(self, status_filter: Optional[str] = None) -> dict:
        """获取已填字段值映射（兼容 @tool 函数调用）"""
        return self.get_field_values(status_filter=status_filter)

    # ── 序列化 ──

    def to_dict(self) -> dict:
        """序列化为字典（可存入缓存或传给前端）"""
        return {
            "template_name": self.template_name,
            "template_path": self.template_path,
            "created_at": self.created_at,
            "updated_at": self.updated_at,
            "fields": {fid: dict(f) for fid, f in self._fields.items()},
            "field_order": self._field_order,
            "progress": self.get_progress(),
        }

    def to_json(self) -> str:
        """序列化为JSON"""
        return json.dumps(self.to_dict(), ensure_ascii=False, indent=2)

    @classmethod
    def from_dict(cls, data: dict) -> "FormFillingState":
        """从字典反序列化"""
        state = cls(
            template_name=data.get("template_name", ""),
            template_path=data.get("template_path", ""),
            session_id=data.get("session_id", ""),
        )
        state.created_at = data.get("created_at", time.time())
        state.updated_at = data.get("updated_at", time.time())
        state._fields = data.get("fields", {})
        state._field_order = data.get("field_order", [])
        return state

    @classmethod
    def from_json(cls, json_str: str) -> "FormFillingState":
        """从JSON反序列化"""
        return cls.from_dict(json.loads(json_str))

    # ── 内部方法 ──

    def _resolve_field_id(self, key: str) -> Optional[str]:
        """将 label 或 field_id 解析为 field_id"""
        # 先精确匹配 field_id
        if key in self._fields:
            return key
        # 再匹配 label / raw_label
        for fid, f in self._fields.items():
            if f["label"] == key or f["raw_label"] == key:
                return fid
        # 模糊匹配
        key_lower = key.lower().strip()
        for fid, f in self._fields.items():
            label_lower = (f["raw_label"] or f["label"]).lower().strip()
            if key_lower in label_lower or label_lower in key_lower:
                return fid
        return None

    def __repr__(self):
        p = self.get_progress()
        return (f"FormFillingState(template={self.template_name!r}, "
                f"progress={p['filled']}/{p['total']} ({p['fill_rate']}%))")
