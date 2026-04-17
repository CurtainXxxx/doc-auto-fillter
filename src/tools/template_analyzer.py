"""模板解析器 - 自动识别Word模板中所有可填充字段，输出精简的用户收集清单"""

import os
import re
import json
import logging
from dataclasses import dataclass, field, asdict
from typing import Optional
from docx import Document
from docx.oxml.ns import qn

logger = logging.getLogger(__name__)

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _get_unique_cells(row) -> list:
    """获取行中的独立单元格 (跳过被合并重复的)"""
    seen = set()
    cells = []
    for cell in row.cells:
        cid = id(cell._element)
        if cid not in seen:
            seen.add(cid)
            cells.append(cell)
    return cells


def _is_section_title_row(unique_cells) -> bool:
    """判断是否是章节标题行"""
    if len(unique_cells) != 1:
        return False
    text = unique_cells[0].text.strip()
    return bool(re.match(r'^[一二三四五六七八九十]+[、.．]', text))


def _is_empty_data_row(unique_cells) -> bool:
    """判断是否是空白数据行"""
    if not unique_cells:
        return True
    empty = sum(1 for c in unique_cells if not c.text.strip())
    return empty >= len(unique_cells) * 0.5


def _is_header_row(unique_cells) -> bool:
    """判断是否是表头行"""
    if not unique_cells:
        return False
    non_empty = sum(1 for c in unique_cells if c.text.strip())
    has_colon = any('：' in c.text or ':' in c.text for c in unique_cells)
    return non_empty >= len(unique_cells) * 0.6 and not has_colon


def _find_header_for_row(table, row_idx) -> Optional[int]:
    """向上查找最近的表头行"""
    for r in range(row_idx - 1, -1, -1):
        unique = _get_unique_cells(table.rows[r])
        if _is_header_row(unique):
            return r
    return None


def _extract_labels_from_cell(cell) -> list[tuple[str, str]]:
    """从单元格中提取所有"标签: 值"对，返回 [(label, existing_value), ...]"""
    results = []
    text = cell.text.strip().replace('\r', '')
    if not text:
        return results
    
    for line in text.split('\n'):
        line = line.strip()
        if not line:
            continue
        match = re.match(r'^(.+?)\s*[：:]\s*(.*)', line)
        if match:
            label = match.group(1).strip()
            value = match.group(2).strip()
            results.append((label, value))
    
    return results


def analyze_template(template_path: str) -> dict:
    """
    解析模板文件，返回精简的字段清单和行组信息。
    
    返回结构:
    {
        "label_fields": [          # 标签字段 (冒号后待填)
            {
                "field_id": "T0_R0_C0",
                "table_idx": 0, "row_idx": 0, "col_idx": 0, "line_idx": 0,
                "label": "课程名称",
                "description": "请填写课程名称",
                "fill_mode": "append"
            }
        ],
        "row_groups": [            # 重复数据行组
            {
                "group_id": "T0_G0",
                "table_idx": 0,
                "start_row": 4,
                "template_row_count": 3,
                "column_labels": ["毕业要求", "毕业要求指标点", "课程目标"],
                "header_row_idx": 3
            }
        ],
        "summary": {
            "total_label_fields": 9,
            "total_row_groups": 5,
            "collection_plan": "..."  # 给Agent看的收集计划
        }
    }
    """
    doc = Document(template_path)
    label_fields = []
    row_groups = []

    for t_idx, table in enumerate(doc.tables):
        # ── 1. 扫描 label_value 字段 ──
        for r_idx, row in enumerate(table.rows):
            unique = _get_unique_cells(row)
            if _is_section_title_row(unique):
                continue
            
            for c_idx, cell in enumerate(unique):
                labels = _extract_labels_from_cell(cell)
                for line_idx, (label, existing_value) in enumerate(labels):
                    field_id = f"T{t_idx}_R{r_idx}_C{c_idx}_L{line_idx}"
                    
                    # 去重: 如果同类标签已经在同一列出现(跨行重复), 只记录一次并标注重复次数
                    label_fields.append({
                        "field_id": field_id,
                        "table_idx": t_idx,
                        "row_idx": r_idx,
                        "col_idx": c_idx,
                        "line_idx": line_idx,
                        "label": label,
                        "existing_value": existing_value,
                        "description": f"请填写{label}" if not existing_value else f"{label}(已有:{existing_value})",
                        "fill_mode": "append",
                    })

        # ── 2. 检测重复行组 ──
        r = 0
        while r < len(table.rows):
            unique = _get_unique_cells(table.rows[r])
            if _is_empty_data_row(unique):
                header_row_idx = _find_header_for_row(table, r)
                if header_row_idx is not None:
                    header_unique = _get_unique_cells(table.rows[header_row_idx])
                    col_labels = []
                    for c in header_unique:
                        text = c.text.strip().replace('\n', '|')
                        col_labels.append(text)
                    
                    # 统计连续空行
                    start = r
                    count = 0
                    while r < len(table.rows):
                        row_unique = _get_unique_cells(table.rows[r])
                        if _is_empty_data_row(row_unique):
                            start_unique = _get_unique_cells(table.rows[start])
                            if len(row_unique) == len(start_unique):
                                count += 1
                                r += 1
                                continue
                        break
                    
                    if count > 0:
                        group_id = f"T{t_idx}_G{len(row_groups)}"
                        row_groups.append({
                            "group_id": group_id,
                            "table_idx": t_idx,
                            "start_row": start,
                            "template_row_count": count,
                            "column_labels": col_labels,
                            "header_row_idx": header_row_idx,
                            "header_text": " | ".join(col_labels),
                        })
                        continue
            r += 1

    # ── 3. 去重 label_fields: 同一表格同一列的重复标签合并 ──
    # 统计每个(table_idx, col_idx, label)出现的行列表
    label_occurrences = {}
    for f in label_fields:
        key = (f["table_idx"], f["col_idx"], f["label"])
        if key not in label_occurrences:
            label_occurrences[key] = []
        label_occurrences[key].append(f)
    
    # 合并: 只保留一条记录，附加 row_indices 信息
    deduped_fields = []
    for key, occurrences in label_occurrences.items():
        first = occurrences[0].copy()
        first["row_indices"] = [o["row_idx"] for o in occurrences]
        first["repeat_count"] = len(occurrences)
        if len(occurrences) > 1:
            first["description"] = f"请填写{first['label']} (共{len(occurrences)}处, 每个课程目标各填一次)"
        deduped_fields.append(first)

    # ── 4. 生成收集计划 ──
    # 按表格和行号排序
    deduped_fields.sort(key=lambda f: (f["table_idx"], min(f["row_indices"]), f["col_idx"]))
    
    plan_lines = []
    plan_lines.append("【标签字段 - 冒号后待填】")
    for f in deduped_fields:
        repeat = f" (x{f['repeat_count']})" if f["repeat_count"] > 1 else ""
        plan_lines.append(f"  - {f['label']}{repeat}: {f['description']}")
    
    plan_lines.append("\n【数据行组 - 表格区域待填】")
    for g in row_groups:
        plan_lines.append(f"  - {g['group_id']}: {g['template_row_count']}行x{len(g['column_labels'])}列, 表头: {g['header_text']}")

    result = {
        "label_fields": deduped_fields,
        "row_groups": row_groups,
        "summary": {
            "total_unique_labels": len(deduped_fields),
            "total_row_groups": len(row_groups),
            "collection_plan": "\n".join(plan_lines),
        }
    }

    return result


if __name__ == "__main__":
    template_path = os.path.join(
        os.getenv("COZE_WORKSPACE_PATH", "/workspace/projects"),
        "assets", "template.docx"
    )
    result = analyze_template(template_path)
    print(json.dumps(result, ensure_ascii=False, indent=2))
