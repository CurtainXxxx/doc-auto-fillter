"""模板解析器 - 自动识别Word模板中所有可填充字段，输出精简的用户收集清单

支持两种字段模式：
1. "标签：值"模式 - 冒号后为空则待填（如 "课程名称："）
2. "标签格+空白格"模式 - 标签文字单独在一个格，紧邻空白格为待填（如 |课程名称|[空白]|）
"""

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

# 复杂度阈值：列数超过此值的行组不自动填充
_MAX_SIMPLE_GROUP_COLS = 15

# 已知的标签文字黑名单（这些不是待填字段，而是表头或装饰文字）
_LABEL_BLACKLIST = {'序号', '编号', '合计', '总计', '备注', '说明', '项', '次', '类', '号'}


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
    """判断是否是空白数据行：大部分格为空，且不含标签模式。
    含"："的行属于待填标签行，不应归入纯数据行组。
    也排除"标签格+空白格"模式的行（有标签但非表头）。
    """
    if not unique_cells:
        return True
    empty = sum(1 for c in unique_cells if not c.text.strip())
    # 含冒号的行 = 标签行
    label_count = sum(1 for c in unique_cells if re.search(r'[：:]', c.text.strip()))
    # 含短标签文字（2-8字）且不含冒号的格 = 可能是标签格
    short_label_count = sum(1 for c in unique_cells
                           if c.text.strip() and 2 <= len(c.text.strip()) <= 8
                           and not re.search(r'[：:]', c.text.strip())
                           and c.text.strip() not in _LABEL_BLACKLIST)
    return empty >= len(unique_cells) * 0.5 and label_count == 0 and short_label_count == 0


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
    """从单元格中提取所有"标签: 值"对（冒号模式），返回 [(label, existing_value), ...]"""
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


def _is_label_cell(cell) -> bool:
    """判断单元格是否是"标签格"（有短文字，不含冒号，不是纯数字）"""
    text = cell.text.strip()
    if not text:
        return False
    # 太长的不算标签
    if len(text) > 15:
        return False
    # 含冒号的由冒号模式处理
    if '：' in text or ':' in text:
        return False
    # 纯数字不是标签
    if re.match(r'^[\d.]+$', text):
        return False
    # 黑名单
    if text in _LABEL_BLACKLIST:
        return False
    # 纯标点符号不算标签（如……、---、===等）
    if re.match(r'^[^\w\u4e00-\u9fff]+$', text):
        return False
    # 常见表头词不算标签
    header_words = {'试题', '题号', '分数', '得分', '评卷人', '项目', '内容', '签名', '日期'}
    if text in header_words:
        return False
    # 选项类标签不算（如"选修""必修""开卷""闭卷""试题库""试卷库""教师组题""是""否"等）
    # 这些是勾选项，不是待填字段
    option_words = {'选修', '必修', '开卷', '闭卷', '试题库', '试卷库', '教师组题', '是', '否',
                    '优秀', '良好', '中等', '及格', '不及格', 'A', 'B', 'C', 'D', '√', '✓', '○', '●'}
    if text in option_words:
        return False
    return True


def _is_vmerge_continue(cell) -> bool:
    """判断单元格是否是垂直合并的延续格（非起始格）"""
    tc = cell._element
    vm = tc.find(qn('w:tcPr'))
    if vm is not None:
        vm_elem = vm.find(qn('w:vMerge'))
        if vm_elem is not None:
            val = vm_elem.get(qn('w:val'))
            # val=None 或 val="continue" 都是延续格
            return val is None or val == "continue"
    return False


def analyze_template(template_path: str) -> dict:
    """
    解析模板文件，返回精简的字段清单和行组信息。

    返回结构:
    {
        "label_fields": [...],
        "row_groups": [...],
        "summary": {...}
    }
    """
    doc = Document(template_path)
    label_fields = []
    row_groups = []

    # 记录已经被行组覆盖的行（避免标签字段和行组重复）
    row_group_rows = {}  # {(table_idx, row_idx): group_id}

    for t_idx, table in enumerate(doc.tables):
        # ── 1. 扫描冒号模式字段 ("标签：值") ──
        for r_idx, row in enumerate(table.rows):
            unique = _get_unique_cells(row)
            if _is_section_title_row(unique):
                continue

            for c_idx, cell in enumerate(unique):
                labels = _extract_labels_from_cell(cell)
                for line_idx, (label, existing_value) in enumerate(labels):
                    field_id = f"T{t_idx}_R{r_idx}_C{c_idx}_L{line_idx}"
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
                        "pattern": "colon",
                    })

        # ── 2. 扫描"标签格+空白格"模式 ──
        for r_idx, row in enumerate(table.rows):
            unique = _get_unique_cells(row)
            if _is_section_title_row(unique):
                continue

            i = 0
            while i < len(unique) - 1:
                cell = unique[i]
                next_cell = unique[i + 1]

                # 当前格是标签格，下一格是空白格（待填）
                if (_is_label_cell(cell)
                    and not next_cell.text.strip()
                    and not _is_vmerge_continue(next_cell)):

                    label = cell.text.strip()
                    # 检查这个标签是否已经被冒号模式识别过
                    already_found = any(
                        f["table_idx"] == t_idx and f["row_idx"] == r_idx
                        and f["label"] == label and f["pattern"] == "colon"
                        for f in label_fields
                    )
                    if not already_found:
                        field_id = f"T{t_idx}_R{r_idx}_C{i+1}"
                        label_fields.append({
                            "field_id": field_id,
                            "table_idx": t_idx,
                            "row_idx": r_idx,
                            "col_idx": i + 1,  # 指向空白格
                            "line_idx": 0,
                            "label": label,
                            "existing_value": "",
                            "description": f"请填写{label}",
                            "fill_mode": "set",  # 直接设置空白格
                            "pattern": "label_blank",
                        })
                i += 1

            # 也检查行末尾的独立空白格（前面有标签格，但标签格和空白格之间可能隔了一格）
            # 这种情况较少，暂不处理

        # ── 3. 检测重复行组 ──
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
                                row_group_rows[(t_idx, r)] = f"T{t_idx}_G{len(row_groups)}"
                                r += 1
                                continue
                        break

                    if count > 0:
                        num_cols = len(col_labels)
                        group_id = f"T{t_idx}_G{len(row_groups)}"
                        row_groups.append({
                            "group_id": group_id,
                            "table_idx": t_idx,
                            "start_row": start,
                            "template_row_count": count,
                            "column_labels": col_labels,
                            "header_row_idx": header_row_idx,
                            "header_text": " | ".join(col_labels),
                            "num_cols": num_cols,
                            "is_complex": num_cols > _MAX_SIMPLE_GROUP_COLS,
                        })
                        # 标记行组覆盖的行
                        for marked_r in range(start, start + count):
                            row_group_rows[(t_idx, marked_r)] = group_id
                        continue
            r += 1

    # ── 4. 去重 label_fields: 同一表格同一列的重复标签合并 ──
    label_occurrences = {}
    for f in label_fields:
        key = (f["table_idx"], f["col_idx"], f["label"])
        if key not in label_occurrences:
            label_occurrences[key] = []
        label_occurrences[key].append(f)

    deduped_fields = []
    for key, occurrences in label_occurrences.items():
        first = occurrences[0].copy()
        first["row_indices"] = [o["row_idx"] for o in occurrences]
        first["repeat_count"] = len(occurrences)
        if len(occurrences) > 1:
            first["description"] = f"请填写{first['label']} (共{len(occurrences)}处)"
        deduped_fields.append(first)

    # ── 5. 过滤掉已有值的字段（不需要用户填写） ──
    deduped_fields = [f for f in deduped_fields if not f["existing_value"]]

    # ── 6. 按表格和行号排序 ──
    deduped_fields.sort(key=lambda f: (f["table_idx"], min(f["row_indices"]), f["col_idx"]))

    # ── 7. 标记行组覆盖的字段 ──
    for f in deduped_fields:
        for ri in f["row_indices"]:
            if (f["table_idx"], ri) in row_group_rows:
                f["in_row_group"] = row_group_rows[(f["table_idx"], ri)]
                break

    # ── 8. 生成收集计划 ──
    plan_lines = []
    plan_lines.append("【标签字段 - 待填写】")
    for f in deduped_fields:
        repeat = f" (x{f['repeat_count']})" if f["repeat_count"] > 1 else ""
        in_rg = f" [属于行组{f['in_row_group']}]" if f.get("in_row_group") else ""
        plan_lines.append(f"  - {f['label']}{repeat}: {f['description']}{in_rg}")

    plan_lines.append("\n【数据行组 - 表格区域待填】")
    simple_groups = [g for g in row_groups if not g.get("is_complex", False)]
    complex_groups = [g for g in row_groups if g.get("is_complex", False)]

    for g in simple_groups:
        plan_lines.append(f"  - {g['group_id']}: {g['template_row_count']}行x{g['num_cols']}列, 表头: {g['header_text']}")
    if complex_groups:
        plan_lines.append(f"\n【复杂行组 - 将跳过不填】")
        for g in complex_groups:
            plan_lines.append(f"  - {g['group_id']}: {g['num_cols']}列(>{_MAX_SIMPLE_GROUP_COLS}列), 表头: {g['header_text']}")

    result = {
        "label_fields": deduped_fields,
        "row_groups": row_groups,
        "summary": {
            "total_unique_labels": len(deduped_fields),
            "total_row_groups": len(row_groups),
            "simple_row_groups": len(simple_groups),
            "complex_row_groups": len(complex_groups),
            "collection_plan": "\n".join(plan_lines),
        }
    }

    return result


if __name__ == "__main__":
    template_path = os.path.join(
        os.getenv("COZE_WORKSPACE_PATH", "/workspace/projects"),
        "assets", "template_exam_lingnan.docx"
    )
    result = analyze_template(template_path)
    print(json.dumps(result, ensure_ascii=False, indent=2))
