"""模板解析器 - 自动识别Word模板中所有可填充字段，输出精简的用户收集清单

支持的字段模式：
1. "标签：值"模式 - 冒号后为空则待填（如 "课程名称："）
2. "标签格+空白格"模式 - 标签文字单独在一个格，紧邻空白格为待填（如 |课程名称|[空白]|）
3. "标签格+占位符格"模式 - 标签文字后有占位符（如 |题号|[%]|）需替换占位符
4. "多列数据行"模式 - 表头行下有空白/占位符行，按表头名识别（如 题型百分比、分数段人数等）
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

# 选项类标签不算待填字段
_OPTION_WORDS = {'选修', '必修', '开卷', '闭卷', '试题库', '试卷库', '教师组题', '是', '否',
                '优秀', '良好', '中等', '及格', '不及格', 'A', 'B', 'C', 'D', '√', '✓', '○', '●',
                '本人阅卷', '同行阅卷', '集体阅卷', '机器阅卷', '其他'}

# 占位符文字（需要替换为实际值的格）
_PLACEHOLDER_TEXTS = {'%', '…', '……', '...', '—', '--', '___', '□', '○'}

# 章节标题行标签（这些行不包含待填字段，只是标题）
_SECTION_TITLE_LABELS = {'试卷分析', '试　卷　分　析', '课程目标与试卷题目的对应关系'}


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
    # 标准编号标题
    if re.match(r'^[一二三四五六七八九十]+[、.．]', text):
        return True
    # 全宽空格标题（如"试　卷　分　析"）
    if text in _SECTION_TITLE_LABELS:
        return True
    # gridSpan跨整行的标题格
    return False


def _is_empty_data_row(unique_cells) -> bool:
    """判断是否是空白数据行：大部分格为空，且不含标签模式。"""
    if not unique_cells:
        return True
    empty = sum(1 for c in unique_cells if not c.text.strip())
    # 含冒号的行 = 标签行
    label_count = sum(1 for c in unique_cells if re.search(r'[：:]', c.text.strip()))
    # 含短标签文字（2-8字）且不含冒号的格 = 可能是标签格
    short_label_count = sum(1 for c in unique_cells
                           if c.text.strip() and 2 <= len(c.text.strip()) <= 8
                           and not re.search(r'[：:]', c.text.strip())
                           and c.text.strip() not in _LABEL_BLACKLIST
                           and c.text.strip() not in _OPTION_WORDS)
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
    # 纯标点符号不算标签
    if re.match(r'^[^\w\u4e00-\u9fff]+$', text):
        return False
    # 黑名单
    if text in _LABEL_BLACKLIST:
        return False
    # 选项类
    if text in _OPTION_WORDS:
        return False
    # 常见表头词不算标签
    header_words = {'试题', '题号', '分数', '得分', '评卷人', '项目', '内容', '签名', '日期'}
    if text in header_words:
        return False
    return True


def _is_placeholder_cell(cell) -> bool:
    """判断单元格是否是占位符格（如%、……等）"""
    text = cell.text.strip()
    return text in _PLACEHOLDER_TEXTS


def _is_vmerge_continue(cell) -> bool:
    """判断单元格是否是垂直合并的延续格（非起始格）"""
    tc = cell._element
    vm = tc.find(qn('w:tcPr'))
    if vm is not None:
        vm_elem = vm.find(qn('w:vMerge'))
        if vm_elem is not None:
            val = vm_elem.get(qn('w:val'))
            return val is None or val == "continue"
    return False


def _get_grid_span(cell) -> int:
    """获取单元格的gridSpan值"""
    tc = cell._element
    tcPr = tc.find(qn('w:tcPr'))
    if tcPr is not None:
        gs = tcPr.find(qn('w:gridSpan'))
        if gs is not None:
            val = gs.get(qn('w:val'))
            if val:
                return int(val)
    return 1


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

        # ── 2. 扫描"标签格+空白格/占位符格"模式 ──
        # 扩展：不仅检查相邻格，还检查中间跳过选项标签格的情况
        # 注意：当一行有"行标签+多列待填"模式时，跳过label_blank扫描（交给multi_col处理）
        for r_idx, row in enumerate(table.rows):
            unique = _get_unique_cells(row)
            if _is_section_title_row(unique):
                continue

            # 检测该行是否是"行标签+多列待填"模式
            # 多列数据行特征：一个标签后连续>=2个空白格/占位符格，且中间没有标签格或选项格
            # 标签-空白交替行特征：标签-空白-标签-空白 交替出现
            # 选项-空白交替行特征：选项-空白-选项-空白 交替出现（勾选框行）
            is_multi_col_row = False
            first_label_idx = -1
            consecutive_fillable = 0
            max_consecutive_fillable = 0
            label_blank_alternating = False
            option_blank_alternating = False
            
            prev_was_label = False
            prev_was_fillable = False
            prev_was_option = False
            for ci, cell in enumerate(unique):
                if _is_vmerge_continue(cell):
                    continue
                text = cell.text.strip()
                is_label = text and _is_label_cell(cell)
                is_option = text and text in _OPTION_WORDS
                is_fillable = not text or _is_placeholder_cell(cell)
                
                if is_label and first_label_idx < 0:
                    first_label_idx = ci
                
                if (is_label or is_option) and prev_was_fillable:
                    if is_label:
                        label_blank_alternating = True
                    if is_option:
                        option_blank_alternating = True
                
                if is_fillable and first_label_idx >= 0:
                    consecutive_fillable += 1
                    max_consecutive_fillable = max(max_consecutive_fillable, consecutive_fillable)
                elif is_label or is_option:
                    consecutive_fillable = 0
                
                prev_was_label = is_label
                prev_was_fillable = is_fillable
                prev_was_option = is_option
            
            # 多列数据行：有行标签，后面连续>=2个待填格，且不是交替模式
            is_multi_col_row = (first_label_idx >= 0 and max_consecutive_fillable >= 2 
                               and not label_blank_alternating and not option_blank_alternating)
            
            i = 0
            while i < len(unique):
                cell = unique[i]
                
                if _is_label_cell(cell):
                    label = cell.text.strip()
                    
                    # 检查这个标签是否已经被冒号模式识别过
                    already_found = any(
                        f["table_idx"] == t_idx and f["row_idx"] == r_idx
                        and f["label"] == label and f["pattern"] == "colon"
                        for f in label_fields
                    )
                    
                    if not already_found and not is_multi_col_row:
                        # 向后寻找第一个空白格或占位符格
                        for j in range(i + 1, min(i + 4, len(unique))):  # 最多往后看3格
                            next_cell = unique[j]
                            next_text = next_cell.text.strip()
                            
                            if _is_vmerge_continue(next_cell):
                                continue
                            
                            if not next_text or _is_placeholder_cell(next_cell):
                                # 找到待填格
                                fill_mode = "set"
                                if _is_placeholder_cell(next_cell):
                                    fill_mode = "replace"  # 替换占位符
                                
                                field_id = f"T{t_idx}_R{r_idx}_C{j}"
                                label_fields.append({
                                    "field_id": field_id,
                                    "table_idx": t_idx,
                                    "row_idx": r_idx,
                                    "col_idx": j,  # 指向待填格
                                    "line_idx": 0,
                                    "label": label,
                                    "existing_value": next_text if next_text else "",
                                    "description": f"请填写{label}",
                                    "fill_mode": fill_mode,
                                    "pattern": "label_blank",
                                })
                                break
                            elif next_text in _OPTION_WORDS:
                                # 跳过选项格继续寻找
                                continue
                            else:
                                # 遇到非选项非空白格，停止寻找
                                break
                i += 1

        # ── 2.5 扫描"表头+数据行"模式（多列数据行） ──
        # 处理试卷分析中的"题型百分比"、"分数段人数/比例"等区域
        # 这些区域特征：表头行有多个标签，下一行有空白格或占位符格待填
        _scan_table_data_sections(table, t_idx, label_fields)

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
    # 注意：占位符字段（fill_mode="replace"）的existing_value是占位符文本，需要保留
    deduped_fields = [f for f in deduped_fields
                      if not f["existing_value"] or f.get("fill_mode") == "replace"]
    
    # ── 5.5 过滤掉完全属于行组区域的multi_col字段 ──
    # 行组区域由行组填充逻辑处理，不需要单独的字段
    # 但保留部分在行组外的multi_col字段
    def _fully_in_row_group(field):
        """如果字段的所有行都在行组区域内，返回True"""
        for ri in field["row_indices"]:
            if (field["table_idx"], ri) not in row_group_rows:
                return False
        return True
    
    deduped_fields = [f for f in deduped_fields if not _fully_in_row_group(f)]

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


def _scan_table_data_sections(table, t_idx: int, label_fields: list):
    """扫描表格中的"表头+数据行"区域，识别多列数据待填字段。
    
    处理的场景：
    1. "标签+空白/占位符"在同行（如 |人数|[空白]|[空白]|...）
    2. "表头行+数据行"配对（如"分数段"行下有"人数"行和"比例"行）
    3. 纯占位符行（如全是%的行，需要替换）
    """
    for r_idx, row in enumerate(table.rows):
        unique = _get_unique_cells(row)
        if len(unique) < 2:
            continue
        if _is_section_title_row(unique):
            continue
        
        # 检测"多列标签+待填"模式
        # 模式A：行内有多个空白格/占位符格，且前面有标签格提供上下文
        # 模式B：行内有"行标签"+"占位符格/空白格"序列（如 |人数|空白|空白|空白|...）
        
        _detect_multi_column_fields(table, t_idx, r_idx, unique, label_fields)


def _detect_multi_column_fields(table, t_idx: int, r_idx: int, unique_cells: list, label_fields: list):
    """检测多列待填字段。
    
    典型场景：
    - 试卷分析行11: |分数分布(续)|人数|空白|空白|空白|空白|空白|
    - 试卷分析行12: |分数分布(续)|比例|%|%|%|%|%|
    - 试卷分析行9: |及百分比|%|%|%|%|%|%|%|%|%|
    """
    # 构建已被label_blank模式覆盖的格集合 (table_idx, row_idx, col_idx)
    covered_cells = set()
    for f in label_fields:
        if f["pattern"] in ("label_blank", "colon"):
            covered_cells.add((f["table_idx"], f["row_idx"], f["col_idx"]))
    
    # 检查此行是否有"行标签+多列待填"模式
    row_label = None
    fillable_cells = []  # (col_idx_in_unique, cell, existing_text)
    
    for ci, cell in enumerate(unique_cells):
        if _is_vmerge_continue(cell):
            continue
        
        text = cell.text.strip()
        
        # 第一个非空非选项格作为行标签
        if row_label is None and text and _is_label_cell(cell):
            row_label = text
            continue
        
        # 跳过已被label_blank模式识别的格
        if (t_idx, r_idx, ci) in covered_cells:
            continue
        
        if not text or _is_placeholder_cell(cell):
            fillable_cells.append((ci, cell, text))
    
    if not fillable_cells or len(fillable_cells) < 2:
        logger.debug(f"multi_col T{t_idx}_R{r_idx}: no fillable cells (count={len(fillable_cells)})")
        return
    
    logger.debug(f"multi_col T{t_idx}_R{r_idx}: row_label={row_label}, fillable_count={len(fillable_cells)}")
    
    # 需要为这些待填格生成有意义的标签
    # 策略1：如果有行标签，用"行标签+列标签"
    # 策略2：看上方行对应的格，获取列标签
    
    if row_label:
        # 清理行标签中的换行符
        row_label = row_label.replace('\n', '').replace('\r', '')
        # 这是一个多列数据行（如人数行、比例行）
        col_labels = _get_column_labels_for_row(table, t_idx, r_idx, unique_cells, fillable_cells)
        
        for idx, (ci, cell, existing) in enumerate(fillable_cells):
            col_label = col_labels[idx] if idx < len(col_labels) else f"第{idx+1}列"
            full_label = f"{row_label}_{col_label}"
            
            # 检查是否已存在
            already = any(f["label"] == full_label and f["table_idx"] == t_idx for f in label_fields)
            if already:
                continue
            
            fill_mode = "replace" if existing else "set"
            field_id = f"T{t_idx}_R{r_idx}_C{ci}"
            label_fields.append({
                "field_id": field_id,
                "table_idx": t_idx,
                "row_idx": r_idx,
                "col_idx": ci,
                "line_idx": 0,
                "label": full_label,
                "existing_value": existing if existing else "",
                "description": f"请填写{row_label}的{col_label}",
                "fill_mode": fill_mode,
                "pattern": "multi_col",
            })
    
    else:
        # 没有行标签，但有多个待填格
        col_labels = _get_column_labels_for_row(table, t_idx, r_idx, unique_cells, fillable_cells)
        above_label = _get_row_label_from_above(table, t_idx, r_idx)
        
        for idx, (ci, cell, existing) in enumerate(fillable_cells):
            col_label = col_labels[idx] if idx < len(col_labels) else f"第{idx+1}列"
            if above_label:
                full_label = f"{above_label}_{col_label}"
            else:
                full_label = col_label
            
            already = any(f["label"] == full_label and f["table_idx"] == t_idx for f in label_fields)
            if already:
                continue
            
            fill_mode = "replace" if existing else "set"
            field_id = f"T{t_idx}_R{r_idx}_C{ci}"
            label_fields.append({
                "field_id": field_id,
                "table_idx": t_idx,
                "row_idx": r_idx,
                "col_idx": ci,
                "line_idx": 0,
                "label": full_label,
                "existing_value": existing if existing else "",
                "description": f"请填写{full_label}",
                "fill_mode": fill_mode,
                "pattern": "multi_col",
            })


def _get_column_labels_for_row(table, t_idx: int, r_idx: int, unique_cells: list, fillable_cells: list) -> list:
    """获取待填格对应的列标签（从上方表头行获取）"""
    col_labels = []
    
    # 向上查找表头行
    for above_r in range(r_idx - 1, max(r_idx - 5, -1), -1):
        above_unique = _get_unique_cells(table.rows[above_r])
        if len(above_unique) < len(fillable_cells):
            continue
        
        for idx, (ci, cell, _) in enumerate(fillable_cells):
            # 尝试从上方行对应位置的格获取列标签
            if ci < len(above_unique):
                above_text = above_unique[ci].text.strip().replace('\n', '')
                if above_text and above_text not in _PLACEHOLDER_TEXTS and len(above_text) <= 15:
                    col_labels.append(above_text)
                else:
                    col_labels.append(f"第{idx+1}列")
            else:
                col_labels.append(f"第{idx+1}列")
        
        # 检查是否获取到有意义的标签
        meaningful = sum(1 for l in col_labels if not l.startswith('第'))
        if meaningful > 0:
            return col_labels
    
    # 没找到，用序号
    return [f"第{idx+1}列" for idx in range(len(fillable_cells))]


def _get_row_label_from_above(table, t_idx: int, r_idx: int) -> Optional[str]:
    """从上方行获取行标签"""
    for above_r in range(r_idx - 1, max(r_idx - 3, -1), -1):
        above_unique = _get_unique_cells(table.rows[above_r])
        if above_unique:
            first_text = above_unique[0].text.strip()
            if first_text and _is_label_cell(above_unique[0]):
                return first_text
    return None


if __name__ == "__main__":
    template_path = os.path.join(
        os.getenv("COZE_WORKSPACE_PATH", "/workspace/projects"),
        "assets", "template_exam_lingnan.docx"
    )
    result = analyze_template(template_path)
    print(json.dumps(result, ensure_ascii=False, indent=2))
