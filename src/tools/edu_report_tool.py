"""
教务报告生成工具 - 核心逻辑
"""
import os
import json
import copy
import re
import tempfile
from typing import Optional

from docx import Document
from docx.oxml.ns import qn
from langchain.tools import tool

from coze_coding_utils.log.write_log import request_context
from coze_coding_utils.runtime_ctx.context import new_context
from storage.memory.memory_saver import get_memory_saver

from tools.template_analyzer import analyze_template


# ── 模板注册表 ──
TEMPLATE_REGISTRY = {
    "评价报告": "assets/2023-2024-2《xxx》 岭南师范学院专业课程目标达成度评价报告模板.docx",
    "试卷分析": "assets/2023-2024-2《xxx》 试卷分析模板.docx",
    "关联矩阵": "assets/2023-2024-2《xxx》岭南师范学院考题与课程目标及毕业要求关联矩阵表模板.docx",
}

# 复杂行组最大列数阈值
_MAX_SIMPLE_GROUP_COLS = 15


def _get_template_path(name: str) -> str:
    workspace = os.getenv("COZE_WORKSPACE_PATH", "/workspace/projects")
    file = TEMPLATE_REGISTRY.get(name)
    if not file:
        avail = ", ".join(TEMPLATE_REGISTRY.keys())
        raise ValueError(f"未找到模板'{name}'，可用模板: [{avail}]")
    return os.path.join(workspace, file)


# ── 单元格操作 ──

def _get_unique_cells(row):
    """获取行中的独立单元格（通过XML元素去重）。"""
    seen = set()
    unique = []
    for cell in row.cells:
        cid = id(cell._element)
        if cid not in seen:
            seen.add(cid)
            unique.append(cell)
    return unique


def _is_vmerge_continue(cell) -> bool:
    """判断单元格是否为 vMerge=continue（延续格）。"""
    tc = cell._element
    tcPr = tc.find(qn("w:tcPr"))
    if tcPr is not None:
        vm = tcPr.find(qn("w:vMerge"))
        if vm is not None:
            val = vm.get(qn("w:val"))
            if val is None or val == "":
                return True
    return False


def _is_vmerge_restart(cell) -> bool:
    """判断单元格是否为 vMerge=restart。"""
    tc = cell._element
    tcPr = tc.find(qn("w:tcPr"))
    if tcPr is not None:
        vm = tcPr.find(qn("w:vMerge"))
        if vm is not None:
            val = vm.get(qn("w:val"))
            if val == "restart":
                return True
    return False


def _set_tc_text(tc, text: str, rPr_source=None):
    """替换 tc 元素中所有文本内容，保留原有格式。

    策略：
    1. 移除所有内容子元素（w:p、w:tbl 等），仅保留 w:tcPr（边框/宽度等属性）
       - 同时移除嵌套表格（w:tbl），防止旧文本藏在内嵌表中
    2. 对含图表（w:drawing / c:chart）的单元格，保留全部原有内容，仅插入新段落
    3. 创建新段落写入文本，从 rPr_source 复制字体/字号等格式
    """
    # 检查单元格是否含有图表/绘图（需保护，不可删除）
    has_drawing = (len(tc.findall('.//' + qn('w:drawing'))) > 0 or
                   len(tc.findall('.//' + qn('c:chart'))) > 0)

    if not has_drawing:
        # 移除所有内容子元素，只保留 tcPr
        to_remove = []
        for child in tc:
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag != 'tcPr':
                to_remove.append(child)
        for child in to_remove:
            tc.remove(child)

    # 在开头插入新段落
    p = tc.makeelement(qn('w:p'), {})
    tc.insert(0, p)

    # 创建 run
    r = p.makeelement(qn('w:r'), {})

    # 从 rPr_source 复制格式（字体、字号等）
    if rPr_source is not None:
        if rPr_source.tag == qn('w:rPr'):
            r.append(copy.deepcopy(rPr_source))
        else:
            src_rPr = rPr_source.find(qn('w:rPr'))
            if src_rPr is not None:
                r.append(copy.deepcopy(src_rPr))

    t_elem = r.makeelement(qn('w:t'), {})
    t_elem.text = str(text)
    t_elem.set(qn('xml:space'), 'preserve')
    r.append(t_elem)

    # 将 run 添加到段落
    p.append(r)


def _set_cell_text(cell, text: str, rPr_source=None):
    """设置单元格文本（清空后写入），保留格式。"""
    _set_tc_text(cell._element, text, rPr_source=rPr_source)


def _append_value_to_tc_after_label(tc, label: str, value: str):
    """在tc元素中，在label后追加value。"""
    for p in tc.findall(qn("w:p")):
        for r in p.findall(qn("w:r")):
            for t in r.findall(qn("w:t")):
                if t.text and label in t.text:
                    t.text = t.text.replace(label, f"{label}{value}", 1)
                    return True
    # 如果没找到label，直接追加
    paragraphs = tc.findall(qn("w:p"))
    if paragraphs:
        r_elem = paragraphs[0].makeelement(qn("w:r"), {})
        t_elem = r_elem.makeelement(qn("w:t"), {})
        t_elem.text = value
        r_elem.append(t_elem)
        paragraphs[0].append(r_elem)
        return True
    return False


def _add_table_row_after(table, after_row_idx):
    """在指定行后复制一行（用于行组填充）。"""
    src_row = table.rows[after_row_idx]
    new_tr = copy.deepcopy(src_row._tr)
    after_row_idx + 1  # noqa
    src_row._tr.addnext(new_tr)
    return new_tr


# ── 选项词集合 ──
_OPTION_WORDS_SET = frozenset({
    "选修", "必修", "开卷", "闭卷", "是", "否",
    "试题库", "试卷库", "教师组题", "本人阅卷", "同行阅卷",
    "集体阅卷", "机器阅卷", "其他",
})


# ── 勾选框行检测和填充 ──

def _detect_checkbox_row(unique_cells):
    """检测一行是否是勾选框行，返回标签组列表。
    
    勾选框行模式: [行标签] [选项1] [空白1] [选项2] [空白2] ... [行标签2] [选项3] [空白3] ...
    返回: [{"label": 行标签, "option_blanks": {选项文字: 空白格索引}}, ...]
    """
    groups = []
    current_label = None
    current_options = {}
    
    for ci in range(len(unique_cells) - 1):
        cell_text = unique_cells[ci].text.strip()
        next_text = unique_cells[ci + 1].text.strip()
        
        if cell_text in _OPTION_WORDS_SET and not next_text:
            # 找到选项+空白格对
            current_options[cell_text] = ci + 1
        elif cell_text and cell_text not in _OPTION_WORDS_SET and not _is_vmerge_continue(unique_cells[ci]):
            # 遇到非选项标签文字
            # 如果当前组已有选项，先保存
            if current_options and current_label:
                groups.append({"label": current_label, "option_blanks": current_options})
            current_label = cell_text
            current_options = {}
    
    # 保存最后一组
    if current_options and current_label:
        groups.append({"label": current_label, "option_blanks": current_options})
    
    return groups


def _fill_checkbox_row(unique_cells, group, user_value):
    """根据用户值在勾选框组打√。
    
    user_value: 如 "必修" 或 "闭卷" （单选）
    """
    selected = [v.strip() for v in user_value.replace("，", ",").split(",")]
    option_blanks = group["option_blanks"]
    
    # 找到同行中第一个有格式的cell作为格式源
    rPr_source = None
    for cell in unique_cells:
        tc = cell._element
        for p in tc.findall(qn("w:p")):
            for r in p.findall(qn("w:r")):
                rPr = r.find(qn("w:rPr"))
                if rPr is not None:
                    rPr_source = rPr
                    break
            if rPr_source is not None:
                break
        if rPr_source is not None:
            break
    
    for opt_text, blank_idx in option_blanks.items():
        if opt_text in selected:
            _set_cell_text(unique_cells[blank_idx], "√", rPr_source=rPr_source)
        # 未选中的空白格保持空白


# ── 标签字段填充 ──

def _find_label_rPr_in_row(tr):
    """从一行中找到第一个有rPr格式的标签格，返回其rPr元素（深拷贝）。
    
    用于空白格填充时继承同行标签格的字体/字号等格式。
    """
    tcs = tr.findall(qn("w:tc"))
    for tc in tcs:
        for p in tc.findall(qn("w:p")):
            for r in p.findall(qn("w:r")):
                rPr = r.find(qn("w:rPr"))
                if rPr is not None:
                    return rPr
    return None


def _fill_label_fields(doc, label_fields, data):
    """填充标签字段，支持多种填充模式。保留原有格式。"""
    for f in label_fields:
        label = f["label"]
        if label not in data:
            continue
        value = data[label]
        fill_mode = f.get("fill_mode", "append")
        pattern = f.get("pattern", "colon")
        
        for ri in f["row_indices"]:
            t_idx = f["table_idx"]
            table = doc.tables[t_idx]
            
            if fill_mode == "set":
                # 标签格+空白格模式：直接设置空白格内容
                col_idx = f["col_idx"]
                # 用tr级别获取tc
                tr = table.rows[ri]._tr
                tcs = tr.findall(qn("w:tc"))
                if col_idx < len(tcs):
                    # 获取同行标签格的格式作为格式源
                    rPr_source = _find_label_rPr_in_row(tr)
                    _set_tc_text(tcs[col_idx], str(value), rPr_source=rPr_source)
            
            elif fill_mode == "replace":
                # 占位符替换模式：清空格内容后写入新值（保留原run格式）
                col_idx = f["col_idx"]
                tr = table.rows[ri]._tr
                tcs = tr.findall(qn("w:tc"))
                if col_idx < len(tcs):
                    rPr_source = _find_label_rPr_in_row(tr)
                    _set_tc_text(tcs[col_idx], str(value), rPr_source=rPr_source)
            
            elif fill_mode == "append":
                # 冒号模式：在标签后追加值
                col_idx = f["col_idx"]
                tr = table.rows[ri]._tr
                tcs = tr.findall(qn("w:tc"))
                if col_idx < len(tcs):
                    _append_value_to_tc_after_label(tcs[col_idx], label, str(value))


def _fill_checkbox_rows_in_table(doc, analysis, data):
    """扫描所有表格行，检测勾选框行并填充。"""
    for t_idx, table in enumerate(doc.tables):
        for ri, row in enumerate(table.rows):
            unique = _get_unique_cells(row)
            groups = _detect_checkbox_row(unique)
            
            if not groups:
                continue
            
            # 对每个标签组，查找data中对应的值并填充
            for group in groups:
                row_label = group["label"]
                user_value = None
                
                # 1. 精确匹配
                if row_label in data:
                    user_value = data[row_label]
                else:
                    # 2. 模糊匹配
                    clean_label = row_label.replace("\n", "").replace(" ", "")
                    for key in data:
                        clean_key = key.replace("\n", "").replace(" ", "")
                        if clean_key == clean_label:
                            user_value = data[key]
                            break
                
                # 3. 检查选项本身是否在data中
                if user_value is None:
                    for opt in group["option_blanks"]:
                        if opt in data:
                            user_value = data[opt]
                            break
                
                if user_value is not None:
                    _fill_checkbox_row(unique, group, user_value)


def _fill_simple_row_groups(doc, groups, data):
    """填充简单行组（重复行数据）。"""
    for g in groups:
        gid = g["group_id"]
        if gid not in data:
            continue
        
        row_data_list = data[gid]
        if not isinstance(row_data_list, list):
            continue
        
        table = doc.tables[g["table_idx"]]
        start_row = g["start_row"]
        template_row_count = g["template_row_count"]
        
        # 填充模板行
        for row_offset, row_data in enumerate(row_data_list):
            if row_offset >= template_row_count:
                break
            actual_row = start_row + row_offset
            
            # 使用tr级别的tc元素
            tr = table.rows[actual_row]._tr
            tcs = tr.findall(qn("w:tc"))
            
            data_idx = 0
            for ti, tc in enumerate(tcs):
                if data_idx >= len(row_data):
                    break
                
                # 检查vMerge
                tcPr = tc.find(qn("w:tcPr"))
                if tcPr is not None:
                    vm = tcPr.find(qn("w:vMerge"))
                    if vm is not None:
                        val = vm.get(qn("w:val"))
                        if val is None or val == "":
                            # vMerge=continue：消耗数据但不填值
                            data_idx += 1
                            continue
                
                # 获取格内文本
                all_text = []
                for p in tc.findall(qn("w:p")):
                    for r in p.findall(qn("w:r")):
                        for t in r.findall(qn("w:t")):
                            if t.text:
                                all_text.append(t.text)
                text = "".join(all_text).strip()
                
                # 空白格或占位符格：填入数据
                if not text or text in ("%", "…", "……"):
                    rPr_source = _find_label_rPr_in_row(tr)
                    _set_tc_text(tc, str(row_data[data_idx]), rPr_source=rPr_source)
                    data_idx += 1
                else:
                    # 有文字的格：如果不是标签则跳过
                    pass
        
        # 如果数据行数 > 模板行数，需要复制行
        if len(row_data_list) > template_row_count:
            for extra_idx in range(template_row_count, len(row_data_list)):
                # 复制最后一行
                last_row = table.rows[start_row + template_row_count - 1]
                new_tr = copy.deepcopy(last_row._tr)
                last_row._tr.addnext(new_tr)
                
                # 填充新行
                tcs = new_tr.findall(qn("w:tc"))
                row_data = row_data_list[extra_idx]
                data_idx = 0
                for ti, tc in enumerate(tcs):
                    if data_idx >= len(row_data):
                        break
                    tcPr = tc.find(qn("w:tcPr"))
                    if tcPr is not None:
                        vm = tcPr.find(qn("w:vMerge"))
                        if vm is not None:
                            val = vm.get(qn("w:val"))
                            if val is None or val == "":
                                data_idx += 1
                                continue
                    all_text = []
                    for p in tc.findall(qn("w:p")):
                        for r in p.findall(qn("w:r")):
                            for t in r.findall(qn("w:t")):
                                if t.text:
                                    all_text.append(t.text)
                    text = "".join(all_text).strip()
                    if not text or text in ("%", "…", "……"):
                        _set_tc_text(tc, str(row_data[data_idx]))
                        data_idx += 1


# ── 用户字段简化 ──

def _simplify_fields(label_fields):
    """将内部字段列表简化为用户友好的字段列表。
    
    将 _第1列 等子字段归组到基础字段下。
    """
    # 子字段后缀列表
    sub_suffixes = ["_第1列", "_第2列", "_第3列", "_第4列", "_第5列",
                    "_一题", "_二题", "_三题", "_四题", "_五题",
                    "_六题", "_七题", "_八题", "_九题",
                    "_<60", "_60-69", "_70-79", "_80-89", "_90-100",
                    "_应到", "_实到", "_缺考", "_缓考", "_作弊", "_取消考试资格"]
    
    groups = {}  # base_label → {"type": "single"/"group", ...}
    field_order = []
    
    for f in label_fields:
        label = f["label"]
        
        base = None
        suffix = None
        for s in sub_suffixes:
            if label.endswith(s):
                base = label[:-len(s)]
                suffix = s[1:]
                break
        
        if base and suffix:
            if base in groups and groups[base]["type"] == "group":
                # 追加到已有组
                groups[base]["sub_labels"].append(suffix)
                groups[base]["sub_fields"].append(f)
            elif base in groups and groups[base]["type"] == "single":
                # 升级为group
                existing = groups[base]["field"]
                groups[base] = {
                    "type": "group",
                    "base_label": base,
                    "sub_fields": [existing, f],
                    "sub_labels": [base, suffix],
                }
            else:
                if base not in groups:
                    field_order.append(base)
                groups[base] = {
                    "type": "group",
                    "base_label": base,
                    "sub_fields": [f],
                    "sub_labels": [suffix],
                }
        else:
            if label not in groups:
                field_order.append(label)
                groups[label] = {
                    "type": "single",
                    "base_label": label,
                    "field": f,
                }
    
    user_fields = []
    for label in field_order:
        g = groups[label]
        if g["type"] == "single":
            f = g["field"]
            user_fields.append({
                "label": f["label"],
                "description": f.get("description", f"请填写{f['label']}"),
                "fill_mode": f["fill_mode"],
            })
        else:
            # 分组字段 - 检查是否需要拆分考勤数据
            subs = g["sub_labels"]
            base = g["base_label"]
            
            # 检查是否包含考勤子字段
            attendance_suffixes = {"应到", "实到", "缺考", "缓考", "作弊", "取消考试资格"}
            has_attendance = any(s in attendance_suffixes for s in subs)
            
            if has_attendance and base != "考勤":
                # 拆分：base字段 + 考勤组
                # 先添加base字段，跳过"_第N列"等内部子字段
                internal_suffixes = {"第1列", "第2列", "第3列", "第4列", "第5列"}
                base_fields = [sf for sf in g["sub_fields"] 
                              if not any(sf["label"].endswith(f"_{s}") for s in attendance_suffixes)
                              and not any(sf["label"].endswith(f"_{s}") for s in internal_suffixes)]
                for sf in base_fields:
                    user_fields.append({
                        "label": sf["label"],
                        "description": sf.get("description", f"请填写{sf['label']}"),
                        "fill_mode": sf["fill_mode"],
                    })
                # 添加考勤组
                att_labels = [s for s in subs if s in attendance_suffixes]
                if att_labels:
                    user_fields.append({
                        "label": "考勤数据",
                        "description": "请填写考勤数据：应到、实到、缺考、缓考、作弊、取消考试资格人数",
                        "fill_mode": "group",
                        "sub_labels": att_labels,
                    })
            else:
                user_fields.append({
                    "label": base,
                    "description": f"请填写{base}的相关数据",
                    "fill_mode": "group",
                    "sub_labels": subs,
                })
    
    return user_fields


# ── 数据扩展 ──

def _expand_report_data(template_path: str, user_data: dict) -> dict:
    """将用户友好的字段数据扩展为完整的内部字段数据。
    
    处理逻辑：
    1. 逗号分隔的分组字段（如"及百分比": "15,15,20,..."）自动展开为子字段
    2. 勾选框字段（如"课程性质": "必修"）由_fill_checkbox_rows_in_table处理，此处原样保留
    3. 考勤数据自动映射到子字段
    4. 分数段比例自动计算（如果用户未提供）
    """
    analysis = analyze_template(template_path)
    label_fields = analysis["label_fields"]
    
    # 构建基础名→子字段映射
    sub_suffixes = ["_第1列", "_第2列", "_第3列", "_第4列", "_第5列",
                    "_一题", "_二题", "_三题", "_四题", "_五题",
                    "_六题", "_七题", "_八题", "_九题",
                    "_<60", "_60-69", "_70-79", "_80-89", "_90-100",
                    "_应到", "_实到", "_缺考", "_缓考", "_作弊", "_取消考试资格"]
    
    base_to_subs = {}
    for f in label_fields:
        label = f["label"]
        for s in sub_suffixes:
            if label.endswith(s):
                base = label[:-len(s)]
                if base not in base_to_subs:
                    base_to_subs[base] = []
                base_to_subs[base].append({"full_label": label, "suffix": s[1:], "field": f})
                break
    
    expanded = {}
    
    # 考勤子字段名映射
    attendance_suffixes = {"应到", "实到", "缺考", "缓考", "作弊", "取消考试资格"}
    
    # 先处理独立的考勤字段（应到=45, 实到=43等），映射到full_label
    for key, value in list(user_data.items()):
        if key in attendance_suffixes:
            # 找到匹配的full_label（如 学生班级_应到）
            for base_name, subs in base_to_subs.items():
                for sub in subs:
                    if sub["suffix"] == key:
                        expanded[sub["full_label"]] = str(value)
                        if key in user_data:
                            del user_data[key]  # 已处理，避免重复
    
    for key, value in user_data.items():
        if key == "考勤数据":
            # 处理考勤数据组
            if isinstance(value, dict):
                for att_key, att_val in value.items():
                    # 查找匹配的考勤子字段
                    for base_name, subs in base_to_subs.items():
                        for sub in subs:
                            if sub["suffix"] == att_key:
                                expanded[sub["full_label"]] = str(att_val)
            elif isinstance(value, str) and "," in value:
                # 逗号分隔格式：应到,实到,缺考,缓考,作弊,取消资格
                parts = [p.strip() for p in value.split(",")]
                # 按考勤子字段的顺序映射
                for base_name, subs in base_to_subs.items():
                    att_subs = [s for s in subs if s["suffix"] in attendance_suffixes]
                    att_subs.sort(key=lambda s: ["应到", "实到", "缺考", "缓考", "作弊", "取消考试资格"].index(s["suffix"]))
                    for i, sub in enumerate(att_subs):
                        if i < len(parts):
                            expanded[sub["full_label"]] = parts[i]
            continue
        
        if key in base_to_subs:
            subs = base_to_subs[key]
            # 检查是否包含考勤子字段
            has_attendance = any(s["suffix"] in attendance_suffixes for s in subs)
            
            if has_attendance:
                # 拆分：先填非考勤子字段，考勤数据另存
                non_att_subs = [s for s in subs if s["suffix"] not in attendance_suffixes]
                if isinstance(value, str) and "," in value and non_att_subs:
                    parts = [p.strip() for p in value.split(",")]
                    for i, sub in enumerate(non_att_subs):
                        if i < len(parts):
                            expanded[sub["full_label"]] = parts[i]
                else:
                    expanded[key] = value
            elif isinstance(value, list) and len(value) == len(subs):
                # 列表值，按顺序映射
                for i, sub in enumerate(subs):
                    expanded[sub["full_label"]] = value[i]
            elif isinstance(value, str) and "," in value:
                # 逗号分隔值，按顺序映射到子字段
                parts = [p.strip() for p in value.split(",")]
                for i, sub in enumerate(subs):
                    if i < len(parts):
                        expanded[sub["full_label"]] = parts[i]
            else:
                # 单个值，保留原样（勾选框逻辑由_fill_checkbox_rows_in_table处理）
                expanded[key] = value
        else:
            # 直接字段，原样保留
            expanded[key] = value
    
    # 自动计算分数段比例（如果用户未提供但有分数段人数）
    dist_bases = [b for b in base_to_subs if b.startswith("分数分布")]
    for base_name in dist_bases:
        subs = base_to_subs[base_name]
        ratio_subs = [s for s in subs if s["suffix"].startswith("第")]
        count_subs = [s for s in subs if s["suffix"] in ["<60", "60-69", "70-79", "80-89", "90-100"]]
        
        if ratio_subs and count_subs:
            # 检查是否已有比例数据
            has_ratio = any(s["full_label"] in expanded for s in ratio_subs)
            if not has_ratio:
                # 计算比例
                total = 0
                counts = []
                for s in count_subs:
                    v = expanded.get(s["full_label"], "0")
                    try:
                        n = float(v)
                    except (ValueError, TypeError):
                        n = 0
                    counts.append(n)
                    total += n
                
                if total > 0:
                    for i, s in enumerate(ratio_subs):
                        if i < len(counts):
                            pct = counts[i] / total * 100
                            expanded[s["full_label"]] = f"{pct:.1f}%"
    
    return expanded


# ── 主生成函数 ──

def _build_report_docx(template_path: str, user_data: dict) -> bytes:
    """根据模板和用户数据生成填充后的docx文件。"""
    analysis = analyze_template(template_path)
    doc = Document(template_path)
    
    # 1. 扩展用户数据
    expanded_data = _expand_report_data(template_path, user_data)
    
    # 2. 识别勾选框行，排除与勾选框冲突的标签字段
    checkbox_rows = set()  # {(table_idx, row_idx)}
    for t_idx, table in enumerate(doc.tables):
        for r_idx, row in enumerate(table.rows):
            unique = _get_unique_cells(row)
            groups = _detect_checkbox_row(unique)
            if groups:
                checkbox_rows.add((t_idx, r_idx))
    
    # 构建勾选框空白格的位置集合 (table_idx, row_idx, col_idx)
    checkbox_blank_cells = set()
    for t_idx, table in enumerate(doc.tables):
        for r_idx, row in enumerate(table.rows):
            unique = _get_unique_cells(row)
            groups = _detect_checkbox_row(unique)
            for g in groups:
                for ci in g["option_blanks"].values():
                    checkbox_blank_cells.add((t_idx, r_idx, ci))

    # 过滤掉勾选框空白格位置的label字段（避免重复填充冲突）
    # 但保留勾选框行中非空白格位置的label字段（如"教师姓名"）
    filtered_fields = []
    for f in analysis["label_fields"]:
        t_idx = f["table_idx"]
        skip = False
        for r_idx in f["row_indices"]:
            if (t_idx, r_idx, f["col_idx"]) in checkbox_blank_cells:
                skip = True
                break
        if not skip:
            filtered_fields.append(f)
    
    # 3. 填充标签字段（不含勾选框行）
    _fill_label_fields(doc, filtered_fields, expanded_data)
    
    # 4. 填充勾选框行（独立于标签字段，智能识别选项+空白格模式）
    _fill_checkbox_rows_in_table(doc, analysis, expanded_data)
    
    # 5. 填充行组
    _fill_simple_row_groups(doc, analysis["row_groups"], expanded_data)
    
    # 保存到临时文件
    tmp = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
    doc.save(tmp.name)
    tmp.seek(0)
    content = tmp.read()
    tmp.close()
    os.unlink(tmp.name)
    
    return content


def _fill_custom_template(template_path: str, user_data: dict) -> bytes:
    """通用模板纯填充：只向空白格/待填格写入文本，绝不改变文档格式。
    
    与 _build_report_docx 的区别：
    - 不调用 _expand_report_data（内置模板专用扩展逻辑）
    - 直接将用户数据映射到识别出的字段位置
    - 完整保留原有字体、字号、加粗、对齐等格式
    """
    analysis = analyze_template(template_path)
    doc = Document(template_path)
    
    # 1. 构建勾选框行检测，排除冲突的标签字段
    checkbox_blank_cells = set()
    for t_idx, table in enumerate(doc.tables):
        for r_idx, row in enumerate(table.rows):
            unique = _get_unique_cells(row)
            groups = _detect_checkbox_row(unique)
            for g in groups:
                for ci in g["option_blanks"].values():
                    checkbox_blank_cells.add((t_idx, r_idx, ci))
    
    # 过滤掉勾选框空白格位置的label字段
    filtered_fields = []
    for f in analysis["label_fields"]:
        t_idx = f["table_idx"]
        skip = False
        for r_idx in f["row_indices"]:
            if (t_idx, r_idx, f["col_idx"]) in checkbox_blank_cells:
                skip = True
                break
        if not skip:
            filtered_fields.append(f)
    
    # 2. 填充标签字段（直接用用户数据，不做expand）
    _fill_label_fields(doc, filtered_fields, user_data)
    
    # 3. 填充勾选框行
    _fill_checkbox_rows_in_table(doc, analysis, user_data)
    
    # 4. 填充行组（直接用用户数据中的行组）
    _fill_simple_row_groups(doc, analysis["row_groups"], user_data)
    
    # 5. 填充多列字段（multi_col字段，如评价报告中的课程目标表）
    for f in analysis["label_fields"]:
        if f.get("pattern") == "multi_col" and f["label"] in user_data:
            _fill_multi_col_field(doc, f, user_data[f["label"]])
    
    # 保存
    tmp = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
    doc.save(tmp.name)
    tmp.seek(0)
    content = tmp.read()
    tmp.close()
    os.unlink(tmp.name)
    
    return content


def _fill_multi_col_field(doc, field, value):
    """填充multi_col模式字段（一行内多个待填列）。"""
    t_idx = field["table_idx"]
    table = doc.tables[t_idx]
    
    # value可能是逗号分隔字符串或列表
    if isinstance(value, str):
        values = [v.strip() for v in value.replace("，", ",").split(",")]
    else:
        values = list(value)
    
    # 获取可填充列索引
    fillable_cols = field.get("fillable_cols", [])
    
    for i, col_idx in enumerate(fillable_cols):
        if i >= len(values):
            break
        for ri in field["row_indices"]:
            tr = table.rows[ri]._tr
            tcs = tr.findall(qn("w:tc"))
            if col_idx < len(tcs):
                rPr_source = _find_label_rPr_in_row(tr)
                _set_tc_text(tcs[col_idx], str(values[i]), rPr_source=rPr_source)


# ── 工具函数 ──

@tool
def list_templates() -> str:
    """列出所有可用的教务报告模板。"""
    templates = []
    for name, path in TEMPLATE_REGISTRY.items():
        templates.append({"name": name, "file": os.path.basename(path)})
    return json.dumps({"success": True, "templates": templates}, ensure_ascii=False)


@tool
def analyze_report_template(template_name: str) -> str:
    """解析指定模板，返回用户友好的字段清单和信息收集指引。
    在开始收集用户信息前必须先调用此工具。

    Args:
        template_name: 模板名称，从list_templates返回的名称中选择
    """
    try:
        path = _get_template_path(template_name)
        analysis = analyze_template(path)
        
        # 简化字段列表
        user_fields = _simplify_fields(analysis["label_fields"])
        
        # 构建收集指引
        guide_parts = []
        guide_parts.append(f"模板【{template_name}】共需填写 {len(user_fields)} 项信息：\n")
        
        for i, f in enumerate(user_fields, 1):
            if f["fill_mode"] == "group":
                subs = f.get("sub_labels", [])
                if len(subs) <= 5:
                    guide_parts.append(f"  {i}. {f['label']}（{', '.join(subs)}）")
                else:
                    guide_parts.append(f"  {i}. {f['label']}（共{len(subs)}项，可用逗号分隔）")
            else:
                guide_parts.append(f"  {i}. {f['label']}")
        
        guide_parts.append('\n提示：多值字段可用逗号分隔一次性提供，如"及百分比: 15,15,20,20,10,10,5,5,0"')
        
        return json.dumps({
            "success": True,
            "template_name": template_name,
            "total_fields": len(user_fields),
            "user_fields": user_fields,
            "row_groups_count": analysis["summary"]["total_row_groups"],
            "collection_guide": "\n".join(guide_parts),
        }, ensure_ascii=False)
        
    except Exception as e:
        return json.dumps({"success": False, "message": f"模板解析失败: {e}"}, ensure_ascii=False)


@tool
def generate_edu_report(template_name: str, report_data: str) -> str:
    """根据预设模板和用户数据生成教务报告文档（仅支持预设的3种模板）。

    Args:
        template_name: 模板名称（评价报告/试卷分析/关联矩阵）
        report_data: JSON格式的报告数据，键为字段名，值为字段值。
                     多值字段用逗号分隔，如 {"及百分比": "15,15,20,20,10,10,5,5,0"}
                     行组数据用二维数组，如 {"T0_G0": [["值1","值2"],...]}
    """
    ctx = request_context.get() or new_context(method="generate_edu_report")
    
    try:
        path = _get_template_path(template_name)
        
        # 解析报告数据
        if isinstance(report_data, str):
            data = json.loads(report_data)
        else:
            data = report_data
        
        # 生成文档
        docx_bytes = _build_report_docx(path, data)
        
        # 上传到对象存储
        from coze_coding_dev_sdk.s3 import S3SyncStorage
        import time
        
        storage = S3SyncStorage(
            endpoint_url=os.getenv("COZE_BUCKET_ENDPOINT_URL"),
            access_key="",
            secret_key="",
            bucket_name=os.getenv("COZE_BUCKET_NAME"),
            region="cn-beijing",
        )
        
        timestamp = time.strftime("%Y%m%d%H%M%S")
        safe_name = template_name.replace(" ", "_")
        file_name = f"edu_report/{safe_name}_{timestamp}.docx"
        
        file_key = storage.upload_file(
            file_content=docx_bytes,
            file_name=file_name,
            content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        url = storage.generate_presigned_url(key=file_key, expire_time=86400)
        
        return json.dumps({
            "success": True,
            "message": "报告已成功生成并上传",
            "file_name": file_name,
            "download_url": url,
        }, ensure_ascii=False)
        
    except Exception as e:
        return json.dumps({
            "success": False,
            "message": f"报告生成失败: {e}",
        }, ensure_ascii=False)


# ═══════════════════════════════════════════════════════════════
# 通用模板工具：支持任意docx文件的自动识别和填充
# ═══════════════════════════════════════════════════════════════

@tool
def analyze_uploaded_template(file_path: str) -> str:
    """分析用户上传的任意Word模板文件，自动识别待填字段。
    支持识别：冒号字段、标签+空白格、勾选框、占位符、多列数据行、行组等。

    Args:
        file_path: 上传的docx文件路径（通常是上传后的临时文件路径）
    """
    ctx = request_context.get() or new_context(method="analyze_uploaded_template")
    
    try:
        import os
        workspace = os.getenv("COZE_WORKSPACE_PATH", "/workspace/projects")
        if not os.path.isabs(file_path):
            full_path = os.path.join(workspace, file_path)
        else:
            full_path = file_path
        
        if not os.path.exists(full_path):
            return json.dumps({"success": False, "message": f"文件不存在: {file_path}"}, ensure_ascii=False)
        
        analysis = analyze_template(full_path)
        user_fields = _simplify_fields(analysis['label_fields'])
        
        # 构建用户友好的字段描述
        field_descriptions = []
        for f in user_fields:
            if f['fill_mode'] == 'group':
                subs = f['sub_labels']
                field_descriptions.append({
                    "label": f['label'],
                    "type": "group",
                    "sub_items": subs,
                    "hint": f"请提供{len(subs)}个值，用逗号分隔",
                })
            else:
                desc = {"label": f['label'], "type": "single"}
                # 添加提示
                label_lower = f['label']
                if any(k in label_lower for k in ['签字', '签名']):
                    desc['hint'] = "请填写姓名"
                elif any(k in label_lower for k in ['日期', '时间']):
                    desc['hint'] = "请填写日期"
                elif any(k in label_lower for k in ['百分比', '比例', '占比']):
                    desc['hint'] = "请填写数值"
                field_descriptions.append(desc)
        
        # 行组信息
        row_group_info = []
        for g in analysis['row_groups']:
            row_group_info.append({
                "group_id": g['group_id'],
                "num_cols": g['num_cols'],
                "template_row_count": g['template_row_count'],
                "column_labels": g.get('column_labels', []),
            })
        
        return json.dumps({
            "success": True,
            "file_name": os.path.basename(full_path),
            "total_fields": len(user_fields),
            "fields": field_descriptions,
            "row_groups": row_group_info,
            "summary": f"识别到{len(user_fields)}个待填字段和{len(row_group_info)}个数据行组",
        }, ensure_ascii=False)
        
    except Exception as e:
        return json.dumps({"success": False, "message": f"模板分析失败: {e}"}, ensure_ascii=False)


@tool
def generate_from_template(file_path: str, report_data: str) -> str:
    """根据用户上传的模板文件和填写数据，生成填充后的文档。
    支持任意docx模板文件的自动填充。

    Args:
        file_path: 模板文件路径（与analyze_uploaded_template使用的路径相同）
        report_data: JSON格式的填写数据，键为字段名，值为字段值。
                     多值字段用逗号分隔，如 {"及百分比": "15,15,20,20,10,10,5,5,0"}
                     行组数据用二维数组，如 {"T0_G0": [["值1","值2"],...]}
    """
    ctx = request_context.get() or new_context(method="generate_from_template")
    
    try:
        import os
        workspace = os.getenv("COZE_WORKSPACE_PATH", "/workspace/projects")
        if not os.path.isabs(file_path):
            full_path = os.path.join(workspace, file_path)
        else:
            full_path = file_path
        
        if not os.path.exists(full_path):
            return json.dumps({"success": False, "message": f"文件不存在: {file_path}"}, ensure_ascii=False)
        
        data = json.loads(report_data)
        
        # 通用模板：直接填充，不调用expand（保留原格式）
        doc_bytes = _fill_custom_template(full_path, data)
        
        # 上传到对象存储
        from coze_coding_dev_sdk.s3 import S3SyncStorage
        import time
        
        storage = S3SyncStorage(
            endpoint_url=os.getenv("COZE_BUCKET_ENDPOINT_URL"),
            access_key="",
            secret_key="",
            bucket_name=os.getenv("COZE_BUCKET_NAME"),
            region="cn-beijing",
        )
        
        timestamp = time.strftime("%Y%m%d%H%M%S")
        safe_name = os.path.splitext(os.path.basename(full_path))[0].replace(" ", "_")
        file_name = f"custom_report/{safe_name}_{timestamp}.docx"
        
        file_key = storage.upload_file(
            file_content=doc_bytes,
            file_name=file_name,
            content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        url = storage.generate_presigned_url(key=file_key, expire_time=86400)
        
        return json.dumps({
            "success": True,
            "message": "文档已成功生成并上传",
            "file_name": file_key,
            "download_url": url,
        }, ensure_ascii=False)
        
    except Exception as e:
        return json.dumps({
            "success": False,
            "message": f"文档生成失败: {e}",
        }, ensure_ascii=False)
