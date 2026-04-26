"""高校教务文档生成工具 - 自动识别模板 + 动态填充（简化版）

核心能力:
1. 自动解析任意Word模板，识别所有可填充字段
2. 支持两种字段模式填充：冒号追加模式和标签格+空白格设置模式
3. 行组填充：处理含vMerge的合并单元格行组
4. 支持多模板，通过template_name参数选择
"""

import os
import io
import re
import json
import logging
import copy
from datetime import datetime

from docx import Document
from docx.oxml.ns import qn
from langchain.tools import tool
from coze_coding_dev_sdk.s3 import S3SyncStorage
from coze_coding_utils.log.write_log import request_context
from coze_coding_utils.runtime_ctx.context import new_context
from tools.template_analyzer import analyze_template

logger = logging.getLogger(__name__)

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# 模板注册表：key=模板名称, value=文件路径
TEMPLATE_REGISTRY = {
    "评价报告": "assets/2023-2024-2《xxx》 岭南师范学院专业课程目标达成度评价报告模板.docx",
    "试卷分析": "assets/2023-2024-2《xxx》 试卷分析模板.docx",
    "关联矩阵": "assets/2023-2024-2《xxx》岭南师范学院考题与课程目标及毕业要求关联矩阵表模板.docx",
}

# 行组复杂度阈值：列数超过此值视为复杂行组，跳过填充
_MAX_SIMPLE_GROUP_COLS = 15

# S3 客户端
_s3_storage = None


def _get_s3_storage() -> S3SyncStorage:
    global _s3_storage
    if _s3_storage is None:
        _s3_storage = S3SyncStorage(
            endpoint_url=os.getenv("COZE_BUCKET_ENDPOINT_URL"),
            access_key="",
            secret_key="",
            bucket_name=os.getenv("COZE_BUCKET_NAME"),
            region="cn-beijing",
        )
    return _s3_storage


def _get_template_path(template_name: str) -> str:
    """根据名称获取模板绝对路径"""
    rel_path = TEMPLATE_REGISTRY.get(template_name)
    if not rel_path:
        raise ValueError(f"未找到模板'{template_name}'，可用模板: {list(TEMPLATE_REGISTRY.keys())}")
    return os.path.join(os.getenv("COZE_WORKSPACE_PATH", "/workspace/projects"), rel_path)


# ──────────────────── 单元格操作基础函数 ────────────────────

def _get_tr_cells(row) -> list:
    """从row._tr获取真正属于这一行的tc元素列表。
    
    注意：不能用 row.cells 或 _get_unique_cells，因为python-docx对合并单元格的处理
    会导致不同行的cell指向同一个XML元素（vMerge=continue格与restart格共享元素）。
    _get_tr_cells 直接遍历 tr 下的 tc 子元素，保证每行获取的元素是独立的。
    """
    return row._tr.findall(qn("w:tc"))


def _get_tc_text(tc) -> str:
    """获取tc元素的文本内容"""
    texts = []
    for p in tc.findall(qn("w:p")):
        for r in p.findall(qn("w:r")):
            for t in r.findall(qn("w:t")):
                if t.text:
                    texts.append(t.text)
    return "".join(texts).strip()


def _is_tc_vmerge_continue(tc) -> bool:
    """判断tc元素是否是垂直合并延续格（vMerge=continue，无val属性）"""
    tcPr = tc.find(qn("w:tcPr"))
    if tcPr is not None:
        vm = tcPr.find(qn("w:vMerge"))
        if vm is not None:
            val = vm.get(qn("w:val"))
            # val=None 表示continue, val="continue"也表示continue
            # val="restart" 表示起始格
            return val is None or val == "continue"
    return False


def _is_tc_vmerge_restart(tc) -> bool:
    """判断tc元素是否是垂直合并起始格（vMerge=restart）"""
    tcPr = tc.find(qn("w:tcPr"))
    if tcPr is not None:
        vm = tcPr.find(qn("w:vMerge"))
        if vm is not None:
            val = vm.get(qn("w:val"))
            return val == "restart"
    return False


def _get_unique_cells(row) -> list:
    """获取行中的独立单元格（python-docx Cell对象），用于标签字段定位。
    
    注意：此函数返回的Cell对象可能被合并单元格的vMerge影响，
    在填充行组数据时不应使用此函数，应使用 _get_tr_cells。
    """
    seen = set()
    cells = []
    for cell in row.cells:
        cid = id(cell._element)
        if cid not in seen:
            seen.add(cid)
            cells.append(cell)
    return cells


def _is_vmerge_continue(cell) -> bool:
    """判断Cell对象是否是垂直合并延续格"""
    return _is_tc_vmerge_continue(cell._element)


# ──────────────────── 填充函数 ────────────────────

def _append_after_label(cell, label: str, value: str):
    """在单元格中找到含 label 且以冒号结尾的 run，在其后追加同格式 run"""
    for para in cell.paragraphs:
        for run in para.runs:
            if label in run.text and run.text.strip().endswith(('：', ':')):
                new_run_elem = copy.deepcopy(run._element)
                for t_elem in new_run_elem.findall(qn("w:t")):
                    t_elem.text = value
                run._element.addnext(new_run_elem)
                return True
    return False


def _set_tc_text(tc, text: str):
    """清空tc元素内容，设置新文本。操作XML层级，不受python-docx合并单元格影响。"""
    # 清空所有段落中的run
    for p in tc.findall(qn("w:p")):
        for r in list(p.findall(qn("w:r"))):
            p.remove(r)
    # 移除多余段落，只保留第一个
    paragraphs = tc.findall(qn("w:p"))
    for p in paragraphs[1:]:
        tc.remove(p)
    
    # 在第一个段落中添加run
    if paragraphs:
        first_p = paragraphs[0]
    else:
        first_p = tc.makeelement(qn("w:p"), {})
        tc.append(first_p)
    
    # 创建run
    r_elem = first_p.makeelement(qn("w:r"), {})
    # 创建rPr（字体设置）
    rPr = r_elem.makeelement(qn("w:rPr"), {})
    rFonts = rPr.makeelement(qn("w:rFonts"), {
        qn("w:ascii"): "FangSong_GB2312",
        qn("w:hAnsi"): "FangSong_GB2312",
        qn("w:eastAsia"): "FangSong_GB2312",
    })
    rPr.append(rFonts)
    r_elem.append(rPr)
    
    # 创建文本元素
    t_elem = r_elem.makeelement(qn("w:t"), {})
    t_elem.text = text
    r_elem.append(t_elem)
    
    first_p.append(r_elem)


def _set_cell_text(cell, text: str):
    """清空单元格内容，设置新文本（通过Cell对象操作）"""
    if not cell.paragraphs:
        return
    for para in cell.paragraphs:
        for run in list(para.runs):
            run._element.getparent().remove(run._element)
    for para in cell.paragraphs[1:]:
        para._element.getparent().remove(para._element)
    new_run = cell.paragraphs[0].add_run(text)
    new_run.font.name = "FangSong_GB2312"
    new_run._element.rPr.rFonts.set(qn("w:eastAsia"), "FangSong_GB2312")


def _add_table_row_after(table, after_row_idx: int):
    """深拷贝模板行，清空文本后插入"""
    src_tr = table.rows[after_row_idx]._element
    new_tr = copy.deepcopy(src_tr)
    for tc in new_tr.findall(qn("w:tc")):
        for p in tc.findall(qn("w:p")):
            for r in p.findall(qn("w:r")):
                for t in r.findall(qn("w:t")):
                    t.text = ""
    src_tr.addnext(new_tr)


def _fill_label_fields(doc, analysis: dict, data: dict):
    """填充标签字段。支持三种模式：
    - fill_mode="append"：在"标签："后追加用户值（冒号模式）
    - fill_mode="set"：直接设置空白格内容（标签格+空白格模式）
    - fill_mode="replace"：清空格内容后设置新值（替换占位符模式，如%→实际值）
    跳过行组覆盖区域。
    """
    row_groups = analysis["row_groups"]

    def _in_row_group(t_idx, r_idx):
        for g in row_groups:
            if g["table_idx"] == t_idx and g["start_row"] <= r_idx < g["start_row"] + g["template_row_count"]:
                return True
        return False

    for field_info in analysis["label_fields"]:
        label = field_info["label"]
        value = data.get(label, "")
        if not value:
            continue

        t_idx = field_info["table_idx"]
        table = doc.tables[t_idx]
        fill_mode = field_info.get("fill_mode", "append")
        row_indices = [r for r in field_info["row_indices"] if not _in_row_group(t_idx, r)]

        if not row_indices:
            continue

        def _fill_cell(cell, val, mode):
            if mode == "replace":
                _set_cell_text(cell, str(val))
            elif mode == "set":
                _set_cell_text(cell, str(val))
            else:
                _append_after_label(cell, label, val)

        if isinstance(value, list):
            for i, row_idx in enumerate(row_indices):
                if i < len(value):
                    unique = _get_unique_cells(table.rows[row_idx])
                    col_idx = field_info["col_idx"]
                    if col_idx < len(unique):
                        _fill_cell(unique[col_idx], value[i], fill_mode)
        else:
            row_idx = row_indices[0]
            unique = _get_unique_cells(table.rows[row_idx])
            col_idx = field_info["col_idx"]
            if col_idx < len(unique):
                _fill_cell(unique[col_idx], value, fill_mode)


def _fill_simple_row_groups(doc, analysis: dict, data: dict):
    """填充简单行组：只处理列数<=阈值且无复杂合并的行组。
    
    关键改进：使用 _get_tr_cells 直接获取每行的tc元素，避免python-docx
    合并单元格共享XML元素导致的填充错位问题。
    
    策略：
    - vMerge=continue 的格：跳过（它是上方restart格的延续，不能独立设值）
    - vMerge=restart 的格：空白则填值，有文字则跳过
    - 普通格：空白则填值，有冒号标签则追加，有固定标签则跳过
    """
    for group_info in analysis["row_groups"]:
        group_id = group_info["group_id"]
        col_count = len(group_info["column_labels"])

        # 跳过复杂行组
        if col_count > _MAX_SIMPLE_GROUP_COLS:
            logger.info(f"跳过复杂行组 {group_id} ({col_count}列)")
            continue

        group_data = data.get(group_id, [])
        if not group_data:
            continue

        t_idx = group_info["table_idx"]
        table = doc.tables[t_idx]
        start_row = group_info["start_row"]
        template_count = group_info["template_row_count"]

        # 扩展行
        needed = len(group_data)
        if needed > template_count:
            last = start_row + template_count - 1
            for i in range(needed - template_count):
                _add_table_row_after(table, last + i)

        # 填入数据：使用tr级别的tc元素，避免合并单元格问题
        for row_offset, row_data in enumerate(group_data):
            if row_offset >= needed:
                break
            actual_row = start_row + row_offset
            if actual_row >= len(table.rows):
                break
            
            # 使用 _get_tr_cells 获取真正属于该行的tc元素
            tr_cells = _get_tr_cells(table.rows[actual_row])

            data_idx = 0
            for tc in tr_cells:
                if data_idx >= len(row_data):
                    break

                text = _get_tc_text(tc).replace('\r', '')
                
                # vMerge=continue的格：它显示restart格的内容，不能独立设值
                # 但用户数据中该列有值，需要消耗掉这个数据索引（丢弃）
                # 因为Word中continue格自动显示restart格的内容
                if _is_tc_vmerge_continue(tc):
                    data_idx += 1  # 消耗数据但不填值
                    continue

                if not text:
                    # 空白格（包括vMerge=restart但文本为空的格）→ 填值
                    _set_tc_text(tc, str(row_data[data_idx]))
                    data_idx += 1
                else:
                    # 含标签格：检查"标签："模式
                    lines = [l.strip() for l in text.split('\n') if l.strip()]
                    for line in lines:
                        m = re.match(r'^(.+?)\s*[：:]\s*(.*)', line)
                        if m and not m.group(2).strip():
                            # "标签："后为空 → 需要追加值
                            _append_value_to_tc_after_label(tc, m.group(1).strip(), str(row_data[data_idx]))
                            data_idx += 1
                        elif m and m.group(2).strip():
                            # "标签：已有值" → 跳过
                            data_idx += 1
                        # 无冒号固定标签 → 不消费数据


def _append_value_to_tc_after_label(tc, label: str, value: str):
    """在tc元素中找到含label且以冒号结尾的run，在其后追加同格式run"""
    for p in tc.findall(qn("w:p")):
        for r in p.findall(qn("w:r")):
            for t in r.findall(qn("w:t")):
                if t.text and label in t.text and t.text.strip().endswith(('：', ':')):
                    new_r = copy.deepcopy(r)
                    for new_t in new_r.findall(qn("w:t")):
                        new_t.text = value
                    r.addnext(new_r)
                    return


def _fill_summary_section(doc, analysis: dict, data: dict):
    """填充课程总结/改进措施等大段文本字段。
    这些字段通常在表格末尾的大合并单元格中，标签后是空行区域。
    """
    for field_info in analysis["label_fields"]:
        label = field_info["label"]
        if label not in ("课程总结", "改进措施"):
            continue
        value = data.get(label, "")
        if not value:
            continue

        t_idx = field_info["table_idx"]
        table = doc.tables[t_idx]
        for row_idx in field_info["row_indices"]:
            if row_idx >= len(table.rows):
                continue
            unique = _get_unique_cells(table.rows[row_idx])
            for cell in unique:
                if label in cell.text:
                    _append_after_label(cell, label, value)
                    break


def _build_report_docx(template_path: str, data: dict) -> bytes:
    """基于模板 + 解析结果 + 用户数据，动态生成文档"""
    analysis = analyze_template(template_path)
    doc = Document(template_path)

    _fill_label_fields(doc, analysis, data)
    _fill_simple_row_groups(doc, analysis, data)
    _fill_summary_section(doc, analysis, data)

    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()


# ──────────────────── 工具定义 ────────────────────

@tool
def list_templates() -> str:
    """列出所有可用的报告模板。当用户开始对话时调用此工具，展示可选模板列表。"""
    ctx = request_context.get() or new_context(method="list_templates")
    result = {
        "success": True,
        "templates": [
            {"name": name, "file": path}
            for name, path in TEMPLATE_REGISTRY.items()
        ],
    }
    return json.dumps(result, ensure_ascii=False, indent=2)


@tool
def analyze_report_template(template_name: str) -> str:
    """解析指定模板，自动识别所有需要填写的字段，返回字段清单和收集计划。
    在开始收集用户信息前必须先调用此工具。

    Args:
        template_name: 模板名称，从list_templates返回的名称中选择

    Returns:
        包含字段清单和收集计划的JSON字符串
    """
    ctx = request_context.get() or new_context(method="analyze_report_template")

    try:
        template_path = _get_template_path(template_name)
        result = analyze_template(template_path)

        # 区分简单行组和复杂行组
        simple_groups = []
        complex_groups = []
        for g in result["row_groups"]:
            if len(g["column_labels"]) <= _MAX_SIMPLE_GROUP_COLS:
                simple_groups.append(g)
            else:
                complex_groups.append(g)

        output = {
            "success": True,
            "template_name": template_name,
            "total_label_fields": result["summary"]["total_unique_labels"],
            "field_labels": [f["label"] for f in result["label_fields"]],
            "repeat_labels": {f["label"]: f["repeat_count"] for f in result["label_fields"] if f["repeat_count"] > 1},
            "simple_row_groups": [
                {
                    "group_id": g["group_id"],
                    "template_row_count": g["template_row_count"],
                    "column_labels": g["column_labels"],
                }
                for g in simple_groups
            ],
            "skipped_complex_groups": len(complex_groups),
            "collection_plan": result["summary"]["collection_plan"],
        }

        return json.dumps(output, ensure_ascii=False, indent=2)

    except Exception as e:
        logger.error(f"模板解析失败: {e}")
        return json.dumps({"success": False, "message": f"模板解析失败: {str(e)}"}, ensure_ascii=False, indent=2)


@tool
def generate_edu_report(template_name: str, report_data: str) -> str:
    """生成教务报告Word文档。自动解析模板，将用户数据填入对应字段，保持模板格式不变，
    上传到对象存储，返回下载链接。

    Args:
        template_name: 模板名称（与analyze_report_template使用相同名称）
        report_data: JSON字符串，key为字段标签名或行组group_id，value为填写内容。
            标签字段: key=标签名, value=字符串; 重复标签: value=数组。
            行组字段: key=group_id, value=二维数组 [[行1列1,...],...]。

    Returns:
        包含文档下载链接的JSON字符串
    """
    ctx = request_context.get() or new_context(method="generate_edu_report")

    try:
        data = json.loads(report_data)
    except json.JSONDecodeError as e:
        return json.dumps({"success": False, "message": f"JSON解析失败: {str(e)}"}, ensure_ascii=False, indent=2)

    try:
        template_path = _get_template_path(template_name)
        logger.info(f"基于模板'{template_name}'生成文档...")
        doc_bytes = _build_report_docx(template_path, data)
        logger.info(f"文档生成完成，大小: {len(doc_bytes)} bytes")

        storage = _get_s3_storage()
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        file_name = f"edu_report/{template_name}_{timestamp}.docx"

        file_key = storage.upload_file(
            file_content=doc_bytes,
            file_name=file_name,
            content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        download_url = storage.generate_presigned_url(key=file_key, expire_time=86400)

        return json.dumps({
            "success": True,
            "message": "报告已成功生成并上传",
            "file_name": file_name,
            "download_url": download_url,
        }, ensure_ascii=False, indent=2)

    except Exception as e:
        logger.error(f"生成报告失败: {e}")
        return json.dumps({"success": False, "message": f"生成报告失败: {str(e)}"}, ensure_ascii=False, indent=2)
