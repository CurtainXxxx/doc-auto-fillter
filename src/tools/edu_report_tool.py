"""高校教务文档生成工具 - 自动识别模板 + 动态填充（简化版）

核心能力:
1. 自动解析任意Word模板，识别所有可填充字段
2. 只填充简单字段（标签冒号后追加值）和简单行组（空白格直接填值）
3. 复杂表格（大量合并、超多列）保持原样不填，避免格式错乱
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
    "评价报告": "assets/template.docx",
    "试卷分析": "assets/template_exam.docx",
    "关联矩阵": "assets/template_matrix.docx",
}

# 行组复杂度阈值：列数超过此值视为复杂行组，跳过填充
_MAX_SIMPLE_GROUP_COLS = 10

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


def _get_unique_cells(row) -> list:
    """获取行中的独立单元格"""
    seen = set()
    cells = []
    for cell in row.cells:
        cid = id(cell._element)
        if cid not in seen:
            seen.add(cid)
            cells.append(cell)
    return cells


def _is_vmerge_continue(cell) -> bool:
    """判断单元格是否是垂直合并延续格"""
    tcPr = cell._element.find(f"{{{_W_NS}}}tcPr")
    if tcPr is not None:
        vm = tcPr.find(f"{{{_W_NS}}}vMerge")
        if vm is not None:
            vm_val = vm.get(f"{{{_W_NS}}}val")
            if vm_val is None or vm_val == "continue":
                return True
    return False


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


def _set_cell_text(cell, text: str):
    """清空单元格内容，设置新文本，保留第一个段落格式"""
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
    """填充标签字段：在"标签："后追加用户值。跳过行组覆盖区域。"""
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
        row_indices = [r for r in field_info["row_indices"] if not _in_row_group(t_idx, r)]

        if not row_indices:
            continue

        if isinstance(value, list):
            for i, row_idx in enumerate(row_indices):
                if i < len(value):
                    unique = _get_unique_cells(table.rows[row_idx])
                    if field_info["col_idx"] < len(unique):
                        _append_after_label(unique[field_info["col_idx"]], label, value[i])
        else:
            row_idx = row_indices[0]
            unique = _get_unique_cells(table.rows[row_idx])
            if field_info["col_idx"] < len(unique):
                _append_after_label(unique[field_info["col_idx"]], label, value)


def _fill_simple_row_groups(doc, analysis: dict, data: dict):
    """填充简单行组：只处理列数<=阈值且无复杂合并的行组。
    策略：只填空白独立单元格，跳过含标签和vMerge的格。
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

        # 填入数据：遍历独立单元格，空白格按顺序填值
        for row_offset, row_data in enumerate(group_data):
            if row_offset >= needed:
                break
            actual_row = start_row + row_offset
            if actual_row >= len(table.rows):
                break
            unique = _get_unique_cells(table.rows[actual_row])

            data_idx = 0
            for cell in unique:
                if data_idx >= len(row_data):
                    break

                # 跳过vMerge=continue的格
                if _is_vmerge_continue(cell):
                    continue

                text = cell.text.strip().replace('\r', '')
                if not text:
                    # 空白格 → 填值
                    _set_cell_text(cell, str(row_data[data_idx]))
                    data_idx += 1
                else:
                    # 含标签格：检查"标签："模式
                    lines = [l.strip() for l in text.split('\n') if l.strip()]
                    for line in lines:
                        m = re.match(r'^(.+?)\s*[：:]\s*(.*)', line)
                        if m and not m.group(2).strip():
                            # "标签："后为空 → 追加值
                            if data_idx < len(row_data) and row_data[data_idx]:
                                _append_after_label(cell, m.group(1).strip(), str(row_data[data_idx]))
                            data_idx += 1
                        elif m and m.group(2).strip():
                            # "标签：已有值" → 跳过
                            data_idx += 1
                        # 无冒号固定标签 → 不消费数据


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
