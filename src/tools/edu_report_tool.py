"""高校教务评价报告生成工具 - 生成《专业课程目标达成度评价报告》Word文档"""

import os
import io
import json
import logging
from datetime import datetime

from docx import Document
from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from langchain.tools import tool
from coze_coding_dev_sdk.s3 import S3SyncStorage
from coze_coding_utils.log.write_log import request_context
from coze_coding_utils.runtime_ctx.context import new_context

logger = logging.getLogger(__name__)

# 全局 S3 客户端（懒初始化）
_s3_storage = None


def _get_s3_storage() -> S3SyncStorage:
    """获取 S3 存储单例"""
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


def _set_cell_text(cell, text, bold=False, font_size=10.5, alignment=WD_ALIGN_PARAGRAPH.CENTER):
    """设置表格单元格文本样式"""
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.alignment = alignment
    run = paragraph.add_run(text)
    run.font.size = Pt(font_size)
    run.font.name = "宋体"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    run.bold = bold


def _set_paragraph_text(paragraph, text, bold=False, font_size=12, alignment=WD_ALIGN_PARAGRAPH.LEFT, font_name="宋体"):
    """设置段落文本样式"""
    paragraph.alignment = alignment
    run = paragraph.add_run(text)
    run.font.size = Pt(font_size)
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
    run.bold = bold


def _build_report_docx(
    teaching_class: str,
    evaluator: str,
    participants: str,
    course_objectives: list[str],
) -> bytes:
    """
    根据收集的信息构建《专业课程目标达成度评价报告》Word文档，
    返回文档的二进制内容。
    """
    doc = Document()

    # ---- 全局默认样式 ----
    style = doc.styles["Normal"]  # type: ignore[union-attr]
    style.font.name = "宋体"  # type: ignore[union-attr]
    style.font.size = Pt(12)  # type: ignore[union-attr]
    style.element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")  # type: ignore[union-attr]
    style.paragraph_format.line_spacing = 1.5  # type: ignore[union-attr]

    # ---- 标题 ----
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("专业课程目标达成度评价报告")
    run.bold = True
    run.font.size = Pt(22)
    run.font.name = "黑体"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "黑体")

    # ---- 基本信息 ----
    doc.add_paragraph()  # 空行
    info_table = doc.add_table(rows=4, cols=4, style="Table Grid")
    info_table.alignment = WD_TABLE_ALIGNMENT.CENTER

    info_data = [
        ("教学班级", teaching_class, "评价责任人", evaluator),
        ("参与人", participants, "评价日期", datetime.now().strftime("%Y年%m月%d日")),
    ]

    for row_idx, (label1, val1, label2, val2) in enumerate(info_data):
        row = info_table.rows[row_idx]
        _set_cell_text(row.cells[0], label1, bold=True, font_size=10.5)
        _set_cell_text(row.cells[1], val1, font_size=10.5, alignment=WD_ALIGN_PARAGRAPH.LEFT)
        _set_cell_text(row.cells[2], label2, bold=True, font_size=10.5)
        _set_cell_text(row.cells[3], val2, font_size=10.5, alignment=WD_ALIGN_PARAGRAPH.LEFT)

    # 合并第三行（"课程目标"标题行）
    row_obj_title = info_table.rows[2]
    row_obj_title.cells[0].merge(row_obj_title.cells[3])
    _set_cell_text(row_obj_title.cells[0], "课程目标", bold=True, font_size=11)

    # 合并第四行（课程目标内容）
    row_obj_content = info_table.rows[3]
    row_obj_content.cells[0].merge(row_obj_content.cells[3])
    obj_text = "\n".join([f"目标{i+1}：{obj}" for i, obj in enumerate(course_objectives)])
    _set_cell_text(row_obj_content.cells[0], obj_text, font_size=10.5, alignment=WD_ALIGN_PARAGRAPH.LEFT)

    # ---- 一、课程目标及支撑毕业要求 ----
    doc.add_paragraph()
    h1 = doc.add_paragraph()
    _set_paragraph_text(h1, "一、课程目标及支撑毕业要求", bold=True, font_size=14, font_name="黑体")

    for i, obj in enumerate(course_objectives):
        p = doc.add_paragraph()
        _set_paragraph_text(p, f"课程目标{i+1}：{obj}", font_size=12)
        # 为每个目标预留支撑毕业要求描述区域
        p2 = doc.add_paragraph()
        _set_paragraph_text(p2, "支撑毕业要求指标点：（待填写）", font_size=10.5)

    # ---- 二、课程目标达成度评价表 ----
    doc.add_paragraph()
    h2 = doc.add_paragraph()
    _set_paragraph_text(h2, "二、课程目标达成度评价表", bold=True, font_size=14, font_name="黑体")

    # 评价表格：每个课程目标一行
    eval_table = doc.add_table(rows=len(course_objectives) + 1, cols=6, style="Table Grid")
    eval_table.alignment = WD_TABLE_ALIGNMENT.CENTER

    headers = ["课程目标", "考核方式", "权重", "达成度评分", "达成度评价结果", "备注"]
    for col_idx, header in enumerate(headers):
        _set_cell_text(eval_table.rows[0].cells[col_idx], header, bold=True, font_size=10)

    for i, obj in enumerate(course_objectives):
        row = eval_table.rows[i + 1]
        _set_cell_text(row.cells[0], f"目标{i+1}", font_size=10)
        _set_cell_text(row.cells[1], "（待填写）", font_size=10)
        _set_cell_text(row.cells[2], "（待填写）", font_size=10)
        _set_cell_text(row.cells[3], "（待填写）", font_size=10)
        _set_cell_text(row.cells[4], "（待填写）", font_size=10)
        _set_cell_text(row.cells[5], "（待填写）", font_size=10)

    # ---- 三、评价分析 ----
    doc.add_paragraph()
    h3 = doc.add_paragraph()
    _set_paragraph_text(h3, "三、评价分析", bold=True, font_size=14, font_name="黑体")

    for i in range(len(course_objectives)):
        p = doc.add_paragraph()
        _set_paragraph_text(p, f"课程目标{i+1}达成情况分析：", bold=True, font_size=12)
        p2 = doc.add_paragraph()
        _set_paragraph_text(p2, "（待填写）", font_size=12)

    # ---- 四、持续改进措施 ----
    doc.add_paragraph()
    h4 = doc.add_paragraph()
    _set_paragraph_text(h4, "四、持续改进措施", bold=True, font_size=14, font_name="黑体")

    p = doc.add_paragraph()
    _set_paragraph_text(p, "（待填写）", font_size=12)

    # ---- 签名区 ----
    doc.add_paragraph()
    doc.add_paragraph()
    sign_table = doc.add_table(rows=2, cols=3, style="Table Grid")
    sign_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    _set_cell_text(sign_table.rows[0].cells[0], "评价责任人", bold=True, font_size=10.5)
    _set_cell_text(sign_table.rows[0].cells[1], "", font_size=10.5)
    _set_cell_text(sign_table.rows[0].cells[2], "日期：", font_size=10.5)
    _set_cell_text(sign_table.rows[1].cells[0], "系（教研室）主任", bold=True, font_size=10.5)
    _set_cell_text(sign_table.rows[1].cells[1], "", font_size=10.5)
    _set_cell_text(sign_table.rows[1].cells[2], "日期：", font_size=10.5)

    # 输出到字节流
    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()


@tool
def generate_edu_report(
    teaching_class: str,
    evaluator: str,
    participants: str,
    course_objective_1: str,
    course_objective_2: str,
    course_objective_3: str,
    course_objective_4: str,
) -> str:
    """
    生成《专业课程目标达成度评价报告》Word文档。
    将收集到的信息渲染为标准格式报告，并上传到对象存储，返回可下载链接。

    Args:
        teaching_class: 教学班级名称，如"软件工程24级1班"
        evaluator: 评价责任人姓名
        participants: 参与人姓名，多人用逗号隔开
        course_objective_1: 课程目标1的内容
        course_objective_2: 课程目标2的内容
        course_objective_3: 课程目标3的内容
        course_objective_4: 课程目标4的内容

    Returns:
        包含文档下载链接的JSON字符串
    """
    ctx = request_context.get() or new_context(method="generate_edu_report")

    try:
        course_objectives = [course_objective_1, course_objective_2, course_objective_3, course_objective_4]

        # 1. 生成 Word 文档
        logger.info("开始生成评价报告Word文档...")
        doc_bytes = _build_report_docx(
            teaching_class=teaching_class,
            evaluator=evaluator,
            participants=participants,
            course_objectives=course_objectives,
        )
        logger.info(f"Word文档生成完成，大小: {len(doc_bytes)} bytes")

        # 2. 上传到对象存储
        storage = _get_s3_storage()
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        file_name = f"edu_report/course_evaluation_report_{timestamp}.docx"

        file_key = storage.upload_file(
            file_content=doc_bytes,
            file_name=file_name,
            content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        logger.info(f"文档已上传，key: {file_key}")

        # 3. 生成签名下载链接（有效期24小时）
        download_url = storage.generate_presigned_url(
            key=file_key,
            expire_time=86400,
        )
        logger.info("签名URL生成完成")

        result = {
            "success": True,
            "message": "评价报告已成功生成并上传",
            "file_name": file_name,
            "download_url": download_url,
            "teaching_class": teaching_class,
            "evaluator": evaluator,
        }

        return json.dumps(result, ensure_ascii=False, indent=2)

    except Exception as e:
        logger.error(f"生成评价报告失败: {e}")
        return json.dumps({
            "success": False,
            "message": f"生成评价报告失败: {str(e)}",
        }, ensure_ascii=False, indent=2)
