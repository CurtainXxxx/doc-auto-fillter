"""高校教务评价报告生成工具 - 基于模板严格匹配格式生成Word文档"""

import os
import io
import json
import logging
import copy
from datetime import datetime

from docx import Document
from docx.shared import Pt, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from langchain.tools import tool
from coze_coding_dev_sdk.s3 import S3SyncStorage
from coze_coding_utils.log.write_log import request_context
from coze_coding_utils.runtime_ctx.context import new_context

logger = logging.getLogger(__name__)

# 全局 S3 客户端（懒初始化）
_s3_storage = None

# 模板文件路径
_TEMPLATE_PATH = os.path.join(
    os.getenv("COZE_WORKSPACE_PATH", "/workspace/projects"),
    "assets",
    "template.docx",
)

# OOXML 命名空间
_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


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


def _find_cell_containing(table, keyword: str):
    """在表格中找到包含指定关键词的单元格，返回 (row_idx, col_idx, cell)"""
    for r_idx, row in enumerate(table.rows):
        for c_idx, cell in enumerate(row.cells):
            if keyword in cell.text:
                return r_idx, c_idx, cell
    return None


def _append_text_to_cell(cell, label: str, value: str):
    """在已有label文本后追加value文本，保持相同的字体格式。
    例如单元格已有"教学班级："，在其后追加具体班级名。
    """
    # 找到包含label的run，在其后追加value
    for para in cell.paragraphs:
        for run in para.runs:
            if label in run.text:
                # 在这个run后面追加一个新run，复制格式
                run_elem = run._element
                parent = run_elem.getparent()
                new_run_elem = copy.deepcopy(run_elem)
                # 清空新run的文本
                for t_elem in new_run_elem.findall(qn("w:t")):
                    t_elem.text = value
                # 插入到当前run之后
                run_elem.addnext(new_run_elem)
                return
    # 如果没找到label run，在最后一个段落末尾追加
    last_para = cell.paragraphs[-1] if cell.paragraphs else cell.add_paragraph()
    new_run = last_para.add_run(value)
    new_run.font.name = "FangSong_GB2312"
    new_run._element.rPr.rFonts.set(qn("w:eastAsia"), "FangSong_GB2312")


def _set_cell_text_keep_format(cell, text: str):
    """清空单元格内容并设置新文本，保持单元格原有的段落格式。
    保留第一个段落的格式属性（对齐等），仅替换文本。
    """
    # 保存第一个段落的格式
    if not cell.paragraphs:
        return
    
    first_para = cell.paragraphs[0]
    
    # 清空所有现有run
    for para in cell.paragraphs:
        for run in list(para.runs):
            run._element.getparent().remove(run._element)
    
    # 删除多余段落，只保留第一个
    for para in cell.paragraphs[1:]:
        para._element.getparent().remove(para._element)
    
    # 在第一个段落中添加新run
    new_run = first_para.add_run(text)
    new_run.font.name = "FangSong_GB2312"
    new_run._element.rPr.rFonts.set(qn("w:eastAsia"), "FangSong_GB2312")


def _add_table_row_after(table, after_row_idx: int):
    """在指定行之后添加一行，复制该行的格式（合并信息、宽度等），但清空内容。
    通过复制XML实现，确保格式完全一致。
    """
    src_row = table.rows[after_row_idx]
    src_tr = src_row._element
    tbl_element = table._element
    
    # 深拷贝源行
    new_tr = copy.deepcopy(src_tr)
    
    # 清空新行中所有单元格的文本内容（保留格式）
    for tc in new_tr.findall(qn("w:tc")):
        for p in tc.findall(qn("w:p")):
            for r in p.findall(qn("w:r")):
                for t in r.findall(qn("w:t")):
                    t.text = ""
    
    # 插入到源行之后
    src_tr.addnext(new_tr)


def _fill_table0(table, teaching_class: str, evaluator: str, participants: str, course_objectives: list):
    """填充表格0：基本信息 + 课程目标与毕业要求的对应关系"""
    
    # ---- Row 1: 教学班级 / 评价责任人 / 参与人 ----
    _append_text_to_cell(table.rows[1].cells[0], "教学班级：", teaching_class)
    _append_text_to_cell(table.rows[1].cells[3], "评价责任人:", evaluator)
    _append_text_to_cell(table.rows[1].cells[6], "参与人：", participants)
    
    # ---- Rows 4+: 课程目标数据行 ----
    # 模板有3个空行(4,5,6)，用户需要4个课程目标
    # 需要再添加1行
    template_data_rows = 3  # 行4,5,6
    needed_rows = len(course_objectives)
    
    if needed_rows > template_data_rows:
        for _ in range(needed_rows - template_data_rows):
            _add_table_row_after(table, table.rows.__len__() - 1)
    
    # 填入课程目标
    # 表格0的数据行结构：
    # 列0: 毕业要求 (1列)
    # 列1-3: 毕业要求指标点 (3列合并)
    # 列4-6: 课程目标 (3列合并)
    data_start_row = 4
    for i, obj in enumerate(course_objectives):
        row_idx = data_start_row + i
        if row_idx >= len(table.rows):
            break
        row = table.rows[row_idx]
        # 课程目标填入合并列(列4-6)
        _set_cell_text_keep_format(row.cells[4], f"课程目标{i+1}：{obj}")


def _fill_table1(table, course_objectives: list):
    """填充表格1：评价依据、考核分布、评价结果、总结改进"""
    
    # ---- Row 1: 课程目标标题行 ----
    # 列1(g s=8): 课程目标1, 列2(gs=9): 课程目标2, 列3(gs=8): 课程目标3, 列4(gs=5): 课程目标4
    # 这些标题已经在模板中，只是序号与用户目标对应
    
    # ---- Rows 10-13: 课程目标1-4行（期末考核分布）----
    # 行10: 课程目标1, 行11: 课程目标2, 行12: 课程目标3, 行13: 课程目标4
    # 这些标签已存在于模板中
    
    # ---- Rows 18-29: 评价结果 ----
    # 每个课程目标3行(期末考试/平时成绩/实验成绩)
    # 行18-20: 课程目标1, 行21-23: 课程目标2, 行24-26: 课程目标3, 行27-29: 课程目标4
    # 在"课程目标"列填入标签
    obj_col = 0  # 第0列
    for i in range(len(course_objectives)):
        base_row = 18 + i * 3
        if base_row < len(table.rows):
            # 课程目标列是垂直合并的，只在第一行（restart）写文本
            _set_cell_text_keep_format(table.rows[base_row].cells[obj_col], f"课程目标{i+1}")
    
    # ---- Row 33: 课程总结与改进措施 ----
    # 保持模板原有格式，无需额外填充


def _build_report_docx(
    teaching_class: str,
    evaluator: str,
    participants: str,
    course_objectives: list,
) -> bytes:
    """
    基于模板文件生成评价报告，确保格式与模板完全一致。
    """
    # 加载模板文档
    doc = Document(_TEMPLATE_PATH)
    
    if len(doc.tables) < 2:
        raise ValueError("模板文件格式异常：未找到预期的2个表格")
    
    # 填充表格0：基本信息
    _fill_table0(doc.tables[0], teaching_class, evaluator, participants, course_objectives)
    
    # 填充表格1：评价相关
    _fill_table1(doc.tables[1], course_objectives)
    
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
    基于岭南师范学院标准模板，将收集到的信息填入模板，严格保持模板格式不变，
    并上传到对象存储，返回可下载链接。

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

        # 1. 基于模板生成 Word 文档
        logger.info("基于模板生成评价报告Word文档...")
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
