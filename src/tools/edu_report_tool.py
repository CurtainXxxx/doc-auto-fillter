"""高校教务评价报告生成工具 - 基于模板严格匹配格式生成Word文档
自动识别模板中所有缺失字段，接受用户填入的完整数据。
"""

import os
import io
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


# ──────────────────── 低级工具函数 ────────────────────

def _append_text_after_run(cell, label: str, value: str):
    """在单元格中找到包含 label 的 run，在其后插入一个同格式的 run 填入 value。"""
    for para in cell.paragraphs:
        for run in para.runs:
            if label in run.text:
                new_run_elem = copy.deepcopy(run._element)
                for t_elem in new_run_elem.findall(qn("w:t")):
                    t_elem.text = value
                run._element.addnext(new_run_elem)
                return


def _set_cell_text(cell, text: str):
    """清空单元格内容（保留第一个段落的格式），设置新文本。"""
    if not cell.paragraphs:
        return
    first_para = cell.paragraphs[0]
    # 清空所有 run
    for para in cell.paragraphs:
        for run in list(para.runs):
            run._element.getparent().remove(run._element)
    # 删除多余段落
    for para in cell.paragraphs[1:]:
        para._element.getparent().remove(para._element)
    new_run = first_para.add_run(text)
    new_run.font.name = "FangSong_GB2312"
    new_run._element.rPr.rFonts.set(qn("w:eastAsia"), "FangSong_GB2312")


def _add_table_row_after(table, after_row_idx: int):
    """深拷贝模板行，清空文本后插入到指定行之后，保持所有格式。"""
    src_tr = table.rows[after_row_idx]._element
    new_tr = copy.deepcopy(src_tr)
    for tc in new_tr.findall(qn("w:tc")):
        for p in tc.findall(qn("w:p")):
            for r in p.findall(qn("w:r")):
                for t in r.findall(qn("w:t")):
                    t.text = ""
    src_tr.addnext(new_tr)


def _set_run_in_multiline_cell(cell, line_keyword: str, value: str):
    """在含多段落的单元格中，找到包含 line_keyword 的段落，在其 run 后追加 value。
    用于类似 '实现途径:|评价方法:' 这样的多行单元格。
    """
    for para in cell.paragraphs:
        full_text = "".join(r.text for r in para.runs if r.text)
        if line_keyword in full_text:
            # 在此段落最后一个 run 后追加
            new_run = para.add_run(value)
            # 复制同段落的字体
            for r in para.runs:
                if r.font.name:
                    new_run.font.name = r.font.name
                    if r._element.rPr is not None and r._element.rPr.rFonts is not None:
                        new_run._element.rPr.rFonts.set(qn("w:eastAsia"),
                            r._element.rPr.rFonts.get(qn("w:eastAsia"), "FangSong_GB2312"))
                    break
            return


# ──────────────────── 表格0 填充 ────────────────────

def _fill_table0(table, data: dict):
    """填充表格0：基本信息 + 课程目标与毕业要求的对应关系

    模板结构（7列）:
      行0: 课程名称(gs=2) | 开课时间(gs=3) | 考试类别/平时/期末 | 参评人数
      行1: 教学班级(gs=3) | 评价责任人(gs=3) | 参与人
      行2: 一、课程目标与毕业要求的对应关系 (gs=7)
      行3: 毕业要求 | 毕业要求指标点(gs=3) | 课程目标(gs=3)   ← 表头
      行4-6: 毕业要求 | 毕业要求指标点(gs=3) | 课程目标(gs=3) ← 3行空白数据行
    """
    # ── 行0: 课程名称 / 开课时间 / 参评人数 ──
    _append_text_after_run(table.rows[0].cells[0], "课程名称:", data.get("course_name", ""))
    _append_text_after_run(table.rows[0].cells[2], "开课时间:", data.get("course_time", ""))
    _append_text_after_run(table.rows[0].cells[6], "参评人数：", data.get("eval_headcount", ""))

    # ── 行1: 教学班级 / 评价责任人 / 参与人 ──
    _append_text_after_run(table.rows[1].cells[0], "教学班级：", data.get("teaching_class", ""))
    _append_text_after_run(table.rows[1].cells[3], "评价责任人:", data.get("evaluator", ""))
    _append_text_after_run(table.rows[1].cells[6], "参与人：", data.get("participants", ""))

    # ── 行4+: 课程目标数据行 ──
    objectives = data.get("objectives", [])
    template_data_rows = 3  # 模板行4,5,6

    # 不足则补行
    if len(objectives) > template_data_rows:
        for _ in range(len(objectives) - template_data_rows):
            _add_table_row_after(table, len(table.rows) - 1)

    for i, obj in enumerate(objectives):
        row_idx = 4 + i
        if row_idx >= len(table.rows):
            break
        row = table.rows[row_idx]
        # 列0: 毕业要求
        _set_cell_text(row.cells[0], obj.get("graduation_requirement", ""))
        # 列1-3(合并): 毕业要求指标点
        _set_cell_text(row.cells[1], obj.get("indicator", ""))
        # 列4-6(合并): 课程目标
        _set_cell_text(row.cells[4], f"课程目标{i+1}：{obj.get('objective', '')}")


# ──────────────────── 表格1 填充 ────────────────────

def _fill_table1(table, data: dict):
    """填充表格1：评价依据 + 考核分布 + 评价结果 + 总结改进

    模板结构（35列, 34行）:
      行0:  二、课程目标评价依据 (gs=35)
      行1:  考核环节 | 课程目标1(gs=8) | 课程目标2(gs=9) | 课程目标3(gs=8) | 课程目标4(gs=5)
      行2-6: 评价依据数据行(5行)
      行7:  三、课程目标期末考核分布 (gs=35)
      行8-15: 考核分布表
      行16: 四、课程教学质量评价结果 (gs=35)
      行17: 表头
      行18-29: 4个目标×3考核方式(期末/平时/实验)
      行30: 课程目标达成评价值
      行31: 课程目标达成分布
      行32: 五、课程总结与改进措施 (gs=35)
      行33: 课程总结 + 改进措施 (gs=35)
    """
    objectives = data.get("objectives", [])

    # ── 行2-6: 评价依据数据 ──
    eval_basis = data.get("evaluation_basis", [])
    for i, row_data in enumerate(eval_basis[:5]):
        row_idx = 2 + i
        if row_idx > 6:
            break
        row = table.rows[row_idx]
        # 5个逻辑列: 考核环节 | 课程目标1 | 课程目标2 | 课程目标3 | 课程目标4
        seen = set()
        logical_col = 0
        for cell in row.cells:
            cid = id(cell._element)
            if cid in seen:
                continue
            seen.add(cid)
            val = row_data[logical_col] if logical_col < len(row_data) else ""
            _set_cell_text(cell, val)
            logical_col += 1

    # ── 行18-29: 教学质量评价结果 ──
    quality_results = data.get("quality_results", [])
    for i, obj_result in enumerate(quality_results[:4]):
        base_row = 18 + i * 3
        # 课程目标标签
        if base_row < len(table.rows):
            _set_cell_text(table.rows[base_row].cells[0], f"课程目标{i+1}")
        # 3行: 期末考试 / 平时成绩 / 实验成绩
        for j, sub_result in enumerate(obj_result.get("details", [])[:3]):
            row_idx = base_row + j
            if row_idx >= len(table.rows):
                break
            row = table.rows[row_idx]
            # 实现途径、评价方法
            _set_run_in_multiline_cell(row.cells[1], "实现途径:", sub_result.get("approach", ""))
            _set_run_in_multiline_cell(row.cells[1], "评价方法:", sub_result.get("method", ""))
            # 目标分值(期末考试/平时成绩/实验成绩下的子列)
            target_score = sub_result.get("target_score", "")
            actual_score = sub_result.get("actual_score", "")
            achievement = sub_result.get("achievement", "")
            # 这些列需要跳到正确的逻辑列位置
            # 行18结构: 课程目标(gs=3) | 实现途径(gs=17) | 目标分值-期末考试(gs=3)+空(gs=4) | 实际平均分(gs=5) | 达成评价值(gs=3)
            # 目标分值区分为"期末考试"标签和数值
            _append_text_after_run(row.cells[2], "期末考试", target_score) if j == 0 and "期末" in row.cells[2].text else None
            # 简化处理：直接在对应列追加
            if target_score:
                _append_text_after_run(row.cells[2], row.cells[2].paragraphs[0].runs[0].text if row.cells[2].paragraphs[0].runs else "", target_score)
            if actual_score:
                seen2 = set()
                for cell in row.cells:
                    cid = id(cell._element)
                    if cid in seen2:
                        continue
                    seen2.add(cid)
                    if "实际平均分" in cell.text or actual_score:
                        _append_text_after_run(cell, cell.text.split("\n")[-1] if cell.text else "", actual_score)
                        break

    # ── 行30: 课程目标达成评价值 ──
    achievement_values = data.get("achievement_values", "")
    if achievement_values:
        _set_cell_text(table.rows[30].cells[1], achievement_values)

    # ── 行33: 课程总结 + 改进措施 ──
    course_summary = data.get("course_summary", "")
    improvement = data.get("improvement", "")
    cell33 = table.rows[33].cells[0]
    for para in cell33.paragraphs:
        full_text = "".join(r.text for r in para.runs if r.text)
        if "课程总结" in full_text and course_summary:
            # 在该段落后追加内容
            new_run = para.add_run(course_summary)
            new_run.font.name = "FangSong_GB2312"
            new_run._element.rPr.rFonts.set(qn("w:eastAsia"), "FangSong_GB2312")
        if "改进措施" in full_text and improvement:
            new_run = para.add_run(improvement)
            new_run.font.name = "FangSong_GB2312"
            new_run._element.rPr.rFonts.set(qn("w:eastAsia"), "FangSong_GB2312")


def _build_report_docx(data: dict) -> bytes:
    """基于模板文件生成评价报告，确保格式与模板完全一致。"""
    doc = Document(_TEMPLATE_PATH)

    if len(doc.tables) < 2:
        raise ValueError("模板文件格式异常：未找到预期的2个表格")

    _fill_table0(doc.tables[0], data)
    _fill_table1(doc.tables[1], data)

    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()


# ──────────────────── 工具定义 ────────────────────

# 模板中所有需要用户填写的字段，用于描述工具参数
_TEMPLATE_FIELDS_DESCRIPTION = """
模板包含以下需要填写的字段（按表格区域划分）：

【表格0 - 基本信息】
- course_name: 课程名称
- course_time: 开课时间
- eval_headcount: 参评人数
- teaching_class: 教学班级
- evaluator: 评价责任人
- participants: 参与人

【表格0 - 课程目标与毕业要求】(每个课程目标包含)
- objectives: 数组，每个元素包含:
  - objective: 课程目标内容
  - graduation_requirement: 对应的毕业要求
  - indicator: 毕业要求指标点

【表格1 - 课程目标评价依据】
- evaluation_basis: 数组(最多5行)，每行为5个字符串 [考核环节, 目标1, 目标2, 目标3, 目标4]

【表格1 - 教学质量评价结果】
- quality_results: 数组(4个目标)，每个元素包含:
  - details: 数组(3项: 期末考试/平时成绩/实验成绩)，每项包含:
    - approach: 实现途径
    - method: 评价方法
    - target_score: 目标分值
    - actual_score: 实际平均分
    - achievement: 达成评价值

【表格1 - 总结改进】
- achievement_values: 课程目标达成评价值（实际值/目标期望值）
- course_summary: 课程总结
- improvement: 改进措施
"""


@tool
def generate_edu_report(report_data: str) -> str:
    """
    生成《专业课程目标达成度评价报告》Word文档。
    基于岭南师范学院标准模板，将用户提供的完整数据填入模板所有字段，严格保持模板格式不变，
    并上传到对象存储，返回可下载链接。

    参数 report_data 是一个 JSON 字符串，包含以下字段:
    {
      "course_name": "课程名称",
      "course_time": "开课时间，如2023-2024-2",
      "eval_headcount": "参评人数",
      "teaching_class": "教学班级",
      "evaluator": "评价责任人",
      "participants": "参与人，多人用逗号隔开",
      "objectives": [
        {
          "objective": "课程目标1内容",
          "graduation_requirement": "对应的毕业要求",
          "indicator": "毕业要求指标点"
        },
        ... 共4个
      ],
      "evaluation_basis": [
        ["考核环节名", "目标1权重", "目标2权重", "目标3权重", "目标4权重"],
        ... 最多5行
      ],
      "quality_results": [
        {
          "details": [
            {"approach":"实现途径","method":"评价方法","target_score":"目标分值","actual_score":"实际平均分","achievement":"达成评价值"},
            ... 3项: 期末考试/平时成绩/实验成绩
          ]
        },
        ... 共4个目标
      ],
      "achievement_values": "课程目标达成评价值",
      "course_summary": "课程总结内容",
      "improvement": "改进措施内容"
    }

    所有字段均可为空字符串，工具会保留模板原样。

    Returns:
        包含文档下载链接的JSON字符串
    """
    ctx = request_context.get() or new_context(method="generate_edu_report")

    try:
        data = json.loads(report_data)
    except json.JSONDecodeError as e:
        return json.dumps({
            "success": False,
            "message": f"report_data JSON解析失败: {str(e)}",
        }, ensure_ascii=False, indent=2)

    try:
        logger.info("基于模板生成评价报告Word文档...")
        doc_bytes = _build_report_docx(data)
        logger.info(f"Word文档生成完成，大小: {len(doc_bytes)} bytes")

        # 上传到对象存储
        storage = _get_s3_storage()
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        file_name = f"edu_report/course_evaluation_report_{timestamp}.docx"

        file_key = storage.upload_file(
            file_content=doc_bytes,
            file_name=file_name,
            content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        logger.info(f"文档已上传，key: {file_key}")

        # 生成签名下载链接（有效期24小时）
        download_url = storage.generate_presigned_url(key=file_key, expire_time=86400)
        logger.info("签名URL生成完成")

        return json.dumps({
            "success": True,
            "message": "评价报告已成功生成并上传",
            "file_name": file_name,
            "download_url": download_url,
        }, ensure_ascii=False, indent=2)

    except Exception as e:
        logger.error(f"生成评价报告失败: {e}")
        return json.dumps({
            "success": False,
            "message": f"生成评价报告失败: {str(e)}",
        }, ensure_ascii=False, indent=2)
