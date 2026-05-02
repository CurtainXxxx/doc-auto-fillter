"""知识文件解析工具 - 解析用户上传的文件，提取文档所需信息"""

import os
import re
import json
from langchain.tools import tool
from coze_coding_utils.log.write_log import request_context
from coze_coding_utils.runtime_ctx.context import new_context


def _extract_text_from_file(file_path: str) -> str:
    """根据文件扩展名提取文本内容"""
    ext = os.path.splitext(file_path)[1].lower()

    if ext == '.txt':
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
            return f.read()

    if ext == '.csv':
        import csv
        lines = []
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
            reader = csv.reader(f)
            for row in reader:
                lines.append(', '.join(row))
        return '\n'.join(lines)

    if ext in ('.docx', '.doc'):
        try:
            from docx import Document
            doc = Document(file_path)
            text_parts = []
            for para in doc.paragraphs:
                if para.text.strip():
                    text_parts.append(para.text.strip())
            for table in doc.tables:
                for row in table.rows:
                    # 用unique cells避免合并单元格重复
                    seen = set()
                    cells = []
                    for cell in row.cells:
                        cid = id(cell._element)
                        if cid not in seen:
                            seen.add(cid)
                            cells.append(cell.text.strip())
                    non_empty = [c for c in cells if c]
                    if non_empty:
                        text_parts.append(' | '.join(non_empty))
            return '\n'.join(text_parts)
        except Exception as e:
            return f"[docx解析失败: {str(e)}]"

    if ext in ('.xlsx', '.xls'):
        try:
            import openpyxl
            wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
            lines = []
            for sheet in wb.worksheets:
                for row in sheet.iter_rows(values_only=True):
                    cells = [str(c) for c in row if c is not None]
                    if cells:
                        lines.append(', '.join(cells))
            wb.close()
            return '\n'.join(lines)
        except Exception as e:
            return f"[xlsx解析失败: {str(e)}]"

    if ext == '.pdf':
        try:
            from PyPDF2 import PdfReader
            reader = PdfReader(file_path)
            text_parts = []
            for page in reader.pages:
                t = page.extract_text()
                if t:
                    text_parts.append(t)
            return '\n'.join(text_parts)
        except Exception as e:
            return f"[PDF解析失败: {str(e)}]"

    return f"[不支持的文件格式: {ext}]"


def extract_text_from_upload(file_path: str) -> dict:
    """
    解析上传文件并提取文本内容（供API调用，非@tool）

    Returns:
        dict: {success: bool, extracted_text: str, error: str}
    """
    try:
        text = _extract_text_from_file(file_path)
        if text.startswith('['):
            return {"success": False, "extracted_text": "", "error": text}
        return {"success": True, "extracted_text": text, "error": ""}
    except Exception as e:
        return {"success": False, "extracted_text": "", "error": str(e)}


def _find_field_in_text(field: str, content: str) -> list:
    """
    在文本内容中智能搜索字段的值。
    支持多种格式：
    - 冒号模式: 课程性质：专业必修课 / 课程性质: 专业必修课
    - 等号模式: 课程性质＝专业必修课 / 课程性质=专业必修课
    - 管道符/表格模式: 课程性质 | 专业必修课
    - "是"模式: 课程性质是专业必修课
    - 【】模式: 【课程性质】专业必修课
    """
    found = []

    # 标准化字段名，去掉可能的空格
    field_clean = field.strip()

    for line in content.split('\n'):
        line = line.strip()
        if not line:
            continue

        # ── 模式1: 冒号/等号分隔 ──
        # 课程性质：专业必修课 / 课程性质: 专业必修课
        for sep in ['：', ':', '＝', '=']:
            # 精确匹配: 行首或管道符后 + 字段名 + 分隔符
            patterns = [
                re.compile(r'(?:^|\|)\s*' + re.escape(field_clean) + r'\s*' + re.escape(sep) + r'\s*(.+?)(?:\s*\||$)'),
            ]
            for pat in patterns:
                m = pat.search(line)
                if m:
                    val = m.group(1).strip().rstrip('，。,.;；')
                    if val and len(val) < 100 and val != field_clean:
                        found.append(val)
                        break
            else:
                continue
            break  # 如果已经匹配到了，跳过其他分隔符

        if found:
            # 已经在这行找到了，跳到下一行
            continue

        # ── 模式2: 管道符/表格格式 ──
        # ... | 课程性质 | 专业必修课 | ...
        if '|' in line:
            parts = [p.strip() for p in line.split('|') if p.strip()]
            for i, part in enumerate(parts):
                if part == field_clean and i + 1 < len(parts):
                    val = parts[i + 1].rstrip('，。,.;；')
                    if val and len(val) < 100 and val != field_clean:
                        found.append(val)
                        break

        if found:
            continue

        # ── 模式3: "是"连接 ──
        # 课程性质是专业必修课
        m = re.search(re.escape(field_clean) + r'是(.+?)(?:[,，。.;；\|]|$)', line)
        if m:
            val = m.group(1).strip()
            if val and len(val) < 100:
                found.append(val)
                continue

        # ── 模式4: 【字段】值 ──
        m = re.search(r'【' + re.escape(field_clean) + r'】\s*(.+?)(?:[,，。.;；\|]|$)', line)
        if m:
            val = m.group(1).strip()
            if val and len(val) < 100:
                found.append(val)
                continue

        # ── 模式5: 逗号分隔键值对 ──
        # 课程性质,专业必修课  (在CSV或简单格式中)
        m = re.search(re.escape(field_clean) + r'[,，]\s*([^,，\n]+)', line)
        if m:
            val = m.group(1).strip().rstrip('，。,.;；')
            if val and len(val) < 100 and val != field_clean:
                found.append(val)
                continue

    return found


@tool
def parse_knowledge_file(file_description: str, file_content: str, missing_fields: str) -> str:
    """
    解析用户上传的知识文件内容，提取文档所需的缺失信息。

    当Agent发现用户缺少某些字段信息，且用户上传了文件时，调用此工具从文件内容中提取缺失字段。

    Args:
        file_description: 文件描述（如文件名、类型等）
        file_content: 文件的文本内容
        missing_fields: 需要提取的缺失字段列表，用逗号分隔

    Returns:
        提取结果JSON字符串，包含每个字段的候选值
    """
    ctx = request_context.get() or new_context(method="parse_knowledge_file")

    fields = [f.strip() for f in missing_fields.split(',') if f.strip()]
    content = file_content[:5000]  # 扩大到5000字符

    result = {
        "file": file_description,
        "missing_fields": fields,
        "extracted": {},
        "summary": f"已从文件中搜索以下字段：{', '.join(fields)}"
    }

    # 使用增强的智能匹配提取
    for field in fields:
        found = _find_field_in_text(field, content)
        if found:
            # 去重
            unique = list(dict.fromkeys(found))
            result["extracted"][field] = unique[:5]  # 最多5个候选

    # 生成候选选择提示
    choice_lines = []
    for field, values in result["extracted"].items():
        if len(values) == 1:
            choice_lines.append(f"✅ {field}：{values[0]}（唯一匹配）")
        elif len(values) > 1:
            options = '、'.join([f"({chr(0x2460+i)}){v}" for i, v in enumerate(values)])
            choice_lines.append(f"❓ {field}：检测到多个候选值 {options}，请用户确认")

    unfound = [f for f in fields if f not in result["extracted"]]
    if unfound:
        choice_lines.append(f"⚠ 未找到：{', '.join(unfound)}")

    result["choice_prompt"] = '\n'.join(choice_lines)

    return json.dumps(result, ensure_ascii=False)
