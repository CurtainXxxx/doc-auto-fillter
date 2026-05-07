# 修改说明

## 版本 1.0.1 - 核心填充逻辑修复 (2026-05-07)

### 修复 1：`_set_tc_text` 彻底重写 (`src/tools/edu_report_tool.py`)

**问题**：原函数只修改单元格内第一个 `w:p` 段落的首个 `w:r` 的文本，其余段落和内嵌表格（`w:tbl`）原封不动。导致复杂单元格（如试卷分析模板的 Row 16 分析区域，含 20+ 段落和 3 个嵌套表格）填充后旧文本依然存在。

**修复**：
- 移除单元格内**所有**内容子元素（`w:p`、`w:tbl` 等），仅保留 `w:tcPr`（边框/宽度/合并信息等属性）
- 对含图表（`w:drawing` / `c:chart`）的单元格，保留全部原有内容，仅在最前插入新文本段落，防止图表被误删
- 创建单个新 `w:p` + `w:r`，从 `rPr_source` 复制字体/字号等格式

### 修复 2：docx 表格解析改用 lxml (`src/tools/knowledge_tool.py`)

**问题**：原函数用 `python-docx` 的 `row.cells` 遍历表格。`python-docx` 对垂直合并（vMerge）的单元格会缓存并返回相同 `Cell` 对象，导致多行共享同一底层 XML 元素，读取时文本重复或丢失。

**修复**：
- 使用 `table._tbl.findall(qn('w:tr'))` 逐行获取实际行 XML
- 使用 `tr.findall(qn('w:tc'))` 获取每行的实际单元格 XML，绕过合并缓存
- 自动识别 2 列 key-value 表格，输出为 `"标签：值"` 格式，方便 LLM 理解和匹配

### 影响范围

所有调用 `_set_tc_text` 的函数均受影响修复：
- `_fill_label_fields` - 标签字段填充
- `_fill_checkbox_rows_in_table` - 勾选框填充
- `_fill_simple_row_groups` - 行组填充
- `_fill_multi_col_field` - 多列字段填充
- `_build_report_docx` - 内置模板生成
- `_fill_custom_template` - 自定义模板生成
