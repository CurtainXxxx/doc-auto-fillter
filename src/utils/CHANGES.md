# 修改说明

## 版本 1.1.2 - 字段标签分组修复 (2026-05-10 第三轮)

### 10. 修复 `_\d+$` 数字后缀字段无法分组
- `_simplify_fields` 和 `_expand_report_data` 新增 `_\d+$` 正则模式
- 如 `课程目标1_1` → base=`课程目标1`, suffix=`1`，自动归组显示
- **Why**: 原有 `_第\d+列` 正则无法匹配不带"第"的数字后缀，导致 `课程目标1_1`~`课程目标1_5` 显示为5个独立字段

### 11. 修复 `_detect_multi_column_fields` 生成垃圾字段标签
- 无 row_label 且无 above_label 时直接跳过，不再创建字段
- 纯数字列标签（如 "2"）替换为 `第N列` 格式
- **Why**: 关联矩阵模板中表头数字"2"被当作列标签，生成 `['2', '第2列', ...]` 等无意义字段

### 影响范围
- `src/tools/edu_report_tool.py` - `_simplify_fields`, `_expand_report_data`
- `src/tools/template_analyzer.py` - `_detect_multi_column_fields`

---

## 版本 1.1.0 - 任意模板自动填写 + 关键Bug修复 (2026-05-10)

### Coze测试反馈修复 (2026-05-10 第二轮)

#### 7. 修复 Agent 输出原始 JSON 过长
- 系统提示词新增严格规则：**严禁**输出工具返回的原始JSON/数据
- analyze_uploaded_template 返回几十个字段时，只告知总数和关键字段类型
- **Why**: Agent 把92个字段的完整JSON输出给用户，严重影响交互体验

#### 8. 修复 `_simplify_fields` 不支持动态列数
- 从硬编码 `_第1列`~`_第5列` 改为正则 `_第\d+列` 动态匹配
- 支持任意宽度的表格（如26列的试题号表格）
- 同步修复 `_expand_report_data` 使用相同正则
- 限制 `sub_labels` 显示数量（最多10个），避免输出过长
- **Why**: 原代码只处理前5列，第6~26列变成垃圾字段标签如"课程目标1_第26列"

#### 9. 修复 `_fill_simple_row_groups` 产生重复行
- 数据行数超过模板行数时，先尝试复用后续已存在的空行
- 只有真的不够用时才复制行
- **Why**: 二次提醒后生成文档出现试题号行重复4次、评价数据行重复，因盲目复制行导致

### 新增功能

#### 1. `auto_fill_from_knowledge` 工具 (`src/tools/edu_report_tool.py`)
一键自动填写：分析模板字段 → 从知识文件批量提取信息 → 填充生成文档。
- 支持多个知识文件（逗号分隔路径）
- LLM批量提取所有字段值，返回已填/未填字段
- 前端上传模板+知识文件后自动触发

#### 2. `_expand_custom_data` 函数 (`src/tools/edu_report_tool.py`)
为自定义模板展开用户数据中的分组字段。
- `"人数": "45,43,2"` → `"人数_第1列": "45", "人数_第2列": "43", ...`
- 支持逗号分隔字符串和列表
- 自动匹配模糊字段名

### Bug修复

#### 3. 修复 `_fill_multi_col_field` 填充链路 (`src/tools/edu_report_tool.py`)
- `_detect_multi_column_fields` 不再只创建单字段，同时存储 `fillable_cols`
- `_fill_custom_template` 先调用 `_expand_custom_data` 展开分组数据
- 展开后的数据由 `_fill_label_fields` 逐个填充

#### 4. 修复 `/upload-template` 表格解析 (`src/main.py`)
- 从 `row.cells` (合并单元格缓存Bug) 改为 lxml `tr.findall(qn('w:tc'))`
- 与 `knowledge_tool.py` 的表格解析保持一致

#### 5. 修复 `/upload` 文件生命周期 (`src/main.py`)
- 知识文件从临时路径改为持久路径 (`uploads/` 目录)
- 返回 `file_path` 供 `auto_fill_from_knowledge` 工具引用

### 前端改进 (`web/index.html`)

#### 6. 自动填写触发逻辑
- 追踪 `currentTemplatePath` 和 `knowledgeFilePaths`
- 模板+知识文件都上传后自动提示Agent调用 `auto_fill_from_knowledge`
- 欢迎消息更新为说明新功能

### 系统提示词更新 (`config/agent_llm_config.json`)

- 新增「自动填写流程」章节
- 添加 `auto_fill_from_knowledge` 工具说明
- 引导Agent在上传模板+知识文件后优先使用一键填写

### 影响范围
- `src/tools/edu_report_tool.py` - 新增 `_expand_custom_data`, `auto_fill_from_knowledge`；修复 `_fill_multi_col_field`, `_fill_custom_template`
- `src/main.py` - 修复 `/upload-template` 表格解析；修复 `/upload` 文件持久化
- `src/agents/agent.py` - 注册 `auto_fill_from_knowledge`
- `config/agent_llm_config.json` - 系统提示词和工具列表更新
- `web/index.html` - 文件路径追踪和自动触发

---

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
