# 项目架构文档

## 一、系统总览

```
┌──────────────────────────────────────────────────────────┐
│                      用户（浏览器）                        │
│                  web/index.html (SPA)                     │
└──────────────┬──────────────────────────┬────────────────┘
               │ HTTP/SSE                 │ 文件上传
               ▼                          ▼
┌──────────────────────────────────────────────────────────┐
│                   FastAPI 服务层                           │
│                      src/main.py                          │
│  ┌──────────┐ ┌──────────┐ ┌──────────┐ ┌──────────┐    │
│  │ /run      │ │ /upload  │ │ /prefill │ │ /preview │    │
│  │ /stream   │ │ /template│ │          │ │ /download│    │
│  └─────┬────┘ └─────┬────┘ └─────┬────┘ └─────┬────┘    │
└────────┼────────────┼───────────┼────────────┼──────────┘
         │            │           │            │
         ▼            ▼           ▼            ▼
┌──────────────────────────────────────────────────────────┐
│                   Agent 层 (LangGraph)                     │
│                   src/agents/agent.py                     │
│                                                          │
│  ┌─────────────────────────────────────────────────┐     │
│  │              工具集 (9 个 @tool)                   │     │
│  │  ┌──────────┐ ┌──────────┐ ┌──────────────┐     │     │
│  │  │模板分析   │ │文档生成   │ │ 知识文件处理  │     │     │
│  │  │analyze   │ │generate  │ │ parse_know   │     │     │
│  │  │list      │ │          │ │ extract_facts│     │     │
│  │  └──────────┘ └──────────┘ └──────────────┘     │     │
│  │  ┌──────────────────────┐ ┌──────────────┐     │     │
│  │  │   AI 预填             │ │ 文档预览     │     │     │
│  │  │ prefill_single/multi │ │ (前端渲染)   │     │     │
│  │  └──────────────────────┘ └──────────────┘     │     │
│  └─────────────────────────────────────────────────┘     │
└──────────────────────────────────────────────────────────┘
         │            │           │
         ▼            ▼           ▼
┌──────────────┐ ┌─────────┐ ┌──────────────┐
│ 模板解析引擎  │ │校验管线  │ │ S3 对象存储   │
│ analyzer.py  │ │validator│ │ (文件存储)    │
└──────────────┘ └─────────┘ └──────────────┘
```

## 二、核心模块详解

### 2.1 模板解析引擎 (`template_analyzer.py` — 1169 行)

**职责**：解析任意 Word 模板，自动识别所有可填充字段。

**字段识别算法**（5 层扫描）：

```
输入：Word 模板 (.docx)
         │
         ▼
  ┌──────────────────┐
  │ Step 1: 表格扫描  │  遍历所有 table → 逐行逐列识别
  │ _scan_tables()   │  
  └────────┬─────────┘
           │
  ┌────────▼─────────┐
  │ Step 2: 段落扫描  │  检测正文中的"标签+下划线"字段
  │ _scan_paragraph  │  如 "教材名称______"
  │ _underline_fields│
  └────────┬─────────┘
           │
  ┌────────▼─────────┐
  │ Step 3: 标签提取  │  从每个单元格中提取标签文字
  │ _extract_labels  │  处理"标签（填写说明）"模式
  │ _from_cell()     │
  └────────┬─────────┘
           │
  ┌────────▼─────────┐
  │ Step 4: 去重合并  │  移除重复字段、合并相邻标签
  │ _deduplicate()   │
  └────────┬─────────┘
           │
  ┌────────▼─────────┐
  │ Step 5: 过滤      │  剔除黑名单、选项词、纯数字
  │ _filter()        │  识别 fill_mode（set/append/
  └────────┬─────────┘  replace/check/group）
           │
           ▼
  输出：字段清单 JSON
```

**支持的 6 种字段模式**：

| 模式 | 示例 | fill_mode |
|------|------|-----------|
| 标签+空白格 | `课程名称` `[_]` | `append` |
| 标签+占位符 | `题号` `[%]` | `set` |
| 标签+填写说明 | `主要教学经历（授课名称、起止时间...）` | `replace` |
| 多列数据行 | 表头下有空白行 | `group` |
| 勾选框行 | 必修/选修 选项行 | `check` |
| 段落下划线 | `教材名称______` | `paragraph_underline` |

### 2.2 文档生成引擎 (`edu_report_tool.py` — 1767 行)

**职责**：接收字段值，填充 Word 模板并生成文档。

**两条生成流程**：

```
内置模板路径：
  generate_edu_report(template_name, report_data)
      │
      ▼
  _get_template_path(name) → 内置模板文件
      │
      ▼
  _build_report_docx(template, data)
      │
      ├─ analyze_template()        → 解析字段
      ├─ _build_field_id_value_map → 用户数据→字段ID映射
      ├─ _fill_label_fields()      → 填充普通字段
      ├─ _fill_paragraph_fields()  → 填充段落下划线字段
      ├─ _fill_simple_row_groups() → 填充数据行组
      ├─ _fill_checkbox_rows()     → 填充勾选框行
      ├─ fix_docx()                → 自动修复结构
      ├─ validate_doc()            → 校验+Diff
      └─ upload_to_s3()            → 上传存储

自定义模板路径：
  generate_from_template(file_path, report_data)
      │
      ▼
  _fill_custom_template(template, data)
      │ （同上流程，仅模板来源不同）
```

**填充策略**：

- **复用 XML 结构**：不创建新 run，只修改现有 run 的文本，保证格式不变
- **安全守卫**：填充前检查字段类型，跳过不匹配的填充模式
- **文本清理**：`sanitize_fill_text()` 清理控制字符和 XML 特殊字符
- **行组扩展**：当数据行数超过模板预留行时，自动克隆行

### 2.3 校验与对比模块 (`docx_validator.py` — 747 行)

**7 层校验管线**：

```
validate_docx(doc)
    │
    ├─ 1. 结构完整性    → 检查文档是否可正常打开、有无损坏
    ├─ 2. 元素顺序      → 检查 OpenXML 子元素顺序合规
    ├─ 3. 单元格段落    → 每个	tc 至少包含一个 p
    ├─ 4. 合并单元格    → 检测合并单元格连续性
    ├─ 5. 表格维度      → 对比模板与生成文档的行列数
    ├─ 6. sectPr 位置   → sectPr 必须是 body 最后子元素
    └─ 7. 格式污染      → 检测直连格式属性污染风险
```

**Diff 对比**：

```
diff_docx(template_path, filled_path)
    │
    ├─ 对比每个表格的每个单元格文本
    ├─ 标记为 filled / still_empty / changed_structure
    └─ 输出填写率摘要 + 详细字段清单
```

**自动修复**：

```
fix_docx(path)
    │
    ├─ 修复缺少段落的单元格（插入空 <w:p>）
    ├─ 修复元素顺序错乱（按 OpenXML 规范重排）
    └─ 合并相邻的连续 run（减少碎片）
```

### 2.4 AI 预填工具 (`prefill_tool.py` — 660 行)

**预填流程**：

```
prefill_from_file_paths(file_paths, template_path)
    │
    ├─ 1. 解析知识文件内容（复用 knowledge_tool）
    ├─ 2. 解析模板字段清单（复用 template_analyzer）
    │
    ├─ 3. LLM 智能提取
    │     └─ 构造 Prompt → LLM 返回 {label, value, confidence, source}
    │
    ├─ 4. 规则匹配兜底
    │     └─ 正则匹配字段名→值（速度更快，覆盖常见模式）
    │
    ├─ 5. 合并结果
    │     └─ 同一字段取置信度最高的值
    │
    └─ 6. 三色分类
          ├─ confidence ≥ 0.8 → confirmed（绿）
          ├─ confidence ≥ 0.4 → review（黄）
          └─ confidence < 0.4 → empty（灰）
```

**多文件合并策略**：
- 每个文件独立提取 → 合并取最高置信度
- 同一字段多个文件值冲突 → 降级为 review

### 2.5 前端 (`web/index.html` — 2591 行)

**两种交互模式**：

```
对话模式（默认）：
  ┌────────────┬──────────────┐
  │  聊天窗口   │   文档预览    │
  │  Agent 问答 │   实时更新    │
  └────────────┴──────────────┘

人机协同模式：
  ┌────────────┬──────────────┐
  │  字段审核   │   文档预览    │
  │  🟢🟡⬜    │   实时更新    │
  │  进度条     │              │
  ├────────────┴──────────────┤
  │ [取消] [交给AI] [确认生成]  │
  └───────────────────────────┘
```

**核心交互机制**：

- `[FIELDS]...[/FIELDS]` 回显：Agent 回显字段值，前端解析后实时更新预览区
- 预览区保护：生成时不替换预览 DOM，用叠加遮罩 + 单元格文本更新
- 文件上传：支持多文件选择，FormData 批量提交
- 预填审核：三色卡片 + 进度条 + 一键确认

### 2.6 Web 服务 (`src/main.py` — 919 行)

**API 架构**：

```
                    ┌─ /run (同步)
                    ├─ /stream_run (SSE 流式)
Agent 调用 ─────────┤
                    ├─ /cancel/{id}
                    └─ /v1/chat/completions (OpenAI 兼容)

                    ├─ /upload (多文件)
文件管理 ───────────┤
                    └─ /upload-template (单文件)

                    ├─ /prefill (AI 预填)
预填/预览 ──────────┤
                    ├─ /template-preview
                    ├─ /generated-preview
                    ├─ /download-docx
                    └─ /convert-pdf
```

**安全防护**：

- `_safe_path()`：白名单路径校验，阻止路径遍历
- 文件大小限制：上传最大 10MB
- CORS：允许所有来源（开发模式）

## 三、数据流

### 3.1 完整对话流程

```
用户选择模板
    │
    ▼
Agent 调用 analyze_report_template("评价报告")
    │
    ▼
template_analyzer.analyze_template(doc)
    │ → 解析出 93 个字段
    │ → 分类：22 个普通字段 + 4 个数据行组 + 勾选行 + 段落字段
    │
    ▼
Agent 分批询问（每批 2-3 个字段）
    │ 用户回答 → Agent 回显 [FIELDS]...[/FIELDS]
    │ → 前端解析 → 更新预览区对应单元格
    │
    ▼ （用户可能上传知识文件）
Agent 调用 parse_knowledge_file / prefill_from_knowledge
    │ → 从文件提取字段值
    │ → Agent 回显预填结果
    │
    ▼ （所有字段收集完毕）
Agent 调用 generate_edu_report(template_name, report_data)
    │
    ▼
_build_report_docx(template, data)
    ├─ 解析字段 → 映射值 → 填充各类字段
    ├─ fix_docx() → 自动修复
    ├─ validate_doc() → 7 层校验 + Diff
    └─ upload_to_s3() → 返回下载 URL
```

### 3.2 预填流程

```
用户上传知识文件
    │
    ▼
前端调用 POST /prefill { file_paths, template_path }
    │
    ▼
prefill_from_file_paths(paths, template)
    ├─ 解析文件内容
    ├─ 解析模板字段
    ├─ LLM 提取 + 规则匹配
    └─ 合并 + 三色分类
    │
    ▼
前端渲染审核面板
    │ 用户确认/修正
    │
    ▼
前端调用 Agent 生成（传入审核后的字段值）
```

## 四、状态管理

### 4.1 Agent 记忆

```python
# agent.py — 滑动窗口机制
MAX_MESSAGES = 40  # 保留最近 40 条消息（约 10-15 轮对话）

class AgentState(MessagesState):
    messages: Annotated[list[AnyMessage], _windowed_messages]

# 窗口裁剪时：
# 1. 丢弃最老的消息
# 2. 清理 reasoning_content（防止序列化错误）
# 3. 修复孤立的 ToolMessage（保持 tool_calls 关联）
```

### 4.2 前端状态

```
字段映射 fieldMap = {
    "课程名称": { existing_value: "数据结构", field_id: "T0_R0_C0_L0" },
    ...
}

预填数据 prefillResult = {
    fields: [...],
    summary: "...",
    fill_rate: 0.6,
    ...
}
```

## 五、关键设计决策

### 5.1 为什么复用 XML 结构而不是创建新元素？

填充 Word 文档有两种策略：
- **创建新 run**：简单但丢失格式（字体、颜色、对齐方式）
- **复用现有 run**：只修改 `<w:t>` 文本节点，格式完整保留

本项目选择复用，因为教务文档对格式要求严格（表格边框、字体大小、单元格对齐）。

### 5.2 为什么模板解析不依赖 LLM？

模板解析是确定性任务（"这个单元格是不是空的？"），用规则算法更可靠：
- 速度快：解析 93 个字段 < 1 秒（LLM 需要 5-10 秒）
- 零幻觉：不会把非空单元格误判为待填
- 可调试：每个判断步骤都有明确规则

LLM 只用于需要语义理解的任务（知识文件提取、字段值推断）。

### 5.3 为什么前端用单文件 HTML 而不是 React/Vue？

- 部署简单：一个文件，无需构建
- 调试方便：浏览器直接查看源码
- 足够用：本项目的交互复杂度不需要框架

## 六、已知限制与改进方向

| 限制 | 说明 | 优先级 |
|------|------|--------|
| 记忆窗口有限 | 40 条消息后早期信息丢失 | P1 |
| 无会话恢复 | 刷新页面丢失所有数据 | P1 |
| prefill 无超时 | LLM 调用卡住时无重试 | P1 |
| 零测试覆盖 | tests/ 目录为空 | P2 |
| 无认证机制 | API 端点无需认证 | P2 |
| 移动端不适配 | 核心布局在手机上挤压 | P3 |

## 七、代码规模

| 模块 | 文件 | 行数 | 职责 |
|------|------|------|------|
| 模板解析 | `template_analyzer.py` | 1169 | 字段识别引擎 |
| 文档生成 | `edu_report_tool.py` | 1767 | 填充+校验+导出 |
| 文档校验 | `docx_validator.py` | 747 | 校验/对比/修复 |
| AI 预填 | `prefill_tool.py` | 660 | 知识文件→字段值 |
| 文档预览 | `docx_preview.py` | 395 | docx→HTML 渲染 |
| 知识解析 | `knowledge_tool.py` | 572 | 文件提取+LLM 提取 |
| Web 服务 | `main.py` | 919 | API 路由+安全防护 |
| Agent | `agent.py` | 160 | LangGraph 状态机 |
| 前端 | `index.html` | 2591 | SPA 交互界面 |
| **合计** | | **8980** | |
