# 教务文档自动填充系统 — 项目架构文档

## 一、文件树

```
/workspace/projects/
│
├── .env                                  # 外部模型API密钥（DeepSeek）
├── .env.example                          # API密钥配置模板
├── .gitignore                            # Git忽略规则
├── pyproject.toml                        # 项目依赖声明（uv管理）
├── requirements.txt                      # pip依赖（兼容）
│
├── config/
│   └── agent_llm_config.json             # LLM配置 + 系统提示词 + 工具注册
│
├── assets/                               # 资源目录
│   ├── ...评价报告模板.docx               # 内置模板1
│   ├── ...试卷分析模板.docx               # 内置模板2
│   └── ...关联矩阵表模板.docx             # 内置模板3
│
├── src/                                  # 核心源码
│   ├── __init__.py
│   │
│   ├── main.py                           # HTTP服务入口（FastAPI，655行）
│   │
│   ├── agents/
│   │   ├── __init__.py
│   │   └── agent.py                      # Agent主逻辑（120行）
│   │
│   ├── tools/
│   │   ├── __init__.py
│   │   ├── template_analyzer.py          # 模板解析器（635行）
│   │   ├── edu_report_tool.py            # 文档生成工具（1286行）
│   │   └── knowledge_tool.py             # 知识文件提取（435行）
│   │
│   ├── storage/
│   │   ├── __init__.py
│   │   ├── memory/
│   │   │   ├── __init__.py
│   │   │   └── memory_saver.py           # 短期记忆（134行）
│   │   └── database/
│   │       ├── __init__.py
│   │       ├── db.py                     # 数据库连接（94行）
│   │       └── shared/
│   │           ├── __init__.py
│   │           └── model.py
│   │
│   └── utils/
│       ├── __init__.py
│       ├── helper.py
│       └── log/
│           ├── __init__.py
│           └── loop_trace.py
│
├── web/
│   └── index.html                        # 前端界面（1203行）
│
└── scripts/                              # 运维脚本
    ├── http_run.sh                       # 启动HTTP服务
    ├── local_run.sh                      # 本地运行
    ├── pack.sh                           # 打包
    ├── setup.sh                          # 环境初始化
    ├── load_env.py                       # 环境变量加载（Python）
    └── load_env.sh                       # 环境变量加载（Shell）
```

---

## 二、核心模块详解

### 2.1 main.py — HTTP服务（655行）

| API路由 | 方法 | 功能 |
|---|---|---|
| `/` | GET | 重定向到 /web/ |
| `/chat` | POST | Agent对话（SSE流式响应） |
| `/upload` | POST | 知识文件上传（返回文件路径+内容） |
| `/upload-template` | POST | 模板文件上传（返回路径+提取文本） |
| `/web/*` | GET | 静态文件（前端） |

---

### 2.2 agent.py — Agent主逻辑（120行）

```
build_agent()
  ├── 加载 .env（override=True）
  ├── 判断外部API or 内置API
  │   ├── 外部: DeepSeek（ChatOpenAI + base_url + api_key）
  │   └── 内置: doubao-seed（COZE_WORKLOAD_IDENTITY_API_KEY）
  ├── _strip_reasoning 中间件 → 清理DeepSeek reasoning_content
  └── create_agent(model, tools, checkpointer, state_schema)
```

**6个注册工具：**
1. `list_templates` — 列出内置模板
2. `analyze_report_template` — 分析内置模板字段
3. `generate_edu_report` — 生成内置模板文档
4. `analyze_uploaded_template` — 分析上传模板字段
5. `generate_from_template` — 生成上传模板文档
6. `parse_knowledge_file` — 从知识文件提取字段值

**短期记忆：** 滑动窗口保留最近20轮对话（40条消息）

---

### 2.3 template_analyzer.py — 模板解析器（635行）

**核心函数：** `analyze_template(path) → {label_fields, row_groups, checkbox_rows}`

**3种字段识别模式：**

| 模式 | 示例 | fill_mode |
|---|---|---|
| label_blank | `|课程名称|[空白]|` | set |
| colon | `课程名称：` | append |
| placeholder | `%` | replace |

**关键子函数：**
- `_is_label_cell()` — 判断标签格（2-15字，非数字，非选项词）
- `_is_placeholder_cell()` — 判断占位符格（%等）
- `_is_vmerge_continue()` — 判断vMerge延续格
- `_detect_multi_column_fields()` — 识别多列行组
- `_scan_table_data_sections()` — 扫描表格数据区

---

### 2.4 edu_report_tool.py — 文档生成工具（1286行）

#### 6个 @tool 工具

| 工具 | 参数 | 返回 |
|---|---|---|
| `list_templates()` | 无 | 3个模板名+文件路径 |
| `analyze_report_template(template_name)` | 模板名 | 简化字段列表 |
| `generate_edu_report(template_name, report_data)` | 模板名+JSON数据 | 下载链接 |
| `analyze_uploaded_template(file_path)` | 文件路径 | 归组字段列表 |
| `generate_from_template(file_path, report_data)` | 路径+JSON数据 | 下载链接 |
| — | — | — |

#### 格式保留核心

```
_set_tc_text(tc, text, rPr_source=None)
  ├── 修改已有run: 保留原rPr，只改w:t文本
  ├── 空白格: 从rPr_source继承格式（字体/字号/加粗）
  └── 保留: tcPr(单元格属性) + pPr(段落属性)

_find_label_rPr_in_row(tr) → 从同行找第一个有rPr的标签格
```

#### 两条填充路径

**内置模板路径：**
```
generate_edu_report
  → _expand_report_data()      # 用户数据→内部子字段（考勤/分数展开）
  → _build_report_docx()
      ├── _fill_label_fields()      # 标签字段填充
      ├── _fill_simple_row_groups() # 行组填充（vMerge安全）
      └── _fill_checkbox_rows()     # 勾选框行自动打√
  → S3上传 → 返回下载链接
```

**通用模板路径：**
```
generate_from_template
  → _expand_generic_data()     # 归组数据展开为内部子字段
  → _fill_custom_template()
      ├── _fill_label_fields()
      ├── _fill_simple_row_groups()
      └── _fill_checkbox_rows()
  → S3上传 → 返回下载链接
```

#### 字段简化

| 函数 | 用途 | 效果 |
|---|---|---|
| `_simplify_fields()` | 内置模板 | 硬编码映射，45→18字段 |
| `_simplify_generic_fields()` | 通用模板 | 前缀自动归组，92→21字段 |

#### 勾选框检测

```
_detect_checkbox_row(unique_cells)
  → 识别"选项+空白格"行（如 必修√ 选修□）
  → 返回 {label, option_blanks: {选项: 列索引}}

_fill_checkbox_rows_in_table(doc, data, analysis)
  → 遍历所有行，检测勾选框行
  → 根据用户值在正确选项后打√
```

---

### 2.5 knowledge_tool.py — 知识文件提取（435行）

**双引擎提取：**

```
parse_knowledge_file(file_path, required_fields)
  ├── _extract_text_from_file()     # 读取docx/pdf/txt内容
  ├── _rule_extract_fields()        # 规则匹配（快、免费）
  │   ├── 20+别名映射（学时数→总学时/学时/课时）
  │   ├── 5种匹配模式（管道符/冒号/等号/是/逗号）
  │   └── 返回: {字段: 值}
  ├── _llm_extract_fields()         # LLM提取（准、理解语义）
  │   ├── DeepSeek ChatOpenAI
  │   ├── 提示词: 根据字段列表从文本提取值
  │   └── 返回: {字段: 值}
  └── 合并: 规则结果 + LLM补充（LLM覆盖规则）
```

---

### 2.6 前端 index.html（1203行）

```
┌─────────────────────────────────────────────────┐
│  教务文档助手                                      │
├──────────────┬──────────────────────────────────┤
│  聊天面板     │  文档预览区                        │
│  (左侧40%)   │  (右侧60%)                        │
│              │                                    │
│  AI消息气泡   │  ┌──────────────────────────────┐ │
│  用户消息     │  │  基本信息                      │ │
│              │  │  课程名称: 待填写              │ │
│  ┌─────────┐│  │  课程代码: 待填写              │ │
│  │内置模板  ││  │  ...                          │ │
│  │上传模板  ││  └──────────────────────────────┘ │
│  │人机协同  ││                                    │
│  │知识上传  ││                                    │
│  └─────────┘│                                    │
│              │                                    │
│  [输入消息]   │                                    │
└──────────────┴──────────────────────────────────┘
```

**4个功能按钮：**
1. 🏛️ 内置模板 → 选择评价报告/试卷分析/关联矩阵
2. 📄 上传模板 → 上传任意.docx文件
3. ✍️ 人机协同编写（开发中）
4. 📎 知识文件上传 → 自动提取字段信息

---

## 三、数据流

```
用户输入
  │
  ▼
main.py (/chat)
  │
  ▼
agent.py (create_agent + LLM)
  │
  ├─── list_templates ──────→ TEMPLATE_REGISTRY ──→ assets/*.docx
  │
  ├─── analyze_report_template ──→ template_analyzer ──→ _simplify_fields() ──→ 18字段
  │
  ├─── generate_edu_report ──→ _expand_report_data() + _build_report_docx()
  │                              ──→ S3上传 ──→ 下载链接
  │
  ├─── analyze_uploaded_template ──→ template_analyzer ──→ _simplify_generic_fields() ──→ 21字段
  │
  ├─── generate_from_template ──→ _expand_generic_data() + _fill_custom_template()
  │                               ──→ S3上传 ──→ 下载链接
  │
  └─── parse_knowledge_file ──→ _rule_extract() + _llm_extract() ──→ {字段: 值}
```

---

## 四、模型配置

| 模块 | 内置模式 | 外部模式（.env） |
|---|---|---|
| Agent主LLM | doubao-seed-1-6-251015 | DeepSeek V4 Pro |
| 知识文件提取 | doubao-seed-1-6-lite | DeepSeek V4 Pro |

**.env 配置项：**
```
EXTERNAL_LLM_API_KEY=sk-xxx       # API密钥
EXTERNAL_LLM_BASE_URL=https://api.deepseek.com  # API地址
EXTERNAL_LLM_MODEL=deepseek-v4-pro  # 模型名
```

---

## 五、关键设计决策

1. **格式100%保留**：`_set_tc_text` 优先修改已有 `w:r` 的 `w:t`，不删除重建；空白格从同行标签格继承 `rPr`
2. **vMerge安全填充**：使用 `row._tr.findall(qn('w:tc'))` 获取每行独立tc元素，continue格消耗索引但不填值
3. **字段归组**：通用模板用前缀自动归组，内置模板用硬编码映射
4. **双引擎知识提取**：先规则匹配（快），再LLM补充（准）
5. **DeepSeek兼容**：`_strip_reasoning` 中间件清理 `reasoning_content`
6. **勾选框自动打√**：`_detect_checkbox_row` 识别选项行，在正确选项后填入√
