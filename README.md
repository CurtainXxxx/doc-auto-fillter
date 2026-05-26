# 智能教务文档填写系统

基于 LLM + LangGraph 的对话式 Word 文档自动填写系统。用户上传教务模板，通过对话或上传知识文件即可自动填充字段，保留原始格式并生成可下载的 docx 文件。

## 核心功能

| 功能 | 说明 |
|------|------|
| **对话式填写** | 上传模板后，Agent 识别字段并分批提问，用户对话完成填写 |
| **知识文件预填** | 上传教学大纲/成绩单，AI 自动提取字段值，用户确认即可 |
| **人机协同审核** | 预填后进入全屏审核界面，三色卡片标记置信度，一键确认生成 |
| **任意模板适配** | 内置 3 套教务模板，也支持上传任意 Word 模板自动识别字段 |
| **格式安全保护** | 填充时复用模板 XML 结构，7 层校验管线确保文档不损坏 |
| **填写对比** | 自动对比模板与生成文档，输出已填/未填/变更清单 |

## 技术栈

- **后端**：Python 3.12 / FastAPI / LangChain 1.0 / LangGraph 1.0
- **前端**：原生 HTML + CSS + JavaScript（单文件 SPA）
- **模型**：豆包 Seed（通过 OpenAI 兼容接口）
- **文档处理**：python-docx / OpenXML
- **存储**：S3 兼容对象存储 / 内存 Checkpointer

## 快速开始

### 环境准备

```bash
# 安装依赖（项目使用 uv 管理依赖，禁止使用 pip）
uv sync
```

### 启动服务

```bash
# 启动 HTTP 服务
bash scripts/http_run.sh -m http -p 5000
```

服务启动后访问 `http://localhost:5000` 即可使用 Web 界面。

### 使用方式

**方式一：对话式填写**
1. 选择内置模板或上传自定义模板
2. Agent 逐批询问字段值（每批 2-3 个）
3. 可随时上传知识文件加速填写
4. 确认后生成 docx 并下载

**方式二：人机协同填写**
1. 点击「人机协同填写」按钮
2. 选择模板 → 上传知识文件
3. AI 自动预填，三色卡片展示结果：
   - 🟢 绿色 = 高置信度预填
   - 🟡 黄色 = 低置信度需确认
   - ⬜ 灰色 = 需手动填写
4. 审核确认后一键生成

## 内置模板

| 模板名称 | 文件 | 字段数 |
|---------|------|--------|
| 评价报告 | 专业课程目标达成度评价报告模板 | 93 |
| 试卷分析 | 试卷分析模板 | 20 |
| 关联矩阵 | 考题与课程目标及毕业要求关联矩阵表 | 7 |

## 项目结构

```
.
├── config/
│   └── agent_llm_config.json     # Agent 模型配置 + 系统提示词 + 工具注册
├── docs/
│   └── ARCHITECTURE.md           # 项目架构文档
├── scripts/                      # 启动/构建脚本
├── assets/                       # 模板文件 + 静态资源
├── src/
│   ├── agents/
│   │   └── agent.py              # Agent 主逻辑（LangGraph 状态机）
│   ├── tools/
│   │   ├── template_analyzer.py  # 模板解析引擎（字段识别核心）
│   │   ├── edu_report_tool.py    # 文档生成引擎（填充+校验+导出）
│   │   ├── docx_validator.py     # 文档校验/对比/修复模块
│   │   ├── docx_preview.py       # 文档预览渲染（docx→HTML）
│   │   ├── prefill_tool.py       # AI 预填工具（知识文件→字段值提取）
│   │   └── knowledge_tool.py     # 知识文件解析工具
│   ├── storage/
│   │   └── memory/
│   │       └── memory_saver.py   # 会话记忆持久化
│   ├── utils/                    # 内置工具函数
│   └── main.py                   # Web 服务入口（FastAPI）
├── web/
│   └── index.html                # 前端 SPA（单文件）
├── tests/                        # 单元测试
└── pyproject.toml                # 项目配置 + 依赖声明
```

## API 接口

| 方法 | 路径 | 说明 |
|------|------|------|
| POST | `/run` | Agent 同步调用 |
| POST | `/stream_run` | Agent 流式调用（SSE） |
| POST | `/upload` | 上传知识文件（支持多文件） |
| POST | `/upload-template` | 上传自定义模板 |
| POST | `/prefill` | AI 预填（知识文件→字段值） |
| GET | `/template-preview` | 模板预览渲染 |
| GET | `/generated-preview` | 生成文档预览 |
| GET | `/download-docx` | 下载生成的 docx |
| GET | `/convert-pdf` | docx 转 PDF |

## 配置说明

### Agent 配置 (`config/agent_llm_config.json`)

```json
{
    "config": {
        "model": "doubao-seed-1-6-251015",
        "temperature": 0.7,
        "top_p": 0.9,
        "max_completion_tokens": 10000,
        "timeout": 600,
        "thinking": "disabled"
    },
    "sp": "系统提示词...",
    "tools": ["list_templates", "analyze_report_template", ...]
}
```

### 环境变量

| 变量 | 说明 | 默认值 |
|------|------|--------|
| `COZE_WORKSPACE_PATH` | 项目根目录 | `/workspace/projects` |
| `COZE_WORKLOAD_IDENTITY_API_KEY` | 模型 API Key | - |
| `COZE_INTEGRATION_MODEL_BASE_URL` | 模型 API 地址 | - |

## 安全特性

- **路径遍历防护**：所有接受文件路径的 API 端点均有白名单校验
- **文件大小限制**：上传文件最大 10MB
- **输入清理**：填入文本自动清理控制字符和 XML 特殊字符
- **格式污染防护**：填充时剥离直连格式属性，防止源文档格式污染模板
