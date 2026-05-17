# 通用模板引擎重构方案

> 目标：将系统从"按字段名猜着填"改为"按位置精准填"，支持任意模板上传。

---

## 当前映射链路（问题所在）

```
用户说话/知识文件
    ↓
Agent 理解后输出 [FIELDS]label=value[/FIELDS]
    ↓
前端 detectAndUpdateFields() → updateFieldByLabel(label, value)
    ↓ label 匹配
labelToFields[label] → [field_id_1, field_id_2, ...]
    ↓ 逐个更新
updateFieldById(field_id, value)

生成时:
Agent 输出 report_data = {"课程名称": "高等数学", ...}
    ↓ label 匹配
_fill_label_fields(doc, label_fields, data) → data[label] 匹配
```

**核心问题**：整条链路的"钥匙"是 `label`（字段名），不是 `field_id`（位置ID）。
同名字段会混淆，不同名但语义相同的字段会匹配不上。

---

## 目标映射链路

```
模板上传 → analyze_template() → 返回 field_id 清单（含语义描述）
    ↓
知识文件上传 → extract_facts() → 返回"事实表"（纯键值对，与模板无关）
    ↓
Agent 做映射：看 field_id 清单 + 事实表 → 输出 field_id=value
    ↓
前端：detectAndUpdateFields() → updateFieldById(field_id, value) 为主
后端：_fill_by_field_id(doc, field_id, value) 为主
    ↓
label 兼容：保留 updateFieldByLabel 作为 fallback
```

---

## 任务一：梳理映射链路

**做什么**：明确当前每个环节用的是什么 key，标注哪些要改。

### 当前各环节 key 使用情况

| 环节 | 文件 | 当前 key | 目标 key |
|------|------|---------|---------|
| 模板分析输出 | `template_analyzer.py` | `field_id` + `label` | 不变，但 `field_id` 语义要增强 |
| analyze 工具返回给 Agent | `edu_report_tool.py` `_simplify_fields()` | 只返回 `label` | 返回 `field_id` + `label` + 语义描述 |
| Agent 提示词 | `agent_llm_config.json` | 要求输出 `label=value` | 要求输出 `field_id=value` |
| Agent 回显 | `[FIELDS]` 协议 | `label=value` | `field_id=value`（兼容 `label=value`） |
| 前端实时更新 | `index.html` `detectAndUpdateFields()` | 先查 `fieldMap[key]` 再 fallback label | 优先 `field_id`，label 兜底 |
| 前端 tool 结果回填 | `applyFilledData()` / `applyPreciseFieldValues()` | 两条路径都有 | 不变，已支持 |
| 生成接口输入 | `report_data` 参数 | `{"label": "value"}` | `{"field_id": "value"}`，兼容 label |
| 生成接口填写 | `_fill_label_fields()` | `data[label]` 匹配 | `data[field_id]` 匹配，label 兜底 |
| 生成接口输出 | `filled_data` + `filled_field_values` | 两者都有 | 不变 |

### 结论

- `field_id` 在底层已经存在，但 **Agent ↔ 系统** 的交互层完全靠 `label`
- 改造重点：**Agent 输出用 field_id、生成接口用 field_id 查找、label 只做兜底**

---

## 任务二：field_id 为主、label 为兼容

### 2.1 增强 field_id 语义描述

**文件**：`template_analyzer.py`

当前 `field_id` 如 `T0_R1_C2_L0` 只表达了位置，Agent 无法从 ID 本身理解"这是填什么的"。

**改法**：analyze_template 输出中增加 `field_description` 字段：

```python
# 当前
{"field_id": "T0_R1_C2_L0", "label": "课程名称", "fill_mode": "append", ...}

# 改后
{
    "field_id": "T0_R1_C2_L0", 
    "label": "课程名称",           # 人类可读名，兼容用
    "field_description": "第1张表第1行第2格，冒号后内容，附近标签为'课程名称：'",  # Agent 用
    "fill_mode": "append",
    ...
}
```

`field_description` 由 `_extract_labels_from_cell` 和位置信息自动拼接，逻辑：

```
field_description = f"第{t_idx+1}张表第{r_idx+1}行第{c_idx+1}格" 
                    + f"，{fill_mode_desc}"
                    + f"，标签为'{label}'"
```

### 2.2 analyze 工具返回 field_id 清单

**文件**：`edu_report_tool.py` — `_simplify_fields()` 和 `analyze_report_template()`

**改法**：返回给 Agent 的字段列表中包含 `field_id` 和 `field_description`：

```python
# 当前 user_fields 只返回 label
{"label": "课程名称", "fill_mode": "set"}

# 改后
{
    "field_id": "T0_R1_C2_L0",
    "label": "课程名称",
    "field_description": "第1张表第1行第2格，冒号后内容，标签为'课程名称：'",
    "fill_mode": "set"
}
```

### 2.3 Agent 提示词改为要求输出 field_id

**文件**：`agent_llm_config.json`

```
# 当前
[FIELDS]
课程名称=高等数学
[/FIELDS]

# 改后（field_id 优先）
[FIELDS]
T0_R1_C2_L0=高等数学
[/FIELDS]

# 兼容写法（Agent 也可以用 label，系统会自动查找）
[FIELDS]
课程名称=高等数学
[/FIELDS]
```

提示词关键变更：
- 明确告知 Agent 每个字段的 `field_id`
- 要求优先用 `field_id=value` 格式输出
- 允许用 `label=value` 作为兜底（系统会查找对应的 field_id）

### 2.4 前端 detectAndUpdateFields 升级

**文件**：`index.html`

当前逻辑已经是先查 `fieldMap[key]` 再 fallback `labelToFields[key]`，只需确保：
- `[FIELDS]` 解析时，key 先当 field_id 查，查不到再当 label 查
- 当前代码已实现此逻辑（`if (fieldMap[key]) { updateFieldById } else { updateFieldByLabel }`）

**无需改动前端**，已兼容。

### 2.5 生成接口支持 field_id 输入

**文件**：`edu_report_tool.py` — `_fill_label_fields()` 和 `_expand_generic_data()`

**改法**：在 `_fill_label_fields` 前增加一层"field_id → label"转换：

```python
def _resolve_data_keys(label_fields, data):
    """将 data 中的 field_id 键转换为 label 键，合并同名值"""
    resolved = {}
    for key, value in data.items():
        if key.startswith("T") and "_R" in key and "_C" in key:
            # 看起来像 field_id，查找对应的 label
            matched = False
            for f in label_fields:
                if f["field_id"] == key:
                    resolved[f["label"]] = value
                    matched = True
                    break
            if not matched:
                resolved[key] = value  # 找不到就原样保留
        else:
            resolved[key] = value
    return resolved
```

在 `generate_edu_report` 和 `generate_from_template` 中，解析 data 后先调用 `_resolve_data_keys`。

### 2.6 涉及文件清单

| 文件 | 改动点 |
|------|--------|
| `template_analyzer.py` | 输出增加 `field_description` 字段 |
| `edu_report_tool.py` | `_simplify_fields()` 返回 `field_id` + `field_description`；新增 `_resolve_data_keys()` |
| `agent_llm_config.json` | 提示词改为要求 field_id 优先输出，告知每个字段的 field_id |
| `index.html` | 无需改动（已兼容） |

---

## 任务三：知识文件事实提取步骤

### 3.1 当前问题

当前 `parse_knowledge_file` 工具直接让 LLM 按 label 列表提取值，是"一步到位"：
- 传入 `["课程名称", "学时数", ...]`
- LLM 直接输出 `{"课程名称": "高等数学", ...}`
- 问题：label 和知识文件中的表述可能不匹配，LLM 容易漏提或错配

### 3.2 改为两步

**第一步：提取原始事实（与模板无关）**

新增工具 `extract_facts(file_path, file_content)`：

```python
@tool
def extract_facts(file_path: str) -> str:
    """从知识文件中提取所有事实信息，返回结构化事实表。
    此步骤与模板无关，只做纯粹的信息提取。
    
    Args:
        file_path: 知识文件路径
    """
```

LLM 提示词：
- 不限定字段名，让 LLM 自由提取文件中所有关键信息
- 输出格式：`{"事实1": "值1", "事实2": "值2", ...}`
- 值要保留原文表述，不做语义转换

**第二步：Agent 做 field_id 映射**

Agent 拿到事实表 + field_id 清单后，自行判断每个 field_id 应填什么值。
这一步不需要新工具，靠 Agent 的推理能力即可。

### 3.3 流程变更

```
# 当前
用户上传文件 → Agent 调用 parse_knowledge_file(字段名列表) → 直接拿到 label=value

# 改后
用户上传文件 → Agent 调用 extract_facts(文件路径) → 拿到事实表
             → Agent 结合 field_id 清单，判断映射关系
             → Agent 输出 [FIELDS] field_id=value [/FIELDS]
```

### 3.4 涉及文件清单

| 文件 | 改动点 |
|------|--------|
| `knowledge_tool.py` | 新增 `extract_facts` 工具函数；保留 `parse_knowledge_file` 兼容 |
| `agent.py` | 注册 `extract_facts` 工具 |
| `agent_llm_config.json` | 提示词增加两步提取流程指引；tools 数组增加 `extract_facts` |

---

## 任务四：生成后文档校验

### 4.1 校验内容

| 校验项 | 方法 | 失败处理 |
|--------|------|---------|
| 文件能否打开 | `Document(docx_bytes)` 不抛异常 | 返回错误，不提供给用户 |
| 表格数量不变 | 生成前后 `len(doc.tables)` 对比 | 返回错误 |
| 行数未暴增 | 每个表行数不超过模板的 120% | 返回警告 |
| 关键字段已填写 | 抽查前 5 个 field_id 对应位置非空 | 返回警告 |

### 4.2 实现

**文件**：`edu_report_tool.py` — 新增 `_validate_docx()`

```python
def _validate_docx(template_path: str, generated_bytes: bytes) -> dict:
    """校验生成的文档，返回 {valid: bool, warnings: [...], errors: [...]}"""
    errors = []
    warnings = []
    
    try:
        doc = Document(io.BytesIO(generated_bytes))
    except Exception as e:
        return {"valid": False, "errors": [f"文件无法打开: {e}"], "warnings": []}
    
    template_doc = Document(template_path)
    
    # 1. 表格数量
    if len(doc.tables) != len(template_doc.tables):
        errors.append(f"表格数量变化: 模板{len(template_doc.tables)}个, 生成{len(doc.tables)}个")
    
    # 2. 行数检查
    for i, (t_table, g_table) in enumerate(zip(template_doc.tables, doc.tables)):
        t_rows = len(t_table.rows)
        g_rows = len(g_table.rows)
        if g_rows > t_rows * 1.2:
            warnings.append(f"表{i}行数异常: 模板{t_rows}行, 生成{g_rows}行")
    
    # 3. 文件大小检查（不应比模板小太多，说明内容丢失）
    if len(generated_bytes) < os.path.getsize(template_path) * 0.3:
        errors.append("生成文件过小，可能内容丢失")
    
    return {
        "valid": len(errors) == 0,
        "errors": errors, 
        "warnings": warnings
    }
```

### 4.3 接入位置

在 `generate_edu_report` 和 `generate_from_template` 中，`_build_report_docx` / `_fill_custom_template` 返回后、上传前调用校验：

```python
validation = _validate_docx(template_path, docx_bytes)
if not validation["valid"]:
    return json.dumps({"success": False, "message": f"文档校验失败: {'; '.join(validation['errors'])}"})
```

### 4.4 涉及文件清单

| 文件 | 改动点 |
|------|--------|
| `edu_report_tool.py` | 新增 `_validate_docx()`；在两个生成函数中调用 |

---

## 任务五：自查与验证

### 5.1 最小语法自查

在生成接口返回前，对 `report_data` 做 JSON 格式校验：
- 已有 `json.loads` 会自动校验
- 增加：field_id 格式校验（以 `T` 开头包含 `_R` `_C`）

### 5.2 关键路径测试用例

| 用例 | 验证内容 |
|------|---------|
| 内置模板"试卷分析"，全部字段用 label 输入 | 兼容性不退化 |
| 内置模板"试卷分析"，部分字段用 field_id 输入 | field_id 通道生效 |
| 上传自定义模板 + 知识文件 → 两步提取 | 新流程跑通 |
| 同名字段模板（如两个"日期"） | field_id 区分不混淆 |
| 生成后校验：人为传入错误数据导致行数暴增 | 校验拦截生效 |

### 5.3 回归测试

每个任务完成后执行 `test_run`，确保：
- 内置模板生成正常
- 前端预览更新正常
- 打印功能正常
- 记忆功能正常

---

## 实施顺序与依赖关系

```
任务一（梳理链路）     ← 纯分析，无代码改动，最先做
     ↓
任务二（field_id 为主）← 核心改造，必须先做
     ↓
任务三（事实提取分离）  ← 依赖任务二的 field_id 清单
     ↓
任务四（生成后校验）   ← 独立，可与任务三并行
     ↓
任务五（自查验证）     ← 最后做，收尾
```

---

## 风险点

| 风险 | 影响 | 缓解 |
|------|------|------|
| Agent 不理解 field_id 格式 | 输出仍用 label | label 兜底机制保证不退化 |
| field_description 自动生成不够准确 | Agent 映射出错 | 允许 Agent 同时看 field_id + label + description |
| 两步提取增加延迟 | 用户体验变慢 | 事实提取可缓存，同文件不重复提取 |
| 校验误报 | 正常文档被拦截 | 只拦截 errors（表格数/文件打不开），warnings 只提示不阻断 |
