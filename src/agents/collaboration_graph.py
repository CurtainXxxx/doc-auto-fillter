"""多智能体协作系统 - 3+1架构 StateGraph
架构: Supervisor(LLM意图理解) + DataAgent(数据理解层) + FillAgent(填充执行层) + DocAgent(输出层)
协商回路: FillAgent审查不通过 → 上报Supervisor → 调度DataAgent补充 → 回到FillAgent
"""

import os
import json
import re
from typing import Annotated, Literal, TypedDict

from dotenv import load_dotenv

_workspace = os.getenv("COZE_WORKSPACE_PATH", "/workspace/projects")
load_dotenv(os.path.join(_workspace, ".env"), override=True)

from langchain.agents import create_agent
from langchain.agents.middleware import wrap_tool_call
from langchain.messages import ToolMessage, AIMessage, HumanMessage, SystemMessage
from langchain_openai import ChatOpenAI
from langchain_core.messages import AnyMessage, BaseMessage
from langgraph.graph import StateGraph, START, END
from langgraph.graph.message import add_messages
from langgraph.graph.state import CompiledStateGraph
from coze_coding_utils.runtime_ctx.context import default_headers

from storage.memory.memory_saver import get_memory_saver
from tools.edu_report_tool import (
    generate_edu_report, analyze_report_template, list_templates,
    analyze_uploaded_template, generate_from_template,
    init_form_filling, get_form_status, update_form_fields, generate_form_document,
    _active_form_states,
)
from tools.knowledge_tool import parse_knowledge_file, extract_facts
from tools.prefill_tool import prefill_from_knowledge, prefill_from_multiple_knowledge
from tools.old_report_extractor import (
    extract_from_old_report, prefill_from_old_report, get_fill_checklist,
    inject_form_states as _inject_form_states,
)
from .specialist_prompts import (
    SUPERVISOR_PROMPT, DATA_AGENT_PROMPT, FILL_AGENT_PROMPT, DOC_AGENT_PROMPT,
)

LLM_CONFIG = "config/agent_llm_config.json"
MAX_MESSAGES = 40
MAX_REVIEW_LOOPS = 3

# ---- 路由决策的合法目标 ----
VALID_ROUTES = ["data_agent", "fill_agent", "doc_agent", "END"]


# ============================================================
# 共享状态
# ============================================================
class CollaborationState(TypedDict):
    messages: Annotated[list[AnyMessage], add_messages]
    # Supervisor 调度
    user_intent: str              # 用户意图: full_fill / modify_field / check_status
    next_agent: str               # Supervisor 决定的下一个 Agent
    supervisor_instruction: str   # Supervisor 给 Worker 的指令
    # DataAgent 产出
    knowledge_cache: dict         # {文件名: {事实列表}}
    template_fields: list         # 模板字段列表
    template_path: str            # 模板文件路径
    # FillAgent 产出
    session_id: str               # 填写会话ID
    fill_report: dict             # 审查报告 {pass, fill_rate, missing_fields, ...}
    confidence_scores: dict       # {字段名: 置信度}
    # 循环控制
    review_loops: int             # 审查循环计数
    retry_history: list           # 审查退回历史 [{loop, reason, missing_fields}]
    task_stage: str               # init / data_ready / filled / reviewed_pass / reviewed_fail / generated / agent_refused
    agent_refused_by: str         # 哪个 Agent 拒绝了任务


# ============================================================
# 工具分组
# ============================================================
DATA_AGENT_TOOLS = [
    list_templates, analyze_report_template, analyze_uploaded_template,
    extract_from_old_report, prefill_from_knowledge, prefill_from_multiple_knowledge,
    parse_knowledge_file, extract_facts, get_fill_checklist,
]

FILL_AGENT_TOOLS = [
    init_form_filling, update_form_fields, get_form_status, get_fill_checklist,
    prefill_from_old_report, prefill_from_knowledge, prefill_from_multiple_knowledge,
]

DOC_AGENT_TOOLS = [
    generate_form_document, generate_edu_report, generate_from_template,
]


# ============================================================
# 消息清理与错误处理（复用自 agent.py）
# ============================================================
def _strip_reasoning(msg):
    if not isinstance(msg, AIMessage):
        return msg
    rc = getattr(msg, "reasoning_content", None)
    if not rc:
        return msg
    try:
        delattr(msg, "reasoning_content")
    except Exception:
        pass
    if hasattr(msg, "additional_kwargs") and "reasoning_content" in msg.additional_kwargs:
        msg.additional_kwargs.pop("reasoning_content", None)
    return msg


def _fix_orphan_tool_messages(messages):
    valid_tool_call_ids = set()
    for m in messages:
        if isinstance(m, AIMessage) and hasattr(m, "tool_calls") and m.tool_calls:
            for tc in m.tool_calls:
                if "id" in tc:
                    valid_tool_call_ids.add(tc["id"])
    return [m for m in messages if not (
        isinstance(m, ToolMessage) and m.tool_call_id not in valid_tool_call_ids
    )]


@wrap_tool_call
def handle_tool_errors(request, handler):
    try:
        return handler(request)
    except Exception as e:
        return ToolMessage(
            content=f"工具执行出错: ({str(e)})",
            tool_call_id=request.tool_call["id"]
        )


# ============================================================
# LLM 构建（复用 agent.py 逻辑）
# ============================================================
def _build_llm(ctx=None):
    workspace_path = os.getenv("COZE_WORKSPACE_PATH", "/workspace/projects")
    config_path = os.path.join(workspace_path, LLM_CONFIG)

    with open(config_path, "r", encoding="utf-8") as f:
        cfg = json.load(f)

    ext_api_key = os.getenv("EXTERNAL_LLM_API_KEY")
    ext_base_url = os.getenv("EXTERNAL_LLM_BASE_URL")

    if ext_api_key and ext_base_url:
        api_key = ext_api_key
        base_url = ext_base_url
        model = os.getenv("EXTERNAL_LLM_MODEL", "deepseek-chat")
    else:
        api_key = os.getenv("COZE_WORKLOAD_IDENTITY_API_KEY")
        base_url = os.getenv("COZE_INTEGRATION_MODEL_BASE_URL")
        model = cfg["config"].get("model", "doubao-seed-1-6-251015")

    return ChatOpenAI(
        model=model,
        api_key=api_key,
        base_url=base_url,
        temperature=cfg["config"].get("temperature", 0.7),
        streaming=True,
        timeout=cfg["config"].get("timeout", 600),
        extra_body=(
            {"thinking": {"type": "disabled"}} if ext_api_key else {
                "thinking": {"type": cfg["config"].get("thinking", "disabled")}
            }
        ),
        default_headers=default_headers(ctx) if ctx and not ext_api_key else {},
    )


# ============================================================
# Supervisor 节点：LLM 意图理解 + 路由决策
# ============================================================
def _build_state_summary(state: CollaborationState) -> str:
    """构建给 Supervisor LLM 看的当前状态摘要"""
    summary_parts = [f"当前阶段: {state.get('task_stage', 'init')}"]
    summary_parts.append(f"用户意图: {state.get('user_intent', '未知')}")
    summary_parts.append(f"审查循环次数: {state.get('review_loops', 0)}/{MAX_REVIEW_LOOPS}")

    if state.get("template_path"):
        summary_parts.append(f"模板路径: {state['template_path']}")
    if state.get("template_fields"):
        fields = state["template_fields"]
        # 区分真实字段数据和兜底占位符
        is_placeholder = (
            len(fields) == 1 and isinstance(fields[0], dict)
            and (fields[0].get("_text_parsed") or fields[0].get("_from_summary"))
        )
        if is_placeholder:
            summary_parts.append("模板字段: ⚠️ 未获取到结构化字段（仅有文本摘要），需要DataAgent重新分析")
        else:
            summary_parts.append(f"模板字段: 已分析 {len(fields)} 个字段")
    if state.get("session_id"):
        summary_parts.append(f"填写会话: {state['session_id']}")
    if state.get("knowledge_cache"):
        summary_parts.append(f"知识缓存: {len(state['knowledge_cache'])} 个来源")

    fill_report = state.get("fill_report")
    if fill_report:
        summary_parts.append(
            f"审查结果: {'通过' if fill_report.get('pass') else '不通过'}, "
            f"填写率 {fill_report.get('fill_rate', '?')}"
        )

    retry_history = state.get("retry_history", [])
    if retry_history:
        summary_parts.append(f"退回历史: {len(retry_history)} 次")

    return "\n".join(f"- {p}" for p in summary_parts)


def _supervisor_node(state: CollaborationState, config):
    """Supervisor 节点：LLM 理解意图 → 决定路由 + 下发指令"""
    llm = getattr(_supervisor_node, "_llm", None)  # 从闭包获取构建时注入的 LLM（带 ctx）
    if llm is None:
        raise RuntimeError("Supervisor LLM not initialized. Call build_collaboration_graph first.")

    state_summary = _build_state_summary(state)
    system_prompt = SUPERVISOR_PROMPT.format(state_summary=state_summary)

    # 获取最近的用户消息作为意图理解依据
    recent_user_msgs = [
        m for m in state.get("messages", [])
        if hasattr(m, "type") and m.type == "human"
    ]
    last_user_msg = recent_user_msgs[-1].content if recent_user_msgs else "无用户消息"

    decision_prompt = f"""用户最新消息: "{last_user_msg}"

请根据当前状态做出路由决策。用以下格式回复（只输出这两行，不要其他内容）：
NEXT_AGENT: <data_agent|fill_agent|doc_agent|END>
INSTRUCTION: <给目标Agent的具体指令>"""

    messages = [
        SystemMessage(content=system_prompt),
        HumanMessage(content=decision_prompt),
    ]

    response = llm.invoke(messages)
    content = response.content if hasattr(response, "content") else str(response)

    # 解析 LLM 决策
    next_agent = "END"
    instruction = ""
    for line in content.strip().split("\n"):
        if line.upper().startswith("NEXT_AGENT:"):
            agent = line.split(":", 1)[1].strip().lower()
            if agent in VALID_ROUTES:
                next_agent = agent
        elif line.upper().startswith("INSTRUCTION:"):
            instruction = line.split(":", 1)[1].strip()

    # ---- 硬约束守卫（覆盖 LLM 不合理的决策） ----
    stage = state.get("task_stage", "init")
    fill_report = state.get("fill_report", {})
    review_loops = state.get("review_loops", 0)

    # 守卫0: 等待用户输入 → 必须返回 END（防止死循环）
    if stage == "waiting_user_input":
        next_agent = "END"

    # 守卫0b: Agent 拒绝了任务 → 根据拒绝者重新路由
    if stage == "agent_refused":
        refused_by = state.get("agent_refused_by", "")
        if refused_by == "fill_agent":
            # FillAgent 拒绝 → 说明任务是 DataAgent 的活，路由到 DataAgent
            next_agent = "data_agent"
            instruction = f"FillAgent表示无法执行该任务（任务超出其职责范围）。请DataAgent完成所需工作。原始指令: {instruction}"
        elif refused_by == "doc_agent":
            next_agent = "fill_agent"
            instruction = "DocAgent表示无法执行。请FillAgent先完成填充和审查。"
        elif refused_by == "data_agent":
            # DataAgent 拒绝（罕见），尝试让 FillAgent 直接处理
            next_agent = "fill_agent"
            instruction = f"DataAgent拒绝，请FillAgent尝试直接处理。原指令: {instruction}"

    # 守卫1: 初始状态，没有模板分析 → 必须走 DataAgent
    if stage == "init" and not state.get("template_fields"):
        next_agent = "data_agent"
        instruction = instruction or "请分析模板结构并提取上传材料中的知识"

    # 守卫2: DataAgent 完成，有模板分析 → 走 FillAgent
    if stage == "data_ready" and state.get("template_fields"):
        next_agent = "fill_agent"
        # ★ 指令审查：如果 LLM 的指令是给 DataAgent 的（分析模板/提取数据），必须替换
        _data_keywords = ["分析模板", "analyze_uploaded_template", "提取字段", "识别字段",
                          "分析.*docx", "extract_from_old_report", "prefill_from_knowledge"]
        _is_data_task = instruction and any(re.search(kw, instruction) for kw in _data_keywords)
        if _is_data_task or not instruction or instruction == "请根据当前状态自主判断需要做什么":
            instruction = (
                f"请初始化填写会话（模板路径: {state.get('template_path', '请从上下文获取')}），"
                "匹配知识到模板字段并批量填入，完成后做质量审查。"
            )

    # 守卫3: FillAgent 审查不通过 → 必须走 DataAgent 补充（关键协商回路）
    if stage == "reviewed_fail" and review_loops < MAX_REVIEW_LOOPS:
        missing = fill_report.get("missing_fields", [])
        reason = fill_report.get("review_note", "需要补充数据")
        # 如果 LLM 已经决策走 data_agent 且给出了具体指令，追加审查上下文而不是覆盖
        review_context = f"\n[审查退回-第{review_loops}次] 缺失: {missing}。{reason}"
        if next_agent == "data_agent" and instruction and instruction != "请根据当前状态自主判断需要做什么":
            instruction = instruction + review_context
        else:
            next_agent = "data_agent"
            instruction = f"审查不通过（第{review_loops}次）。缺失字段: {missing}。审查意见: {reason}。请针对性重新提取数据。"

    # 守卫3b: FillAgent 审查报告无法解析（stage=filled），当作通过处理，直接生成
    if stage == "filled":
        next_agent = "doc_agent"
        instruction = instruction or "填充已完成，请生成最终文档"

    # 守卫4: FillAgent 审查通过 → 走 DocAgent
    if stage == "reviewed_pass" or (stage == "reviewed_fail" and review_loops >= MAX_REVIEW_LOOPS):
        next_agent = "doc_agent"
        if review_loops >= MAX_REVIEW_LOOPS:
            instruction = f"审查已达上限（{MAX_REVIEW_LOOPS}次），强制生成。未通过项已在报告中标注。"
        else:
            instruction = instruction or "审查已通过，请生成最终文档"

    # 守卫5: DocAgent 完成 → END
    if stage == "generated":
        next_agent = "END"

    # 推断用户意图
    user_intent = state.get("user_intent", "full_fill")
    if any(kw in last_user_msg for kw in ["改", "修改", "更新", "换"]):
        user_intent = "modify_field"
    elif any(kw in last_user_msg for kw in ["查", "看", "状态", "进度"]):
        user_intent = "check_status"

    # ★ 提取模板路径：扫描所有消息和指令中的 /tmp/*.docx 路径
    template_path = state.get("template_path", "")
    if not template_path:
        # 扫描所有消息内容
        for m in state.get("messages", []):
            try:
                content = m.content if hasattr(m, "content") and isinstance(m.content, str) else ""
            except Exception:
                content = ""
            found = re.findall(r'/tmp/[\w.-]+\.docx?', content)
            if found:
                template_path = found[0]
                break
        # 也扫描 LLM 指令
        if not template_path:
            found = re.findall(r'/tmp/[\w.-]+\.docx?', instruction)
            if found:
                template_path = found[0]

    # 输出决策日志（评审可见）
    agent_names = {
        "data_agent": "📄 DataAgent（数据理解）",
        "fill_agent": "✏️ FillAgent（填充审查）",
        "doc_agent": "📦 DocAgent（文档生成）",
        "END": "✅ 任务完成",
    }
    decision_msg = (
        f"**[Supervisor]** → {agent_names.get(next_agent, next_agent)}\n"
        f"> {instruction}" if instruction else f"**[Supervisor]** → {agent_names.get(next_agent, next_agent)}"
    )

    return {
        "user_intent": user_intent,
        "next_agent": next_agent,
        "supervisor_instruction": instruction or "请根据当前状态自主判断需要做什么",
        "template_path": template_path,
        "messages": [AIMessage(content=decision_msg)],
    }


# ============================================================
# Worker 节点工厂
# ============================================================
def _make_worker_node(agent_graph: CompiledStateGraph, agent_name: str):
    """创建 Worker 节点：调用 create_agent 子图执行专项任务"""

    def _worker_node(state: CollaborationState, config):
        instruction = state.get("supervisor_instruction", "请开始工作")

        # 补充上下文给 Worker
        context_parts = [f"## Supervisor 指令\n{instruction}"]

        # ★ 关键：明确传递模板路径，防止 Agent 用错模板
        template_path = state.get("template_path", "")
        if template_path:
            context_parts.insert(0, f"## ⚠️ 当前模板文件（必须使用此路径！）\n{template_path}")
        if state.get("template_fields"):
            context_parts.append(f"## 模板字段\n{json.dumps(state['template_fields'], ensure_ascii=False)[:3000]}")
        if state.get("session_id"):
            context_parts.append(f"## 填写会话ID\n{state['session_id']}")
        if state.get("knowledge_cache"):
            context_parts.append(f"## 已有知识缓存\n{json.dumps(state['knowledge_cache'], ensure_ascii=False)[:3000]}")
        if state.get("fill_report"):
            context_parts.append(f"## 上次审查报告\n{json.dumps(state['fill_report'], ensure_ascii=False)[:2000]}")

        worker_input = "\n\n".join(context_parts)

        # 调用子 Agent
        result = agent_graph.invoke(
            {"messages": [HumanMessage(content=worker_input)]},
            config,
        )

        # 提取 Worker 的最终回复
        final_messages = result.get("messages", [])
        last_ai_msg = None
        for m in reversed(final_messages):
            if isinstance(m, AIMessage) and m.content and not (
                hasattr(m, "tool_calls") and m.tool_calls
            ):
                last_ai_msg = m.content
                break
        if last_ai_msg is None and final_messages:
            last_ai_msg = str(final_messages[-1].content) if hasattr(final_messages[-1], "content") else ""

        # ★ 只返回最后一条 AI 结论，不暴露中间工具调用/失败过程
        visible_msg = AIMessage(
            content=f"**[{agent_name.replace('_', ' ').title()}]**\n{last_ai_msg or '(任务完成)'}"
        )
        updates = {"messages": [visible_msg]}

        # ★ 检测 Agent 是否明确拒绝了任务
        _refusal_patterns = [
            r'超出.*职责范围', r'请派.*处理', r'无权', r'无法执行该任务',
            r'不在.*能力范围', r'此任务超出', r'职责范围.*请派',
        ]
        if last_ai_msg and any(re.search(p, last_ai_msg) for p in _refusal_patterns):
            updates["task_stage"] = "agent_refused"
            updates["agent_refused_by"] = agent_name
            return updates

        if agent_name == "data_agent":
            # 判断 DataAgent 是否真正完成了数据准备
            has_session = False
            has_template_path = False

            # 从 ToolMessage 解析结构化数据（兼容 label_fields / template_fields / summary）
            for m in final_messages:
                if isinstance(m, ToolMessage) and m.content:
                    try:
                        parsed = json.loads(m.content)
                        if isinstance(parsed, dict):
                            if "session_id" in parsed:
                                updates["session_id"] = parsed["session_id"]
                                has_session = True
                            # analyze_uploaded_template 返回 label_fields，不是 template_fields
                            if "label_fields" in parsed:
                                updates["template_fields"] = parsed["label_fields"]
                            elif "template_fields" in parsed:
                                updates["template_fields"] = parsed["template_fields"]
                            elif "summary" in parsed and isinstance(parsed.get("summary"), dict):
                                updates["template_fields"] = [{"_from_summary": True, "total": parsed["summary"].get("total_fields", 0)}]
                            if "template_path" in parsed:
                                updates["template_path"] = parsed["template_path"]
                                has_template_path = True
                    except (json.JSONDecodeError, TypeError):
                        pass

            # ★ 尝试从 [DATA_OUTPUT] 文本解析真实字段名
            if last_ai_msg and not updates.get("template_fields", None):
                parsed_fields = _parse_data_output(last_ai_msg)
                existing_fields = state.get("template_fields") or []
                if parsed_fields:
                    updates["template_fields"] = parsed_fields
                elif not existing_fields:
                    # 真没有字段数据才用兜底占位符
                    if re.search(r'(模板分析|字段总数|字段清单)', last_ai_msg):
                        updates["template_fields"] = [{"_text_parsed": True}]

            # 根据是否有数据准备结果来判定阶段
            if has_session and updates.get("template_fields"):
                updates["task_stage"] = "data_ready"
            elif updates.get("template_fields"):
                updates["task_stage"] = "data_ready"  # 模板已分析，即使无session也算ready
            else:
                updates["task_stage"] = "waiting_user_input"

        elif agent_name == "fill_agent":
            # 解析 FillAgent 的审查报告
            report = _parse_fill_report(last_ai_msg or "")
            if report:
                updates["fill_report"] = report
                reviews = state.get("review_loops", 0)
                if report.get("pass"):
                    updates["task_stage"] = "reviewed_pass"
                else:
                    updates["task_stage"] = "reviewed_fail"
                    updates["review_loops"] = reviews + 1
                    history = list(state.get("retry_history", []))
                    history.append({
                        "loop": reviews + 1,
                        "reason": report.get("review_note", ""),
                        "missing_fields": report.get("missing_fields", []),
                    })
                    updates["retry_history"] = history
                if report.get("confidence_scores"):
                    updates["confidence_scores"] = report["confidence_scores"]
            else:
                # 没有审查报告：判断是否只是问用户问题
                # 检查是否有 ToolMessage（说明调用了工具），如果没有则等用户输入
                has_tool_call = any(isinstance(m, ToolMessage) for m in final_messages)
                if has_tool_call:
                    updates["task_stage"] = "filled"
                else:
                    updates["task_stage"] = "waiting_user_input"

            # 提取 session_id
            for m in final_messages:
                if isinstance(m, ToolMessage) and m.content:
                    try:
                        parsed = json.loads(m.content)
                        if isinstance(parsed, dict) and "session_id" in parsed:
                            updates["session_id"] = parsed["session_id"]
                    except (json.JSONDecodeError, TypeError):
                        pass

        elif agent_name == "doc_agent":
            updates["task_stage"] = "generated"

        return updates

    return _worker_node


def _parse_data_output(content: str) -> list:
    """从 DataAgent 的 [DATA_OUTPUT] 文本中解析字段名列表，替代兜底占位符。
    策略：找 markdown 表格，识别"字段名"列位置，只提取该列的值。
    """
    block_match = re.search(r'\[DATA_OUTPUT\](.*?)\[/DATA_OUTPUT\]', content, re.DOTALL | re.IGNORECASE)
    text = block_match.group(1) if block_match else content

    fields = []
    field_name_col_idx = None  # "字段名"列的位置（0-based）
    lines = text.split('\n')

    for i, line in enumerate(lines):
        stripped = line.strip()
        if not stripped.startswith('|'):
            continue
        # 跳过分隔行
        if re.match(r'^\|[\s\-:|]+\|$', stripped):
            continue

        cells = [c.strip() for c in stripped.split('|')[1:-1]]

        # 检测表头：找"字段名"列
        if field_name_col_idx is None:
            for j, cell in enumerate(cells):
                if cell == '字段名':
                    field_name_col_idx = j
                    break
            if field_name_col_idx is None:
                # 可能没有"字段名"表头，检查是否是序号+字段名的表格（| # | 字段名 | ...）
                if len(cells) >= 2 and cells[0] == '#' and cells[1] == '字段名':
                    field_name_col_idx = 1
                elif len(cells) >= 2 and cells[0].isdigit() and 2 <= len(cells[1]) <= 30:
                    # 推测第2列是字段名（| 1 | 课程名称 | ...）
                    field_name_col_idx = 1
            continue  # 跳过表头行

        if field_name_col_idx is None or field_name_col_idx >= len(cells):
            continue

        candidate = cells[field_name_col_idx]

        # 过滤：不是有效字段名
        if not candidate or len(candidate) < 2 or len(candidate) > 40:
            continue
        if re.match(r'^\d+$', candidate):  # 纯数字
            continue
        if re.match(r'^\d+列$', candidate):  # 如"5列"
            continue
        if re.match(r'^\d+行$', candidate):  # 如"5行"
            continue
        if re.match(r'^T\d+_G\d+$', candidate):  # 行组ID如 T1_G1
            continue
        if candidate in ('字段名', '类型', '说明', '提示', '序号', '#', '行组ID', '列数', '模板行数', '列名'):
            continue
        # 过滤包含顿号的长字符串（如"考核环节、课程目标1、课程目标2"）
        if '、' in candidate and len(candidate) > 20:
            continue

        fields.append(candidate)

    # 去重保持顺序
    seen = set()
    result = []
    for f in fields:
        if f not in seen:
            seen.add(f)
            result.append(f)
    return result


def _parse_fill_report(content: str) -> dict | None:
    """从 FillAgent 的输出中解析 [FILL_REPORT]...[/FILL_REPORT] 块"""
    pattern = r'\[FILL_REPORT\]\s*(.*?)\s*\[/FILL_REPORT\]'
    match = re.search(pattern, content, re.DOTALL | re.IGNORECASE)
    if not match:
        return None

    report_text = match.group(1)
    report = {}

    # 解析审查结论
    if re.search(r'审查结论[：:]\s*通过', report_text):
        report["pass"] = True
    elif re.search(r'审查结论[：:]\s*不通过', report_text):
        report["pass"] = False
    else:
        return None

    # 解析填写率
    fill_match = re.search(r'填写率[：:]\s*([\d.]+)%', report_text)
    if fill_match:
        report["fill_rate"] = float(fill_match.group(1))

    # 解析缺失字段（兼容3种格式）
    missing_fields = _parse_missing_fields(report_text)
    report["missing_fields"] = missing_fields

    # 解析审查意见（支持多行，直到下一个标签或块结束）
    note_match = re.search(r'审查意见[：:]\s*(.+?)$', report_text, re.MULTILINE | re.DOTALL)
    if note_match:
        note = note_match.group(1).strip()
        # 截取到第一个有意义换行+标点前的内容，但要足够长
        # 只需前200字符作为摘要
        report["review_note"] = note[:300].replace("\n", " ").strip()
    else:
        report["review_note"] = ""

    return report


def _parse_missing_fields(report_text: str) -> list:
    """从审查报告文本中解析缺失字段列表，兼容3种格式：
    格式A（标准）: 缺失必填字段: [字段1, 字段2, 字段3]
    格式B（中文列表）: 缺失必填字段: 字段1、字段2、字段3
    格式C（Markdown列表）:
        缺失必填字段:
        · 字段1 (T0_R0_C1)
        - 字段2 (T0_R1_C1)
    """
    # 格式A: 方括号包裹 [字段1, 字段2]
    bracket_match = re.search(r'缺失必填字段[：:]\s*\[(.+?)\]', report_text, re.DOTALL)
    if bracket_match:
        return [f.strip() for f in bracket_match.group(1).split(",") if f.strip()]

    # 格式C: Markdown/符号列表（- 或 · 开头），在"缺失必填字段"到下一个标签之间
    mf_start = re.search(r'缺失必填字段[：:]', report_text)
    if mf_start:
        after_label = report_text[mf_start.end():]
        # 截取到下一个标签（如"审查意见""逻辑错误"等）或 [/FILL_REPORT]
        next_tag = re.search(r'\n(?:审查意见|逻辑错误|已填[：:]|\Z)', after_label)
        section = after_label[:next_tag.start()] if next_tag else after_label

        # 提取符号列表项
        bullets = re.findall(r'^\s*[-·•]\s*(.+?)(?:\s*\([^)]*\))?\s*$', section, re.MULTILINE)
        if bullets:
            return [b.strip() for b in bullets if b.strip()]

        # 格式B: 顿号/逗号/中文逗号分隔的枚举
        # 如: "全部14个字段均未填写，包括：字段1、字段2、字段3"
        enum_text = section.strip()
        # 移除引导语（"全部X个字段均未填写" 等）
        enum_text = re.sub(r'^全部\d+个字段均未填写[，,：:]?\s*(?:包括[：:]?\s*)?', '', enum_text)
        if enum_text and len(enum_text) < 500:
            # 用中文/英文逗号、顿号分隔
            parts = re.split(r'[，,、]', enum_text)
            fields = [p.strip() for p in parts if p.strip() and len(p.strip()) < 50]
            if fields:
                return fields

    return []


# ============================================================
# 路由决策
# ============================================================
def _route_decision(state: CollaborationState) -> Literal["data_agent", "fill_agent", "doc_agent", "END"]:
    """根据 Supervisor 的 next_agent 决策路由"""
    next_agent = state.get("next_agent", "END")
    if next_agent in VALID_ROUTES and next_agent != "END":
        return next_agent  # type: ignore
    return "END"


# ============================================================
# 构建多智能体协作图
# ============================================================
def build_collaboration_graph(ctx=None) -> CompiledStateGraph:
    """构建 3+1 多智能体协作 StateGraph"""
    llm = _build_llm(ctx)

    # 注入共享状态
    _inject_form_states(_active_form_states)

    # 构建专项 Agent 子图
    data_agent = create_agent(
        model=llm,
        system_prompt=DATA_AGENT_PROMPT,
        tools=DATA_AGENT_TOOLS,
        middleware=[handle_tool_errors],
    )

    fill_agent = create_agent(
        model=llm,
        system_prompt=FILL_AGENT_PROMPT,
        tools=FILL_AGENT_TOOLS,
        middleware=[handle_tool_errors],
    )

    doc_agent = create_agent(
        model=llm,
        system_prompt=DOC_AGENT_PROMPT,
        tools=DOC_AGENT_TOOLS,
        middleware=[handle_tool_errors],
    )

    # 构建协作图
    workflow = StateGraph(CollaborationState)

    # 注入 LLM 到 Supervisor 节点（避免节点内部重新创建丢失 ctx）
    setattr(_supervisor_node, "_llm", llm)

    workflow.add_node("supervisor", _supervisor_node)
    workflow.add_node("data_agent", _make_worker_node(data_agent, "data_agent"))
    workflow.add_node("fill_agent", _make_worker_node(fill_agent, "fill_agent"))
    workflow.add_node("doc_agent", _make_worker_node(doc_agent, "doc_agent"))

    workflow.add_edge(START, "supervisor")

    # 所有 Worker 完成后回到 Supervisor
    workflow.add_edge("data_agent", "supervisor")
    workflow.add_edge("fill_agent", "supervisor")
    workflow.add_edge("doc_agent", "supervisor")

    # Supervisor 条件路由
    workflow.add_conditional_edges("supervisor", _route_decision, {
        "data_agent": "data_agent",
        "fill_agent": "fill_agent",
        "doc_agent": "doc_agent",
        "END": END,
    })

    return workflow.compile(checkpointer=get_memory_saver())
