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
    task_stage: str               # init / data_ready / filled / reviewed_pass / reviewed_fail / generated


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
        summary_parts.append(f"模板字段: 已分析 {len(state['template_fields'])} 个字段")
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


async def _supervisor_node(state: CollaborationState, config):
    """Supervisor 节点：LLM 理解意图 → 决定路由 + 下发指令"""
    llm = _supervisor_node._llm  # 从闭包获取构建时注入的 LLM（带 ctx）

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

    response = await llm.ainvoke(messages)
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

    # 守卫1: 初始状态，没有模板分析 → 必须走 DataAgent
    if stage == "init" and not state.get("template_fields"):
        next_agent = "data_agent"
        instruction = instruction or "请分析模板结构并提取上传材料中的知识"

    # 守卫2: DataAgent 完成，有模板分析 → 走 FillAgent
    if stage == "data_ready" and state.get("template_fields"):
        next_agent = "fill_agent"
        instruction = instruction or "请初始化填写会话，匹配知识到模板字段，并做质量审查"

    # 守卫3: FillAgent 审查不通过 → 必须走 DataAgent 补充（关键协商回路）
    if stage == "reviewed_fail" and review_loops < MAX_REVIEW_LOOPS:
        missing = fill_report.get("missing_fields", [])
        reason = fill_report.get("review_note", "需要补充数据")
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
        "messages": [AIMessage(content=decision_msg)],
    }


# ============================================================
# Worker 节点工厂
# ============================================================
def _make_worker_node(agent_graph: CompiledStateGraph, agent_name: str):
    """创建 Worker 节点：调用 create_agent 子图执行专项任务"""

    async def _worker_node(state: CollaborationState, config):
        instruction = state.get("supervisor_instruction", "请开始工作")

        # 补充上下文给 Worker
        context_parts = [f"## Supervisor 指令\n{instruction}"]
        if state.get("template_path"):
            context_parts.append(f"## 模板路径\n{state['template_path']}")
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
        result = await agent_graph.ainvoke(
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

        # 根据 Agent 类型更新状态
        updates = {"messages": final_messages}

        if agent_name == "data_agent":
            # 从 DataAgent 输出中解析模板和知识
            updates["task_stage"] = "data_ready"
            # 尝试从消息中提取 session_id 和 template_path
            for m in final_messages:
                if isinstance(m, ToolMessage) and m.content:
                    try:
                        parsed = json.loads(m.content)
                        if isinstance(parsed, dict):
                            if "session_id" in parsed:
                                updates["session_id"] = parsed["session_id"]
                            if "template_fields" in parsed:
                                updates["template_fields"] = parsed["template_fields"]
                            if "template_path" in parsed:
                                updates["template_path"] = parsed["template_path"]
                    except (json.JSONDecodeError, TypeError):
                        pass

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
                updates["task_stage"] = "filled"

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

    # 解析缺失字段
    missing_match = re.search(r'缺失必填字段[：:]\s*\[(.+?)\]', report_text, re.DOTALL)
    if missing_match:
        report["missing_fields"] = [f.strip() for f in missing_match.group(1).split(",") if f.strip()]
    else:
        report["missing_fields"] = []

    # 解析审查意见
    note_match = re.search(r'审查意见[：:]\s*(.+?)(?:\n|$)', report_text)
    if note_match:
        report["review_note"] = note_match.group(1).strip()

    return report


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
    _supervisor_node._llm = llm

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
