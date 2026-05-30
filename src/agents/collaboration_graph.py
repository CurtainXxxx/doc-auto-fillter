"""多智能体协作系统 - 3+1架构 StateGraph (v2)

架构: Supervisor(LLM意图理解) + DataAgent(数据理解层) + FillAgent(填充执行层) + DocAgent(输出层)
协商回路: FillAgent审查不通过 → 上报Supervisor → 调度DataAgent补充 → 回到FillAgent

v2 改进:
- Command(goto=...) 动态路由替代 conditional_edges
- _resolve_transition() 跳转表替代 7 个顺序守卫
- with_structured_output(Pydantic) 替代正则解析文本块
- agent_refused 仅在 structured_output 失败时触发（不通过文本模式匹配）
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
from langgraph.types import Command
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
from .schemas import (
    DataAgentOutput, FillReport as FillReportSchema, DocGenerationResult,
)

LLM_CONFIG = "config/agent_llm_config.json"
MAX_MESSAGES = 40
MAX_REVIEW_LOOPS = 3

# ---- Worker 输出 Schema 映射 ----
WORKER_OUTPUT_SCHEMAS = {
    "data_agent": DataAgentOutput,
    "fill_agent": FillReportSchema,
    "doc_agent": DocGenerationResult,
}

# ---- 结构化提取的系统提示词 ----
EXTRACTION_PROMPTS = {
    "data_agent": (
        "根据以上对话中的工具调用结果，提取模板分析和知识提取的结构化信息。\n"
        "- 从 analyze_uploaded_template / analyze_report_template 的输出中提取 template_type、total_fields、fields\n"
        "- 从 extract_from_old_report / prefill_from_knowledge / extract_facts 的输出中提取 facts\n"
        "- 置信度从高/中/低映射为 high/medium/low\n"
        "- source_files 填写所有被处理的文件路径"
    ),
    "fill_agent": (
        "根据以上对话中的工具调用结果，生成填充质量审查报告。\n"
        "- 从 get_form_status 输出中判断填写率（filled/total）\n"
        "- 审查通过标准：填写率 ≥ 80% 且所有必填字段已填\n"
        "- missing_fields: 列出所有未填写的必填字段名\n"
        "- logic_errors: 列出数值不一致、前后矛盾等逻辑问题\n"
        "- review_note: 如果不通过，明确说明需要DataAgent补充什么数据"
    ),
    "doc_agent": (
        "根据以上对话中的工具调用结果，提取文档生成结果。\n"
        "- 从 generate_form_document / generate_edu_report / generate_from_template 的输出中提取 file_path\n"
        "- success=false 时填写 error_message"
    ),
}

# ---- Agent 中文名（日志用） ----
AGENT_DISPLAY_NAMES = {
    "data_agent": "DataAgent (数据理解)",
    "fill_agent": "FillAgent (填充审查)",
    "doc_agent": "DocAgent (文档生成)",
}


# ============================================================
# 共享状态
# ============================================================
class CollaborationState(TypedDict):
    messages: Annotated[list[AnyMessage], add_messages]
    # Supervisor 调度
    user_intent: str              # 用户意图: full_fill / modify_field / check_status
    supervisor_instruction: str   # Supervisor 给 Worker 的指令
    # DataAgent 产出
    knowledge_cache: dict         # {文件名: {事实列表}}
    template_fields: list         # 模板字段列表
    template_path: str            # 模板文件路径
    # FillAgent 产出
    session_id: str               # 填写会话ID
    fill_report: dict             # 审查报告 (FillReportSchema.model_dump())
    # DocAgent 产出
    doc_result: dict              # 文档生成结果 (DocGenerationResult.model_dump())
    # 循环控制
    review_loops: int             # 审查循环计数
    retry_history: list           # 审查退回历史 [{loop, reason, missing_fields}]
    task_stage: str               # init / data_ready / filled / reviewed_pass / reviewed_fail / generated / agent_refused / waiting_user_input
    agent_refused_by: str         # 哪个 Agent 的 structured_output 失败了


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
# 消息清理与错误处理
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
# LLM 构建
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
# 跳转表：单一入口，替代 7 个顺序守卫
# ============================================================
def _resolve_transition(stage: str, state: CollaborationState) -> str:
    """根据当前 stage + state 决定下一个 Agent。
    这是系统的硬路由规则，LLM 不会直接指定路由目标。
    """
    review_loops = state.get("review_loops", 0)
    user_intent = state.get("user_intent", "full_fill")

    # ---- 修改/查询意图：跳过流程，直接到 FillAgent ----
    if user_intent in ("modify_field", "check_status"):
        if stage in ("reviewed_pass", "generated", "reviewed_fail", "data_ready"):
            return "fill_agent"

    # ---- 按阶段决定路由 ----
    if stage == "init":
        return "data_agent" if not state.get("template_fields") else "fill_agent"

    if stage == "data_ready":
        return "fill_agent"

    if stage == "filled":
        return "doc_agent"

    if stage == "reviewed_pass":
        return "doc_agent"

    if stage == "reviewed_fail":
        if review_loops < MAX_REVIEW_LOOPS:
            return "data_agent"
        else:
            return "doc_agent"  # 超过最大循环，强制生成

    if stage == "generated":
        return "END"

    if stage == "waiting_user_input":
        return "END"

    if stage == "agent_refused":
        refused_by = state.get("agent_refused_by", "")
        if refused_by == "fill_agent":
            return "data_agent"
        if refused_by in ("doc_agent", "data_agent"):
            return "fill_agent"

    return "END"


# ============================================================
# Supervisor 节点：LLM 意图理解 → Command 动态路由
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
        is_placeholder = (
            len(fields) == 1 and isinstance(fields[0], dict)
            and (fields[0].get("_text_parsed") or fields[0].get("_from_summary"))
        )
        if is_placeholder:
            summary_parts.append("模板字段: 未获取到结构化字段（仅有文本摘要），需要DataAgent重新分析")
        else:
            summary_parts.append(f"模板字段: 已分析 {len(fields)} 个字段")
    if state.get("session_id"):
        summary_parts.append(f"填写会话: {state['session_id']}")
    if state.get("knowledge_cache"):
        summary_parts.append(f"知识缓存: {len(state['knowledge_cache'])} 个来源")

    fill_report = state.get("fill_report")
    if fill_report:
        summary_parts.append(
            f"审查结果: {'通过' if fill_report.get('passed') else '不通过'}, "
            f"填写率 {fill_report.get('fill_rate', '?')}%"
        )

    retry_history = state.get("retry_history", [])
    if retry_history:
        summary_parts.append(f"退回历史: {len(retry_history)} 次")

    return "\n".join(f"- {p}" for p in summary_parts)


def _extract_template_path_from_messages(state: CollaborationState, instruction: str) -> str:
    """从消息和指令中提取 /tmp/*.docx 模板路径"""
    template_path = state.get("template_path", "")
    if template_path:
        return template_path

    # 扫描所有消息内容
    for m in state.get("messages", []):
        try:
            content = m.content if hasattr(m, "content") and isinstance(m.content, str) else ""
        except Exception:
            content = ""
        found = re.findall(r'/tmp/[\w.-]+\.docx?', content)
        if found:
            return found[0]

    # 也扫描指令
    found = re.findall(r'/tmp/[\w.-]+\.docx?', instruction)
    if found:
        return found[0]

    return ""


def _detect_user_intent(state: CollaborationState) -> str:
    """从最近用户消息推断意图"""
    recent_user_msgs = [
        m for m in state.get("messages", [])
        if hasattr(m, "type") and m.type == "human"
    ]
    last_user_msg = recent_user_msgs[-1].content if recent_user_msgs else ""

    if any(kw in last_user_msg for kw in ["改", "修改", "更新", "换"]):
        return "modify_field"
    elif any(kw in last_user_msg for kw in ["查", "看", "状态", "进度"]):
        return "check_status"
    return state.get("user_intent", "full_fill")


def _make_supervisor_node(llm):
    """创建 Supervisor 节点（工厂函数，llm 显式传入避免闭包陷阱）"""

    def _supervisor_node(state: CollaborationState, config):
        state_summary = _build_state_summary(state)
        system_prompt = SUPERVISOR_PROMPT.format(state_summary=state_summary)

        # 获取最近的用户消息
        recent_user_msgs = [
            m for m in state.get("messages", [])
            if hasattr(m, "type") and m.type == "human"
        ]
        last_user_msg = recent_user_msgs[-1].content if recent_user_msgs else "无用户消息"

        # 推断用户意图
        user_intent = _detect_user_intent(state)

        # 提取模板路径
        instruction_hint = ""
        template_path = _extract_template_path_from_messages(state, instruction_hint)

        # LLM 分析意图并生成指令（不再解析路由目标，路由由跳转表决定）
        decision_prompt = (
            f"用户最新消息: \"{last_user_msg}\"\n\n"
            "请根据当前状态，为目标 Agent 写一条清晰的工作指令（不要指定路由目标，系统会自动路由）。\n"
            "格式：直接写自然语言指令即可，无需 NEXT_AGENT 标记。"
        )

        messages = [
            SystemMessage(content=system_prompt),
            HumanMessage(content=decision_prompt),
        ]

        response = llm.invoke(messages)
        instruction = response.content if hasattr(response, "content") else str(response)

        # 清理 LLM 可能输出的路由标记
        instruction = re.sub(r'\bNEXT_AGENT\s*:.*', '', instruction).strip()

        # ---- 跳转表决定路由目标 ----
        stage = state.get("task_stage", "init")
        target = _resolve_transition(stage, state)

        # 根据目标补充/调整指令
        if target == "data_agent" and stage == "init":
            if not instruction or len(instruction) < 10:
                instruction = "请分析模板结构并提取上传材料中的知识"
        elif target == "data_agent" and stage == "reviewed_fail":
            fill_report = state.get("fill_report", {})
            missing = fill_report.get("missing_fields", [])
            review_note = fill_report.get("review_note", "")
            review_loops = state.get("review_loops", 0)
            instruction = (
                f"审查不通过（第{review_loops}次）。缺失字段: {missing}。"
                f"审查意见: {review_note}。请针对性重新提取数据。"
            )
        elif target == "fill_agent" and stage in ("data_ready", "init"):
            instruction = (
                f"请初始化填写会话（模板路径: {template_path or '请从上下文获取'}），"
                "匹配知识到模板字段并批量填入，完成后做质量审查。"
            )
        elif target == "doc_agent":
            if stage == "reviewed_fail" and state.get("review_loops", 0) >= MAX_REVIEW_LOOPS:
                instruction = f"审查已达上限（{MAX_REVIEW_LOOPS}次），强制生成。未通过项已在报告中标注。"
            else:
                instruction = instruction or "审查已通过，请生成最终文档"

        # 输出决策日志（评审可见）
        display_name = AGENT_DISPLAY_NAMES.get(target, target)
        decision_msg = (
            f"**[Supervisor]** → {display_name}\n> {instruction}"
            if instruction else f"**[Supervisor]** → {display_name}"
        )

        return Command(
            goto=target,
            update={
                "user_intent": user_intent,
                "supervisor_instruction": instruction or "请根据当前状态自主判断需要做什么",
                "template_path": template_path,
                "messages": [AIMessage(content=decision_msg)],
            }
        )

    return _supervisor_node


# ============================================================
# Worker 节点工厂：ReAct + with_structured_output
# ============================================================
def _build_worker_context(state: CollaborationState, instruction: str) -> str:
    """构建 Worker 的上下文输入"""
    context_parts = [f"## Supervisor 指令\n{instruction}"]

    template_path = state.get("template_path", "")
    if template_path:
        context_parts.insert(0, f"## 当前模板文件（必须使用此路径！）\n{template_path}")

    if state.get("template_fields"):
        context_parts.append(
            f"## 模板字段\n{json.dumps(state['template_fields'], ensure_ascii=False)[:3000]}"
        )
    if state.get("session_id"):
        context_parts.append(f"## 填写会话ID\n{state['session_id']}")
    if state.get("knowledge_cache"):
        context_parts.append(
            f"## 已有知识缓存\n{json.dumps(state['knowledge_cache'], ensure_ascii=False)[:3000]}"
        )
    if state.get("fill_report"):
        context_parts.append(
            f"## 上次审查报告\n{json.dumps(state['fill_report'], ensure_ascii=False)[:2000]}"
        )

    return "\n\n".join(context_parts)


def _make_worker_node(agent_graph: CompiledStateGraph, agent_name: str,
                      llm: ChatOpenAI, output_schema: type):
    """创建 Worker 节点（工厂函数，llm/output_schema 显式传入避免闭包陷阱）

    流程:
    1. ReAct 循环：agent_graph.invoke() 执行工具调用
    2. with_structured_output(schema)：从对话中提取结构化数据
    3. 失败处理：structured_output 异常时设 agent_refused（不静默吞掉）
    """

    async def _worker_node(state: CollaborationState, config):
        instruction = state.get("supervisor_instruction", "请开始工作")
        worker_input = _build_worker_context(state, instruction)

        # ================================================================
        # 阶段 1：ReAct 循环（工具调用）
        # ================================================================
        result = agent_graph.invoke(
            {"messages": [HumanMessage(content=worker_input)]},
            config,
        )

        final_messages = result.get("messages", [])

        # 提取最后一条非工具调用的 AI 消息（用户可见）
        last_ai_msg = None
        for m in reversed(final_messages):
            if isinstance(m, AIMessage) and m.content and not (
                hasattr(m, "tool_calls") and m.tool_calls
            ):
                last_ai_msg = m.content
                break
        if last_ai_msg is None and final_messages:
            last_ai_msg = str(final_messages[-1].content) if hasattr(final_messages[-1], "content") else ""

        # ================================================================
        # 阶段 2：结构化输出提取（with_structured_output）
        # ================================================================
        structured_data = None
        try:
            structured_llm = llm.with_structured_output(output_schema)
            extraction_prompt = EXTRACTION_PROMPTS.get(agent_name, "请提取结构化信息。")

            extraction_messages = [
                SystemMessage(content=extraction_prompt),
            ] + final_messages[-30:]  # 最近30条消息作为上下文

            structured_result = structured_llm.invoke(extraction_messages)
            structured_data = structured_result.model_dump() if hasattr(structured_result, "model_dump") else structured_result
        except Exception as e:
            # structured_output 失败 → agent_refused（不静默吞掉）
            visible_msg = AIMessage(
                content=(
                    f"**[{AGENT_DISPLAY_NAMES.get(agent_name, agent_name)}]**\n"
                    f"{last_ai_msg or '(任务完成)'}\n\n"
                    f"⚠️ 结构化输出提取失败: {str(e)}"
                )
            )
            return Command(
                goto="supervisor",
                update={
                    "task_stage": "agent_refused",
                    "agent_refused_by": agent_name,
                    "messages": [visible_msg],
                }
            )

        # ================================================================
        # 阶段 3：处理结构化数据 → 状态更新
        # ================================================================
        visible_msg = AIMessage(
            content=f"**[{AGENT_DISPLAY_NAMES.get(agent_name, agent_name)}]**\n{last_ai_msg or '(任务完成)'}"
        )
        updates: dict = {"messages": [visible_msg]}

        if agent_name == "data_agent":
            if isinstance(structured_data, dict):
                if structured_data.get("fields"):
                    updates["template_fields"] = structured_data["fields"]
                if structured_data.get("template_type"):
                    pass  # template_type 记录在结构化数据中
                if structured_data.get("facts"):
                    # 合并到 knowledge_cache
                    cache = dict(state.get("knowledge_cache", {}))
                    for fact in structured_data["facts"]:
                        source = fact.get("source", "unknown")
                        if source not in cache:
                            cache[source] = []
                        cache[source].append(fact)
                    updates["knowledge_cache"] = cache
                if structured_data.get("source_files"):
                    pass  # 可在消息中展示

            # 判断 DataAgent 完成度
            if updates.get("template_fields"):
                updates["task_stage"] = "data_ready"
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

        elif agent_name == "fill_agent":
            if isinstance(structured_data, dict):
                updates["fill_report"] = structured_data
                reviews = state.get("review_loops", 0)
                if structured_data.get("passed"):
                    updates["task_stage"] = "reviewed_pass"
                else:
                    updates["task_stage"] = "reviewed_fail"
                    updates["review_loops"] = reviews + 1
                    history = list(state.get("retry_history", []))
                    history.append({
                        "loop": reviews + 1,
                        "reason": structured_data.get("review_note", ""),
                        "missing_fields": structured_data.get("missing_fields", []),
                    })
                    updates["retry_history"] = history

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
            if isinstance(structured_data, dict):
                updates["doc_result"] = structured_data

        return Command(goto="supervisor", update=updates)

    return _worker_node


# ============================================================
# 构建多智能体协作图
# ============================================================
def build_collaboration_graph(ctx=None) -> CompiledStateGraph:
    """构建 3+1 多智能体协作 StateGraph (v2)

    路由机制:
    - Supervisor 返回 Command(goto=跳转表目标) 实现动态路由
    - Worker 返回 Command(goto="supervisor") 回到协调者
    - 仅用 add_edge 做静态回退，不依赖 conditional_edges
    """
    llm = _build_llm(ctx)

    # 注入共享状态
    _inject_form_states(_active_form_states)

    # 构建专项 Agent 子图（ReAct 循环用）
    data_agent_graph = create_agent(
        model=llm,
        system_prompt=DATA_AGENT_PROMPT,
        tools=DATA_AGENT_TOOLS,
        middleware=[handle_tool_errors],
    )

    fill_agent_graph = create_agent(
        model=llm,
        system_prompt=FILL_AGENT_PROMPT,
        tools=FILL_AGENT_TOOLS,
        middleware=[handle_tool_errors],
    )

    doc_agent_graph = create_agent(
        model=llm,
        system_prompt=DOC_AGENT_PROMPT,
        tools=DOC_AGENT_TOOLS,
        middleware=[handle_tool_errors],
    )

    # 构建协作图
    workflow = StateGraph(CollaborationState)

    # Supervisor 节点（llm 显式传入）
    workflow.add_node("supervisor", _make_supervisor_node(llm))

    # Worker 节点（llm + output_schema 显式传入）
    workflow.add_node("data_agent", _make_worker_node(
        data_agent_graph, "data_agent", llm, DataAgentOutput
    ))
    workflow.add_node("fill_agent", _make_worker_node(
        fill_agent_graph, "fill_agent", llm, FillReportSchema
    ))
    workflow.add_node("doc_agent", _make_worker_node(
        doc_agent_graph, "doc_agent", llm, DocGenerationResult
    ))

    # 静态边：START → supervisor，所有 Worker → supervisor
    workflow.add_edge(START, "supervisor")
    workflow.add_edge("supervisor", END)  # 回退边（Command 优先）
    workflow.add_edge("data_agent", "supervisor")
    workflow.add_edge("fill_agent", "supervisor")
    workflow.add_edge("doc_agent", "supervisor")

    return workflow.compile(checkpointer=get_memory_saver())
