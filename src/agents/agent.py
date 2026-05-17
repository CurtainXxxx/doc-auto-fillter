"""高校教务办公数字员工 - 专业课程目标达成度评价报告生成Agent"""

import os
import json
from typing import Annotated
from dotenv import load_dotenv

# 加载项目根目录的 .env 文件（用于配置外部模型API密钥等）
_workspace = os.getenv("COZE_WORKSPACE_PATH", "/workspace/projects")
load_dotenv(os.path.join(_workspace, ".env"), override=True)

from langchain.agents import create_agent
from langchain.agents.middleware import wrap_tool_call
from langchain.messages import ToolMessage, AIMessage
from langchain_openai import ChatOpenAI
from langgraph.graph import MessagesState
from langgraph.graph.message import add_messages
from langchain_core.messages import AnyMessage
from coze_coding_utils.runtime_ctx.context import default_headers
from storage.memory.memory_saver import get_memory_saver
from tools.edu_report_tool import (
    generate_edu_report, analyze_report_template, list_templates,
    analyze_uploaded_template, generate_from_template,
)
from tools.knowledge_tool import parse_knowledge_file, extract_facts

LLM_CONFIG = "config/agent_llm_config.json"

# 默认保留最近 20 轮对话 (40 条消息)
MAX_MESSAGES = 40


def _strip_reasoning(msg):
    """清理 DeepSeek 等模型返回的 reasoning_content，防止多轮对话报错"""
    if not isinstance(msg, AIMessage):
        return msg
    rc = getattr(msg, "reasoning_content", None)
    if not rc:
        return msg
    # 直接删除 reasoning_content 属性，而不是重建消息对象（避免丢失 id/tool_calls 关联）
    try:
        delattr(msg, "reasoning_content")
    except Exception:
        pass
    # 同时清理 additional_kwargs 中的 reasoning_content
    if hasattr(msg, "additional_kwargs") and "reasoning_content" in msg.additional_kwargs:
        msg.additional_kwargs.pop("reasoning_content", None)
    return msg


def _windowed_messages(old, new):
    """滑动窗口: 只保留最近 MAX_MESSAGES 条消息，并清理 reasoning_content"""
    merged = add_messages(old, new)[-MAX_MESSAGES:]  # type: ignore
    merged = [_strip_reasoning(m) for m in merged]
    # 修复滑动窗口裁剪后 ToolMessage 无对应 tool_calls 的问题
    merged = _fix_orphan_tool_messages(merged)
    return merged


def _fix_orphan_tool_messages(messages):
    """删除没有对应 AIMessage.tool_calls 的 ToolMessage，防止 API 400 错误"""
    # 收集所有有效的 tool_call_id
    valid_tool_call_ids = set()
    for m in messages:
        if isinstance(m, AIMessage) and hasattr(m, "tool_calls") and m.tool_calls:
            for tc in m.tool_calls:
                if "id" in tc:
                    valid_tool_call_ids.add(tc["id"])
    # 过滤掉孤立的 ToolMessage
    result = []
    for m in messages:
        if isinstance(m, ToolMessage):
            if m.tool_call_id not in valid_tool_call_ids:
                continue  # 跳过孤立的 ToolMessage
        result.append(m)
    return result


class AgentState(MessagesState):
    messages: Annotated[list[AnyMessage], _windowed_messages]


@wrap_tool_call
def handle_tool_errors(request, handler):
    """工具执行错误处理"""
    try:
        return handler(request)
    except Exception as e:
        return ToolMessage(
            content=f"工具执行出错: ({str(e)})",
            tool_call_id=request.tool_call["id"]
        )


@wrap_tool_call
def sanitize_before_llm(request, handler):
    """发送给LLM前清理孤立ToolMessage，防止400错误"""
    if hasattr(request, 'messages') and request.messages:
        valid_ids = set()
        for m in request.messages:
            for tc in (getattr(m, "tool_calls", None) or []):
                valid_ids.add(tc.get("id"))
        request.messages = [
            m for m in request.messages
            if getattr(m, "type", "") != "tool"
            or (getattr(m, "tool_call_id", None) in valid_ids)
        ]
    return handler(request)


def build_agent(ctx=None):
    workspace_path = os.getenv("COZE_WORKSPACE_PATH", "/workspace/projects")
    config_path = os.path.join(workspace_path, LLM_CONFIG)

    with open(config_path, "r", encoding="utf-8") as f:
        cfg = json.load(f)

    # 支持外部模型API：设置 EXTERNAL_LLM_API_KEY 即可切换
    ext_api_key = os.getenv("EXTERNAL_LLM_API_KEY")
    ext_base_url = os.getenv("EXTERNAL_LLM_BASE_URL")

    if ext_api_key and ext_base_url:
        # 使用外部模型API（如 DeepSeek）
        api_key = ext_api_key
        base_url = ext_base_url
        model = os.getenv("EXTERNAL_LLM_MODEL", "deepseek-chat")
    else:
        # 使用平台内置模型
        api_key = os.getenv("COZE_WORKLOAD_IDENTITY_API_KEY")
        base_url = os.getenv("COZE_INTEGRATION_MODEL_BASE_URL")
        model = cfg["config"].get("model", "doubao-seed-1-6-251015")

    llm = ChatOpenAI(
        model=model,
        api_key=api_key,
        base_url=base_url,
        temperature=cfg["config"].get("temperature", 0.7),
        streaming=True,
        timeout=cfg["config"].get("timeout", 600),
        extra_body=(
            {"thinking": {"type": "disabled"}} if ext_api_key else {
                "thinking": {
                    "type": cfg["config"].get("thinking", "disabled")
                }
            }
        ),
        default_headers=default_headers(ctx) if ctx and not ext_api_key else {},
    )

    return create_agent(
        model=llm,
        system_prompt=cfg.get("sp"),
        tools=[list_templates, analyze_report_template, generate_edu_report,
               parse_knowledge_file, extract_facts, analyze_uploaded_template, generate_from_template],
        checkpointer=get_memory_saver(),
        state_schema=AgentState,
        middleware=[handle_tool_errors, sanitize_before_llm],
    )
