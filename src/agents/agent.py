"""高校教务办公数字员工 - 专业课程目标达成度评价报告生成Agent"""

import os
import json
from typing import Annotated
from dotenv import load_dotenv

# 加载项目根目录的 .env 文件（用于配置外部模型API密钥等）
_workspace = os.getenv("COZE_WORKSPACE_PATH", "/workspace/projects")
load_dotenv(os.path.join(_workspace, ".env"), override=True)

from langchain.agents import create_agent
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
from tools.knowledge_tool import parse_knowledge_file

LLM_CONFIG = "config/agent_llm_config.json"

# 默认保留最近 20 轮对话 (40 条消息)
MAX_MESSAGES = 40


def _windowed_messages(old, new):
    """滑动窗口: 只保留最近 MAX_MESSAGES 条消息"""
    return add_messages(old, new)[-MAX_MESSAGES:]  # type: ignore


class AgentState(MessagesState):
    messages: Annotated[list[AnyMessage], _windowed_messages]


def build_agent(ctx=None):
    workspace_path = os.getenv("COZE_WORKSPACE_PATH", "/workspace/projects")
    config_path = os.path.join(workspace_path, LLM_CONFIG)

    with open(config_path, "r", encoding="utf-8") as f:
        cfg = json.load(f)

    # 支持外部模型API：设置 EXTERNAL_LLM_API_KEY 即可切换
    # 例如 DeepSeek: EXTERNAL_LLM_API_KEY=sk-xxx EXTERNAL_LLM_BASE_URL=https://api.deepseek.com
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
        extra_body={
            "thinking": {
                "type": cfg["config"].get("thinking", "disabled")
            }
        } if not ext_api_key else {},  # 外部API不需要thinking参数
        default_headers=default_headers(ctx) if ctx and not ext_api_key else {},
    )

    return create_agent(
        model=llm,
        system_prompt=cfg.get("sp"),
        tools=[list_templates, analyze_report_template, generate_edu_report,
               parse_knowledge_file, analyze_uploaded_template, generate_from_template],
        checkpointer=get_memory_saver(),
        state_schema=AgentState,
    )
