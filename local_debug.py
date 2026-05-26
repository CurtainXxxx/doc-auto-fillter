"""
本地调试入口 — 直接在终端与 Agent 对话

用法:
    uv run python local_debug.py

注意:
    - 需要 .env 文件配置 DEEPSEEK_API_KEY
    - coze_coding_utils / coze_coding_dev_sdk 在本地不可用，
      此脚本会自动 mock 这些模块
"""

import os
import sys

# ── 1. Mock 平台专属模块（必须在 import agent 之前） ──────────────

import types
from unittest.mock import MagicMock

# 创建 mock 模块树
def _make_mock_module(name: str) -> types.ModuleType:
    mod = types.ModuleType(name)
    mod.__path__ = []  # 标记为 package
    mod.__all__ = []
    return mod

# coze_coding_utils
_cozu = _make_mock_module("coze_coding_utils")
_cozu_log = _make_mock_module("coze_coding_utils.log")
_cozu_log_write = _make_mock_module("coze_coding_utils.log.write_log")
_cozu_ctx = _make_mock_module("coze_coding_utils.runtime_ctx")
_cozu_ctx_context = _make_mock_module("coze_coding_utils.runtime_ctx.context")
_cozu_helper = _make_mock_module("coze_coding_utils.helper")
_cozu_helper_agent = _make_mock_module("coze_coding_utils.helper.agent_helper")
_cozu_helper_stream = _make_mock_module("coze_coding_utils.helper.stream_runner")

# coze_coding_dev_sdk
_sdk = _make_mock_module("coze_coding_dev_sdk")
_sdk_s3 = _make_mock_module("coze_coding_dev_sdk.s3")

# 注入 mock 的 request_context
from contextvars import ContextVar
_cozu_log_write.request_context = ContextVar("request_context", default=None)

# 注入 mock 的 default_headers / new_context
_cozu_ctx_context.default_headers = lambda ctx=None: {}
_cozu_ctx_context.new_context = lambda method="": MagicMock()

# 注入 mock 的 S3SyncStorage
class _MockS3SyncStorage:
    """本地 mock：文件保存到 /tmp，URL 返回 file:// 路径"""
    def __init__(self, **kwargs):
        self._tmp_dir = "/tmp"

    def upload_file(self, file_content: bytes, file_name: str, content_type: str = "") -> str:
        local_path = os.path.join(self._tmp_dir, os.path.basename(file_name))
        with open(local_path, "wb") as f:
            f.write(file_content)
        return file_name

    def generate_presigned_url(self, key: str, expire_time: int = 3600) -> str:
        return f"file:///tmp/{os.path.basename(key)}"

_sdk_s3.S3SyncStorage = _MockS3SyncStorage

# 注册所有 mock 模块到 sys.modules
sys.modules.update({
    "coze_coding_utils": _cozu,
    "coze_coding_utils.log": _cozu_log,
    "coze_coding_utils.log.write_log": _cozu_log_write,
    "coze_coding_utils.runtime_ctx": _cozu_ctx,
    "coze_coding_utils.runtime_ctx.context": _cozu_ctx_context,
    "coze_coding_utils.helper": _cozu_helper,
    "coze_coding_utils.helper.agent_helper": _cozu_helper_agent,
    "coze_coding_utils.helper.stream_runner": _cozu_helper_stream,
    "coze_coding_dev_sdk": _sdk,
    "coze_coding_dev_sdk.s3": _sdk_s3,
})

# ── 2. 加载环境变量 ──────────────────────────────────────────

from dotenv import load_dotenv
load_dotenv()

# 确保工作目录正确
os.environ.setdefault("COZE_WORKSPACE_PATH", os.path.dirname(os.path.abspath(__file__)))

# ── 3. 构建并运行 Agent ──────────────────────────────────────

from agents.agent import build_agent


def main():
    print("=" * 50)
    print("  教务文档自动填写助手 — 本地调试模式")
    print("  输入 'quit' 退出，'reset' 重置对话")
    print("=" * 50)
    print()

    agent = build_agent()
    thread_id = "local-debug"
    config = {"configurable": {"thread_id": thread_id}}

    while True:
        try:
            user_input = input("你: ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\n再见！")
            break

        if not user_input:
            continue
        if user_input.lower() in ("quit", "exit", "q"):
            print("再见！")
            break
        if user_input.lower() == "reset":
            thread_id = f"local-debug-{os.getpid()}"
            config = {"configurable": {"thread_id": thread_id}}
            print("对话已重置\n")
            continue

        try:
            result = agent.invoke(
                {"messages": [("user", user_input)]},
                config=config,
            )
            # 打印最后一条 AI 回复
            for msg in reversed(result.get("messages", [])):
                if msg.type == "ai" and msg.content:
                    print(f"\nAgent: {msg.content}\n")
                    break
        except Exception as e:
            print(f"\n❌ 调用出错: {e}\n")


if __name__ == "__main__":
    main()
