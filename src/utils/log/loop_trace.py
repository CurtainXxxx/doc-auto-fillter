# 桥接模块：coze_coding_utils/openai/handler.py 内部使用 from utils.log.loop_trace import ...
# 将请求转发到实际的 coze_coding_utils.log.loop_trace 模块
from coze_coding_utils.log.loop_trace import init_agent_config, init_run_config

__all__ = ['init_agent_config', 'init_run_config']
