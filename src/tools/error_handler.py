"""
error_handler — @tool 函数统一错误边界

职责：
1. 配置标准 Python logging（时间 / 级别 / 模块名 / 堆栈）
2. 提供 @tool_error_boundary 装饰器，替代 @tool 函数内部裸写的 try-except
3. 对常见业务异常（JSON解析、文件不存在等）做细化处理，其余异常记录完整堆栈

用法：
    from tools.error_handler import tool_error_boundary

    @tool
    @tool_error_boundary
    def my_tool(param: str) -> str:
        ...  # 不再需要 try-except
"""

import json
import logging
import functools
import traceback
from typing import Optional

# ──────────────────────────────────────────────────────────
# 1. 标准 logging 配置
# ──────────────────────────────────────────────────────────

logger = logging.getLogger("edu_agent.tools")

# 仅在无 handler 时初始化（防止重复添加）
if not logger.handlers:
    formatter = logging.Formatter(
        fmt="%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    # 控制台 handler（开发环境）
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    console_handler.setLevel(logging.DEBUG)
    logger.addHandler(console_handler)

    # 文件 handler（生产环境，写入标准日志目录）
    try:
        from pathlib import Path
        log_dir = Path("/app/work/logs/bypass")
        log_dir.mkdir(parents=True, exist_ok=True)
        file_handler = logging.FileHandler(log_dir / "app.log", encoding="utf-8")
        file_handler.setFormatter(formatter)
        file_handler.setLevel(logging.INFO)
        logger.addHandler(file_handler)
    except OSError:
        # 日志目录不可写时静默降级，不阻断业务
        pass

    logger.setLevel(logging.DEBUG)


# ──────────────────────────────────────────────────────────
# 2. 业务异常分类 & 用户友好消息映射
# ──────────────────────────────────────────────────────────

class _ErrorCategory:
    """异常分类：不同类别返回不同粒度的用户提示"""

    # 用户输入问题 → 提示具体原因
    INPUT = "input"
    # 系统基础设施问题 → 提示通用信息，隐藏内部细节
    INFRA = "infra"
    # 未知异常 → 提示通用信息，日志记录完整堆栈
    UNKNOWN = "unknown"

    @classmethod
    def values(cls):
        return {cls.INPUT, cls.INFRA, cls.UNKNOWN}


_EXCEPTION_MAP: dict[type, tuple[_ErrorCategory, str]] = {
    # ── 用户输入类 ──
    json.JSONDecodeError:  (_ErrorCategory.INPUT, "JSON格式错误，请检查输入数据"),
    FileNotFoundError:     (_ErrorCategory.INPUT, "文件不存在，请确认路径"),
    PermissionError:       (_ErrorCategory.INPUT, "文件权限不足，无法读取"),
    IsADirectoryError:     (_ErrorCategory.INPUT, "路径是目录而非文件"),
    ValueError:            (_ErrorCategory.INPUT, "参数值无效"),
    TypeError:             (_ErrorCategory.INPUT, "参数类型错误"),
    KeyError:              (_ErrorCategory.INPUT, "缺少必要字段"),

    # ── 基础设施类 ──
    ConnectionError:       (_ErrorCategory.INFRA, "网络连接失败，请稍后重试"),
    TimeoutError:          (_ErrorCategory.INFRA, "请求超时，请稍后重试"),
    OSError:               (_ErrorCategory.INFRA, "系统IO异常"),
    MemoryError:           (_ErrorCategory.INFRA, "内存不足"),
}


def _classify_exception(exc: Exception) -> tuple[_ErrorCategory, str]:
    """将异常分类并返回用户友好消息。

    精确匹配优先；若无精确匹配，按 MRO 顺序查找父类。
    """
    exc_type = type(exc)

    # 1. 精确匹配
    if exc_type in _EXCEPTION_MAP:
        category, friendly_msg = _EXCEPTION_MAP[exc_type]
        return category, f"{friendly_msg}（{exc}）"

    # 2. 按 MRO 向上查找父类匹配
    for parent in exc_type.__mro__[1:]:
        if parent in _EXCEPTION_MAP:
            category, friendly_msg = _EXCEPTION_MAP[parent]
            return category, f"{friendly_msg}（{exc}）"

    # 3. 完全未知的异常
    return _ErrorCategory.UNKNOWN, f"操作失败（{exc}）"


# ──────────────────────────────────────────────────────────
# 3. @tool_error_boundary 装饰器
# ──────────────────────────────────────────────────────────

def tool_error_boundary(func):
    """@tool 函数统一错误边界装饰器。

    功能：
    - 捕获所有未处理异常，防止 Agent 循环中断
    - 使用 logging.exception 记录完整堆栈（含 traceback）
    - 按异常分类返回用户友好的 JSON 错误消息
    - 保留原函数的 __name__、__doc__ 等元信息

    用法：
        @tool
        @tool_error_boundary
        def my_tool(param: str) -> str:
            # 不再需要 try-except，异常自动被边界捕获
            data = json.loads(param)
            ...

    注意：
    - 装饰器必须放在 @tool **下方**（@tool_error_boundary 先执行包装）
    - 如果函数内部需要自定义错误处理逻辑，不要使用此装饰器
    """

    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)

        except Exception as exc:
            func_name = func.__qualname__
            category, user_msg = _classify_exception(exc)

            # ── 日志记录（完整堆栈） ──
            # logger.exception 会自动附加 traceback，是最详细的日志
            logger.exception(
                "[%s] %s raised %s.%s",
                category,
                func_name,
                type(exc).__module__,
                type(exc).__qualname__,
            )

            # ── 补充上下文信息到 debug 日志 ──
            logger.debug(
                "[%s] args=%s kwargs_keys=%s",
                func_name,
                [type(a).__name__ for a in args],
                list(kwargs.keys()),
            )

            # ── 返回用户友好的 JSON ──
            # INPUT 类：暴露具体原因（用户能修复）
            # INFRA / UNKNOWN 类：隐藏内部细节（用户无法修复）
            result = {
                "success": False,
                "message": user_msg,
                "error_type": type(exc).__name__,
            }

            if category == _ErrorCategory.INPUT:
                # 用户输入错误：附加原始错误信息帮助排查
                result["detail"] = str(exc)

            return json.dumps(result, ensure_ascii=False)

    return wrapper
