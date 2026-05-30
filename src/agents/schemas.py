"""Agent 间结构化通信协议 — Pydantic 强类型约束。

DataAgent → TemplateAnalysis / KnowledgeExtraction
FillAgent → FillReport
DocAgent → DocGenerationResult

所有 Worker 的最终产出都通过这些 Schema 校验，
消除正则解析文本块的不稳定性。
"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field


# ============================================================
# 通用
# ============================================================
class FieldDef(BaseModel):
    """单个模板字段定义"""
    name: str
    field_id: str = ""
    type: Literal["single", "group", "datarow"] = "single"


# ============================================================
# DataAgent 产出
# ============================================================
class ExtractedFact(BaseModel):
    """从材料中提取的单条事实"""
    field_name: str
    value: Any
    confidence: Literal["high", "medium", "low"] = "medium"
    source: str = ""  # 来源文件名


class DataAgentOutput(BaseModel):
    """DataAgent 的统一产出（模板分析 + 知识提取）"""
    template_type: str = ""
    total_fields: int = 0
    fields: list[FieldDef] = []
    facts: list[ExtractedFact] = []
    source_files: list[str] = []


# ============================================================
# FillAgent 产出
# ============================================================
class FillReport(BaseModel):
    """FillAgent 的填充质量审查报告"""
    passed: bool
    fill_rate: float = 0.0
    filled_count: int = 0
    high_conf_count: int = 0
    medium_conf_count: int = 0
    low_conf_count: int = 0
    missing_fields: list[str] = []
    logic_errors: list[str] = []
    review_note: str = ""


# ============================================================
# DocAgent 产出
# ============================================================
class DocGenerationResult(BaseModel):
    """DocAgent 的文档生成结果"""
    file_path: str = ""
    template_type: str = ""
    filled_count: int = 0
    success: bool = True
    error_message: str = ""
