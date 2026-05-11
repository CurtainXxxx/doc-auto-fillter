#!/usr/bin/env python3
"""
本地测试脚本 —— 在本地完整运行 auto_fill 流程，无需上传 Coze。

用法:
  python local_test.py <模板文件.docx> <知识文件1.docx> [知识文件2.docx ...]

示例:
  # 一键自动填写
  python local_test.py "测试文件/评价报告模板.docx" "知识文件/示例数据.docx"

  # 仅分析模板（不填）
  python local_test.py --analyze-only "测试文件/评价报告模板.docx"

  # 分析知识文件（看提取内容）
  python local_test.py --extract-only "知识文件/示例数据.docx"

输出:
  - 生成的 docx 文件保存在 output/ 目录
"""

import sys
import os

# ── 0. 设置本地环境 ──────────────────────────────────────────
PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(PROJECT_DIR)

# 注入顺序很重要：_mocks 必须先于 src，才能拦截 Coze SDK 等模块
sys.path.insert(0, os.path.join(PROJECT_DIR, "_mocks"))
sys.path.insert(1, os.path.join(PROJECT_DIR, "src"))
sys.path.insert(2, PROJECT_DIR)

# 加载 .env
from dotenv import load_dotenv
env_file = os.path.join(PROJECT_DIR, ".env")
if os.path.exists(env_file):
    load_dotenv(env_file, override=True)
    print(f"[OK] 已加载 .env: {env_file}")
else:
    print("[!] 未找到 .env 文件，请先创建并填入 API Key")

# 设置 workspace 路径
workspace = os.getenv("COZE_WORKSPACE_PATH", PROJECT_DIR)
os.environ["COZE_WORKSPACE_PATH"] = workspace

# 注入 mock 模块（必须在导入真实代码之前）
mock_dir = os.path.join(PROJECT_DIR, "_mocks")
sys.path.insert(0, mock_dir)

# ── 1. 导入真实模块 ──────────────────────────────────────────
from tools.template_analyzer import analyze_template
from tools.knowledge_tool import _extract_text_from_file, _llm_extract_fields, _rule_extract_fields
from tools.edu_report_tool import (
    _simplify_fields, _expand_custom_data, _fill_custom_template,
    analyze_uploaded_template,
)

API_KEY = os.getenv("EXTERNAL_LLM_API_KEY", "")
BASE_URL = os.getenv("EXTERNAL_LLM_BASE_URL", "https://api.deepseek.com")
MODEL = os.getenv("EXTERNAL_LLM_MODEL", "deepseek-v4-pro")


def resolve_path(file_path: str) -> str:
    """将相对路径或绝对路径解析为绝对路径"""
    if os.path.isabs(file_path):
        return file_path
    # 尝试相对于当前目录
    full = os.path.join(PROJECT_DIR, file_path)
    if os.path.exists(full):
        return os.path.abspath(full)
    return os.path.abspath(file_path)


def cmd_analyze_template(template_path: str):
    """分析模板字段"""
    full = resolve_path(template_path)
    if not os.path.exists(full):
        print(f"[!] 文件不存在: {full}")
        return

    print(f"\n{'='*60}")
    print(f"分析模板: {os.path.basename(full)}")
    print(f"{'='*60}")

    analysis = analyze_template(full)
    user_fields = _simplify_fields(analysis["label_fields"])

    singles = [f for f in user_fields if f["fill_mode"] != "group"]
    groups = [f for f in user_fields if f["fill_mode"] == "group"]
    row_groups = analysis["row_groups"]

    print(f"\n字段总数: {len(user_fields)}（{len(singles)} 单字段 + {len(groups)} 分组字段）")
    print(f"行组: {len(row_groups)}")

    print("\n── 单字段 ──")
    for f in singles:
        print(f"  [{f['fill_mode']}] {f['label']}")

    print("\n── 分组字段 ──")
    for g in groups:
        subs = g.get("sub_labels", [])
        print(f"  [group] {g['label']} ({len(subs)}项): {subs[:10]}{'...' if len(subs) > 10 else ''}")

    if row_groups:
        print("\n── 行组 ──")
        for rg in row_groups:
            print(f"  {rg['group_id']}: {rg['template_row_count']}行×{rg['num_cols']}列, 表头: {rg['header_text'][:80]}")


def cmd_extract_knowledge(knowledge_path: str):
    """分析知识文件，显示提取内容"""
    full = resolve_path(knowledge_path)
    if not os.path.exists(full):
        print(f"[!] 文件不存在: {full}")
        return

    print(f"\n{'='*60}")
    print(f"知识文件: {os.path.basename(full)}")
    print(f"{'='*60}")

    text = _extract_text_from_file(full)
    if text.startswith("["):
        print(f"[!] 提取失败: {text}")
        return

    print(f"\n文本长度: {len(text)} 字符")
    print(f"\n── 内容预览（前2000字符）──")
    print(text[:2000])
    if len(text) > 2000:
        print(f"\n...（共 {len(text)} 字符，已截断）")


def cmd_auto_fill(template_path: str, knowledge_paths: list[str]):
    """完整自动填写流程"""
    template_full = resolve_path(template_path)
    if not os.path.exists(template_full):
        print(f"[!] 模板文件不存在: {template_full}")
        return

    knowledge_full_paths = []
    for kp in knowledge_paths:
        kfull = resolve_path(kp)
        if os.path.exists(kfull):
            knowledge_full_paths.append(kfull)
        else:
            print(f"[!] 知识文件不存在: {kp}")

    if not knowledge_full_paths:
        print("[!] 没有有效的知识文件")
        return

    # ── 步骤1: 分析模板 ──
    print(f"\n{'='*60}")
    print(f"步骤1: 分析模板")
    print(f"{'='*60}")
    print(f"模板: {os.path.basename(template_full)}")

    analysis = analyze_template(template_full)
    raw_label_fields = analysis["label_fields"]
    row_groups = analysis["row_groups"]
    user_fields = _simplify_fields(raw_label_fields)

    singles = sum(1 for f in user_fields if f["fill_mode"] != "group")
    groups = sum(1 for f in user_fields if f["fill_mode"] == "group")
    print(f"识别到 {len(user_fields)} 个简化字段（{singles}单 + {groups}分组）, {len(row_groups)} 行组")

    # 使用原始字段名提取（LLM 对具体字段名匹配更好）
    all_field_names = [f["label"] for f in raw_label_fields]

    # ── 步骤2: 读取知识文件 ──
    print(f"\n{'='*60}")
    print(f"步骤2: 读取知识文件")
    print(f"{'='*60}")

    all_content_parts = []
    for kpath in knowledge_full_paths:
        print(f"读取: {os.path.basename(kpath)}")
        text = _extract_text_from_file(kpath)
        if text and not text.startswith("["):
            all_content_parts.append(f"=== 文件: {os.path.basename(kpath)} ===\n{text}")
        else:
            print(f"  [!] 提取失败")

    if not all_content_parts:
        print("[!] 所有知识文件提取失败")
        return

    combined_content = "\n\n".join(all_content_parts)
    print(f"总内容: {len(combined_content)} 字符")

    # ── 步骤3: 提取字段值 ──
    print(f"\n{'='*60}")
    print(f"步骤3: 提取字段值（{len(all_field_names)}个原始字段）")
    print(f"{'='*60}")

    extracted_data = {}

    # 3a. LLM 智能提取
    if API_KEY and API_KEY != "你的DeepSeek_API_Key":
        print("使用 LLM 智能提取...")
        try:
            llm_result = _llm_extract_fields(all_field_names, combined_content)
            extracted_data.update(llm_result)
            print(f"  LLM 提取: {len(llm_result)} 个字段")
            for k, v in list(llm_result.items())[:10]:
                display_v = v[:80] + "..." if len(v) > 80 else v
                print(f"    {k}: {display_v}")
            if len(llm_result) > 10:
                print(f"    ...（共 {len(llm_result)} 个）")
        except Exception as e:
            import traceback
            print(f"  LLM 提取失败: {e}")
            traceback.print_exc()
    else:
        print("  [!] 未配置 API Key，跳过 LLM 提取")

    # 3b. 规则提取（补充 LLM 未提取到的字段）
    print("使用规则匹配补充...")
    rule_result = _rule_extract_fields(all_field_names, combined_content)
    for field, value in rule_result.items():
        if field not in extracted_data:
            extracted_data[field] = value
    new_from_rule = sum(1 for k in rule_result if k not in extracted_data)
    print(f"  规则补充: {len(rule_result)} 个字段（新增 {new_from_rule} 个）")

    print(f"\nLLM+规则合计: {len(extracted_data)}/{len(all_field_names)} 个字段")

    # ── 步骤4: 填充模板 ──
    print(f"\n{'='*60}")
    print(f"步骤4: 填充模板并生成文档")
    print(f"{'='*60}")

    try:
        # 展开数据（含比例自动计算等）
        expanded_data = _expand_custom_data(analysis, extracted_data)

        doc_bytes = _fill_custom_template(template_full, extracted_data)

        output_dir = os.path.join(PROJECT_DIR, "output")
        os.makedirs(output_dir, exist_ok=True)

        base_name = os.path.splitext(os.path.basename(template_full))[0]
        output_file = os.path.join(output_dir, f"{base_name}_已填充.docx")

        with open(output_file, "wb") as f:
            f.write(doc_bytes)

        final_filled = sum(1 for f in all_field_names if f in expanded_data)
        final_unfilled = [f for f in all_field_names if f not in expanded_data]

        print(f"\n[OK] 文档已生成: {output_file}")
        print(f"[OK] 大小: {len(doc_bytes):,} bytes")
        print(f"[OK] 填充率: {final_filled}/{len(all_field_names)} ({100*final_filled//max(1,len(all_field_names))}%)")
        if final_unfilled:
            print(f"未填: {len(final_unfilled)} 个")
            for f in final_unfilled[:8]:
                print(f"  - {f}")
            if len(final_unfilled) > 8:
                print(f"  ... 共 {len(final_unfilled)} 个")

        summary = {
            "template": os.path.basename(template_full),
            "knowledge_files": [os.path.basename(k) for k in knowledge_full_paths],
            "output": output_file,
            "filled_count": final_filled,
            "total_fields": len(all_field_names),
            "unfilled_fields": final_unfilled[:20] if final_unfilled else [],
            "extracted_data": {k: v for k, v in list(expanded_data.items())[:30]},
        }
        summary_file = output_file.replace(".docx", "_summary.json")
        import json
        with open(summary_file, "w", encoding="utf-8") as f:
            json.dump(summary, f, ensure_ascii=False, indent=2)
        print(f"[OK] 摘要: {summary_file}")

    except Exception as e:
        import traceback
        print(f"[!] 文档生成失败: {e}")
        traceback.print_exc()
        traceback.print_exc()


# ── 入口 ──────────────────────────────────────────────────────
if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        print("用法:")
        print("  python local_test.py <模板.docx> <知识文件1.docx> [知识文件2.docx ...]")
        print("  python local_test.py --analyze-only <模板.docx>")
        print("  python local_test.py --extract-only <知识文件.docx>")
        sys.exit(1)

    if sys.argv[1] == "--analyze-only":
        cmd_analyze_template(sys.argv[2])
    elif sys.argv[1] == "--extract-only":
        cmd_extract_knowledge(sys.argv[2])
    else:
        template = sys.argv[1]
        knowledge_files = sys.argv[2:]
        cmd_auto_fill(template, knowledge_files)
