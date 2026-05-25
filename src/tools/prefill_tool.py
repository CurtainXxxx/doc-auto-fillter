"""
prefill_tool.py — AI预填工具：从知识文件自动提取字段值并生成预填结果

核心流程：
1. 解析知识文件内容（复用 knowledge_tool 的文件提取能力）
2. 获取模板字段清单（复用 template_analyzer 的分析能力）
3. LLM智能提取字段值 + 评估置信度
4. 规则匹配兜底
5. 多文件结果合并
6. 返回带置信度的预填结果，供前端审核界面展示
"""

import os
import re
import json
from dotenv import load_dotenv

# 加载项目根目录的 .env 文件
_workspace = os.getenv("COZE_WORKSPACE_PATH", "/workspace/projects")
load_dotenv(os.path.join(_workspace, ".env"), override=True)

from langchain.tools import tool
from langchain_core.messages import SystemMessage, HumanMessage
from coze_coding_utils.log.write_log import request_context
from coze_coding_utils.runtime_ctx.context import new_context

from tools.knowledge_tool import _extract_text_from_file, _rule_extract_fields
from tools.template_analyzer import analyze_template


# ── 置信度阈值 ──
CONFIDENCE_CONFIRMED = 0.8    # ≥ 此值标记为 confirmed（绿）
CONFIDENCE_REVIEW = 0.4       # ≥ 此值标记为 review（黄），< 此值标记为 empty（灰）

# ── 单次LLM提取最大字段数（防止prompt过长） ──
MAX_FIELDS_PER_LLM_CALL = 40


def _classify_status(confidence: float) -> str:
    """根据置信度分类字段状态。"""
    if confidence >= CONFIDENCE_CONFIRMED:
        return "confirmed"
    elif confidence >= CONFIDENCE_REVIEW:
        return "review"
    else:
        return "empty"


def _llm_prefill_fields(field_list: list, file_content: str, ctx=None) -> list:
    """使用LLM从文件内容中提取字段值，带置信度评估。

    与 knowledge_tool._llm_extract_fields 的区别：
    - 返回带置信度的结构化结果（不只是 {字段: 值}）
    - 分批处理大量字段（避免prompt过长）
    - 更严格的提取Prompt（明确区分直接提取/推断/猜测）

    Args:
        field_list: 字段信息列表，每个元素为 {"label": str, "field_id": str}
        file_content: 知识文件文本内容
        ctx: 请求上下文

    Returns:
        list: [{"label", "field_id", "value", "confidence", "source"}]
    """
    # 限制文件内容长度
    max_chars = 10000
    if len(file_content) > max_chars:
        file_content = file_content[:max_chars] + "\n...(内容过长已截断)"

    # 分批处理
    all_results = []
    batches = [field_list[i:i + MAX_FIELDS_PER_LLM_CALL]
               for i in range(0, len(field_list), MAX_FIELDS_PER_LLM_CALL)]

    system_prompt = """你是一个教务文档信息提取专家，擅长从教学大纲、成绩单、课程计划等文件中精确提取结构化信息。

# 任务
根据模板字段清单，从提供的文件内容中提取每个字段的值，并评估置信度。

# 置信度定义
- 1.0：直接提取 — 文件中有完全匹配的原文（如标题写"数据结构与算法"→课程名称）
- 0.8：近似提取 — 文件中有语义相近的表述（如"专业核心课"→课程性质=必修）
- 0.6：合理推断 — 从多个信息综合推断（如多门课都是闭卷→推断考试形式）
- 0.3：低置信猜测 — 仅有微弱依据的猜测
- 0.0：未找到 — 文件中完全无关

# 提取规则
1. 直接提取：文件中有明确原文的，直接引用，confidence ≥ 0.8
2. 合理推断：文件中有间接依据的，标注推断逻辑，confidence 0.4-0.7
3. 无法提取：文件中完全无关的，value填null，confidence 0
4. 绝不编造不存在的值
5. 数字类字段只提取数字，不添加单位
6. 日期字段保持原文格式

# 输出格式
严格返回JSON数组，每个元素：
{
  "label": "字段名",
  "value": "提取的值或null",
  "confidence": 0.0到1.0,
  "source": "信息来源描述"
}"""

    for batch in batches:
        fields_desc = "\n".join(
            f"- {f['label']} (field_id: {f['field_id']})"
            for f in batch
        )

        user_prompt = f"""请从以下文件内容中提取这些字段的值：

字段清单：
{fields_desc}

文件内容：
{file_content}

请返回JSON数组，只返回JSON，不要其他文字。"""

        try:
            content_str = _call_llm(system_prompt, user_prompt, ctx)

            # 解析JSON
            batch_result = _parse_json_array(content_str)
            if batch_result:
                # 与输入字段匹配
                for f in batch:
                    matched = _find_matching_result(batch_result, f['label'], f['field_id'])
                    if matched:
                        all_results.append({
                            "label": f['label'],
                            "field_id": f['field_id'],
                            "value": matched.get("value"),
                            "confidence": float(matched.get("confidence", 0)),
                            "source": matched.get("source", ""),
                        })
                    else:
                        all_results.append({
                            "label": f['label'],
                            "field_id": f['field_id'],
                            "value": None,
                            "confidence": 0,
                            "source": "",
                        })
            else:
                # JSON解析失败，所有字段标记为空
                for f in batch:
                    all_results.append({
                        "label": f['label'],
                        "field_id": f['field_id'],
                        "value": None,
                        "confidence": 0,
                        "source": "",
                    })

        except Exception as e:
            import logging
            logging.getLogger(__name__).error(f"LLM预填提取失败: {e}")
            for f in batch:
                all_results.append({
                    "label": f['label'],
                    "field_id": f['field_id'],
                    "value": None,
                    "confidence": 0,
                    "source": "",
                })

    return all_results


def _call_llm(system_prompt: str, user_prompt: str, ctx=None) -> str:
    """调用LLM，支持外部API和平台内置LLM。"""
    ext_api_key = os.getenv("EXTERNAL_LLM_API_KEY")
    ext_base_url = os.getenv("EXTERNAL_LLM_BASE_URL")

    if ext_api_key and ext_base_url:
        from langchain_openai import ChatOpenAI
        ext_model = os.getenv("EXTERNAL_LLM_MODEL", "deepseek-chat")
        ext_llm = ChatOpenAI(
            model=ext_model,
            api_key=ext_api_key,
            base_url=ext_base_url,
            temperature=0.1,
            max_tokens=4096,
        )
        from langchain_core.messages import SystemMessage as SM, HumanMessage as HM
        response = ext_llm.invoke([SM(content=system_prompt), HM(content=user_prompt)])
        return response.content
    else:
        from coze_coding_dev_sdk import LLMClient
        client = LLMClient(ctx=ctx)
        messages = [
            SystemMessage(content=system_prompt),
            HumanMessage(content=user_prompt),
        ]
        response = client.invoke(
            messages=messages,
            model="doubao-seed-1-6-lite-251015",
            temperature=0.1,
            max_completion_tokens=4096,
        )
        content = response.content
        if isinstance(content, list):
            return " ".join(
                item.get("text", "") for item in content
                if isinstance(item, dict) and item.get("type") == "text"
            )
        elif isinstance(content, str):
            return content
        else:
            return str(content)


def _parse_json_array(text: str) -> list:
    """从LLM返回文本中解析JSON数组。"""
    text = text.strip()

    # 尝试直接解析
    try:
        result = json.loads(text)
        if isinstance(result, list):
            return result
    except json.JSONDecodeError:
        pass

    # 尝试提取JSON代码块
    json_match = re.search(r'```(?:json)?\s*([\s\S]*?)```', text)
    if json_match:
        try:
            result = json.loads(json_match.group(1).strip())
            if isinstance(result, list):
                return result
        except json.JSONDecodeError:
            pass

    # 尝试提取方括号内容
    bracket_match = re.search(r'\[[\s\S]*\]', text)
    if bracket_match:
        try:
            result = json.loads(bracket_match.group(0))
            if isinstance(result, list):
                return result
        except json.JSONDecodeError:
            pass

    return []


def _find_matching_result(results: list, label: str, field_id: str) -> dict:
    """在LLM返回的结果列表中找到匹配label的结果。"""
    # 精确匹配label
    for r in results:
        if isinstance(r, dict) and r.get("label") == label:
            return r

    # 模糊匹配（去除空格、标点后比较）
    import unicodedata
    def _normalize(s):
        s = unicodedata.normalize('NFKC', str(s))
        s = re.sub(r'[\s\u3000：:，,。.、]', '', s)
        return s

    norm_label = _normalize(label)
    for r in results:
        if isinstance(r, dict):
            r_label = _normalize(r.get("label", ""))
            if r_label and (r_label == norm_label or norm_label in r_label or r_label in norm_label):
                return r

    return None


def _rule_prefill_fields(field_list: list, file_content: str) -> list:
    """使用规则匹配预填（兜底方案），给规则匹配的结果一个固定置信度。"""
    labels = [f['label'] for f in field_list]
    rule_extracted = _rule_extract_fields(labels, file_content)

    results = []
    for f in field_list:
        if f['label'] in rule_extracted:
            results.append({
                "label": f['label'],
                "field_id": f['field_id'],
                "value": rule_extracted[f['label']],
                "confidence": 0.75,  # 规则匹配给0.75置信度（中等偏高，但需审核）
                "source": "规则匹配",
            })
        else:
            results.append({
                "label": f['label'],
                "field_id": f['field_id'],
                "value": None,
                "confidence": 0,
                "source": "",
            })

    return results


def _merge_prefill_results(llm_results: list, rule_results: list) -> list:
    """合并LLM和规则匹配的预填结果，LLM优先。"""
    # 以LLM结果为基础
    merged = {r['field_id']: r for r in llm_results}

    # 规则匹配补充LLM未找到的字段
    for r in rule_results:
        fid = r['field_id']
        if fid not in merged or (merged[fid]['value'] is None and r['value'] is not None):
            merged[fid] = r
        elif merged[fid]['value'] is None and r['value'] is not None:
            # LLM没找到，但规则找到了
            merged[fid] = r

    return list(merged.values())


def _merge_multi_file_results(all_file_results: list) -> list:
    """合并多个文件的预填结果。

    策略：同一字段取置信度最高的值；冲突时标记为review。
    """
    if not all_file_results:
        return []

    # 收集所有field_id
    field_ids = set()
    for results in all_file_results:
        for r in results:
            field_ids.add(r['field_id'])

    merged = []
    for fid in field_ids:
        candidates = []
        for results in all_file_results:
            for r in results:
                if r['field_id'] == fid and r['value'] is not None:
                    candidates.append(r)

        if not candidates:
            # 所有文件都没找到
            label = ""
            for results in all_file_results:
                for r in results:
                    if r['field_id'] == fid:
                        label = r['label']
                        break
            merged.append({
                "label": label,
                "field_id": fid,
                "value": None,
                "confidence": 0,
                "source": "",
            })
        elif len(candidates) == 1:
            merged.append(candidates[0])
        else:
            # 多个文件都有值
            # 检查值是否一致
            values = set(r['value'] for r in candidates)
            if len(values) == 1:
                # 值一致，取置信度最高的
                best = max(candidates, key=lambda r: r['confidence'])
                sources = [r['source'] for r in candidates if r['source']]
                best['source'] = " + ".join(sources) if sources else best['source']
                merged.append(best)
            else:
                # 值冲突，取置信度最高的但降低置信度
                best = max(candidates, key=lambda r: r['confidence'])
                best['confidence'] = min(best['confidence'], 0.6)  # 冲突降级
                best['source'] = f"多文件有不同值({'+'.join(str(v)[:20] for v in values)})"
                merged.append(best)

    return merged


# ── 对外工具函数 ──

@tool
def prefill_from_knowledge(file_path: str, template_fields_json: str) -> str:
    """从知识文件中提取模板字段值，生成带置信度的预填结果。
    
    用户上传知识文件（教学大纲、成绩单等）后，调用此工具自动提取所有字段的值。
    返回的预填结果包含置信度，前端据此展示不同颜色标记：
    - confirmed(绿): 置信度≥0.8，可直接使用
    - review(黄): 置信度0.4-0.8，需用户确认
    - empty(灰): 置信度<0.4，需用户手动填写

    Args:
        file_path: 知识文件路径
        template_fields_json: 模板字段清单JSON（来自analyze_template或analyze_uploaded_template的label_fields）
    """
    ctx = request_context.get() or new_context(method="prefill_from_knowledge")

    try:
        # 1. 解析文件路径
        workspace = os.getenv("COZE_WORKSPACE_PATH", "/workspace/projects")
        if not os.path.isabs(file_path):
            full_path = os.path.join(workspace, file_path)
        else:
            full_path = file_path

        if not os.path.exists(full_path):
            return json.dumps({"success": False, "message": f"文件不存在: {file_path}"}, ensure_ascii=False)

        # 2. 提取文件文本内容
        file_content = _extract_text_from_file(full_path)
        if file_content.startswith('['):
            return json.dumps({"success": False, "message": f"文件解析失败: {file_content}"}, ensure_ascii=False)

        # 3. 解析模板字段
        if isinstance(template_fields_json, str):
            template_fields = json.loads(template_fields_json)
        else:
            template_fields = template_fields_json

        # 构建 field_list
        field_list = []
        for f in template_fields:
            label = f.get("raw_label") or f.get("label", "")
            field_id = f.get("field_id", "")
            if label:
                field_list.append({"label": label, "field_id": field_id})

        if not field_list:
            return json.dumps({"success": False, "message": "字段清单为空"}, ensure_ascii=False)

        # 4. LLM智能提取
        llm_results = _llm_prefill_fields(field_list, file_content, ctx=ctx)

        # 5. 规则匹配兜底
        rule_results = _rule_prefill_fields(field_list, file_content)

        # 6. 合并结果
        merged = _merge_prefill_results(llm_results, rule_results)

        # 7. 标记状态
        for r in merged:
            r['status'] = _classify_status(r['confidence'])

        # 8. 统计
        confirmed = sum(1 for r in merged if r['status'] == 'confirmed')
        review = sum(1 for r in merged if r['status'] == 'review')
        empty = sum(1 for r in merged if r['status'] == 'empty')
        total = len(merged)

        result = {
            "success": True,
            "file_name": os.path.basename(full_path),
            "template_fields": total,
            "prefilled": confirmed + review,
            "needs_review": review,
            "still_empty": empty,
            "fill_rate": f"{confirmed + review}/{total}",
            "fill_rate_pct": round((confirmed + review) / total * 100, 1) if total > 0 else 0,
            "fields": merged,
            "summary": f"已预填 {confirmed + review}/{total} 个字段（{confirmed}个高置信，{review}个需审核，{empty}个未填）"
        }

        return json.dumps(result, ensure_ascii=False)

    except Exception as e:
        import logging
        logging.getLogger(__name__).error(f"预填工具执行失败: {e}", exc_info=True)
        return json.dumps({"success": False, "message": f"预填失败: {e}"}, ensure_ascii=False)


@tool
def prefill_from_multiple_knowledge(file_paths_json: str, template_fields_json: str) -> str:
    """从多个知识文件中联合提取模板字段值，生成带置信度的预填结果。

    与 prefill_from_knowledge 的区别：支持多个知识文件，自动合并结果。
    多文件场景下，同一字段的值取置信度最高的；冲突时降级标记。

    Args:
        file_paths_json: 知识文件路径列表JSON，如 '["/tmp/file1.docx", "/tmp/file2.pdf"]'
        template_fields_json: 模板字段清单JSON
    """
    ctx = request_context.get() or new_context(method="prefill_from_multiple_knowledge")

    try:
        # 解析参数
        if isinstance(file_paths_json, str):
            file_paths = json.loads(file_paths_json)
        else:
            file_paths = file_paths_json

        if isinstance(template_fields_json, str):
            template_fields = json.loads(template_fields_json)
        else:
            template_fields = template_fields_json

        if not file_paths:
            return json.dumps({"success": False, "message": "未提供文件路径"}, ensure_ascii=False)

        # 构建 field_list
        field_list = []
        for f in template_fields:
            label = f.get("raw_label") or f.get("label", "")
            field_id = f.get("field_id", "")
            if label:
                field_list.append({"label": label, "field_id": field_id})

        if not field_list:
            return json.dumps({"success": False, "message": "字段清单为空"}, ensure_ascii=False)

        # 逐文件提取
        all_file_results = []
        file_names = []
        for fp in file_paths:
            workspace = os.getenv("COZE_WORKSPACE_PATH", "/workspace/projects")
            if not os.path.isabs(fp):
                full_path = os.path.join(workspace, fp)
            else:
                full_path = fp

            if not os.path.exists(full_path):
                continue

            file_content = _extract_text_from_file(full_path)
            if file_content.startswith('['):
                continue

            file_names.append(os.path.basename(full_path))

            # LLM + 规则 提取
            llm_results = _llm_prefill_fields(field_list, file_content, ctx=ctx)
            rule_results = _rule_prefill_fields(field_list, file_content)
            merged = _merge_prefill_results(llm_results, rule_results)
            all_file_results.append(merged)

        if not all_file_results:
            return json.dumps({"success": False, "message": "所有文件解析失败"}, ensure_ascii=False)

        # 多文件合并
        final_results = _merge_multi_file_results(all_file_results)

        # 标记状态
        for r in final_results:
            r['status'] = _classify_status(r['confidence'])

        # 统计
        confirmed = sum(1 for r in final_results if r['status'] == 'confirmed')
        review = sum(1 for r in final_results if r['status'] == 'review')
        empty = sum(1 for r in final_results if r['status'] == 'empty')
        total = len(final_results)

        result = {
            "success": True,
            "file_names": file_names,
            "file_count": len(file_names),
            "template_fields": total,
            "prefilled": confirmed + review,
            "needs_review": review,
            "still_empty": empty,
            "fill_rate": f"{confirmed + review}/{total}",
            "fill_rate_pct": round((confirmed + review) / total * 100, 1) if total > 0 else 0,
            "fields": final_results,
            "summary": f"从{len(file_names)}个文件中预填 {confirmed + review}/{total} 个字段（{confirmed}个高置信，{review}个需审核，{empty}个未填）"
        }

        return json.dumps(result, ensure_ascii=False)

    except Exception as e:
        import logging
        logging.getLogger(__name__).error(f"多文件预填工具执行失败: {e}", exc_info=True)
        return json.dumps({"success": False, "message": f"多文件预填失败: {e}"}, ensure_ascii=False)


# ── 非工具函数（供API直接调用） ──

def prefill_from_file_paths(file_paths: list, template_path: str) -> dict:
    """直接API调用：从文件路径列表和模板路径生成预填结果。

    不经过Agent，直接供 /prefill API 调用。

    Args:
        file_paths: 知识文件路径列表
        template_path: 模板文件路径

    Returns:
        dict: 预填结果
    """
    try:
        # 1. 分析模板字段
        analysis = analyze_template(template_path)
        template_fields = analysis.get("label_fields", [])

        if not template_fields:
            return {"success": False, "message": "模板字段识别为空"}

        # 2. 构建field_list
        field_list = []
        for f in template_fields:
            label = f.get("raw_label") or f.get("label", "")
            field_id = f.get("field_id", "")
            if label:
                field_list.append({"label": label, "field_id": field_id})

        # 3. 逐文件提取
        all_file_results = []
        file_names = []
        workspace = os.getenv("COZE_WORKSPACE_PATH", "/workspace/projects")

        for fp in file_paths:
            if not os.path.isabs(fp):
                full_path = os.path.join(workspace, fp)
            else:
                full_path = fp

            if not os.path.exists(full_path):
                continue

            file_content = _extract_text_from_file(full_path)
            if file_content.startswith('['):
                continue

            file_names.append(os.path.basename(full_path))

            # LLM + 规则 提取
            ctx = new_context(method="prefill_api")
            llm_results = _llm_prefill_fields(field_list, file_content, ctx=ctx)
            rule_results = _rule_prefill_fields(field_list, file_content)
            merged = _merge_prefill_results(llm_results, rule_results)
            all_file_results.append(merged)

        if not all_file_results:
            return {"success": False, "message": "所有知识文件解析失败"}

        # 4. 多文件合并
        final_results = _merge_multi_file_results(all_file_results)

        # 5. 标记状态
        for r in final_results:
            r['status'] = _classify_status(r['confidence'])

        # 6. 统计
        confirmed = sum(1 for r in final_results if r['status'] == 'confirmed')
        review = sum(1 for r in final_results if r['status'] == 'review')
        empty = sum(1 for r in final_results if r['status'] == 'empty')
        total = len(final_results)

        return {
            "success": True,
            "template_name": os.path.basename(template_path),
            "file_names": file_names,
            "file_count": len(file_names),
            "template_fields": total,
            "prefilled": confirmed + review,
            "needs_review": review,
            "still_empty": empty,
            "fill_rate": f"{confirmed + review}/{total}",
            "fill_rate_pct": round((confirmed + review) / total * 100, 1) if total > 0 else 0,
            "fields": final_results,
            "summary": f"从{len(file_names)}个文件中预填 {confirmed + review}/{total} 个字段"
        }

    except Exception as e:
        import logging
        logging.getLogger(__name__).error(f"API预填执行失败: {e}", exc_info=True)
        return {"success": False, "message": f"预填失败: {e}"}
