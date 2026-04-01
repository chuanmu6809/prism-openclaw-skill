"""
layout_planner.py
调用 AI（排版导演），为整个演示文稿规划每页布局意图。
"""
import json
import re
from pathlib import Path

from openai import OpenAI

from backend.core.paths import CONFIG_DIR
from backend.core.llm_client import get_active_config

# ─── 配置加载 ──────────────────────────────────────────────────────────────────

_ROOT = Path(__file__).parent.parent
_prompt_cfg = json.loads((CONFIG_DIR / "prompts.json").read_text(encoding="utf-8"))

_client = None
_current_key = None


def _get_sync_client():
    """获取同步 OpenAI 客户端（配置变更时自动重建）"""
    global _client, _current_key
    cfg = get_active_config()
    key = (cfg["base_url"], cfg["api_key"])
    if _client is None or key != _current_key:
        _client = OpenAI(base_url=cfg["base_url"], api_key=cfg["api_key"])
        _current_key = key
    return _client, cfg


# ─── 摘要生成 ──────────────────────────────────────────────────────────────────

def _build_page_summary(page: dict) -> str:
    """将单页信息格式化为供 AI 阅读的自然语言摘要。"""
    cp = page.get("content_profile", {})
    parts = [f'第{page["page_number"]}页：标题「{page["title"]}」']

    if page.get("subtitle"):
        parts.append(f'副标题「{page["subtitle"]}」')

    if cp.get("is_first"):
        parts.append("（封面页）")
    if cp.get("is_last"):
        parts.append("（末页）")

    if cp.get("has_background"):
        parts.append("全屏背景图")

    pair_count = cp.get("pair_count", 0)
    if pair_count > 0:
        parts.append(f"{pair_count}组图文配对")

    unpaired = cp.get("unpaired_image_count", 0)
    if unpaired > 0:
        parts.append(f"{unpaired}张独立图片")

    if cp.get("has_key_number"):
        num = cp.get("key_number", "")
        unit = cp.get("key_number_unit", "")
        parts.append(f"核心数字 {num}{unit}")

    parts.append(f'文字量{cp.get("text_weight", "balanced")}')

    return "、".join(parts)


# ─── 容错兜底 ──────────────────────────────────────────────────────────────────

def _fallback_intent(page: dict) -> str:
    """AI 失败时按规则生成默认布局意图。"""
    cp = page.get("content_profile", {})

    if cp.get("is_first") or cp.get("is_last"):
        if cp.get("has_background"):
            return "全屏背景图铺满，主标题超大居中，副标题细字居中，暗色叠加层增强文字可读性"
        return "主标题超大居中，大量留白，简洁封面风格"

    if cp.get("has_key_number"):
        num = cp.get("key_number", "")
        unit = cp.get("key_number_unit", "")
        return f"核心数字「{num}{unit}」极大居中占主体，标题小字置顶，大量留白强调数字冲击力"

    pair_count = cp.get("pair_count", 0)
    if pair_count >= 3:
        return "多组图文以卡片网格排布，每卡片图上文下，标题置顶"
    if pair_count == 2:
        return "两组图文左右并列分栏，各占50%宽，标题居顶"
    if pair_count == 1:
        return "左侧图片占55%宽，右侧文字上下排布，标题置顶"

    if cp.get("has_background"):
        return "全屏背景图，文字叠加居中，半透明遮罩层"

    if cp.get("text_weight") == "heavy":
        return "双栏紧凑排版，标题置顶横贯全宽，正文分两栏"

    return "文字居中排布，标题大字，正文适中，适量留白"


# ─── AI 调用 ───────────────────────────────────────────────────────────────────

def build_layout_plan_prompts(pages: list, style: dict) -> tuple[str, str]:
    """
    构造布局规划所需的 system/user prompt。
    供 CLI 的 host 模式和外部编排复用。
    """
    planner_cfg = _prompt_cfg["layout_planner"]
    style_name = style.get("name", "发布会暗色科技感")
    is_mixed = style.get("mixed", False)

    summaries = "\n".join(_build_page_summary(p) for p in pages)

    sys_lines = planner_cfg["system"]
    if isinstance(sys_lines, list):
        sys_content = "\n".join(sys_lines)
    else:
        sys_content = sys_lines

    if is_mixed and "system_mixed_tone" in planner_cfg:
        mixed_lines = planner_cfg["system_mixed_tone"]
        if isinstance(mixed_lines, list):
            sys_content += "\n" + "\n".join(mixed_lines)
        else:
            sys_content += "\n" + mixed_lines

    tone_instruction = ""
    if is_mixed:
        tone_instruction = "\n此风格为【明暗混合】模式，请为每页输出 contrast_affinity 字段（0、1 或 2）。\n"

    user_prompt = planner_cfg["user_template"].format(
        style_name=style_name,
        tone_instruction=tone_instruction,
        page_summaries=summaries,
    )
    return sys_content, user_prompt


def parse_layout_plan_response(raw: str, pages: list, style: dict) -> list:
    """
    解析布局规划模型输出，统一补齐缺失页与 mixed 模式字段。
    """
    is_mixed = style.get("mixed", False)
    m = re.search(r'\[.*\]', raw, re.DOTALL)
    if not m:
        raise ValueError(f"AI 返回内容中未找到 JSON 数组:\n{raw}")
    intents = json.loads(m.group())

    page_nums = {p["page_number"] for p in pages}
    result = []
    for item in intents:
        if "page" not in item or "layout_intent" not in item:
            raise ValueError(f"AI 返回格式错误: {item}")
        if item["page"] not in page_nums:
            raise ValueError(f"AI 返回了未知页码: {item['page']}")
        entry = {"page": item["page"], "layout_intent": item["layout_intent"]}
        if is_mixed:
            ca = item.get("contrast_affinity", 0)
            entry["contrast_affinity"] = max(0, min(2, int(ca)))
        result.append(entry)

    returned = {r["page"] for r in result}
    for p in pages:
        if p["page_number"] not in returned:
            entry = {"page": p["page_number"], "layout_intent": _fallback_intent(p)}
            if is_mixed:
                entry["contrast_affinity"] = 0
            result.append(entry)

    result.sort(key=lambda x: x["page"])
    return result

def plan_layout(pages: list, style: dict) -> list:
    """
    调用 AI 为所有页面规划布局意图。
    返回：[{"page": N, "layout_intent": "...", "contrast_affinity": 0-2}, ...]
    当 style 含 mixed=True 时，AI 会额外输出每页的反差适宜度评分。
    """
    is_mixed = style.get("mixed", False)
    sys_content, user_prompt = build_layout_plan_prompts(pages, style)

    try:
        client, cfg = _get_sync_client()
        resp = client.chat.completions.create(
            model=cfg["model"],
            messages=[
                {"role": "system", "content": sys_content},
                {"role": "user", "content": user_prompt},
            ],
            max_tokens=cfg.get("max_tokens", 4096),
            timeout=cfg.get("timeout", 60),
        )
        raw = resp.choices[0].message.content.strip()

        result = parse_layout_plan_response(raw, pages, style)

        if is_mixed:
            ca_summary = ", ".join(f"P{r['page']}:CA{r.get('contrast_affinity', 0)}" for r in result)
            print(f"  🎯 反差适宜度: {ca_summary}")

        return result

    except Exception as e:
        print(f"[layout_planner] AI 调用失败，使用兜底规则: {e}")
        return [
            {"page": p["page_number"], "layout_intent": _fallback_intent(p),
             **({"contrast_affinity": 0} if is_mixed else {})}
            for p in pages
        ]
