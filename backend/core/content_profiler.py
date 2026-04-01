"""
content_profiler.py
分析单页内容，生成 content_profile 字典。
"""
import re

# 单位正则（用于核心数字识别）
_UNIT_PAT = re.compile(
    r'(\d+\.?\d*)\s*'
    r'(秒|s|ms|分钟|小时|h|km|m(?!s)|mm|cm|kg|kW|kWh|万|亿|元|%|倍|人|台|次|°C|fps|TOPS)'
)


def _collect_texts(page: dict) -> list[str]:
    """收集页面所有正文文字（不含标题）。"""
    texts = []
    if page["mode"] == "explicit":
        for block in page.get("blocks", []):
            if block["type"] == "standalone_text":
                if block["text"]:
                    texts.append(block["text"])
            else:
                for pair in block.get("pairs", []):
                    if pair["text"]:
                        texts.append(pair["text"])
    else:
        texts.extend(p for p in page.get("paragraphs", []) if p)
    return texts


def _text_weight(texts: list[str]) -> str:
    total = sum(len(t) for t in texts)
    list_items = sum(1 for t in texts if t.startswith("• "))
    if total > 100 or list_items > 3:
        return "heavy"
    if total >= 30:
        return "balanced"
    return "light"


def _find_key_number(title: str, subtitle: str, texts: list[str]) -> tuple:
    """在标题、副标题、正文中寻找最显著的数字+单位。"""
    # 优先在标题/副标题里找
    for src in [title, subtitle] + texts:
        m = _UNIT_PAT.search(src)
        if m:
            return m.group(1), m.group(2)
    return None, None


def generate_profile(page: dict, is_first: bool = False, is_last: bool = False) -> dict:
    """
    分析页面内容，返回 content_profile 字典。
    is_first / is_last 由调用方（run_phase2.py）传入。
    """
    images = page.get("images", {})
    has_background = any(v.get("is_background") for v in images.values())

    # 图文对数量 & 无配对图片数
    pair_count = 0
    unpaired_image_count = 0
    if page["mode"] == "explicit":
        for block in page.get("blocks", []):
            if block["type"] == "paired":
                for pair in block.get("pairs", []):
                    if pair.get("image_ids"):
                        pair_count += 1
    else:
        # 自由模式：无配对概念，非背景图均视为无配对
        unpaired_image_count = sum(
            1 for v in images.values() if not v.get("is_background")
        )

    texts = _collect_texts(page)
    title = page.get("title", "")
    subtitle = page.get("subtitle", "")

    key_number, key_unit = _find_key_number(title, subtitle, texts)

    return {
        "has_background": has_background,
        "pair_count": pair_count,
        "unpaired_image_count": unpaired_image_count,
        "text_weight": _text_weight(texts),
        "has_key_number": key_number is not None,
        "key_number": key_number,
        "key_number_unit": key_unit,
        "is_first": is_first,
        "is_last": is_last,
    }
