"""
prompt_builder.py
构建 Phase 3 HTML 生成的 System Prompt 和 User Prompt。
"""


def build_system_prompt(prompts_cfg: dict) -> str:
    """从 prompts.json 读取 html_generator.system，拼成字符串。"""
    lines = prompts_cfg["html_generator"]["system"]
    return "\n".join(lines) if isinstance(lines, list) else lines


def build_user_prompt(page: dict, layout_intent: str, style: dict, total_pages: int) -> str:
    """为单页构建 User Prompt。"""
    colors = style["colors"]
    typo = style["typography"]
    pn = page["page_number"]

    lines = []

    # ── 风格 ──────────────────────────────────────────────────────────────────
    lines.append("## 风格")
    lines.append(f"名称：{style['name']}")
    lines.append(
        f"背景：{colors['bg']} | "
        f"标题：{colors['text']} | "
        f"正文：{colors.get('text_body', colors['text'])} | "
        f"强调色：{colors['accent']}"
    )
    lines.append(
        f"字体：{typo['display_font']}（大标题）"
        f"/ {typo['heading2_font']}（副标题）"
        f"/ {typo['body_font']}（正文）"
        f"/ {typo['footnote_font']}（注释）"
    )

    # ── 布局意图 ──────────────────────────────────────────────────────────────
    lines.append("")
    lines.append("## 布局意图")
    lines.append(layout_intent)

    # ── 页面内容 ──────────────────────────────────────────────────────────────
    lines.append("")
    lines.append(f"## 本页内容（第 {pn} 页 / 共 {total_pages} 页）")
    lines.append(f"section class: page-{pn}")
    lines.append(f"标题：{page['title']}")
    if page.get("subtitle"):
        lines.append(f"副标题：{page['subtitle']}")

    images = page.get("images", {})
    bg_images = {k: v for k, v in images.items() if v.get("is_background")}
    content_images = {k: v for k, v in images.items() if not v.get("is_background")}

    # 背景图单独说明
    for img_id in bg_images:
        lines.append(
            f"背景图：{{{{{img_id}}}}}  "
            f"（铺满整个 section，用 <img> position:absolute top:0 left:0 width:1280px height:720px object-fit:cover）"
        )

    lines.append("")

    # 正文内容
    if page["mode"] == "explicit":
        lines.append(f"内容块（共 {len(page['blocks'])} 块）：")
        for i, block in enumerate(page["blocks"], 1):
            if block["type"] == "standalone_text":
                lines.append(f"  块{i}【纯文字】：{block['text']}")
            else:
                tag = "【多组图文并列】" if block.get("multi_image") else "【图文配对】"
                lines.append(f"  块{i}{tag}：")
                for j, pair in enumerate(block["pairs"], 1):
                    img_ids = pair.get("image_ids", [])
                    text_str = pair["text"] or "（无文字）"
                    lines.append(f"    对{j} 文字：{text_str}")
                    if img_ids:
                        img_str = "  ".join(f"{{{{{iid}}}}}" for iid in img_ids)
                        lines.append(f"         图片：{img_str}")
    else:
        paragraphs = page.get("paragraphs", [])
        if paragraphs:
            lines.append("正文段落：")
            for p in paragraphs:
                lines.append(f"  {p}")
        if content_images:
            img_str = "  ".join(f"{{{{{k}}}}}" for k in content_images)
            lines.append(f"配图：{img_str}")

    return "\n".join(lines)
