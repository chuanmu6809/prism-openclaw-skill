"""
html_builder.py
将 AI 生成的 <section> 片段拼装成完整 HTML 文件。
- MiSans 字体以 base64 嵌入，HTML 完全自包含
- 图片占位符 {{IMG_XX_XX}} 替换为 base64 data URL
"""
import base64
import re
from pathlib import Path

from backend.core.paths import PROJECT_ROOT, ASSETS_DIR

_ROOT = PROJECT_ROOT

# 兼容 AI 偶发格式偏差：{IMG_xx}、{{IMG_xx}}、{{ IMG_xx }} 均可匹配
_IMG_PLACEHOLDER = re.compile(r'\{+\s*(IMG_[A-Z0-9_]+)\s*\}+')

# markdown 代码块标记（AI 有时会包裹输出）
_CODE_BLOCK = re.compile(r'^```[\w]*\n?|```$', re.MULTILINE)

# 字体名称 → CSS font-weight 映射
_FONT_WEIGHTS = {
    "MiSans-Thin":       100,
    "MiSans-ExtraLight": 200,
    "MiSans-Light":      300,
    "MiSans-Normal":     350,
    "MiSans-Regular":    400,
    "MiSans-Medium":     500,
    "MiSans-Semibold":   600,
    "MiSans-Demibold":   650,
    "MiSans-Bold":       700,
    "MiSans-Heavy":      900,
}


# ─── 字体加载 ──────────────────────────────────────────────────────────────────

def _load_fonts_b64(style: dict) -> dict:
    """读取 style.font_paths 中的字体文件并 base64 编码。"""
    result = {}
    for font_name, rel_path in style.get("font_paths", {}).items():
        # rel_path 形如 "assets/fonts/MiSans-Bold.ttf"
        # 打包后 assets 在 ASSETS_DIR，需要去掉 "assets/" 前缀
        if rel_path.startswith("assets/"):
            abs_path = ASSETS_DIR / rel_path[len("assets/"):]
        else:
            abs_path = _ROOT / rel_path
        if abs_path.exists():
            result[font_name] = base64.b64encode(abs_path.read_bytes()).decode()
        else:
            # 回退：尝试原始路径
            fallback = _ROOT / rel_path
            if fallback.exists():
                result[font_name] = base64.b64encode(fallback.read_bytes()).decode()
            else:
                print(f"[html_builder] 字体文件不存在，跳过：{abs_path}")
    return result


def _build_font_face_css(fonts_b64: dict) -> str:
    """生成 @font-face CSS 声明块。"""
    lines = []
    for font_name, b64 in fonts_b64.items():
        weight = _FONT_WEIGHTS.get(font_name, 400)
        lines.append(
            f"@font-face {{\n"
            f"  font-family: '{font_name}';\n"
            f"  font-weight: {weight};\n"
            f"  src: url('data:font/truetype;base64,{b64}') format('truetype');\n"
            f"}}"
        )
    return "\n".join(lines)


# ─── Section 清洗 ──────────────────────────────────────────────────────────────

def _clean_section(raw: str) -> str:
    """去除 AI 输出中可能夹带的 markdown 代码块标记。"""
    cleaned = _CODE_BLOCK.sub("", raw).strip()
    # 确保以 <section 开头
    idx = cleaned.lower().find("<section")
    if idx > 0:
        cleaned = cleaned[idx:]
    return cleaned


# 匹配 .page-N 选择器块内的 position/top/left 声明（仅针对顶层页面容器）
_PAGE_ABS_PROPS = re.compile(
    r'(\.page-\d+\s*\{[^}]*?)'           # .page-N { ... 的前半部分
    r'(?:position\s*:\s*absolute\s*;?\s*'  # position: absolute
    r'|top\s*:\s*0(?:px)?\s*;?\s*'         # top: 0 / top: 0px
    r'|left\s*:\s*0(?:px)?\s*;?\s*)',      # left: 0 / left: 0px
    re.DOTALL
)


def _fix_page_positioning(html: str) -> str:
    """
    去除 .page-N 选择器内的 position:absolute / top:0 / left:0，
    防止 section 脱离文档流导致堆叠。
    只影响 .page-N 顶层选择器，不影响 .page-N .child 等子选择器。
    """
    # 找每个 <style>...</style> 块，在其中处理
    def fix_style_block(m):
        style_content = m.group(0)
        # 只修改「.page-N {」这种直接选择器（后面没有空格+子选择器）
        def remove_abs(block_match):
            block = block_match.group(0)
            # 去掉 position: absolute（含分号和空白）
            block = re.sub(r'position\s*:\s*absolute\s*;?\s*', '', block)
            # 去掉 top: 0 / top: 0px（只去 0 值，不去其他值）
            block = re.sub(r'(?<![.\w-])top\s*:\s*0(?:px)?\s*;?\s*', '', block)
            # 去掉 left: 0 / left: 0px
            block = re.sub(r'(?<![.\w-])left\s*:\s*0(?:px)?\s*;?\s*', '', block)
            return block
        # 只匹配 .page-N { ... }（中间无空格子选择器）
        style_content = re.sub(
            r'\.page-\d+\s*\{[^}]*\}',
            remove_abs,
            style_content
        )
        return style_content

    return re.sub(r'<style>.*?</style>', fix_style_block, html, flags=re.DOTALL)


# ─── 图片替换 ──────────────────────────────────────────────────────────────────

def _replace_placeholders(html: str, images_b64: dict) -> str:
    """将 {{IMG_XX_XX}} 替换为 base64 data URL。"""
    def replacer(m):
        img_id = m.group(1)
        b64 = images_b64.get(img_id)
        if b64:
            return f"data:image/jpeg;base64,{b64}"
        print(f"[html_builder] 图片占位符未匹配到文件：{img_id}")
        return ""
    return _IMG_PLACEHOLDER.sub(replacer, html)


# ─── 主入口 ────────────────────────────────────────────────────────────────────

def assemble_html(sections: list, images_b64: dict, style: dict) -> str:
    """
    拼装完整 HTML 文件。
    sections: AI 生成的 <section> 片段列表
    images_b64: {img_id: base64_string}
    style: 来自 config/styles/*.json
    """
    fonts_b64 = _load_fonts_b64(style)
    font_css = _build_font_face_css(fonts_b64)
    bg = style["colors"]["bg"]

    processed = []
    for section in sections:
        clean = _clean_section(section)
        final = _replace_placeholders(clean, images_b64)
        processed.append(final)

    sections_html = "\n\n".join(processed)

    full_html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <style>
{font_css}

    * {{ margin: 0; padding: 0; box-sizing: border-box; }}
    body {{ width: 1280px; background: {bg}; }}
    section {{
      width: 1280px;
      height: 720px;
      position: relative;
      overflow: hidden;
      display: block;
    }}
  </style>
</head>
<body>
{sections_html}
</body>
</html>"""

    # 修复 AI 在 .page-N 上写的 position:absolute 导致堆叠问题
    return _fix_page_positioning(full_html)
