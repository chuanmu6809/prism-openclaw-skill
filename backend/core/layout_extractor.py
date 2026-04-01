from __future__ import annotations
"""
layout_extractor.py
用 Playwright 渲染 HTML，提取每页元素的精确布局/样式数据。
输出供 Phase 5 (pptx_builder) 使用。
"""
import re
import json
from pathlib import Path
from playwright.sync_api import sync_playwright

# ─── 单位换算 ──────────────────────────────────────────────────────────────────

EMU_PER_PX = 9525  # 914400 / 96

def px_to_emu(px: float) -> int:
    return round(px * EMU_PER_PX)

def px_to_pt(px: float) -> float:
    return round(px * 72 / 96, 1)

# ─── 颜色工具 ──────────────────────────────────────────────────────────────────

_RGBA_RE = re.compile(r'rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*(?:,\s*([\d.]+))?\s*\)')

def _is_transparent(color_str: str) -> bool:
    """判断颜色是否完全透明。"""
    if not color_str or color_str == 'transparent':
        return True
    m = _RGBA_RE.match(color_str)
    if m and m.group(4) is not None:
        return float(m.group(4)) == 0
    return False

def _color_to_hex(color_str: str) -> str:
    """将 rgb()/rgba() 转为 #RRGGBB 格式，保留 alpha 信息。"""
    if not color_str or color_str == 'transparent':
        return 'transparent'
    m = _RGBA_RE.match(color_str)
    if m:
        r, g, b = int(m.group(1)), int(m.group(2)), int(m.group(3))
        alpha = float(m.group(4)) if m.group(4) is not None else 1.0
        hex_color = f'#{r:02X}{g:02X}{b:02X}'
        if alpha < 1.0:
            return f'{hex_color}:{alpha:.2f}'  # #RRGGBB:0.70
        return hex_color
    return color_str

# ─── 渐变解析 ──────────────────────────────────────────────────────────────────

_GRADIENT_ANGLE_RE = re.compile(r'linear-gradient\(\s*(\d+)deg')
_GRADIENT_COLOR_RE = re.compile(r'(rgba?\([^)]+\))')


def _parse_gradient(bg_image: str) -> dict | None:
    """解析 linear-gradient，返回 {angle, color1, color2} 或 None。

    支持任意数量的色标（取首尾两色用于 PPT 两色渐变），
    支持带百分比的色标（如 rgb(0,0,0) 50%）。
    """
    if not bg_image or 'linear-gradient' not in bg_image:
        return None

    # 提取角度
    angle_m = _GRADIENT_ANGLE_RE.search(bg_image)
    angle = int(angle_m.group(1)) if angle_m else 180

    # 提取所有颜色值
    colors = _GRADIENT_COLOR_RE.findall(bg_image)
    if len(colors) < 2:
        return None

    color1_hex = _color_to_hex(colors[0])
    color2_hex = _color_to_hex(colors[-1])

    # 如果首末色完全相同，这不是一个有效渐变
    if color1_hex == color2_hex:
        return None

    # 如果末色 alpha 很低（≤0.15），渐变效果极微弱，视为纯色
    alpha_m = re.search(r':(\d+\.?\d*)', color2_hex)
    if alpha_m and float(alpha_m.group(1)) <= 0.15:
        return None

    return {
        'angle': angle,
        'color1': color1_hex,
        'color2': color2_hex,
    }

# ─── 图片 ID 提取 ──────────────────────────────────────────────────────────────

_IMG_ID_RE = re.compile(r'(IMG_\d{2}_(?:BG|\d{2}))')

def _extract_img_id(src: str) -> str | None:
    """从 base64 data URL 的上下文中提取图片 ID（HTML 生成时已替换占位符）。"""
    # 无法从 base64 本身提取 ID，需要从元素属性中找
    return None  # 会在 JS 端通过 alt/data 属性或匹配逻辑处理


# ─── Playwright 提取 JS ───────────────────────────────────────────────────────

EXTRACT_JS = """
(sectionEl) => {
    const sectionRect = sectionEl.getBoundingClientRect();
    const results = [];
    const extractedTexts = new Set();  // 文字去重
    let elIndex = 0;  // 为 shape 元素编号，便于后续 DOM 定位
    
    // 直接提取 section 自身的背景色（section 就是 .page-N）
    const sectionStyles = getComputedStyle(sectionEl);
    results.push({
        _role: 'section_bg',
        bg_color: sectionStyles.backgroundColor,
        bg_image: sectionStyles.backgroundImage,
    });
    
    // 获取文字的实际渲染宽度
    function getTextWidth(el) {
        if (!el.innerText || !el.innerText.trim()) return 0;
        const range = document.createRange();
        range.selectNodeContents(el);
        const rects = range.getClientRects();
        if (rects.length === 0) return 0;
        let minLeft = Infinity, maxRight = -Infinity;
        for (const r of rects) {
            minLeft = Math.min(minLeft, r.left);
            maxRight = Math.max(maxRight, r.right);
        }
        return maxRight - minLeft;
    }
    
    // 检查元素是否是叶子文字节点
    // 叶子 = 文字不来自块级子元素（p/div/li/h1-h6），而是直接文本或行内元素
    // 同时，行内子元素（span/strong/b/em 等）若文字是父元素文字的子集则跳过，避免混合字体重复提取
    function isLeafText(el) {
        const text = (el.innerText || '').trim();
        if (!text) return false;

        // 行内元素且文字被父元素包含 → 不是叶子（由父元素统一提取）
        const inlineTags = new Set(['span','strong','b','em','i','a','mark','small',
                                     'sub','sup','code','u','s','del','ins','label','time']);
        const selfTag = el.tagName.toLowerCase();
        if (inlineTags.has(selfTag) && el.parentElement) {
            const parentText = (el.parentElement.innerText || '').trim();
            if (parentText && parentText.includes(text)) return false;
        }

        const blockTags = new Set(['div','p','li','ul','ol','h1','h2','h3','h4','h5','h6',
                                    'section','article','blockquote','table','tr','td','th',
                                    'header','footer','main','nav','aside','figure']);
        for (const child of el.children) {
            const tag = child.tagName.toLowerCase();
            if (tag === 'style') continue;
            const childText = (child.innerText || '').trim();
            // 如果子元素完全等于父元素文字，文字来自这个子元素
            if (childText === text) return false;
            // 如果子元素是块级元素且有文字，父元素不是叶子
            if (blockTags.has(tag) && childText.length > 0) return false;
        }
        return true;
    }
    
    // 解析 border —— 分别读取四边，四边不一致时取最细的
    function parseBorder(styles) {
        const sides = ['Top','Right','Bottom','Left'];
        const parsed = [];
        const activeSides = [];
        
        for (const s of sides) {
            const w = parseFloat(styles['border'+s+'Width']) || 0;
            const c = styles['border'+s+'Color'];
            const st = styles['border'+s+'Style'];
            if (w >= 0.5 && st !== 'none' && c && c !== 'rgba(0, 0, 0, 0)' && c !== 'transparent') {
                parsed.push({ width: w, color: c, style: st, side: s });
                activeSides.push(s);
            }
        }
        if (parsed.length === 0) return null;
        
        // 取最细的边框（通常是主体轮廓，视觉效果最接近 HTML）
        parsed.sort((a, b) => a.width - b.width);
        const result = parsed[0];
        result.sideCount = parsed.length;
        result.activeSides = activeSides;
        return result;
    }
    
    function walk(el) {
        const styles = getComputedStyle(el);
        
        // 跳过不可见元素
        if (styles.display === 'none' || styles.visibility === 'hidden') return;
        if (parseFloat(styles.opacity) === 0) return;
        
        const rect = el.getBoundingClientRect();
        const x = rect.left - sectionRect.left;
        const y = rect.top - sectionRect.top;
        const w = rect.width;
        const h = rect.height;
        
        // 跳过太小的元素（但仍遍历子元素，因为容器 div 可能因
        // 所有子元素均为 position:absolute 而自身高度为 0）
        if (w < 2 || h < 2) {
            for (const child of el.children) {
                if (child.tagName.toLowerCase() !== 'style') walk(child);
            }
            return;
        }
        
        const tag = el.tagName.toLowerCase();
        const isImg = tag === 'img';
        const text = isImg ? '' : (el.innerText || '').trim();
        const bgColor = styles.backgroundColor;
        const bgImage = styles.backgroundImage;
        const hasGradient = bgImage && bgImage.includes('linear-gradient');
        const isTransparent = bgColor === 'rgba(0, 0, 0, 0)' || bgColor === 'transparent';
        const border = parseBorder(styles);
        const hasBorder = border !== null;
        
        // 判断是否是叶子文字节点（文字不来自子元素）
        const isLeaf = isLeafText(el);
        
        // 跳过完全透明、无内容、无图片、无渐变、无边框的元素
        if (isTransparent && !text && !isImg && !hasGradient && !hasBorder) {
            for (const child of el.children) walk(child);
            return;
        }
        
        const entry = {
            tag,
            x: Math.round(x * 100) / 100,
            y: Math.round(y * 100) / 100,
            w: Math.round(w * 100) / 100,
            h: Math.round(h * 100) / 100,
            opacity: parseFloat(styles.opacity),
        };
        
        if (isImg) {
            const idx = elIndex++;
            el.setAttribute('data-shape-idx', idx);
            entry._role = 'image';
            entry._shape_idx = idx;
            entry.src = el.src.substring(0, 80);
            entry.object_fit = styles.objectFit || 'fill';
            // 图片圆角：优先用自身的，如果自身为 0 则检查父元素
            // （HTML 常见做法：父 div 设 border-radius + overflow:hidden 来裁切图片）
            let imgBorderRadius = styles.borderRadius;
            if ((!imgBorderRadius || imgBorderRadius === '0px') && el.parentElement) {
                const parentStyles = getComputedStyle(el.parentElement);
                const parentOverflow = parentStyles.overflow;
                if ((parentOverflow === 'hidden' || parentOverflow === 'clip') && 
                    parentStyles.borderRadius && parentStyles.borderRadius !== '0px') {
                    imgBorderRadius = parentStyles.borderRadius;
                }
            }
            entry.border_radius = imgBorderRadius;
            const altMatch = (el.alt || '').match(/IMG_\\d{2}_(?:BG|\\d{2})/);
            if (altMatch) entry.img_id = altMatch[0];
            results.push(entry);
            return;
        }
        
        let addedShape = false;
        
        // 形状（有背景/渐变/边框）
        if (!isTransparent || hasGradient || hasBorder) {
            const idx = elIndex++;
            el.setAttribute('data-shape-idx', idx);
            entry._role = 'shape';
            entry._shape_idx = idx;
            entry.bg_color = bgColor;
            if (hasGradient) entry.bg_gradient = bgImage;
            entry.border_radius = styles.borderRadius;
            entry.box_shadow = styles.boxShadow !== 'none' ? styles.boxShadow : null;
            if (hasBorder) entry.border = border;
            
            // 检测复杂装饰元素（应截图而非逆向重建）
            const isComplexDeco = (
                // 渐变包含透明（如 linear-gradient(blue, transparent)）
                (hasGradient && (bgImage.includes('transparent') || /rgba\([^)]*,\s*0\)/.test(bgImage))) ||
                // 有 box-shadow
                (styles.boxShadow !== 'none' && styles.boxShadow !== '') ||
                // 非渐变的 background-image（如 url(...)）
                (bgImage && bgImage !== 'none' && !hasGradient && bgImage.includes('url('))
            );
            if (isComplexDeco) entry._complex_deco = true;
            
            results.push({...entry});
            addedShape = true;
        }

        // 检查 ::before / ::after 伪元素的可见装饰
        for (const pseudo of ['::before', '::after']) {
            const ps = getComputedStyle(el, pseudo);
            const content = ps.content;
            if (!content || content === 'none' || content === 'normal') continue;

            const pBg = ps.backgroundColor;
            const pBgImg = ps.backgroundImage;
            const pIsTransparent = pBg === 'rgba(0, 0, 0, 0)' || pBg === 'transparent';
            const pHasGradient = pBgImg && pBgImg.includes('linear-gradient');
            const pBorder = parseBorder(ps);
            const pHasBorder = pBorder !== null;

            if (pIsTransparent && !pHasGradient && !pHasBorder) continue;

            // 伪元素的尺寸：用 computed width/height，位置基于父元素 + position
            let pw = parseFloat(ps.width) || 0;
            let ph = parseFloat(ps.height) || 0;
            if (pw < 1 && ph < 1) continue;
            // 如果宽度是 100%，取父元素宽度
            if (pw === 0 || ps.width === 'auto') pw = w;
            if (ph === 0 || ps.height === 'auto') ph = 2; // 默认线高

            // 位置偏移
            let px = x + (parseFloat(ps.left) || 0);
            let py = y + (parseFloat(ps.top) || 0);

            const pseudoEntry = {
                tag: tag + pseudo,
                _role: 'shape',
                _shape_idx: elIndex++,
                x: Math.round(px * 100) / 100,
                y: Math.round(py * 100) / 100,
                w: Math.round(pw * 100) / 100,
                h: Math.round(ph * 100) / 100,
                opacity: parseFloat(ps.opacity),
                bg_color: pBg,
            };
            if (pHasGradient) pseudoEntry.bg_gradient = pBgImg;
            pseudoEntry.border_radius = ps.borderRadius;
            if (pHasBorder) pseudoEntry.border = pBorder;
            results.push(pseudoEntry);
        }
        
        // 文字（只有叶子节点才提取文字，避免父子重复）
        if (text && isLeaf) {
            const textKey = text + '|' + x + '|' + y;
            if (!extractedTexts.has(textKey)) {
                extractedTexts.add(textKey);

                // 补偿 padding：getBoundingClientRect 返回 border-box，
                // 但文字实际从 content-box 开始渲染
                const padTop = parseFloat(styles.paddingTop) || 0;
                const padRight = parseFloat(styles.paddingRight) || 0;
                const padBottom = parseFloat(styles.paddingBottom) || 0;
                const padLeft = parseFloat(styles.paddingLeft) || 0;
                const textX = x + padLeft;
                const textY = y + padTop;
                const textW = w - padLeft - padRight;
                const textH = h - padTop - padBottom;
                
                // 计算文字的实际渲染宽度
                const actualTextWidth = getTextWidth(el);
                const fontSize = parseFloat(styles.fontSize);
                const isSingleLine = textH <= fontSize * 2;
                
                let safeW;
                const isCentered = styles.textAlign === 'center';
                if (isSingleLine && !isCentered) {
                    // 单行非居中文字：加 15% 余量，但不超过 section 宽度
                    safeW = Math.min(
                        Math.ceil(Math.max(actualTextWidth, textW) * 1.15),
                        sectionRect.width
                    );
                } else {
                    // 多行文字或居中文字：保持容器宽度
                    safeW = textW;
                }
                
                const textEntry = {
                    tag,
                    _role: 'text',
                    x: Math.round(textX * 100) / 100,
                    y: Math.round(textY * 100) / 100,
                    w: Math.round(safeW * 100) / 100,
                    h: Math.round(textH * 100) / 100,
                    opacity: parseFloat(styles.opacity),
                    text: text,
                    color: styles.color,
                    font_size: parseFloat(styles.fontSize),
                    font_family: styles.fontFamily,
                    font_weight: styles.fontWeight,
                    text_align: styles.textAlign,
                    line_height: styles.lineHeight,
                };

                // 检测纵向对齐
                // 1. 自身是 flex 且 align-items: center
                // 2. 父元素是 flex 且 align-items: center
                // 3. line-height 等于或接近元素高度（单行居中技巧）
                let vAlign = 'top';
                const display = styles.display;
                const alignItems = styles.alignItems;
                if ((display === 'flex' || display === 'inline-flex') && alignItems === 'center') {
                    vAlign = 'middle';
                }
                if (vAlign === 'top' && el.parentElement) {
                    const parentStyles = getComputedStyle(el.parentElement);
                    if ((parentStyles.display === 'flex' || parentStyles.display === 'inline-flex')
                        && parentStyles.alignItems === 'center') {
                        vAlign = 'middle';
                    }
                }
                if (vAlign === 'top') {
                    const lh = parseFloat(styles.lineHeight);
                    if (!isNaN(lh) && textH > 0 && Math.abs(lh - textH) < 2) {
                        vAlign = 'middle';
                    }
                }
                textEntry.vertical_align = vAlign;

                // 检查是否有行内子元素导致混合字体
                // 如果有，提取 runs 数组（每个 run 有自己的文字和字体信息）
                const inlineTags = new Set(['span','strong','b','em','i','a','mark','small',
                                             'sub','sup','code','u','s','del','ins','label','time']);
                const hasInlineChildren = Array.from(el.children).some(c => 
                    inlineTags.has(c.tagName.toLowerCase()) && (c.innerText || '').trim()
                );

                if (hasInlineChildren) {
                    const runs = [];
                    function extractRuns(node) {
                        for (const child of node.childNodes) {
                            if (child.nodeType === 3) {
                                // 纯文字节点，继承父元素样式
                                const t = child.textContent;
                                if (t && t.trim()) {
                                    const ps = getComputedStyle(node);
                                    runs.push({
                                        text: t,
                                        color: ps.color,
                                        font_size: parseFloat(ps.fontSize),
                                        font_family: ps.fontFamily,
                                        font_weight: ps.fontWeight,
                                    });
                                }
                            } else if (child.nodeType === 1) {
                                const childTag = child.tagName.toLowerCase();
                                if (childTag === 'br') {
                                    runs.push({ text: String.fromCharCode(10) });
                                } else if (childTag === 'style') {
                                    continue;
                                } else if (inlineTags.has(childTag)) {
                                    // 行内元素：用自身的样式
                                    const cs = getComputedStyle(child);
                                    const ct = child.innerText || '';
                                    if (ct) {
                                        runs.push({
                                            text: ct,
                                            color: cs.color,
                                            font_size: parseFloat(cs.fontSize),
                                            font_family: cs.fontFamily,
                                            font_weight: cs.fontWeight,
                                        });
                                    }
                                } else {
                                    // 其他元素（如 div/p 内嵌），递归提取
                                    extractRuns(child);
                                }
                            }
                        }
                    }
                    extractRuns(el);

                    if (runs.length > 0) {
                        textEntry.runs = runs;
                    }
                }

                results.push(textEntry);
            }
        }
        
        // 继续遍历子元素
        for (const child of el.children) {
            if (child.tagName.toLowerCase() === 'style') continue;
            walk(child);
        }
    }
    
    for (const child of sectionEl.children) {
        if (child.tagName.toLowerCase() === 'style') continue;
        walk(child);
    }
    
    return results;
}
"""


# ─── 图片 ID 匹配 ──────────────────────────────────────────────────────────────

def _match_img_ids(elements: list, html_content: str, page_num: int, html_path: str = None) -> None:
    """
    为图片元素匹配 IMG_ID。
    策略：按图片在 HTML 中出现的顺序，与提取到的图片元素顺序对应。
    """
    # 从 HTML 中按顺序找出该页的所有图片 ID
    page_pattern = re.compile(
        rf'<section\s+class="page-{page_num}".*?</section>',
        re.DOTALL
    )
    page_match = page_pattern.search(html_content)
    if not page_match:
        return

    page_html = page_match.group()
    img_elements = [e for e in elements if e.get('_role') == 'image']

    # 从 html_path 所在目录的 images/ 子目录查找
    if html_path:
        images_dir = Path(html_path).parent / 'images'
    else:
        images_dir = Path(__file__).parent.parent / 'output' / 'images'
    page_prefix = f'IMG_{page_num:02d}_'
    available = sorted([
        f.stem for f in images_dir.glob(f'{page_prefix}*.jpg')
    ])

    # 按在 HTML 中的纵向位置（y 坐标）排序图片元素
    img_elements.sort(key=lambda e: (e.get('y', 0), e.get('x', 0)))

    # 背景图（通常最大或 y=0）
    bg_id = f'IMG_{page_num:02d}_BG'
    if bg_id in [a for a in available]:
        # 找面积最大的图片元素
        if img_elements:
            biggest = max(img_elements, key=lambda e: e.get('w', 0) * e.get('h', 0))
            biggest['img_id'] = bg_id
            available.remove(bg_id)
            img_elements = [e for e in img_elements if e is not biggest]

    # 剩余图片按顺序匹配
    content_ids = [a for a in available if not a.endswith('_BG')]
    for img_el, img_id in zip(img_elements, content_ids):
        img_el['img_id'] = img_id


# ─── 后处理 ────────────────────────────────────────────────────────────────────

def _post_process(raw_elements: list) -> list:
    """将原始提取数据转换为 Phase 5 所需的标准格式。"""
    result = []

    for el in raw_elements:
        role = el.get('_role')
        if role == 'section_bg':
            continue  # section 背景单独处理

        base = {
            'x_emu': px_to_emu(el.get('x', 0)),
            'y_emu': px_to_emu(el.get('y', 0)),
            'w_emu': px_to_emu(el.get('w', 0)),
            'h_emu': px_to_emu(el.get('h', 0)),
        }

        if el.get('opacity', 1.0) < 1.0:
            base['opacity'] = el['opacity']

        if role == 'image':
            base['type'] = 'image'
            base['img_id'] = el.get('img_id', 'UNKNOWN')
            base['object_fit'] = el.get('object_fit', 'cover')
            br = el.get('border_radius', '0px')
            try:
                base['border_radius_px'] = int(float(br.replace('px', '').split()[0]))
            except (ValueError, IndexError):
                base['border_radius_px'] = 0
            result.append(base)

        elif role == 'shape':
            shape = {**base, 'type': 'shape'}
            gradient = _parse_gradient(el.get('bg_gradient', ''))
            if gradient:
                shape['gradient'] = gradient
            else:
                shape['bg_color'] = _color_to_hex(el.get('bg_color', ''))
            
            br = el.get('border_radius', '0px')
            try:
                shape['border_radius_px'] = int(float(br.replace('px', '').split()[0]))
            except (ValueError, IndexError):
                shape['border_radius_px'] = 0

            if el.get('box_shadow'):
                shape['box_shadow'] = el['box_shadow']

            # border 属性（parseBorder 已取最细边框）
            border = el.get('border')
            if border:
                side_count = border.get('sideCount', 4)
                active_sides = border.get('activeSides', [])
                border_color = _color_to_hex(border['color'])
                border_w_pt = round(border['width'] * 72 / 96, 1)

                if side_count >= 3:
                    # 3-4 边 border → 正常应用为矩形边框
                    shape['border'] = {
                        'width_pt': border_w_pt,
                        'color': border_color,
                    }
                else:
                    # 1-2 边 border → 每条边生成独立的窄色条
                    bw_emu = px_to_emu(border['width'])
                    first = True
                    for side in active_sides:
                        if first:
                            # 第一条边直接修改当前 shape
                            target = shape
                            first = False
                        else:
                            # 后续边生成额外的 shape
                            target = {**base, 'type': 'shape', 'border_radius_px': 0}
                            shape.setdefault('_extra_shapes', []).append(target)
                        
                        if side == 'Left':
                            target['w_emu'] = bw_emu
                        elif side == 'Right':
                            target['x_emu'] = base['x_emu'] + base['w_emu'] - bw_emu
                            target['w_emu'] = bw_emu
                        elif side == 'Top':
                            target['h_emu'] = bw_emu
                        elif side == 'Bottom':
                            target['y_emu'] = base['y_emu'] + base['h_emu'] - bw_emu
                            target['h_emu'] = bw_emu
                        target['bg_color'] = border_color
                        target.pop('gradient', None)

            result.append(shape)
            # 追加多边框产生的额外 shape
            for extra in shape.pop('_extra_shapes', []):
                result.append(extra)

        elif role == 'text':
            txt = {**base, 'type': 'text'}
            txt['text'] = el['text']
            txt['color'] = _color_to_hex(el.get('color', ''))
            txt['font_size_pt'] = px_to_pt(el.get('font_size', 15))
            txt['font_family'] = _clean_font_family(el.get('font_family', ''))
            txt['font_weight'] = el.get('font_weight', '400')
            txt['text_align'] = el.get('text_align', 'left')
            txt['vertical_align'] = el.get('vertical_align', 'top')

            # 混合字体 runs（如果有行内子元素导致不同字体）
            if el.get('runs'):
                txt['runs'] = []
                for r in el['runs']:
                    run_data = {'text': r['text']}
                    if r.get('color'):
                        run_data['color'] = _color_to_hex(r['color'])
                    if r.get('font_size'):
                        run_data['font_size_pt'] = px_to_pt(r['font_size'])
                    if r.get('font_family'):
                        run_data['font_family'] = _clean_font_family(r['font_family'])
                    if r.get('font_weight'):
                        run_data['font_weight'] = r['font_weight']
                    txt['runs'].append(run_data)

            result.append(txt)

    return result


def _clean_font_family(raw: str) -> str:
    """提取第一个字体名。"""
    if not raw:
        return 'MiSans-Regular'
    first = raw.split(',')[0].strip().strip("'\"")
    return first


# ─── 容器截图 ──────────────────────────────────────────────────────────────────

def _is_contained(inner, outer) -> bool:
    """判断 inner 元素是否大部分在 outer 元素内部。
    
    使用面积重叠比例判断（≥75% 即视为包含），而非严格边界包含。
    因为 HTML 中装饰元素（如大引号）可能因 overflow 等原因
    略微超出容器边界，严格包含判定会遗漏它们。
    """
    ix, iy = inner.get('x', 0), inner.get('y', 0)
    iw, ih = inner.get('w', 0), inner.get('h', 0)
    ox, oy = outer.get('x', 0), outer.get('y', 0)
    ow, oh = outer.get('w', 0), outer.get('h', 0)

    # 计算交集区域
    inter_x1 = max(ix, ox)
    inter_y1 = max(iy, oy)
    inter_x2 = min(ix + iw, ox + ow)
    inter_y2 = min(iy + ih, oy + oh)

    if inter_x2 <= inter_x1 or inter_y2 <= inter_y1:
        return False  # 无交集

    inter_area = (inter_x2 - inter_x1) * (inter_y2 - inter_y1)
    inner_area = max(iw * ih, 1)

    return inter_area / inner_area >= 0.75


def _identify_containers(raw_elements: list) -> list:
    """
    识别需要截图的 shape 元素：
    1. 容器：有装饰样式且包含子文本的 shape
    2. 复杂装饰元素：渐变到透明、box-shadow 等 PPTX 难以还原的 shape
    3. 图片包装器：仅包含图片不包含文字的 shape（标记为待删除）
    
    排除全屏背景 shape（面积 > 页面 80%）。
    返回: [(shape_idx, shape_el, contained_text_indices, contained_shape_indices)]
    contained_texts 为 ['__deco__'] 表示复杂装饰元素（非真正容器）。
    """
    PAGE_AREA = 1280 * 720
    shapes = [(i, el) for i, el in enumerate(raw_elements)
              if el.get('_role') == 'shape']
    texts = [(i, el) for i, el in enumerate(raw_elements)
             if el.get('_role') == 'text']
    images = [(i, el) for i, el in enumerate(raw_elements)
              if el.get('_role') == 'image']

    containers = []  # [(shape_idx, shape_el, [contained_text_indices], [contained_shape_indices])]

    for si, shape in shapes:
        area = shape.get('w', 0) * shape.get('h', 0)
        # 跳过全屏/接近全屏的背景 shape
        if area > PAGE_AREA * 0.8:
            continue

        # 复杂装饰元素：不管面积大小，只要标记了就截图
        if shape.get('_complex_deco'):
            # 跳过宽和高都太小的（两个维度都 < 3px）
            if shape.get('w', 0) < 3 and shape.get('h', 0) < 3:
                continue
            containers.append((si, shape, ['__deco__'], []))
            continue

        # 跳过太小的装饰条
        if area < 2000:
            continue

        # 找包含的文字和图片
        contained_texts = [ti for ti, txt in texts if _is_contained(txt, shape)]
        contained_images = [ii for ii, img in images if _is_contained(img, shape)]

        # 仅包含图片、不包含文字的 shape → 图片包装器，标记为待删除
        if not contained_texts and contained_images:
            contained_shapes_inner = []
            for si2, shape2 in shapes:
                if si2 != si and _is_contained(shape2, shape):
                    contained_shapes_inner.append(si2)
            # 用 contained_texts=[] 标记为 "图片包装器"
            containers.append((si, shape, [], contained_shapes_inner))
            continue

        if not contained_texts:
            continue  # 不包含文字也不包含图片的 shape 不处理

        # 找包含的子 shape
        contained_shapes = []
        for si2, shape2 in shapes:
            if si2 != si and _is_contained(shape2, shape):
                contained_shapes.append(si2)

        containers.append((si, shape, contained_texts, contained_shapes))

    return containers


def _screenshot_containers(section, page, containers: list, page_num: int,
                           screenshots_dir: Path) -> dict:
    """
    对每个容器和复杂装饰元素截图。
    
    策略：先隐藏 section 内所有子元素，然后逐个显示目标元素并截图，
    最后恢复所有元素。这同时解决了文字残留和截图串扰两个问题。
    """
    # 收集需截图的元素（容器 + 复杂装饰）
    screenshot_targets = []
    for si, shape, contained_texts, _ in containers:
        # 跳过图片包装器（无文字的容器）
        if not contained_texts:
            continue
        shape_idx = shape.get('_shape_idx')
        if shape_idx is not None:
            screenshot_targets.append((si, shape_idx, 'container'))

    result = {}
    if not screenshot_targets:
        return result

    # ── 隐藏 section 内所有直接子元素 ──
    page.evaluate("""
        (sectionEl) => {
            const children = Array.from(sectionEl.children);
            for (const child of children) {
                if (child.tagName === 'STYLE') continue;
                child.dataset.origVis = child.style.visibility || '';
                child.style.visibility = 'hidden';
            }
        }
    """, section)
    page.wait_for_timeout(100)

    # ── 逐个显示目标、截图、隐藏 ──
    for si, shape_idx, kind in screenshot_targets:
        el_handle = section.query_selector(f'[data-shape-idx="{shape_idx}"]')
        if not el_handle:
            continue

        # 显示目标元素及其所有子元素（但不显示文字）
        page.evaluate("""
            (args) => {
                const [sectionEl, shapeIdx] = args;
                const target = sectionEl.querySelector('[data-shape-idx="' + shapeIdx + '"]');
                if (!target) return;
                // 显示目标自身（沿 DOM 树向上直到 section）
                let el = target;
                while (el && el !== sectionEl) {
                    el.style.visibility = 'visible';
                    el = el.parentElement;
                }
                // 显示目标的所有子元素（但跳过文字节点的父元素）
                target.querySelectorAll('*').forEach(c => {
                    if (c.tagName === 'STYLE') return;
                    const hasText = c.childNodes.length > 0 && 
                        Array.from(c.childNodes).some(n => n.nodeType === 3 && n.textContent.trim());
                    if (!hasText) {
                        c.style.visibility = 'visible';
                    }
                });
            }
        """, [section, shape_idx])
        page.wait_for_timeout(50)

        box = el_handle.bounding_box()
        if not box or (box['width'] < 1 and box['height'] < 1):
            # 隐藏回去
            page.evaluate("""
                (args) => {
                    const [sectionEl, shapeIdx] = args;
                    const target = sectionEl.querySelector('[data-shape-idx="' + shapeIdx + '"]');
                    if (!target) return;
                    let el = target;
                    while (el && el !== sectionEl) { el.style.visibility = 'hidden'; el = el.parentElement; }
                    target.querySelectorAll('*').forEach(c => { c.style.visibility = ''; });
                }
            """, [section, shape_idx])
            continue

        img_path = str(screenshots_dir / f'{kind}_p{page_num:02d}_{shape_idx}.png')
        try:
            el_handle.screenshot(path=img_path, timeout=10000)
            result[si] = img_path
        except Exception:
            pass

        # 隐藏回去
        page.evaluate("""
            (args) => {
                const [sectionEl, shapeIdx] = args;
                const target = sectionEl.querySelector('[data-shape-idx="' + shapeIdx + '"]');
                if (!target) return;
                let el = target;
                while (el && el !== sectionEl) { el.style.visibility = 'hidden'; el = el.parentElement; }
                target.querySelectorAll('*').forEach(c => { c.style.visibility = ''; });
            }
        """, [section, shape_idx])

    # ── 恢复所有子元素可见性 ──
    page.evaluate("""
        (sectionEl) => {
            const children = Array.from(sectionEl.children);
            for (const child of children) {
                if (child.tagName === 'STYLE') continue;
                child.style.visibility = child.dataset.origVis || '';
                delete child.dataset.origVis;
            }
        }
    """, section)

    return result


def _apply_container_conversion(elements: list, raw_elements: list,
                                containers: list, container_screenshots: dict) -> list:
    """
    将容器 shape 转为 container_image 类型，移除被容器覆盖的子 shape。
    图片包装器（仅含图片无文字的容器）直接删除。
    """
    # 收集索引
    child_shape_indices = set()
    container_indices = set()  # 有截图的容器
    image_wrapper_indices = set()  # 图片包装器（无文字，仅含图片）
    decorative_text_indices = set()  # 装饰性文字（如引号），已在截图中

    for si, shape, contained_texts, contained_shapes in containers:
        if not contained_texts:
            # 图片包装器：直接删除 shape，也删除子 shape
            image_wrapper_indices.add(si)
            child_shape_indices.update(contained_shapes)
        elif contained_texts == ['__deco__']:
            # 复杂装饰元素：直接截图替换
            if si in container_screenshots:
                container_indices.add(si)
        elif si in container_screenshots:
            container_indices.add(si)
            child_shape_indices.update(contained_shapes)
            # 检查被包含的文字是否为装饰性纯符号（如引号），这些已在截图中
            for ti in contained_texts:
                if ti < len(raw_elements):
                    txt = raw_elements[ti].get('text', '').strip()
                    # 纯装饰符号：引号、星号等单字符或极短非内容文字
                    if len(txt) <= 2 and not txt.isalnum():
                        decorative_text_indices.add(ti)

    result = []
    raw_idx = 0
    for el in elements:
        # 找到对应的 raw_element
        while raw_idx < len(raw_elements):
            raw_el = raw_elements[raw_idx]
            if raw_el.get('_role') in ('section_bg',):
                raw_idx += 1
                continue
            break

        if raw_idx in image_wrapper_indices:
            # 图片包装器：删除 shape（图片元素会独立保留）
            pass
        elif raw_idx in container_indices and raw_idx in container_screenshots:
            # 将此 shape 替换为 container_image
            el['type'] = 'container_image'
            el['screenshot_path'] = container_screenshots[raw_idx]
            for key in ['bg_color', 'gradient', 'border', 'box_shadow']:
                el.pop(key, None)
            result.append(el)
        elif raw_idx in child_shape_indices:
            # 子 shape 被容器截图覆盖，跳过
            pass
        elif raw_idx in decorative_text_indices:
            # 装饰性文字已在容器截图中，跳过（避免重叠）
            pass
        else:
            result.append(el)

        raw_idx += 1

    # 如果还有剩余的 elements（raw_idx 不匹配），全部保留
    # 这种情况不应发生，但防御性编程
    return result


# ─── 主入口 ────────────────────────────────────────────────────────────────────

def extract_layout(html_path: str, progress_file: str = None) -> list:
    """
    渲染 HTML，逐页提取布局数据。
    
    核心策略："逐元素截图 + 可编辑文字"
    1. 提取所有元素
    2. 清空所有 TextNode（最强文字消除保证）
    3. 逐个 shape 元素独立截图
    4. 恢复文字
    5. 输出：shape 截图 + 可编辑文字
    
    返回: [{page, screenshot, elements}, ...]
    """
    html_path = str(Path(html_path).resolve())
    screenshots_dir = Path(html_path).parent / 'screenshots'
    screenshots_dir.mkdir(exist_ok=True)

    results = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(viewport={'width': 1280, 'height': 720})
        page.goto(f'file://{html_path}', wait_until='networkidle', timeout=60000)

        # 等字体加载
        page.wait_for_timeout(2000)

        sections = page.query_selector_all('section')
        print(f'  找到 {len(sections)} 个 section')

        total_sections = len(sections)
        for idx, section in enumerate(sections):
            page_num = idx + 1
            print(f'  → 提取第 {page_num} 页...')

            # 写进度
            if progress_file:
                import json as _json
                with open(progress_file, 'w') as _pf:
                    _pf.write(_json.dumps({'current': idx, 'total': total_sections, 'stage': 'screenshot'}))

            # ── 完整截图（含文字，用于预览缩略图） ──
            screenshot_path = str(screenshots_dir / f'page_{page_num:02d}.png')
            section.screenshot(path=screenshot_path)

            # ── 提取元素 ──
            raw = section.evaluate(EXTRACT_JS)

            # 收集所有需截图的元素（shape + image）的 _shape_idx
            shape_indices = []
            for el in raw:
                if el.get('_role') in ('shape', 'image') and el.get('_shape_idx') is not None:
                    shape_indices.append(el['_shape_idx'])

            # ── 准备截图环境 ──
            # 1. 清空所有 TextNode（最强文字消除保证）
            # 2. 隐藏所有 section 直接子元素（防止截图串扰）
            page.evaluate("""
                (sectionEl) => {
                    // 清空 TextNode（跳过 style/script）
                    const walker = document.createTreeWalker(
                        sectionEl, NodeFilter.SHOW_TEXT);
                    const savedTexts = [];
                    const skipTags = new Set(['STYLE', 'SCRIPT']);
                    while (walker.nextNode()) {
                        const tn = walker.currentNode;
                        if (tn.parentElement && skipTags.has(tn.parentElement.tagName)) continue;
                        if (tn.textContent && tn.textContent.trim()) {
                            savedTexts.push({ node: tn, text: tn.textContent });
                            tn.textContent = '';
                        }
                    }
                    sectionEl.__savedTexts = savedTexts;

                    // 隐藏所有直接子元素
                    const children = Array.from(sectionEl.children);
                    for (const child of children) {
                        if (child.tagName === 'STYLE') continue;
                        child.dataset.origVis = child.style.visibility || '';
                        child.style.visibility = 'hidden';
                    }
                }
            """, section)
            page.wait_for_timeout(100)

            # ── 识别背景元素（面积 ≥ 80% 页面）──
            section_area = 1280 * 720
            bg_shape_indices = set()
            for el in raw:
                role = el.get('_role')
                if role in ('shape', 'image') and el.get('_shape_idx') is not None:
                    el_area = el.get('w', 0) * el.get('h', 0)
                    if el_area >= section_area * 0.8:
                        bg_shape_indices.add(el['_shape_idx'])

            # ── 背景合成截图 ──
            bg_composite_path = None
            if bg_shape_indices:
                # 显示所有背景元素（文字已清空，小元素已隐藏）
                page.evaluate("""
                    (args) => {
                        const [sectionEl, indices] = args;
                        const indexSet = new Set(indices);
                        // section 本身也恢复可见（保留 CSS 背景）
                        sectionEl.style.visibility = 'visible';
                        for (const child of sectionEl.children) {
                            if (child.tagName === 'STYLE') continue;
                            const idx = child.dataset?.shapeIdx;
                            if (idx !== undefined && indexSet.has(parseInt(idx))) {
                                child.style.visibility = 'visible';
                                child.querySelectorAll('*').forEach(
                                    c => { c.style.visibility = 'visible'; });
                            }
                        }
                    }
                """, [section, list(bg_shape_indices)])

                bg_composite_path = str(screenshots_dir / f'page_{page_num:02d}_bg_composite.png')
                section.screenshot(path=bg_composite_path)
                print(f'    🎨 背景合成截图（{len(bg_shape_indices)} 个背景元素）')

                # 重新隐藏背景元素
                page.evaluate("""
                    (args) => {
                        const [sectionEl, indices] = args;
                        const indexSet = new Set(indices);
                        for (const child of sectionEl.children) {
                            if (child.tagName === 'STYLE') continue;
                            const idx = child.dataset?.shapeIdx;
                            if (idx !== undefined && indexSet.has(parseInt(idx))) {
                                child.style.visibility = 'hidden';
                                child.querySelectorAll('*').forEach(
                                    c => { c.style.visibility = ''; });
                            }
                        }
                    }
                """, [section, list(bg_shape_indices)])

            # ── 逐个元素：显示 → 截图 → 隐藏（跳过背景元素）──
            shape_screenshots = {}  # {_shape_idx: screenshot_path}
            for si in shape_indices:
                if si in bg_shape_indices:
                    continue  # 背景元素已在合成截图中，跳过单独截图
                el_handle = section.query_selector(f'[data-shape-idx="{si}"]')
                if not el_handle:
                    continue
                try:
                    # 显示目标元素（沿 DOM 树向上到 section）
                    page.evaluate("""
                        (args) => {
                            const [sectionEl, idx] = args;
                            const target = sectionEl.querySelector(
                                '[data-shape-idx="' + idx + '"]');
                            if (!target) return;
                            let el = target;
                            while (el && el !== sectionEl) {
                                el.style.visibility = 'visible';
                                el = el.parentElement;
                            }
                            // 显示目标的所有子元素
                            target.querySelectorAll('*').forEach(
                                c => { c.style.visibility = 'visible'; });
                        }
                    """, [section, si])

                    box = el_handle.bounding_box()
                    if not box or (box['width'] < 1 and box['height'] < 1):
                        raise ValueError('too small')

                    img_path = str(screenshots_dir / f'shape_p{page_num:02d}_{si}.png')
                    el_handle.screenshot(path=img_path, timeout=3000)
                    shape_screenshots[si] = img_path
                except Exception:
                    pass
                finally:
                    # 隐藏回去
                    page.evaluate("""
                        (args) => {
                            const [sectionEl, idx] = args;
                            const target = sectionEl.querySelector(
                                '[data-shape-idx="' + idx + '"]');
                            if (!target) return;
                            let el = target;
                            while (el && el !== sectionEl) {
                                el.style.visibility = 'hidden';
                                el = el.parentElement;
                            }
                            target.querySelectorAll('*').forEach(
                                c => { c.style.visibility = ''; });
                        }
                    """, [section, si])

            if shape_screenshots:
                print(f'    🖼️ {len(shape_screenshots)} 个视觉元素已截图')

            # ── 恢复环境 ──
            page.evaluate("""
                (sectionEl) => {
                    // 恢复子元素可见性
                    const children = Array.from(sectionEl.children);
                    for (const child of children) {
                        if (child.tagName === 'STYLE') continue;
                        child.style.visibility = child.dataset.origVis || '';
                        delete child.dataset.origVis;
                    }
                    // 恢复文字
                    const saved = sectionEl.__savedTexts || [];
                    for (const { node, text } of saved) {
                        node.textContent = text;
                    }
                    delete sectionEl.__savedTexts;
                }
            """, section)

            # ── 后处理：shape → 截图元素，text → 可编辑文字 ──
            # 过滤掉背景元素（它们已在合成截图中）
            filtered_raw = [el for el in raw
                            if el.get('_shape_idx') not in bg_shape_indices]
            elements = _post_process_hybrid(filtered_raw, shape_screenshots)

            # 提取 section 背景色（兜底，有合成截图时优先用截图）
            page_bg_color = ''
            for el in raw:
                if el.get('_role') == 'section_bg':
                    page_bg_color = _color_to_hex(el.get('bg_color', ''))
                    break

            page_entry = {
                'page': page_num,
                'screenshot': screenshot_path,
                'elements': elements,
            }
            if bg_composite_path:
                page_entry['bg_screenshot'] = bg_composite_path
            elif page_bg_color:
                page_entry['bg_color'] = page_bg_color
            results.append(page_entry)

            print(f'    ✓ {len(elements)} 个元素')

        browser.close()

    return results


def _post_process_hybrid(raw_elements: list, shape_screenshots: dict) -> list:
    """
    混合后处理：
    - shape 元素 → 如果有截图则转为 container_image，否则跳过
    - text 元素 → 保留为可编辑文字
    - image 元素 → 跳过（已包含在 shape 截图中）
    """
    from PIL import Image as PILImage

    result = []

    for el in raw_elements:
        role = el.get('_role')
        if role == 'section_bg':
            continue

        base = {
            'x_emu': px_to_emu(el.get('x', 0)),
            'y_emu': px_to_emu(el.get('y', 0)),
            'w_emu': px_to_emu(el.get('w', 0)),
            'h_emu': px_to_emu(el.get('h', 0)),
        }

        if el.get('opacity', 1.0) < 1.0:
            base['opacity'] = el['opacity']

        if role == 'shape' or role == 'image':
            shape_idx = el.get('_shape_idx')
            if shape_idx is not None and shape_idx in shape_screenshots:
                item = {**base, 'type': 'container_image'}
                screenshot_path = shape_screenshots[shape_idx]
                item['screenshot_path'] = screenshot_path
                if role == 'image':
                    item['_is_image'] = True

                # ── 截图尺寸校正 ──
                # 截图时 visibility 会沿 DOM 向上传播，导致截图区域可能
                # 远小于父元素的布局边界。如果截图实际像素尺寸远小于布局
                # 声明的尺寸，用截图实际尺寸，避免小截图拉伸后覆盖文字。
                try:
                    img = PILImage.open(screenshot_path)
                    img_w, img_h = img.size
                    img.close()
                    layout_w_px = el.get('w', 0)
                    layout_h_px = el.get('h', 0)
                    # 如果截图面积不到布局面积的 50%，说明截图只截到了子元素
                    layout_area = layout_w_px * layout_h_px
                    img_area = img_w * img_h
                    if layout_area > 0 and img_area < layout_area * 0.5:
                        # 用截图实际尺寸，并居中放置在原布局区域中
                        old_x = el.get('x', 0)
                        old_y = el.get('y', 0)
                        center_x = old_x + layout_w_px / 2
                        center_y = old_y + layout_h_px / 2
                        item['x_emu'] = px_to_emu(center_x - img_w / 2)
                        item['y_emu'] = px_to_emu(center_y - img_h / 2)
                        item['w_emu'] = px_to_emu(img_w)
                        item['h_emu'] = px_to_emu(img_h)
                except Exception:
                    pass  # 读取失败就用原始尺寸

                result.append(item)

        elif role == 'text':
            txt = {**base, 'type': 'text'}
            txt['text'] = el['text']
            txt['color'] = _color_to_hex(el.get('color', ''))
            txt['font_size_pt'] = px_to_pt(el.get('font_size', 15))
            txt['font_family'] = _clean_font_family(el.get('font_family', ''))
            txt['font_weight'] = el.get('font_weight', '400')
            txt['text_align'] = el.get('text_align', 'left')
            txt['vertical_align'] = el.get('vertical_align', 'top')

            # 混合字体 runs
            if el.get('runs'):
                txt['runs'] = []
                for r in el['runs']:
                    run_data = {'text': r['text']}
                    if r.get('color'):
                        run_data['color'] = _color_to_hex(r['color'])
                    if r.get('font_size'):
                        run_data['font_size_pt'] = px_to_pt(r['font_size'])
                    if r.get('font_family'):
                        run_data['font_family'] = _clean_font_family(r['font_family'])
                    if r.get('font_weight'):
                        run_data['font_weight'] = r['font_weight']
                    txt['runs'].append(run_data)

            result.append(txt)

    return result


