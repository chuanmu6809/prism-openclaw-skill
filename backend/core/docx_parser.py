from pathlib import Path
from io import BytesIO
from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph as DocxParagraph
from docx.table import Table as DocxTable
from PIL import Image


# ─── XML 检测工具 ──────────────────────────────────────────────────────────────

def _get_heading_level(para) -> int:
    """返回段落标题级别 1/2/3，普通段落返回 0。"""
    pPr = para._p.find(qn('w:pPr'))
    if pPr is not None:
        outline = pPr.find(qn('w:outlineLvl'))
        if outline is not None:
            try:
                lvl = int(outline.get(qn('w:val'))) + 1  # val=0 → H1
                if 1 <= lvl <= 3:
                    return lvl
            except (TypeError, ValueError):
                pass
    # 兼容标准 Heading N 样式名
    style = para.style.name if para.style else ''
    for i in range(1, 4):
        if f'Heading {i}' in style or f'标题 {i}' in style:
            return i
    return 0


def _is_horizontal_rule(para) -> bool:
    """检测飞书分隔线（w:pBdr/w:bottom 元素）。"""
    pPr = para._p.find(qn('w:pPr'))
    if pPr is not None:
        pBdr = pPr.find(qn('w:pBdr'))
        if pBdr is not None and pBdr.find(qn('w:bottom')) is not None:
            return True
    return False


def _is_list_item(para) -> bool:
    """检测列表项（w:numPr 元素）。"""
    pPr = para._p.find(qn('w:pPr'))
    if pPr is not None:
        return pPr.find(qn('w:numPr')) is not None
    return False


def _get_image_bytes(para, doc):
    """提取段落中的图片二进制数据，无图片返回 None。"""
    blip_ns = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    rel_ns = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    for blip in para._p.iter(f'{{{blip_ns}}}blip'):
        r_id = blip.get(f'{{{rel_ns}}}embed')
        if r_id and r_id in doc.part.related_parts:
            return doc.part.related_parts[r_id].blob
    return None


# ─── 图片存储 ──────────────────────────────────────────────────────────────────

def _save_image(img_bytes: bytes, img_id: str, output_dir: str) -> str:
    """将图片转存为 JPEG，返回路径字符串。"""
    images_dir = Path(output_dir) / 'images'
    images_dir.mkdir(parents=True, exist_ok=True)
    img = Image.open(BytesIO(img_bytes))
    if img.mode not in ('RGB',):
        img = img.convert('RGB')
    path = images_dir / f'{img_id}.jpg'
    img.save(str(path), 'JPEG', quality=90)
    return str(path)


# ─── Body 迭代（兼容表格）──────────────────────────────────────────────────────

def _iter_body_items(doc):
    """
    按顺序迭代 document body 顶层元素，返回 DocxParagraph 或 DocxTable。
    用于替代 doc.paragraphs，以捕获表格内容（飞书并列图片使用表格存储）。
    注意：parent 必须传 doc._body（_Body 包装对象），而非原始 XML element，
    否则 para.style 无法向上找到 DocumentPart。
    """
    body_wrapper = doc._body  # _Body 包装对象，持有正确的 .part 链
    for child in doc.element.body.iterchildren():
        if child.tag == qn('w:p'):
            yield DocxParagraph(child, body_wrapper)
        elif child.tag == qn('w:tbl'):
            yield DocxTable(child, body_wrapper)


def _extract_table_image_bytes(table, doc) -> list:
    """从表格所有单元格中提取图片二进制列表（用于并列图片组）。"""
    blip_ns = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    rel_ns = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    imgs = []
    for blip in table._tbl.iter(f'{{{blip_ns}}}blip'):
        r_id = blip.get(f'{{{rel_ns}}}embed')
        if r_id and r_id in doc.part.related_parts:
            imgs.append(doc.part.related_parts[r_id].blob)
    return imgs


# ─── 容错后处理 ────────────────────────────────────────────────────────────────

def _split_bg_markers(classified: list) -> list:
    """
    处理飞书软换行写法：用户在同一段落内换行写了 [背景]，
    例如 para.text = "开启AI智能体新纪元\n[背景]"。
    将含 [背景] 的文字段落拆成：stripped_text + bg_marker 两个 item。
    """
    result = []
    for item in classified:
        if item['kind'] == 'text' and '[背景]' in item['text']:
            clean = item['text'].replace('[背景]', '').strip()
            if clean:
                result.append({'kind': 'text', 'text': clean, 'image_bytes': None})
            result.append({'kind': 'bg_marker', 'text': '[背景]', 'image_bytes': None})
        else:
            result.append(item)
    return result


# ─── 段落分类 ──────────────────────────────────────────────────────────────────

def _classify(para, doc) -> dict:
    """将段落分类，返回 {kind, text, image_bytes}。"""
    if _is_horizontal_rule(para):
        return {'kind': 'rule', 'text': '', 'image_bytes': None}

    lvl = _get_heading_level(para)
    if lvl > 0:
        return {'kind': f'h{lvl}', 'text': para.text.strip(), 'image_bytes': None}

    img_bytes = _get_image_bytes(para, doc)
    if img_bytes:
        return {'kind': 'image', 'text': '', 'image_bytes': img_bytes}

    text = para.text.strip()
    if text == '[背景]':
        return {'kind': 'bg_marker', 'text': text, 'image_bytes': None}
    if not text:
        return {'kind': 'empty', 'text': '', 'image_bytes': None}
    if _is_list_item(para):
        return {'kind': 'text', 'text': f'• {text}', 'image_bytes': None}
    return {'kind': 'text', 'text': text, 'image_bytes': None}


# ─── 内容块处理 ────────────────────────────────────────────────────────────────

def _process_block(items: list, page_num: int, img_counter: list, output_dir: str) -> tuple:
    """
    处理单个内容块，返回 (block_dict, images_dict)。
    img_counter: [int]，页内图片编号计数器（可变列表，跨 block 共享）。
    """
    content = []  # ('text', str) | ('image', img_id, path, is_bg)
    next_is_bg = False

    for item in items:
        k = item['kind']
        if k == 'bg_marker':
            next_is_bg = True
        elif k == 'image':
            if next_is_bg:
                img_id = f'IMG_{page_num:02d}_BG'
                is_bg = True
                next_is_bg = False
            else:
                img_counter[0] += 1
                img_id = f'IMG_{page_num:02d}_{img_counter[0]:02d}'
                is_bg = False
            path = _save_image(item['image_bytes'], img_id, output_dir)
            content.append(('image', img_id, path, is_bg))
        elif k == 'image_group':
            # 表格并列图片：展开为连续 image 条目，后续贪心匹配自动合组
            for img_bytes in item['image_bytes_list']:
                img_counter[0] += 1
                img_id = f'IMG_{page_num:02d}_{img_counter[0]:02d}'
                path = _save_image(img_bytes, img_id, output_dir)
                content.append(('image', img_id, path, False))
        elif k in ('text', 'h3'):  # h3 在 block 内视为正文
            content.append(('text', item['text']))
        # empty 忽略

    images_in_block = [(c[1], c[2], c[3]) for c in content if c[0] == 'image']
    images_dict = {img_id: {'path': path, 'is_background': is_bg}
                   for img_id, path, is_bg in images_in_block}

    if not images_in_block:
        texts = [c[1] for c in content if c[0] == 'text']
        return {'type': 'standalone_text', 'text': '\n'.join(texts)}, images_dict

    # 贪心匹配：连续图片合并为一组，与其前方文字配对（1:N 关系）
    pairs = []
    pending = []
    i = 0
    while i < len(content):
        c = content[i]
        if c[0] == 'text':
            pending.append(c[1])
            i += 1
        elif c[0] == 'image':
            # 收集连续图片组
            img_group = []
            while i < len(content) and content[i][0] == 'image':
                img_group.append(content[i][1])  # img_id
                i += 1
            pairs.append({
                'text': '\n'.join(pending).strip(),
                'image_ids': img_group  # 统一为列表，单图时也是 [img_id]
            })
            pending = []

    # 最后一组图之后的多余文字追加到最后一个 pair
    if pending and pairs:
        suffix = '\n'.join(pending).strip()
        sep = '\n' if pairs[-1]['text'] else ''
        pairs[-1]['text'] = pairs[-1]['text'] + sep + suffix

    block = {'type': 'paired', 'pairs': pairs}
    # 块内有多个 pair，或某个 pair 含多张图，均标记 multi_image
    if len(pairs) > 1 or any(len(p['image_ids']) > 1 for p in pairs):
        block['multi_image'] = True

    return block, images_dict


# ─── 入口函数 ──────────────────────────────────────────────────────────────────

def parse_docx(docx_path: str, output_dir: str) -> dict:
    """
    解析飞书导出的 DOCX，返回 {"total_pages": N, "pages": [...]}。
    图片保存至 output_dir/images/。
    """
    doc = Document(docx_path)
    Path(output_dir).mkdir(parents=True, exist_ok=True)

    # 迭代 body 顶层元素（段落 + 表格），表格视为图片组
    classified = []
    for item in _iter_body_items(doc):
        if isinstance(item, DocxParagraph):
            classified.append(_classify(item, doc))
        elif isinstance(item, DocxTable):
            imgs = _extract_table_image_bytes(item, doc)
            if imgs:
                classified.append({'kind': 'image_group', 'image_bytes_list': imgs, 'text': ''})

    # 容错后处理（软换行写法的 [背景] 标记）
    classified = _split_bg_markers(classified)

    # 按 H1 切页，H1 之前内容忽略
    raw_pages = []
    current = None
    for item in classified:
        if item['kind'] == 'h1':
            if current is not None:
                raw_pages.append(current)
            current = {'title': item['text'], 'items': []}
        elif current is not None:
            current['items'].append(item)
    if current is not None:
        raw_pages.append(current)

    # 逐页处理
    pages = []
    for idx, raw in enumerate(raw_pages):
        page_num = idx + 1
        items = raw['items']

        # 副标题：第一个 H2
        subtitle = next((i['text'] for i in items if i['kind'] == 'h2'), '')

        # 有分隔线 → 显式分块模式，否则 → 自由模式
        has_rule = any(i['kind'] == 'rule' for i in items)
        img_counter = [0]

        if has_rule:
            # ── 显式分块模式 ──
            blocks_raw = []
            current_block = []
            for item in items:
                if item['kind'] == 'h2':
                    continue  # 已提取为副标题，不进入 block
                if item['kind'] == 'rule':
                    blocks_raw.append(current_block)
                    current_block = []
                else:
                    current_block.append(item)
            blocks_raw.append(current_block)

            # 过滤纯空块
            blocks_raw = [b for b in blocks_raw
                          if any(i['kind'] not in ('empty',) for i in b)]

            blocks = []
            all_images = {}
            for block_items in blocks_raw:
                block, imgs = _process_block(block_items, page_num, img_counter, output_dir)
                blocks.append(block)
                all_images.update(imgs)

            pages.append({
                'page_number': page_num,
                'title': raw['title'],
                'subtitle': subtitle,
                'mode': 'explicit',
                'blocks': blocks,
                'images': all_images,
            })

        else:
            # ── 自由模式 ──
            paragraphs = []
            all_images = {}
            next_is_bg = False
            for item in items:
                k = item['kind']
                if k == 'h2':
                    continue
                elif k == 'h3':
                    paragraphs.append(item['text'])
                elif k == 'bg_marker':
                    next_is_bg = True
                elif k == 'image':
                    if next_is_bg:
                        img_id = f'IMG_{page_num:02d}_BG'
                        is_bg = True
                        next_is_bg = False
                    else:
                        img_counter[0] += 1
                        img_id = f'IMG_{page_num:02d}_{img_counter[0]:02d}'
                        is_bg = False
                    path = _save_image(item['image_bytes'], img_id, output_dir)
                    all_images[img_id] = {'path': path, 'is_background': is_bg}
                elif k == 'image_group':
                    for img_bytes in item['image_bytes_list']:
                        img_counter[0] += 1
                        img_id = f'IMG_{page_num:02d}_{img_counter[0]:02d}'
                        path = _save_image(img_bytes, img_id, output_dir)
                        all_images[img_id] = {'path': path, 'is_background': False}
                elif k == 'text':
                    paragraphs.append(item['text'])
                # empty 忽略

            pages.append({
                'page_number': page_num,
                'title': raw['title'],
                'subtitle': subtitle,
                'mode': 'free',
                'paragraphs': paragraphs,
                'images': all_images,
            })

    return {'total_pages': len(pages), 'pages': pages}
