"""
font_embedder.py
将字体文件嵌入已生成的 PPTX（符合 OOXML ECMA-376 §22.2 标准）。

用法（Phase 4 PPT 生成后调用）：
    from src.font_embedder import embed_fonts
    embed_fonts("output/result.pptx", style)   # style 来自 config/styles/*.json
"""
import uuid
import zipfile
import shutil
from pathlib import Path
from lxml import etree

# ─── XML 命名空间 ──────────────────────────────────────────────────────────────

_NS_PKG_REL  = "http://schemas.openxmlformats.org/package/2006/relationships"
_NS_PKG_CT   = "http://schemas.openxmlformats.org/package/2006/content-types"
_NS_PML      = "http://schemas.openxmlformats.org/presentationml/2006/main"
_NS_R        = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_FONT_CT     = "application/vnd.openxmlformats-officedocument.obfuscatedFont"
_FONT_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font"


# ─── OOXML 字体混淆（ECMA-376 §22.2.3）─────────────────────────────────────────

def _guid_to_key(guid_str: str) -> bytes:
    """
    将 GUID 字符串转为 16 字节混淆密钥。
    GUID 格式：{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}
    前三组按 little-endian（字节倒序），后两组按 big-endian（原序）。
    """
    parts = guid_str.strip("{}").split("-")
    # 4 bytes, 2 bytes, 2 bytes → little-endian
    b0 = bytes.fromhex(parts[0])[::-1]
    b1 = bytes.fromhex(parts[1])[::-1]
    b2 = bytes.fromhex(parts[2])[::-1]
    # 2 bytes + 6 bytes → big-endian（原序）
    b3 = bytes.fromhex(parts[3])
    b4 = bytes.fromhex(parts[4])
    return b0 + b1 + b2 + b3 + b4  # 16 bytes


def _obfuscate(font_bytes: bytes, guid_str: str) -> bytes:
    """对字体前 32 字节做 XOR 混淆（密钥 = GUID 派生的 16 字节，重复两次）。"""
    key = _guid_to_key(guid_str)
    data = bytearray(font_bytes)
    for i in range(min(32, len(data))):
        data[i] ^= key[i % 16]
    return bytes(data)


# ─── 主入口 ────────────────────────────────────────────────────────────────────

def embed_fonts(pptx_path: str, style: dict) -> None:
    """
    将 style["font_paths"] 中所有字体嵌入 PPTX 文件（原地修改）。

    style 结构示例（来自 config/styles/xiaomi-dark.json）：
        {
          "typography": {"display_font": "MiSans-Bold", ...},
          "font_paths": {"MiSans-Bold": "assets/fonts/MiSans-Bold.ttf", ...}
        }
    """
    font_paths: dict = style.get("font_paths", {})
    typography: dict = style.get("typography", {})

    # 收集本次 PPT 实际用到的字重（去重）
    used_fonts = set()
    for key in ("display_font", "heading1_font", "heading2_font",
                "body_font", "caption_font", "footnote_font"):
        name = typography.get(key)
        if name and name in font_paths:
            used_fonts.add(name)

    if not used_fonts:
        print("[font_embedder] 无需嵌入字体（font_paths 为空或无匹配）")
        return

    pptx_path = Path(pptx_path)
    root_dir = pptx_path.parent.parent  # MVP01 根目录（output 的上级）
    tmp_path = pptx_path.with_suffix(".tmp.pptx")

    with zipfile.ZipFile(pptx_path, "r") as zin, \
         zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:

        names = set(zin.namelist())

        # 读取需要修改的三个核心 XML
        ct_xml   = zin.read("[Content_Types].xml")
        rels_xml = zin.read("ppt/_rels/presentation.xml.rels")
        prs_xml  = zin.read("ppt/presentation.xml")

        ct_tree   = etree.fromstring(ct_xml)
        rels_tree = etree.fromstring(rels_xml)
        prs_tree  = etree.fromstring(prs_xml)

        # 计算当前最大 rId，后续新 ID 从此递增
        existing_ids = []
        for el in rels_tree:
            rid = el.get("Id", "")
            if rid.startswith("rId"):
                try:
                    existing_ids.append(int(rid[3:]))
                except ValueError:
                    pass
        next_id = max(existing_ids, default=0) + 1

        # 找到或创建 <p:embeddedFontLst>
        font_lst = prs_tree.find(f"{{{_NS_PML}}}embeddedFontLst")
        if font_lst is None:
            font_lst = etree.SubElement(prs_tree, f"{{{_NS_PML}}}embeddedFontLst")

        new_font_files: dict[str, bytes] = {}  # zip内路径 → 混淆字节

        for typeface, rel_ttf_path in font_paths.items():
            if typeface not in used_fonts:
                continue

            abs_ttf = root_dir / rel_ttf_path
            if not abs_ttf.exists():
                print(f"[font_embedder] 字体文件不存在，跳过：{abs_ttf}")
                continue

            ttf_bytes = abs_ttf.read_bytes()
            guid_str  = "{" + str(uuid.uuid4()).upper() + "}"
            obf_bytes = _obfuscate(ttf_bytes, guid_str)

            r_id          = f"rId{next_id}"
            font_filename = f"font{next_id}.fntdata"
            zip_font_path = f"ppt/fonts/{font_filename}"
            next_id += 1

            # 添加 Relationship
            rel_el = etree.SubElement(rels_tree, f"{{{_NS_PKG_REL}}}Relationship")
            rel_el.set("Id",     r_id)
            rel_el.set("Type",   _FONT_REL_TYPE)
            rel_el.set("Target", f"fonts/{font_filename}")

            # 添加 Content Type
            override_el = etree.SubElement(ct_tree, f"{{{_NS_PKG_CT}}}Override")
            override_el.set("PartName",    f"/ppt/fonts/{font_filename}")
            override_el.set("ContentType", _FONT_CT)

            # 添加 <p:embeddedFont> 声明
            font_el  = etree.SubElement(font_lst, f"{{{_NS_PML}}}embeddedFont")
            font_tag = etree.SubElement(font_el,  f"{{{_NS_PML}}}font")
            font_tag.set("typeface", typeface)
            font_tag.set("{%s}guid" % _NS_R, guid_str)
            reg_el = etree.SubElement(font_el, f"{{{_NS_PML}}}regular")
            reg_el.set(f"{{{_NS_R}}}id", r_id)

            new_font_files[zip_font_path] = obf_bytes
            print(f"[font_embedder] 嵌入字体：{typeface} → {font_filename}")

        # 写出所有原始文件（跳过需替换的三个 XML）
        skip = {"[Content_Types].xml",
                "ppt/_rels/presentation.xml.rels",
                "ppt/presentation.xml"}
        for name in zin.namelist():
            if name not in skip:
                zout.writestr(name, zin.read(name))

        xml_opts = dict(xml_declaration=True, encoding="UTF-8", standalone=True)
        zout.writestr("[Content_Types].xml",
                      etree.tostring(ct_tree,   **xml_opts))
        zout.writestr("ppt/_rels/presentation.xml.rels",
                      etree.tostring(rels_tree, **xml_opts))
        zout.writestr("ppt/presentation.xml",
                      etree.tostring(prs_tree,  **xml_opts))

        for zip_path, data in new_font_files.items():
            zout.writestr(zip_path, data)

    shutil.move(str(tmp_path), str(pptx_path))
    print(f"[font_embedder] 字体嵌入完成：{pptx_path.name}")
