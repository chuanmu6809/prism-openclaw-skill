#!/usr/bin/env python3
"""
prism_cli.py
Prism CLI — 独立命令行入口，用于 OpenClaw Skill 调用。
将 .docx 转换为高质量 PPT，无需 Web 服务/数据库。

用法：
  python3 prism_cli.py --input file.docx --style xiaomi-dark --output-dir ./output
  python3 prism_cli.py --input file.docx --style xiaomi-light --api-key sk-xxx --base-url https://api.openai.com/v1
  python3 prism_cli.py --list-styles
  python3 prism_cli.py --check-deps
"""
from __future__ import annotations

import os
import sys
import json
import asyncio
import argparse
import base64
import shutil
import tempfile
import traceback
from pathlib import Path
from datetime import datetime

# 项目根目录
PROJECT_ROOT = Path(__file__).resolve().parent

# 确保项目根目录在 sys.path 中
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))


# ─── 工具函数 ──────────────────────────────────────────────────────────────────

def log(msg: str, level: str = "info"):
    """带时间戳的日志输出"""
    ts = datetime.now().strftime("%H:%M:%S")
    prefix = {"info": "[*]", "ok": "[+]", "err": "[!]", "warn": "[~]"}.get(level, "[*]")
    print(f"  {prefix} {ts} {msg}")


def list_styles() -> list[dict]:
    """列出所有可用风格"""
    from backend.core.paths import CONFIG_DIR
    styles_dir = CONFIG_DIR / "styles"
    styles = []
    for f in sorted(styles_dir.glob("*.json")):
        data = json.loads(f.read_text(encoding="utf-8"))
        styles.append({
            "id": f.stem,
            "name": data.get("name", f.stem),
            "mixed": data.get("mixed", False),
        })
    return styles


def check_dependencies() -> dict:
    """检查所有必要依赖"""
    results = {}

    # Python 包
    for pkg in ["docx", "pptx", "PIL", "openai", "lxml", "playwright"]:
        try:
            __import__(pkg)
            results[pkg] = True
        except ImportError:
            results[pkg] = False

    # Playwright 浏览器
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            browser = p.chromium.launch()
            browser.close()
        results["playwright_chromium"] = True
    except Exception:
        results["playwright_chromium"] = False

    # 字体文件
    from backend.core.paths import ASSETS_DIR
    font_dir = ASSETS_DIR / "fonts"
    results["fonts"] = font_dir.exists() and any(font_dir.glob("*.ttf"))

    return results


# ─── 核心 Pipeline（无数据库依赖）──────────────────────────────────────────────

async def run_cli_pipeline(
    docx_path: str,
    style_name: str,
    output_dir: str,
    export_modes: list[str] = None,
):
    """
    完整的 CLI pipeline：解析 DOCX → AI 生成 HTML → 导出 PPTX。
    不依赖 database.py / routes.py。
    """
    if export_modes is None:
        export_modes = ["image", "editable"]

    from backend.core.paths import CONFIG_DIR

    os.makedirs(output_dir, exist_ok=True)

    # ── 准备工作目录 ──
    work_dir = tempfile.mkdtemp(prefix="prism_")
    log(f"工作目录: {work_dir}")

    try:
        # ── Phase 1: 解析 DOCX ──
        log("Phase 1: 解析 DOCX 文件...")
        from backend.core.docx_parser import parse_docx
        pages_data = parse_docx(docx_path, work_dir)
        pages = pages_data.get("pages", [])
        total_pages = len(pages)

        if total_pages == 0:
            log("DOCX 中未找到任何页面（需要 H1 标题分页）", "err")
            return False

        log(f"解析完成：共 {total_pages} 页", "ok")

        # 保存 pages.json（供后续步骤使用）
        pages_path = os.path.join(work_dir, "pages.json")
        with open(pages_path, "w", encoding="utf-8") as f:
            json.dump(pages_data, f, ensure_ascii=False, indent=2)

        # ── 加载风格 ──
        style_path = CONFIG_DIR / "styles" / f"{style_name}.json"
        if not style_path.exists():
            available = [s["id"] for s in list_styles()]
            log(f"风格 '{style_name}' 不存在。可用风格: {', '.join(available)}", "err")
            return False
        style = json.loads(style_path.read_text(encoding="utf-8"))
        log(f"使用风格: {style.get('name', style_name)}")

        # ── Phase 2: 内容分析 + 布局规划 ──
        log("Phase 2: 分析内容...")
        from backend.core.content_profiler import generate_profile
        for i, page in enumerate(pages):
            profile = generate_profile(page, is_first=(i == 0), is_last=(i == total_pages - 1))
            page["content_profile"] = profile

        log("Phase 2: AI 规划布局...")
        from backend.core.layout_planner import plan_layout
        intents = await asyncio.to_thread(plan_layout, pages, style)
        intent_map = {item["page"]: item["layout_intent"] for item in intents}
        log("布局规划完成", "ok")

        # 色调分配（混合模式）
        from backend.core.pipeline import _rhythm_engine, _build_page_style
        if style.get("mixed"):
            page_tones = _rhythm_engine(total_pages, style, intents)
        else:
            page_tones = ["default"] * total_pages

        # ── Phase 3: 逐页生成 HTML ──
        log(f"Phase 3: 并发生成 {total_pages} 页 HTML...")
        from backend.core.prompt_builder import build_system_prompt, build_user_prompt
        from backend.core.llm_client import call_llm

        prompts_cfg = json.loads((CONFIG_DIR / "prompts.json").read_text(encoding="utf-8"))
        semaphore = asyncio.Semaphore(3)

        async def gen_page(idx: int) -> dict:
            async with semaphore:
                page = pages[idx]
                intent = intent_map.get(page["page_number"],
                                        "文字居中排布，标题大字，正文适中，适量留白")
                page_style = _build_page_style(style, page_tones[idx])
                sys_prompt = build_system_prompt(prompts_cfg)
                user_prompt = build_user_prompt(page, intent, page_style, total_pages)

                for attempt in range(3):
                    try:
                        log(f"  第 {idx+1}/{total_pages} 页生成中...")
                        html = await call_llm(sys_prompt, user_prompt)
                        # 保存单页 HTML
                        sections_dir = os.path.join(work_dir, "sections")
                        os.makedirs(sections_dir, exist_ok=True)
                        with open(os.path.join(sections_dir, f"page_{idx+1}.html"), "w", encoding="utf-8") as f:
                            f.write(html)
                        log(f"  第 {idx+1}/{total_pages} 页完成", "ok")
                        return {"idx": idx, "html": html, "ok": True}
                    except Exception as e:
                        if attempt < 2:
                            log(f"  第 {idx+1} 页失败，重试... ({e})", "warn")
                            await asyncio.sleep(3)
                        else:
                            log(f"  第 {idx+1} 页最终失败: {e}", "err")
                            return {"idx": idx, "html": None, "ok": False}

        results = await asyncio.gather(*(gen_page(i) for i in range(total_pages)))

        sections = [None] * total_pages
        ok_count = 0
        for r in results:
            if r["ok"]:
                sections[r["idx"]] = r["html"]
                ok_count += 1

        if ok_count == 0:
            log("所有页面生成失败", "err")
            return False

        log(f"HTML 生成完成: {ok_count}/{total_pages} 页成功", "ok")

        # ── Phase 4: 组装完整 HTML ──
        log("Phase 4: 组装完整 HTML...")
        images_dir = os.path.join(work_dir, "images")
        images_b64 = {}
        if os.path.exists(images_dir):
            for f in os.listdir(images_dir):
                if f.endswith(('.jpg', '.jpeg', '.png')):
                    img_id = os.path.splitext(f)[0]
                    images_b64[img_id] = base64.b64encode(
                        open(os.path.join(images_dir, f), "rb").read()
                    ).decode()

        from backend.core.html_builder import assemble_html
        successful_sections = [s for s in sections if s is not None]
        full_html = assemble_html(successful_sections, images_b64, style)

        html_path = os.path.join(work_dir, "output.html")
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(full_html)
        log("HTML 组装完成", "ok")

        # ── Phase 5: 导出 PPTX ──
        output_files = []

        if "image" in export_modes:
            log("Phase 5: 导出图片版 PPTX...")
            pptx_image_path = os.path.join(output_dir, "presentation_image.pptx")
            try:
                await asyncio.to_thread(_build_image_pptx_standalone, html_path, pptx_image_path)
                output_files.append(pptx_image_path)
                log(f"图片版导出完成: {pptx_image_path}", "ok")
            except Exception as e:
                log(f"图片版导出失败: {e}", "err")
                traceback.print_exc()

        if "editable" in export_modes:
            log("Phase 5: 导出可编辑版 PPTX...")
            pptx_edit_path = os.path.join(output_dir, "presentation_editable.pptx")
            try:
                from backend.core.layout_extractor import extract_layout
                layout_data = await asyncio.to_thread(extract_layout, html_path)

                from backend.core.pptx_builder import build_pptx
                await asyncio.to_thread(
                    build_pptx, layout_data, style, pptx_edit_path,
                    os.path.join(work_dir, "images")
                )
                output_files.append(pptx_edit_path)
                log(f"可编辑版导出完成: {pptx_edit_path}", "ok")
            except Exception as e:
                log(f"可编辑版导出失败: {e}", "err")
                traceback.print_exc()

        if output_files:
            log(f"全部完成！输出文件:", "ok")
            for f in output_files:
                log(f"  -> {f}")
            return True
        else:
            log("所有导出均失败", "err")
            return False

    finally:
        # 清理工作目录
        shutil.rmtree(work_dir, ignore_errors=True)


# ─── 独立的图片版 PPTX 构建（不依赖 routes.py）──────────────────────────────────

def _build_image_pptx_standalone(html_path: str, pptx_path: str):
    """用 Playwright 截图每个 section，构建图片版 PPTX。"""
    from playwright.sync_api import sync_playwright
    from pptx import Presentation
    from pptx.util import Emu

    screenshots = []
    tmp_dir = tempfile.mkdtemp()

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch()
            page = browser.new_page(viewport={"width": 1280, "height": 720})
            page.goto(f"file://{os.path.abspath(html_path)}")
            page.wait_for_load_state("networkidle")

            page_sections = page.query_selector_all("section")
            total = len(page_sections)
            log(f"  截图: 共 {total} 页")

            for i, section in enumerate(page_sections):
                img_path = os.path.join(tmp_dir, f"slide_{i+1:02d}.png")
                section.screenshot(path=img_path)
                screenshots.append(img_path)
                log(f"  截图: {i+1}/{total}")

            browser.close()

        # 构建 PPTX
        prs = Presentation()
        prs.slide_width = Emu(12192000)   # 13.333"
        prs.slide_height = Emu(6858000)   # 7.5"
        blank_layout = prs.slide_layouts[6]

        for img_path in screenshots:
            slide = prs.slides.add_slide(blank_layout)
            slide.shapes.add_picture(
                img_path,
                left=Emu(0), top=Emu(0),
                width=prs.slide_width, height=prs.slide_height
            )

        prs.save(pptx_path)
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


# ─── CLI 入口 ──────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Prism CLI — AI-powered DOCX to PPT converter",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  # 基本用法（使用环境变量中的 API key）
  python3 prism_cli.py --input slides.docx --output-dir ./output

  # 指定风格
  python3 prism_cli.py --input slides.docx --style xiaomi-light --output-dir ./output

  # 使用自定义 API
  python3 prism_cli.py --input slides.docx --api-key sk-xxx --base-url https://api.openai.com/v1 --model gpt-4o

  # 仅导出图片版
  python3 prism_cli.py --input slides.docx --export image

  # 查看可用风格
  python3 prism_cli.py --list-styles

  # 检查依赖
  python3 prism_cli.py --check-deps
""")

    parser.add_argument("--input", "-i", help="输入 .docx 文件路径")
    parser.add_argument("--output-dir", "-o", default="./output", help="输出目录 (默认: ./output)")
    parser.add_argument("--style", "-s", default="xiaomi-dark", help="风格名称 (默认: xiaomi-dark)")
    parser.add_argument("--export", "-e", default="both",
                        choices=["image", "editable", "both"],
                        help="导出模式 (默认: both)")

    # API 配置（可选，覆盖环境变量和配置文件）
    api_group = parser.add_argument_group("API 配置（可选，覆盖默认配置）")
    api_group.add_argument("--api-key", help="API Key")
    api_group.add_argument("--base-url", help="API Base URL")
    api_group.add_argument("--model", help="模型名称")
    api_group.add_argument("--timeout", type=int, help="超时时间（秒）")

    # 工具命令
    parser.add_argument("--list-styles", action="store_true", help="列出所有可用风格")
    parser.add_argument("--check-deps", action="store_true", help="检查依赖安装情况")

    args = parser.parse_args()

    # ── 工具命令 ──
    if args.list_styles:
        styles = list_styles()
        print("\n可用风格:")
        for s in styles:
            mixed_tag = " [混合明暗]" if s["mixed"] else ""
            print(f"  {s['id']:30s} {s['name']}{mixed_tag}")
        return

    if args.check_deps:
        print("\n依赖检查:")
        deps = check_dependencies()
        all_ok = True
        for name, ok in deps.items():
            status = "OK" if ok else "MISSING"
            if not ok:
                all_ok = False
            print(f"  {name:25s} {status}")
        if not all_ok:
            print("\n部分依赖缺失，请运行 setup.sh 安装")
            sys.exit(1)
        else:
            print("\n所有依赖已就绪")
        return

    # ── 主流程 ──
    if not args.input:
        parser.error("需要指定 --input 参数")

    if not os.path.exists(args.input):
        print(f"错误: 文件不存在: {args.input}")
        sys.exit(1)

    if not args.input.lower().endswith('.docx'):
        print(f"错误: 请提供 .docx 文件")
        sys.exit(1)

    # 设置 API 配置覆盖
    if args.api_key or args.base_url or args.model:
        from backend.core.llm_client import set_runtime_api_config
        set_runtime_api_config(
            api_key=args.api_key,
            base_url=args.base_url,
            model=args.model,
            timeout=args.timeout,
        )

    # 确定导出模式
    export_modes = {
        "image": ["image"],
        "editable": ["editable"],
        "both": ["image", "editable"],
    }[args.export]

    print(f"\n{'='*60}")
    print(f"  Prism CLI — DOCX to PPT")
    print(f"{'='*60}")
    print(f"  输入: {args.input}")
    print(f"  风格: {args.style}")
    print(f"  导出: {', '.join(export_modes)}")
    print(f"  输出: {os.path.abspath(args.output_dir)}")
    print(f"{'='*60}\n")

    success = asyncio.run(
        run_cli_pipeline(
            docx_path=args.input,
            style_name=args.style,
            output_dir=args.output_dir,
            export_modes=export_modes,
        )
    )

    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
