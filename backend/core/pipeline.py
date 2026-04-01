"""
pipeline.py
Phase 4A: 完整生成流水线 — 支持 SSE 进度推送、单页失败隔离、单页重试。
"""
from __future__ import annotations

import os
import json
import asyncio
import base64
import random
import traceback
from pathlib import Path
from typing import Dict, Any, Optional
from datetime import datetime

# Skill 模式：不依赖数据库，DATA_DIR 和 update_task 均为兼容 stub
import tempfile as _tempfile

# CLI/Skill 场景下任务目录由调用方管理，DATA_DIR 仅作兜底备用
DATA_DIR = os.environ.get("PRISM_DATA_DIR", os.path.join(os.path.expanduser("~"), ".prism", "tasks"))

async def update_task(task_id: str, data: dict):
    """Skill 模式下的空实现，原 Web 数据库写入在此静默忽略。"""
    # 将状态写入任务目录下的 meta.json，方便调试
    task_dir = os.path.join(DATA_DIR, task_id)
    if os.path.isdir(task_dir):
        meta_path = os.path.join(task_dir, "meta.json")
        try:
            meta = {}
            if os.path.exists(meta_path):
                with open(meta_path, "r", encoding="utf-8") as _f:
                    meta = json.load(_f)
            meta.update(data)
            with open(meta_path, "w", encoding="utf-8") as _f:
                json.dump(meta, _f, ensure_ascii=False, indent=2)
        except Exception:
            pass

from backend.core.paths import CONFIG_DIR

_ROOT = Path(__file__).parent.parent


# ── 进度状态管理 ─────────────────────────────────────────────────────────────────

class GenerationProgress:
    """Thread-safe progress tracker for a generation task."""
    
    def __init__(self, task_id: str, total_pages: int):
        self.task_id = task_id
        self.total_pages = total_pages
        self.page_states: Dict[int, Dict[str, Any]] = {}
        self.overall_phase = "initializing"
        self.cancelled = False
        self.cancelled_pages: set = set()
        self.created_at = datetime.utcnow().isoformat()
        self._events: asyncio.Queue = asyncio.Queue()
        
        # Initialize page states
        for i in range(total_pages):
            self.page_states[i] = {
                "status": "queued",
                "phase": "",
                "message": "等待中",
                "error": None,
                "started_at": None,
            }
    
    def emit(self, event: dict):
        """Push an event to the SSE queue."""
        event["timestamp"] = datetime.utcnow().isoformat()
        self._events.put_nowait(event)
    
    def update_page(self, page_idx: int, status: str, phase: str, message: str, error: str = None):
        prev = self.page_states.get(page_idx, {})
        started_at = prev.get("started_at")
        # 进入 in_progress 时记录开始时间
        if status == "in_progress" and not started_at:
            started_at = datetime.utcnow().isoformat()
        # 完成或失败时清除
        if status in ("completed", "failed"):
            started_at = None
        self.page_states[page_idx] = {
            "status": status,
            "phase": phase,
            "message": message,
            "error": error,
            "started_at": started_at,
        }
        completed = sum(1 for s in self.page_states.values() if s["status"] == "completed")
        failed = sum(1 for s in self.page_states.values() if s["status"] == "failed")
        percent = int((completed + failed) / self.total_pages * 100)
        
        self.emit({
            "type": "page_progress",
            "page_index": page_idx,
            "phase": phase,
            "status": status,
            "message": message,
            "error": error,
            "overall_percent": percent,
            "completed": completed,
            "failed": failed,
        })
    
    def cancel_page(self, page_idx: int):
        """Mark a page as cancelled by user."""
        self.cancelled_pages.add(page_idx)
        self.update_page(page_idx, "failed", "html_generation",
                        f"第 {page_idx+1} 页：已终止")
    
    def is_page_cancelled(self, page_idx: int) -> bool:
        return page_idx in self.cancelled_pages
    
    def finish(self, status: str = "completed"):
        self.emit({
            "type": "generation_complete",
            "status": status,
            "overall_percent": 100,
        })
    
    async def events(self):
        """Async generator for SSE events."""
        while True:
            try:
                event = await asyncio.wait_for(self._events.get(), timeout=30)
                yield event
                if event.get("type") == "generation_complete":
                    break
            except asyncio.TimeoutError:
                yield {"type": "heartbeat"}


# ── 全局进度注册表 ────────────────────────────────────────────────────────────────

_active_tasks: Dict[str, GenerationProgress] = {}
_llm_semaphore = asyncio.Semaphore(3)  # 最多 3 页同时调用 LLM


def get_progress(task_id: str) -> Optional[GenerationProgress]:
    return _active_tasks.get(task_id)


# ── 单页生成 ─────────────────────────────────────────────────────────────────────

async def _generate_with_semaphore(page, page_idx, intent, style, total_pages, task_dir, progress):
    """Acquire semaphore then generate. Keeps queued pages visible."""
    async with _llm_semaphore:
        return await _generate_single_page(page, page_idx, intent, style, total_pages, task_dir, progress)


async def _generate_single_page(
    page: dict,
    page_idx: int,
    layout_intent: str,
    style: dict,
    total_pages: int,
    task_dir: str,
    progress: GenerationProgress,
    max_retries: int = 2,
):
    """Run the full pipeline for a single page, with auto-retry."""
    from backend.core.prompt_builder import build_system_prompt, build_user_prompt
    from backend.core.llm_client import call_llm
    
    prompts_cfg = json.loads((CONFIG_DIR / "prompts.json").read_text(encoding="utf-8"))
    
    for attempt in range(1 + max_retries):
        try:
            label = f"第 {page_idx+1} 页"
            if attempt > 0:
                label += f"（第 {attempt+1} 次尝试）"
            
            progress.update_page(page_idx, "in_progress", "html_generation",
                               f"{label}：正在生成 HTML...")
            
            sys_prompt = build_system_prompt(prompts_cfg)
            user_prompt = build_user_prompt(page, layout_intent, style, total_pages)
            # 检查是否被用户取消
            if progress.is_page_cancelled(page_idx):
                return {"page_idx": page_idx, "success": False, "error": "cancelled"}
            
            section_html = await call_llm(sys_prompt, user_prompt)
            
            # LLM 返回后再检查一次
            if progress.is_page_cancelled(page_idx):
                return {"page_idx": page_idx, "success": False, "error": "cancelled"}
            
            # Save individual section
            sections_dir = os.path.join(task_dir, "sections")
            os.makedirs(sections_dir, exist_ok=True)
            section_path = os.path.join(sections_dir, f"page_{page_idx+1}.html")
            with open(section_path, "w", encoding="utf-8") as f:
                f.write(section_html)
            
            progress.update_page(page_idx, "completed", "html_generation",
                               f"第 {page_idx+1} 页：完成")
            
            return {"page_idx": page_idx, "section_html": section_html, "success": True}
            
        except Exception as e:
            traceback.print_exc()
            err_str = str(e)
            
            # 判断是否为 API 认证/额度错误（不可重试）
            is_auth_error = any(kw in err_str for kw in [
                "AuthenticationError", "401", "额度已用尽", "Unauthorized",
                "invalid_api_key", "insufficient_quota"
            ])
            
            if is_auth_error:
                user_msg = f"第 {page_idx+1} 页：API 认证失败（Key 无效或额度用尽）"
                progress.update_page(page_idx, "failed", "html_generation", user_msg, err_str)
                return {"page_idx": page_idx, "success": False, "error": user_msg}
            
            if attempt < max_retries:
                progress.update_page(page_idx, "in_progress", "html_generation",
                                   f"第 {page_idx+1} 页失败，{3}s 后重试...", err_str)
                await asyncio.sleep(3)
            else:
                progress.update_page(page_idx, "failed", "html_generation",
                                   f"第 {page_idx+1} 页最终失败", err_str)
                return {"page_idx": page_idx, "success": False, "error": err_str}


# ── 明暗混合：节奏引擎 ──────────────────────────────────────────────────────────

def _rhythm_engine(total_pages: int, style: dict, intents: list) -> list:
    """
    两阶段色调分配：AI 语义建议 + 确定性节奏兜底。

    Phase 1：读取 AI 返回的 contrast_affinity 分数（0-2）
    Phase 2：按窗口分段，每窗口内选最佳反差页；无候选时强制分配

    参数：
      style: 含 primary_tone ("dark"/"light"), mixed=True
      intents: AI 返回的 [{page, layout_intent, contrast_affinity}, ...]

    返回：["dark", "light", "dark", ...] 长度=total_pages
    """
    primary = style.get("primary_tone", "dark")
    contrast = "light" if primary == "dark" else "dark"

    # 所有页面先设为主色调
    tones = [primary] * total_pages

    if total_pages <= 2:
        # 2 页及以下不做混合
        return tones

    # 读取 AI 的 contrast_affinity 分数
    ca_map = {}
    for item in intents:
        page_num = item.get("page", 0)
        ca_map[page_num] = item.get("contrast_affinity", 0)

    # 构造每页的分数列表（0-indexed）
    # 注意：page_number 是 1-indexed
    scores = []
    for i in range(total_pages):
        page_num = i + 1
        score = ca_map.get(page_num, 0)
        scores.append(score)

    # 封面和结尾锁定为主色调（分数设为 -1 表示不可选）
    scores[0] = -1
    scores[-1] = -1

    # 中间页范围
    mid_start = 1
    mid_end = total_pages - 1  # exclusive
    mid_count = mid_end - mid_start

    if mid_count <= 0:
        return tones

    # 目标反差页数：约 25%，至少 1 页
    target_contrast = max(1, round(mid_count * 0.25))

    # 窗口大小：将中间页均匀分成 target_contrast 个窗口
    window_size = max(2, mid_count // target_contrast)

    # 按窗口分段，每窗口选一个反差页
    contrast_indices = []
    pos = mid_start
    while pos < mid_end:
        win_end = min(pos + window_size, mid_end)
        window = list(range(pos, win_end))

        if not window:
            break

        # 在窗口内找分数最高的页面
        best_idx = max(window, key=lambda i: scores[i])

        if scores[best_idx] > 0:
            # 有语义适合的页面
            contrast_indices.append(best_idx)
        else:
            # 无适合的 → 强制选窗口中间位置（保证节奏）
            mid = window[len(window) // 2]
            contrast_indices.append(mid)

        pos = win_end

    # 应用反差色调
    for idx in contrast_indices:
        tones[idx] = contrast

    # 约束：不允许连续超过 4 页相同色调（额外修正）
    consecutive = 1
    for i in range(1, total_pages):
        if tones[i] == tones[i - 1]:
            consecutive += 1
            if consecutive > 4 and i != total_pages - 1 and i != 0:
                tones[i] = contrast if tones[i] == primary else primary
                consecutive = 1
        else:
            consecutive = 1

    tone_summary = ", ".join(f"P{i+1}:{t[0].upper()}" for i, t in enumerate(tones))
    ca_summary = ", ".join(f"P{i+1}:CA{scores[i]}" for i in range(total_pages))
    print(f"  🎯 AI 反差分数: {ca_summary}")
    print(f"  🎨 最终色调分配: {tone_summary}")

    return tones


def _build_page_style(style: dict, tone: str) -> dict:
    """
    根据页面色调，构造带正确 colors 的 style dict。
    - "default"：直接返回原 style（非混合风格）
    - "dark"：用 colors_dark
    - "light"：用 colors_light
    """
    if tone == "default" or not style.get("mixed"):
        return style

    result = {**style}
    if tone == "dark" and "colors_dark" in style:
        result["colors"] = style["colors_dark"]
        result["name"] = style["name"] + "（暗色页）"
    elif tone == "light" and "colors_light" in style:
        result["colors"] = style["colors_light"]
        result["name"] = style["name"] + "（亮色页）"
    return result


# ── 主流水线 ─────────────────────────────────────────────────────────────────────

async def run_pipeline(task_id: str, style_name: str):
    """Run the full generation pipeline for a task."""
    task_dir = os.path.join(DATA_DIR, task_id)
    
    # Load pages data
    pages_path = os.path.join(task_dir, "pages.json")
    with open(pages_path, "r", encoding="utf-8") as f:
        pages_data = json.load(f)
    pages = pages_data.get("pages", [])
    total_pages = len(pages)
    
    # Load style
    style_path = CONFIG_DIR / "styles" / f"{style_name}.json"
    if not style_path.exists():
        raise FileNotFoundError(f"Style '{style_name}' not found")
    style = json.loads(style_path.read_text(encoding="utf-8"))
    
    # Create progress tracker
    progress = GenerationProgress(task_id, total_pages)
    _active_tasks[task_id] = progress
    
    try:
        await update_task(task_id, {"status": "generating", "style": style_name})
        
        # ── Phase 2: Content Profiling ──
        progress.emit({"type": "phase_start", "phase": "content_profiling", "message": "正在分析页面内容..."})
        
        from backend.core.content_profiler import generate_profile
        for i, page in enumerate(pages):
            profile = generate_profile(page, is_first=(i == 0), is_last=(i == total_pages - 1))
            page["content_profile"] = profile
        
        # ── Phase 2: Layout Planning ──
        progress.emit({"type": "phase_start", "phase": "layout_planning", "message": "AI 正在规划布局..."})
        
        from backend.core.layout_planner import plan_layout
        intents = await asyncio.to_thread(plan_layout, pages, style)
        intent_map = {item["page"]: item["layout_intent"] for item in intents}
        
        # 色调分配：节奏引擎（mixed 模式下基于 AI 的 contrast_affinity 分数）
        if style.get("mixed"):
            page_tones = _rhythm_engine(total_pages, style, intents)
        else:
            page_tones = ["default"] * total_pages
        
        # Save layout intents for later retry
        intents_path = os.path.join(task_dir, "intents.json")
        with open(intents_path, "w", encoding="utf-8") as f:
            json.dump(intent_map, f, ensure_ascii=False, indent=2)
        
        # Save tones assignment
        tones_path = os.path.join(task_dir, "page_tones.json")
        with open(tones_path, "w", encoding="utf-8") as f:
            json.dump(page_tones, f, ensure_ascii=False, indent=2)
        
        # Save profiled pages
        with open(pages_path, "w", encoding="utf-8") as f:
            json.dump(pages_data, f, ensure_ascii=False, indent=2)
        
        # ── Phase 3: HTML Generation (concurrent per page) ──
        progress.emit({"type": "phase_start", "phase": "html_generation", "message": "正在并行生成 HTML..."})
        
        tasks = []
        for i, page in enumerate(pages):
            if progress.cancelled:
                break
            intent = intent_map.get(page["page_number"], "文字居中排布，标题大字，正文适中，适量留白")
            page_style = _build_page_style(style, page_tones[i])
            task = _generate_with_semaphore(page, i, intent, page_style, total_pages, task_dir, progress)
            tasks.append(task)
        
        results = await asyncio.gather(*tasks)
        
        # Collect successful sections
        sections = [None] * total_pages
        for r in results:
            if r["success"]:
                sections[r["page_idx"]] = r["section_html"]
        
        # ── Phase 4: Assemble HTML ──
        successful_sections = [s for s in sections if s is not None]
        
        if not successful_sections:
            progress.finish("failed")
            await update_task(task_id, {"status": "failed"})
            return
        
        progress.emit({"type": "phase_start", "phase": "assembling", "message": "正在组装完整 HTML..."})
        
        # Load images as base64
        images_dir = os.path.join(task_dir, "images")
        images_b64 = {}
        if os.path.exists(images_dir):
            for f in os.listdir(images_dir):
                if f.endswith(('.jpg', '.jpeg', '.png')):
                    img_id = os.path.splitext(f)[0]
                    img_path = os.path.join(images_dir, f)
                    images_b64[img_id] = base64.b64encode(
                        open(img_path, "rb").read()
                    ).decode()
        
        from backend.core.html_builder import assemble_html
        full_html = assemble_html(successful_sections, images_b64, style)
        
        html_path = os.path.join(task_dir, "output.html")
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(full_html)
        
        # ── Done ──
        # HTML 组装完成即标记 completed（Playwright + PPTX 在用户点导出时按需执行）
        await update_task(task_id, {"status": "completed"})
        progress.finish("completed")

        # 生成缩略图（用 Playwright 截取第一页）
        try:
            from backend.server.routes import _generate_thumbnail
            thumb_path = os.path.join(task_dir, "thumbnail.png")
            await asyncio.to_thread(_generate_thumbnail, task_id, task_dir, thumb_path)
        except Exception:
            pass  # 缩略图失败不影响主流程（Skill 模式下 routes 模块不存在，跳过即可）

        # 清理源文件节省磁盘空间（内容已提取到 images/ 和 pages.json）
        source_docx = os.path.join(task_dir, "source.docx")
        if os.path.exists(source_docx):
            os.remove(source_docx)
        
    except Exception as e:
        traceback.print_exc()
        progress.emit({"type": "error", "message": str(e)})
        progress.finish("failed")
        await update_task(task_id, {"status": "failed"})
    finally:
        # Keep progress for a while so client can catch up
        await asyncio.sleep(5)
        _active_tasks.pop(task_id, None)


# ── 单页重试 ─────────────────────────────────────────────────────────────────────

async def retry_page(task_id: str, page_idx: int, keep_layout: bool = False):
    """Retry generation for a single page."""
    task_dir = os.path.join(DATA_DIR, task_id)
    pages_path = os.path.join(task_dir, "pages.json")
    
    with open(pages_path, "r", encoding="utf-8") as f:
        pages_data = json.load(f)
    pages = pages_data.get("pages", [])
    
    if page_idx < 0 or page_idx >= len(pages):
        raise ValueError(f"page_idx {page_idx} out of range")
    
    # Load style from task meta.json（Skill 模式下无数据库，直接读文件）
    meta_path = os.path.join(task_dir, "meta.json")
    style_name = "xiaomi-dark"
    if os.path.exists(meta_path):
        try:
            with open(meta_path, "r", encoding="utf-8") as _f:
                _meta = json.load(_f)
            style_name = _meta.get("style", "xiaomi-dark")
        except Exception:
            pass
    style_path = CONFIG_DIR / "styles" / f"{style_name}.json"
    style = json.loads(style_path.read_text(encoding="utf-8"))
    
    # 读取已保存的色调分配
    tones_path = os.path.join(task_dir, "page_tones.json")
    if os.path.exists(tones_path):
        with open(tones_path, "r", encoding="utf-8") as f:
            page_tones = json.load(f)
    else:
        page_tones = ["default"] * len(pages)
    
    page_style = _build_page_style(style, page_tones[page_idx] if page_idx < len(page_tones) else "default")
    
    page = pages[page_idx]
    
    # Load saved layout intent
    intents_path = os.path.join(task_dir, "intents.json")
    saved_intents = {}
    if os.path.exists(intents_path):
        with open(intents_path, "r", encoding="utf-8") as f:
            saved_intents = json.load(f)
    
    original_intent = saved_intents.get(str(page["page_number"]),
                         saved_intents.get(page["page_number"],
                           "文字居中排布，标题大字，正文适中，适量留白"))
    
    # 决定布局意图
    if keep_layout:
        intent = original_intent
    else:
        intent = (f"原始布局意图：{original_intent}\n"
                  f"请使用不同的布局方式重新设计此页。")
    
    # Create a temporary progress tracker if not exists
    progress = _active_tasks.get(task_id)
    if not progress:
        progress = GenerationProgress(task_id, len(pages))
        _active_tasks[task_id] = progress
    
    result = await _generate_single_page(
        page, page_idx, intent, page_style, len(pages), task_dir, progress
    )
    
    if result["success"]:
        progress.update_page(page_idx, "completed", "html_generation",
                           f"第 {page_idx+1} 页重试成功")
        # 同步更新 sections_backup（如果存在），防止换色时从旧备份恢复
        backup_dir = os.path.join(task_dir, "sections_backup")
        if os.path.exists(backup_dir):
            section_file = f"page_{page_idx+1}.html"
            src = os.path.join(task_dir, "sections", section_file)
            dst = os.path.join(backup_dir, section_file)
            if os.path.exists(src):
                import shutil
                shutil.copy2(src, dst)
    
    return result
