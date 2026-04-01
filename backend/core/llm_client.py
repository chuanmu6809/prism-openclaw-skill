from __future__ import annotations

"""
llm_client.py
异步 LLM 调用，支持多页并发。使用 OpenAI 兼容协议。
支持多 API 配置切换。

API 配置优先级（从高到低）：
1. 运行时注入的覆盖配置（CLI --api-key / --base-url / --model）
2. 环境变量 PRISM_API_KEY / PRISM_BASE_URL / PRISM_MODEL
3. 环境变量 OPENAI_API_KEY（兼容 OpenClaw 注入的 key）
4. 配置文件 api_profiles.json / api.json
"""
import asyncio
import json
from pathlib import Path

import os
from openai import AsyncOpenAI

_ROOT = Path(__file__).parent.parent

# 优先使用环境变量指定的可写配置目录（Electron 打包后设置）
# 如果没有设置，回退到项目内的 config 目录（开发模式）
_CONFIG_DIR = Path(os.environ.get("PRISM_CONFIG_DIR", _ROOT / "config"))
_PROFILES_PATH = _CONFIG_DIR / "api_profiles.json"
_LEGACY_PATH = _CONFIG_DIR / "api.json"

# 如果可写目录下没有配置文件，从只读模板复制一份
_BUNDLED_CONFIG = _ROOT / "config"
if not _PROFILES_PATH.exists() and not _LEGACY_PATH.exists():
    _CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    for src in [_BUNDLED_CONFIG / "api_profiles.json", _BUNDLED_CONFIG / "api.json"]:
        if src.exists():
            import shutil
            shutil.copy2(src, _CONFIG_DIR / src.name)

# ─── 运行时覆盖（供 CLI 模式使用）─────────────────────────────────────────────

_runtime_override: dict | None = None


def set_runtime_api_config(api_key: str = None, base_url: str = None,
                           model: str = None, timeout: int = None,
                           max_tokens: int = None):
    """设置运行时 API 配置覆盖（CLI 模式调用）。设置后优先级最高。"""
    global _runtime_override, _async_client, _current_key
    _runtime_override = {}
    if api_key:
        _runtime_override["api_key"] = api_key
    if base_url:
        _runtime_override["base_url"] = base_url
    if model:
        _runtime_override["model"] = model
    if timeout:
        _runtime_override["timeout"] = timeout
    if max_tokens:
        _runtime_override["max_tokens"] = max_tokens
    # 清除客户端缓存，下次调用时用新配置重建
    _async_client = None
    _current_key = None


# ─── 配置加载 ─────────────────────────────────────────────────────────────────

def _load_profiles():
    """加载 API profiles 配置"""
    if _PROFILES_PATH.exists():
        return json.loads(_PROFILES_PATH.read_text(encoding="utf-8"))
    # 兼容旧 api.json
    if _LEGACY_PATH.exists():
        old = json.loads(_LEGACY_PATH.read_text(encoding="utf-8"))
        return {"active": 0, "profiles": [{**old, "name": "默认"}]}
    raise FileNotFoundError("No API config found")


def _save_profiles(data):
    _PROFILES_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def _get_env_config() -> dict | None:
    """尝试从环境变量构建 API 配置。返回 None 表示环境变量不足。"""
    api_key = os.environ.get("PRISM_API_KEY") or os.environ.get("OPENAI_API_KEY")
    if not api_key:
        return None
    return {
        "name": "环境变量",
        "api_key": api_key,
        "base_url": os.environ.get("PRISM_BASE_URL", "https://api.openai.com/v1"),
        "model": os.environ.get("PRISM_MODEL", "gpt-4o"),
        "timeout": int(os.environ.get("PRISM_TIMEOUT", "120")),
        "max_tokens": int(os.environ.get("PRISM_MAX_TOKENS", "4096")),
    }


def get_active_config() -> dict:
    """
    获取当前激活的 API 配置。
    优先级：运行时覆盖 > 环境变量 > 配置文件
    """
    # 1. 运行时覆盖（CLI 传入的参数）
    if _runtime_override:
        # 以环境变量或配置文件为基础，用运行时覆盖的字段覆盖
        base = _get_env_config() or _load_file_config()
        base.update({k: v for k, v in _runtime_override.items() if v is not None})
        return base

    # 2. 环境变量
    env_cfg = _get_env_config()
    if env_cfg:
        return env_cfg

    # 3. 配置文件
    return _load_file_config()


def _load_file_config() -> dict:
    """从配置文件加载激活的 API 配置"""
    data = _load_profiles()
    idx = data.get("active", 0)
    profiles = data.get("profiles", [])
    if not profiles:
        raise ValueError("No API profiles configured")
    return profiles[min(idx, len(profiles) - 1)]


def get_all_profiles() -> dict:
    """获取所有配置（供 API 用）"""
    return _load_profiles()


def save_all_profiles(data: dict):
    """保存所有配置"""
    _save_profiles(data)


# ─── 客户端管理 ────────────────────────────────────────────────────────────────

_async_client = None
_current_key = None


def _get_client() -> tuple:
    """获取或创建 AsyncOpenAI 客户端（配置变更时自动重建）"""
    global _async_client, _current_key
    cfg = get_active_config()
    key = (cfg["base_url"], cfg["api_key"])
    if _async_client is None or key != _current_key:
        _async_client = AsyncOpenAI(base_url=cfg["base_url"], api_key=cfg["api_key"])
        _current_key = key
    return _async_client, cfg


async def call_llm(system_prompt: str, user_prompt: str, retries: int = 2) -> str:
    """单次异步 LLM 调用，超时自动重试。"""
    client, cfg = _get_client()
    timeout = cfg.get("timeout", 120)
    for attempt in range(retries + 1):
        try:
            resp = await client.chat.completions.create(
                model=cfg["model"],
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt},
                ],
                max_tokens=cfg.get("max_tokens", 4096),
                timeout=timeout,
            )
            return resp.choices[0].message.content.strip()
        except Exception as e:
            if attempt < retries:
                print(f"  [重试 {attempt + 1}/{retries}] {type(e).__name__}")
            else:
                raise


async def call_all_pages(prompt_pairs: list, max_concurrent: int = 3) -> list:
    """
    并发调用所有页，最多同时 max_concurrent 个请求。
    prompt_pairs: [(system_prompt, user_prompt), ...]
    返回：[response_text, ...]，顺序与输入一致。
    """
    semaphore = asyncio.Semaphore(max_concurrent)

    async def call_with_limit(idx: int, system_p: str, user_p: str) -> tuple:
        async with semaphore:
            print(f"  → 第{idx + 1}页 开始生成...")
            result = await call_llm(system_p, user_p)
            print(f"  ✓ 第{idx + 1}页 完成")
            return idx, result

    tasks = [call_with_limit(i, s, u) for i, (s, u) in enumerate(prompt_pairs)]
    results = await asyncio.gather(*tasks)

    # 按原始顺序排列
    ordered = [None] * len(prompt_pairs)
    for idx, text in results:
        ordered[idx] = text
    return ordered
