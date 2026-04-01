"""
paths.py
统一的路径解析模块。
开发模式下使用项目目录结构，Electron 打包后通过环境变量指定可写目录。
"""
import os
from pathlib import Path

# 项目根目录（开发模式的 fallback）
# paths.py 在 backend/core/ 下，所以 parent.parent.parent 是项目根目录
_FALLBACK_ROOT = Path(__file__).resolve().parent.parent.parent

# 三个核心目录，优先使用环境变量
CONFIG_DIR = Path(os.environ.get("PRISM_CONFIG_DIR", _FALLBACK_ROOT / "config"))
DATA_DIR = Path(os.environ.get("PRISM_DATA_DIR", _FALLBACK_ROOT / "data"))
ASSETS_DIR = Path(os.environ.get("PRISM_ASSETS_DIR", _FALLBACK_ROOT / "assets"))

# 向后兼容：PROJECT_ROOT 仍指向原始包位置（PyInstaller 场景下为 _MEIPASS）
PROJECT_ROOT = _FALLBACK_ROOT
