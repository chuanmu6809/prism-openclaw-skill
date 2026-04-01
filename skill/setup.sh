#!/usr/bin/env bash
#
# setup.sh — Prism Skill 依赖安装脚本
# 在 OpenClaw workspace 环境中自动安装所有依赖
#
set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"

echo "============================================"
echo "  Prism — 依赖安装"
echo "============================================"
echo "  项目目录: $PROJECT_DIR"
echo ""

# ─── 1. 检查 Python ──
echo "[1/4] 检查 Python..."
if ! command -v python3 &> /dev/null; then
    echo "  错误: python3 未安装"
    echo "  请先安装 Python 3.10+: brew install python3"
    exit 1
fi
PYTHON_VERSION=$(python3 --version 2>&1)
echo "  $PYTHON_VERSION"

# ─── 2. 安装 Python 依赖 ──
echo ""
echo "[2/4] 安装 Python 依赖..."

# Skill 模式下只需核心依赖，不需要 FastAPI/uvicorn 等 Web 层
CORE_DEPS=(
    "python-docx>=1.0.0"
    "python-pptx>=1.0.0"
    "Pillow>=10.0.0"
    "playwright>=1.40.0"
    "lxml>=4.9.0"
    "openai>=1.0.0"
)

pip3 install --quiet "${CORE_DEPS[@]}"
echo "  Python 依赖安装完成"

# ─── 3. 安装 Playwright + Chromium ──
echo ""
echo "[3/4] 安装 Playwright Chromium 浏览器..."
python3 -m playwright install chromium
echo "  Playwright Chromium 安装完成"

# ─── 4. 检查字体文件 ──
echo ""
echo "[4/4] 检查字体文件..."
FONTS_DIR="$PROJECT_DIR/assets/fonts"
if [ -d "$FONTS_DIR" ] && [ "$(ls -1 "$FONTS_DIR"/*.ttf 2>/dev/null | wc -l)" -gt 0 ]; then
    FONT_COUNT=$(ls -1 "$FONTS_DIR"/*.ttf | wc -l)
    echo "  字体目录正常: $FONT_COUNT 个字体文件"
else
    echo "  警告: 字体目录 $FONTS_DIR 不存在或为空"
    echo "  PPT 中将使用系统默认字体"
fi

# ─── 验证 ──
echo ""
echo "============================================"
echo "  安装完成！验证依赖..."
echo "============================================"
python3 "$PROJECT_DIR/prism_cli.py" --check-deps

echo ""
echo "安装完成！你可以通过 /prism 命令使用了。"
