# 🎯 Prism — AI PPT Generator

将 `.docx` 文档转换为高保真 PPT，支持多种风格和导出模式。

> 这是一个 [OpenClaw（小龙虾）](https://openclaw.ai) Skill，安装后可在 OpenClaw 对话中通过 `/prism` 命令使用。

## ✨ 功能特点

- **AI 驱动**：基于大语言模型自动分析内容、规划布局、生成幻灯片
- **高保真还原**：支持图片版（100% 视觉还原）和可编辑版（文字/布局可编辑）两种导出
- **4 种风格**：深色 / 浅色 / 深色主导混合 / 浅色主导混合
- **智能排版**：自动识别内容类型（数据页、时间线、对比等），选择最佳布局
- **字体嵌入**：自动嵌入 MiSans 字体，确保跨平台一致显示

## 📦 安装

在 OpenClaw 中安装此 Skill：

```
/skill install github:chuanmu6809/prism-openclaw-skill
```

首次使用时运行 `/prism config` 完成配置。

## 🚀 使用方法

1. 在 OpenClaw 对话中发送 `.docx` 文件
2. 选择 PPT 风格（或使用默认风格）
3. 等待 AI 生成（约 3-8 分钟）
4. 收到两个版本的 PPTX 文件

### 可用命令

| 命令 | 说明 |
|------|------|
| `/prism` | 显示帮助 |
| `/prism config` | 配置 API 和默认风格 |
| `/prism config show` | 查看当前配置 |
| `/prism config reset` | 重置配置 |
| `/prism styles` | 列出可用风格 |
| `/prism deps` | 检查依赖状态 |

## 📝 文档格式要求

- 使用 **一级标题 (`# `)** 分页，每个 H1 = 一张幻灯片
- 使用 **二级标题 (`## `)** 作为副标题（可选）
- 在图片前写 `[背景]` 将其设为全屏背景图
- 使用分隔线 `---` 精确控制图文配对关系

## 🎨 可用风格

| 风格 ID | 说明 |
|---------|------|
| `xiaomi-dark` | 深蓝底 + 白字，科技感 |
| `xiaomi-light` | 白底 + 深色字，清爽简洁 |
| `xiaomi-dark-dominant` | 暗色为主，部分页面切换亮色 |
| `xiaomi-light-dominant` | 亮色为主，部分页面切换暗色 |

## 🔧 依赖

- Python 3.10+
- Playwright + Chromium
- LLM API（OpenAI 兼容协议）

依赖会在首次使用时通过 `setup.sh` 自动安装。

## 📁 项目结构

```
├── skill/                  # OpenClaw Skill 入口
│   ├── SKILL.md            # Skill 描述与使用指引
│   ├── setup.sh            # 依赖安装脚本
│   └── prism_config.json   # 默认配置模板
├── prism_cli.py            # CLI 入口
├── requirements.txt        # Python 依赖
├── backend/core/           # 核心生成逻辑
├── config/
│   ├── styles/             # 风格配置 (4 套)
│   └── prompts.json        # AI 提示词模板
└── assets/fonts/           # MiSans 字体 (10 个字重)
```

## 📄 License

MIT
