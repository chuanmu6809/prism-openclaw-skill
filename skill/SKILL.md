---
name: prism
description: Convert .docx documents to high-fidelity PPT presentations using AI. Supports multiple styles and both image-based and editable export modes.
metadata: {"openclaw":{"emoji":"🎯","requires":{"bins":["python3"],"env":[]},"primaryEnv":"OPENAI_API_KEY","install":[{"id":"setup","kind":"download","label":"Run Prism setup script","url":"https://github.com/prism-ppt/prism/releases/latest/download/setup.sh"}]}}
---

# Prism — AI PPT Generator

Prism converts .docx files (exported from Feishu/Lark, Google Docs, or any word processor) into professional, high-fidelity PPT presentations using AI.

## Installation Directory

The Prism engine is installed at: `{baseDir}/..`

The CLI entry point is: `{baseDir}/../prism_cli.py`

## Configuration

Prism configuration is stored at `{baseDir}/prism_config.json`.

### First-time setup or `/prism config`

When the user invokes `/prism config` or when config file doesn't exist, guide them through configuration:

1. **API Configuration**: Ask the user:
   > Prism 需要调用 LLM API 来生成 PPT 内容。请选择 API 配置方式：
   > 1. **使用小龙虾当前的 API**（读取环境变量中的 OPENAI_API_KEY）— 推荐，无需额外配置
   > 2. **为 Prism 单独配置 API**（需要提供 base_url、api_key、model）

   - If option 1: set `api_mode` to `"openclaw"` in config. No extra env vars needed.
   - If option 2: ask for `base_url`, `api_key`, and `model`, then save to config as `api_mode: "custom"`.

2. **Default Style**: Ask the user to pick a default style:
   > 请选择默认 PPT 风格：
   > 1. xiaomi-dark — 基础暗色（深蓝底 + 白字，科技感）
   > 2. xiaomi-light — 基础亮色（白底 + 深色字，清爽简洁）
   > 3. xiaomi-dark-dominant — 暗色主导混合（以暗色为主，部分页面切换亮色增强节奏感）
   > 4. xiaomi-light-dominant — 亮色主导混合（以亮色为主，部分页面切换暗色增强层次）

3. Save the config to `{baseDir}/prism_config.json`:
   ```json
   {
     "api_mode": "openclaw",
     "custom_api_key": "",
     "custom_base_url": "",
     "custom_model": "",
     "default_style": "xiaomi-dark",
     "default_export": "both"
   }
   ```

### `/prism config show`

Read and display the current config from `{baseDir}/prism_config.json`.

### `/prism config reset`

Delete `{baseDir}/prism_config.json` and tell the user config has been reset. Next run will trigger setup.

## Usage — Converting a Document

When the user sends a .docx file (as an attachment) or asks to convert a document to PPT:

### Step 1: Check config exists

Read `{baseDir}/prism_config.json`. If it doesn't exist, run the configuration flow first (see above).

### Step 2: Check dependencies

Run:
```bash
python3 {baseDir}/../prism_cli.py --check-deps
```

If any dependency is missing, run the setup:
```bash
cd {baseDir}/.. && bash skill/setup.sh
```

### Step 3: Ask for style preference (if no default or user wants to change)

If the user hasn't specified a style in their message, and there's a `default_style` in config, use that. Otherwise ask:

> 请选择 PPT 风格：
> 1. xiaomi-dark — 基础暗色
> 2. xiaomi-light — 基础亮色
> 3. xiaomi-dark-dominant — 暗色主导混合
> 4. xiaomi-light-dominant — 亮色主导混合
>
> （直接回复数字即可，或输入风格 ID）

### Step 4: Run the conversion

Load config from `{baseDir}/prism_config.json` and build the command:

**If api_mode is "openclaw"** (uses the environment's OPENAI_API_KEY):
```bash
python3 {baseDir}/../prism_cli.py \
  --input "<path_to_docx>" \
  --style "<style_name>" \
  --export both \
  --output-dir "/tmp/prism_output_$(date +%s)"
```

**If api_mode is "custom"**:
```bash
python3 {baseDir}/../prism_cli.py \
  --input "<path_to_docx>" \
  --style "<style_name>" \
  --export both \
  --output-dir "/tmp/prism_output_$(date +%s)" \
  --api-key "<custom_api_key>" \
  --base-url "<custom_base_url>" \
  --model "<custom_model>"
```

### Step 5: Report results

The generation takes 3-8 minutes depending on page count. Inform the user:

> 正在生成 PPT，预计需要 3-8 分钟，请稍候...

After completion, the output directory will contain:
- `presentation_image.pptx` — 图片版（100% 还原视觉效果，不可编辑文字）
- `presentation_editable.pptx` — 可编辑版（文字可编辑，布局可调整）

Send both files to the user with a brief explanation:
> PPT 生成完成！为你生成了两个版本：
> 📄 **图片版** — 完美还原视觉效果，适合直接演示
> ✏️ **可编辑版** — 文字和布局可编辑，适合二次修改

### Step 6: Cleanup

After sending the files, clean up the temp output directory:
```bash
rm -rf "<output_dir>"
```

## Tool Commands Summary

| Command | Description |
|---------|-------------|
| `/prism` | Show help and usage |
| `/prism config` | Configure API and default style |
| `/prism config show` | Show current configuration |
| `/prism config reset` | Reset configuration to defaults |
| `/prism styles` | List available PPT styles |
| `/prism deps` | Check and install dependencies |

### `/prism styles`

```bash
python3 {baseDir}/../prism_cli.py --list-styles
```

### `/prism deps`

```bash
python3 {baseDir}/../prism_cli.py --check-deps
```

If deps are missing:
```bash
cd {baseDir}/.. && bash skill/setup.sh
```

## Error Handling

- If the .docx has no H1 headings: Tell the user "文档中需要使用一级标题(H1)来分页，请检查文档格式"
- If API call fails: Check if api_mode is correct, suggest running `/prism config`
- If Playwright/Chromium missing: Run `python3 -m playwright install chromium`
- If Python deps missing: Run `cd {baseDir}/.. && pip install -r requirements.txt`

## Notes

- The .docx should use H1 headings to separate slides. Each H1 becomes a new slide.
- Images in the document are automatically extracted and embedded in the PPT.
- The `[背景]` marker before an image makes it a full-page background.
- Horizontal rules (---) in the document create content blocks within a slide.
- Generation time depends on the number of slides and the LLM model speed.
