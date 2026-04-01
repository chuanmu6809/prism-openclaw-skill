---
name: prism
description: Convert .docx documents to high-fidelity PPT presentations using AI. Supports multiple styles and both image-based and editable export modes.
metadata: {"openclaw":{"emoji":"🎯","requires":{"bins":["python3"],"env":[]},"install":[{"id":"setup","kind":"download","label":"Run Prism setup script","url":"https://github.com/chuanmu6809/prism-openclaw-skill/releases/latest/download/setup.sh"}]}}
---

# Prism — AI PPT Generator

Prism converts .docx files (exported from Feishu/Lark, Google Docs, or any word processor) into professional, high-fidelity PPT presentations using AI.

默认情况下，Prism 优先使用当前 OpenClaw 宿主会话的大模型来完成布局规划与单页 HTML 生成；用户也可以改为配置自己的 OpenAI-compatible API。

## Installation Directory

The Prism engine is installed at: `{baseDir}/..`

The CLI entry point is: `{baseDir}/../prism_cli.py`

## Configuration

Prism configuration is stored at `{baseDir}/prism_config.json`.

### First-time setup or `/prism config`

When the user invokes `/prism config` or when config file doesn't exist, guide them through configuration:

1. **API Configuration**: Ask the user:
   > Prism 需要模型来完成排版规划和 HTML 生成。请选择运行方式：
   > 1. **使用当前龙虾会话的大模型** — 推荐，无需额外配置
   > 2. **为 Prism 单独配置 API**（需要提供 base_url、api_key、model）

   - If option 1: set `api_mode` to `"host"`.
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
     "api_mode": "host",
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

Load config from `{baseDir}/prism_config.json`.

#### Mode A: `api_mode = "host"` (recommended)

Use the current OpenClaw conversation model for all reasoning steps, and only use Prism CLI for deterministic processing.

1. Create work directories:
```bash
mkdir -p "/tmp/prism_work_$(date +%s)"
mkdir -p "/tmp/prism_output_$(date +%s)"
```

2. Parse the DOCX:
```bash
python3 {baseDir}/../prism_cli.py \
  --parse-docx \
  --input "<path_to_docx>" \
  --work-dir "<work_dir>"
```

3. Emit the layout-planning prompt bundle:
```bash
python3 {baseDir}/../prism_cli.py \
  --emit-layout-plan \
  --work-dir "<work_dir>" \
  --style "<style_name>" \
  --prompt-output "<work_dir>/layout_prompt.json"
```

4. Read `<work_dir>/layout_prompt.json`, then use the **current host model** to answer it.
   - Pass `system_prompt` as system instruction and `user_prompt` as user content.
   - Save the model output as `<work_dir>/intents.json`.
   - The content must be a JSON array like:
   ```json
   [{"page":1,"layout_intent":"..."},{"page":2,"layout_intent":"..."}]
   ```
   - For mixed-tone styles, each object should also include `contrast_affinity`.

5. Generate each page HTML one by one:
   - For page `N`, emit a prompt bundle:
   ```bash
   python3 {baseDir}/../prism_cli.py \
     --emit-page-prompt \
     --work-dir "<work_dir>" \
     --style "<style_name>" \
     --page N \
     --intents-file "<work_dir>/intents.json" \
     --prompt-output "<work_dir>/page_N_prompt.json"
   ```
   - Read `page_N_prompt.json`, use the **current host model** to answer it, and save the output HTML to the `output_path` field shown in the JSON.
   - The model output must be exactly one `<section class="page-N">...</section>`.

6. Assemble the full HTML:
```bash
python3 {baseDir}/../prism_cli.py \
  --assemble-html \
  --work-dir "<work_dir>" \
  --style "<style_name>"
```

7. Export PPTX:
```bash
python3 {baseDir}/../prism_cli.py \
  --export-from-workdir \
  --work-dir "<work_dir>" \
  --style "<style_name>" \
  --export both \
  --output-dir "<output_dir>"
```

#### Mode B: `api_mode = "custom"`

Use Prism's built-in API client:
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

After sending the files, clean up the temp work/output directories:
```bash
rm -rf "<output_dir>" "<work_dir>"
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

If the check still reports `playwright_chromium` missing, run:
```bash
python3 -m playwright install chromium
```

On Linux servers, if Chromium installation fails due to missing system libraries, run:
```bash
python3 -m playwright install-deps chromium
python3 -m playwright install chromium
```

## Error Handling

- If the .docx has no H1 headings: Tell the user "文档中需要使用一级标题(H1)来分页，请检查文档格式"
- If host-mode generation fails: inspect whether the current conversation model can follow long structured prompts and output valid HTML/JSON; if needed, retry page-by-page
- If custom API call fails: Check if api_mode is correct, suggest running `/prism config`
- If Playwright/Chromium missing: Run `python3 -m playwright install chromium`
- If Python deps missing: Run `cd {baseDir}/.. && pip install -r requirements.txt`

## Notes

- The .docx should use H1 headings to separate slides. Each H1 becomes a new slide.
- Images in the document are automatically extracted and embedded in the PPT.
- The `[背景]` marker before an image makes it a full-page background.
- Horizontal rules (---) in the document create content blocks within a slide.
- Generation time depends on the number of slides and the LLM model speed.
