---
name: glmocr-local-ollama
description: |
  Run GLM-OCR locally via Ollama for document parsing (OCR, tables, formulas, handwriting).
  No cloud API key needed. Uses uv to run inside the glm-ocr project directory.

  Trigger when: user wants OCR on images/PDFs AND has local Ollama running with glm-ocr model,
  OR user mentions "local OCR", "ollama ocr", "離線 OCR", "本地 OCR".

  NOT for: cloud/MaaS mode (use the `glmocr` skill instead), real-time camera feeds, audio.
metadata:
  openclaw:
    requires:
      bins:
        - uv
        - ollama
    emoji: "🔍"
---

# GLM-OCR Local Ollama Skill

Parse documents (images, PDFs) using GLM-OCR with a local Ollama server. No cloud API key required.

## Prerequisites

1. **Ollama running** with `glm-ocr` model pulled
2. **uv** installed
3. **glm-ocr repo** cloned with dependencies synced

## Preflight Check (MUST run before first parse)

```bash
# 1. Ollama alive?
curl -sf http://127.0.0.1:11434/api/tags | head -c 200

# 2. glm-ocr model available?
ollama list | grep glm-ocr

# 3. uv can import glmocr?
uv run python -c "import glmocr; print('OK')"
```

If any check fails, fix it before proceeding:

| Check failed | Fix |
|---|---|
| Ollama not running | `ollama serve` (or start Ollama app) |
| Model not found | `ollama pull glm-ocr:bf16` |
| uv import fails | `uv sync` in the glm-ocr project directory |

## Usage

### CLI (preferred for agents)

```bash
# Single image
uv run glmocr parse image.png --output ./results/

# Single image, print to stdout only (no files)
uv run glmocr parse image.png --stdout --no-save

# Directory of images
uv run glmocr parse ./scans/ --output ./results/

# JSON only to stdout (pipe-friendly)
uv run glmocr parse image.png --stdout --json-only --no-save

# With debug logging
uv run glmocr parse image.png --log-level DEBUG

# Skip layout visualization (faster)
uv run glmocr parse image.png --no-layout-vis --output ./results/
```

### Python API (for programmatic use)

```python
uv run python -c "
from glmocr import GlmOcr

with GlmOcr() as parser:
    result = parser.parse('image.png')
    print(result.markdown_result)
"
```

```python
uv run python -c "
from glmocr import GlmOcr
import json

with GlmOcr() as parser:
    result = parser.parse('document.pdf')
    for i, page in enumerate(result.json_result):
        print(f'Page {i}: {len(page)} regions')
    result.save(output_dir='./results/')
"
```

## How It Works

Three-stage async pipeline, all local:

```
Stage 1: PageLoader        PDF/image -> per-page PIL images
Stage 2: PP-DocLayoutV3    Object detection -> region bounding boxes
Stage 3: Ollama OCR        Each cropped region -> text/table/formula via glm-ocr VLM
```

- Stage 2 (layout) runs PP-DocLayoutV3 locally (auto-downloads from HuggingFace on first run)
- Stage 3 sends cropped regions to Ollama in parallel (default 2 workers)
- All three stages run concurrently via thread queues

## Config

The project's `glmocr/config.yaml` is pre-configured for Ollama:

```yaml
pipeline:
  maas:
    enabled: false            # local mode, not cloud
  ocr_api:
    api_host: 127.0.0.1
    api_port: 11434           # Ollama default port
    api_path: /api/generate
    api_mode: ollama_generate # Ollama native endpoint
    model: glm-ocr:bf16
    connect_timeout: 600
    request_timeout: 600
```

Override any config value via `--set`:

```bash
# Use a different model name
uv run glmocr parse image.png --set pipeline.ocr_api.model glm-ocr:latest

# Use a different port
uv run glmocr parse image.png --set pipeline.ocr_api.api_port 11435

# Force layout model to CPU (keep GPU for OCR)
uv run glmocr parse image.png --layout-device cpu

# Increase parallel workers
uv run glmocr parse image.png --set pipeline.max_workers 4
```

## Output

### PipelineResult fields

| Field | Type | Description |
|---|---|---|
| `markdown_result` | `str` | Full document as Markdown (tables as HTML, formulas as LaTeX) |
| `json_result` | `list[list[dict]]` | Per-page list of regions with label, content, bbox |
| `original_images` | `list[str]` | Input file paths |

### json_result region schema

```json
{
  "index": 0,
  "label": "text",
  "content": "recognized text content...",
  "bbox_2d": [x1, y1, x2, y2]
}
```

- `bbox_2d`: normalised 0-1000 coordinates
- `label`: `text`, `table`, `formula`, `image`, `seal`, `doc_title`, etc.

### Saved output structure

```
results/
  <stem>/
    <stem>.json       # structured regions
    <stem>.md         # full Markdown
    imgs/             # cropped figure images
    layout_vis/       # layout detection overlay
```

## Supported Inputs

`.jpg`, `.jpeg`, `.png`, `.bmp`, `.gif`, `.webp`, `.pdf`

## Performance Reference (Apple Silicon)

| Input | Time | Regions |
|---|---|---|
| Single image (code) | ~3s | 5-10 |
| Single image (paper with formulas) | ~12s | ~30 |
| 41-page PDF | ~223s | ~350 |

## Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| `TimeoutError: Failed to connect` | Ollama not running | `ollama serve` |
| `OSError: Repo id must be in the form` | Layout model path wrong | Set `layout.model_dir` to `PaddlePaddle/PP-DocLayoutV3_safetensors` |
| Slow first run | Layout model downloading from HuggingFace | Wait; cached after first download |
| `CUDA out of memory` | GPU overloaded | Use `--layout-device cpu` or reduce `--set pipeline.max_workers 1` |
| Ollama returns empty | Model not loaded | `ollama pull glm-ocr:bf16` |
