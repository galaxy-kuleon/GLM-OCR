# AGENTS.md

This project provides two AI Agent Skills that form a document processing pipeline:
**PDF → editable DOCX → translated DOCX**.

---

## Skill 1: `/pdf-to-docx`

Converts a PDF into a high-fidelity editable DOCX.

```
/pdf-to-docx <pdf-path> [--output <dir>]
```

### Pipeline

1. **Environment check** — validates `pdftocairo`, `pdfinfo`, `pdffonts`, `pdftotext`, `uv`
2. **PDF → PNGs** — renders reference images at 200 DPI via `pdftocairo`
3. **Metadata extraction** — page dimensions, fonts, full text
4. **OCR** — `uv run glmocr parse` produces structured JSON + markdown + cropped images
5. **Style extraction** — VLM (Poe AI) infers font sizes, colors, alignment per region
6. **XML DSL generation** — fixed script converts OCR + styles into per-page XML files
7. **VLM review** — compares XML DSL against original images; agent fixes issues
8. **DSL → DOCX** — deterministic assembly via fixed `dsl_to_docx.py`
9. **Visual verification** (optional) — renders DOCX back to images for VLM comparison

### Key Details

- All scripts are in `.claude/skills/pdf-to-docx/scripts/` — no dynamic code generation
- Workspace: `<output>/<pdf-stem>-docx-workspace/`
- XML DSL in `dsl/page-{N}.xml` is the agent-editable intermediate format
- OCR `bbox_2d` values are normalized 0–1000, not pixels
- Cropped image index is a sequential counter per page, not the region index
- `POE_API_KEY` env var enables style extraction + VLM review; pipeline works without it (falls back to defaults)

---

## Skill 2: `/docx-translate-to-docx`

Translates an existing pdf-to-docx workspace into a target language, preserving all layout and formatting.

```
/docx-translate-to-docx <workspace-path> --lang <target-language> [--style <style-notes>]
```

**Examples:**
```
/docx-translate-to-docx output/report-docx-workspace --lang English
/docx-translate-to-docx output/report-docx-workspace --lang 日本語 --style "丁寧語（です・ます調）"
```

### Prerequisites

- A completed `/pdf-to-docx` workspace (must contain `dsl/page-*.xml`)
- LM Studio running on `localhost:1234` with `qwen/qwen3-4b-2507` loaded

### Pipeline

1. **Validate workspace** — checks for `input.json`, `input.md`, `dsl/page-*.xml`
2. **Extract translatable text** — collects `<run>` and `<cell>` text from XML DSL (skips `<run is-math="true">`)
3. **Translate via LM Studio** — agent generates `translate_content.py`, sends JSON payloads per page
4. **DSL → DOCX** — reuses fixed `dsl_to_docx.py --dsl-dir dsl-translated`
5. **Verify output** — confirms translated DOCX exists and is non-empty

### Key Details

- Only text content is translated; all XML attributes (font-size, color, bold, alignment) are preserved
- Style JSON files are not touched (visual styles are language-independent)
- Math formulas (`is-math="true"`) are never translated
- Output: `<workspace>/translation/translated-output.docx`

---

## Typical End-User Workflow

```
# Step 1: Convert PDF to editable DOCX
/pdf-to-docx ./contract.pdf --output ./output

# Step 2: Translate the resulting workspace
/docx-translate-to-docx ./output/contract-docx-workspace --lang English
```
