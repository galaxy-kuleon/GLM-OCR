# AGENTS.md

This project provides AI Agent Skills for document processing:
**PDF → editable DOCX → translated DOCX**, plus a standalone **DOCX → translated DOCX** pipeline.

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
5. **Style extraction** — VLM (OpenCode Zen, kimi-k2.5) infers font sizes, colors, alignment per region
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
- `OPENCODE_ZEN_API_KEY` env var enables style extraction + VLM review; pipeline works without it (falls back to defaults)

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

## Skill 3: `/another-pure-pure-docx-translate-to-docx`

Translates a DOCX (or DOC) file directly into a target language by parsing and modifying OOXML. **No dependency on PDF, OCR, or the pdf-to-docx pipeline** — works on any standard DOCX file.

```
/another-pure-pure-docx-translate-to-docx <input-file.docx> --lang <target-language> [--style <style-notes>] [--output <output-path>]
```

**Examples:**
```
/another-pure-pure-docx-translate-to-docx report.docx --lang 繁體中文
/another-pure-pure-docx-translate-to-docx contract.doc --lang "formal English" --style "keep legal terms in Latin"
/another-pure-pure-docx-translate-to-docx input/manual.docx --lang 日本語 --output output/manual-ja.docx
```

### Prerequisites

- `uv` (Python package runner)
- `soffice` (LibreOffice, only needed if input is `.doc` not `.docx`)

### Pipeline

1. **Convert DOC → DOCX** (if needed) — `soffice --headless --convert-to docx`
2. **Extract translatable text** — fixed `extract_docx_texts.py` parses DOCX ZIP/XML, outputs `texts.json`
3. **Translate via AI** — Claude agent translates segments in batches of ~50
4. **Apply translations** — fixed `apply_docx_translations.py` writes translated text back into DOCX
5. **Verify output** — confirms translated DOCX exists and is non-empty

### Key Details

- Operates directly on DOCX ZIP/XML using `lxml` — no `python-docx` dependency
- Handles: paragraphs, headings, tables, textboxes (`wps:txbx` + `v:textbox`), image alt text, chart text, headers, footers, footnotes, endnotes
- Only `w:t` text nodes are modified — all formatting (font, size, bold, color, alignment) is preserved (zero-loss round-trip)
- Translation unit is the paragraph (`w:p`); translated text goes into the first run, subsequent runs are cleared
- Scripts are in `.claude/skills/another-pure-pure-docx-translate-to-docx/scripts/`
- Output: `<input-name>.appdttd.docx` (next to the input file by default)

---

## Skill 4: `/image-to-docx`

Converts one or more scanned/photographed images (representing document pages) into a high-fidelity editable DOCX. Performs per-page OCR with glm-ocr, then reuses the full `pdf-to-docx` pipeline for style extraction, DSL generation, VLM review, and DOCX assembly.

```
/image-to-docx <images> [--output <dir>]
```

**`<images>`** accepts:
- Single file: `/image-to-docx scan.jpg`
- Directory (auto-finds jpg/jpeg/png/bmp/gif/webp): `/image-to-docx scans/`
- Space-separated files: `/image-to-docx p1.jpg p2.jpg p3.jpg`

**Examples:**
```
/image-to-docx image_input/20260302_010005.jpg
/image-to-docx image_input/ --output ./output
/image-to-docx page1.jpg page2.jpg page3.jpg --output ./output
```

### Prerequisites

- `uv` (Python package runner) — required
- `OPENCODE_ZEN_API_KEY` env var — optional; enables VLM style extraction and review (improves quality)
- `soffice` (LibreOffice) — optional; enables visual verification step

### Pipeline

1. **Prepare images** — natural-sort images, convert to PNG for VLM reference, measure dimensions in pts, write `image-info.json`
2. **Per-page OCR** — `uv run glmocr parse` on each page, producing structured JSON + markdown + cropped images
3. **Consolidate OCR** — merge per-page results into pdf-to-docx compatible `ocr-output/input/` structure, rename crop images with correct page indices
4. **Style extraction** — VLM (OpenCode Zen, kimi-k2.5) infers font sizes, colors, alignment per region *(reuses pdf-to-docx)*
5. **XML DSL generation** — converts OCR + styles into per-page XML files with page dimensions from `image-info.json` *(reuses pdf-to-docx)*
6. **VLM review** — compares XML DSL against original images; agent fixes issues *(reuses pdf-to-docx)*
7. **DSL → DOCX** — deterministic assembly *(reuses pdf-to-docx)*
8. **Visual verification** (optional) — renders DOCX back to images for VLM comparison *(reuses pdf-to-docx)*

### Key Details

- Workspace: `<output>/<first-image-stem>-img-docx-workspace/`
- New scripts in `.claude/skills/image-to-docx/scripts/` handle image→workspace conversion
- All pdf-to-docx scripts (extract_styles, build_page_dsl, review_dsl, dsl_to_docx, verify_docx_visual) are reused unchanged
- OCR `bbox_2d` values are normalized 0–1000, not pixels
- Page dimensions (pts) are derived from image DPI; fallback is 200 DPI when EXIF is missing
- Output: `<workspace>/final-output.docx`

---

## Typical End-User Workflows

### PDF → DOCX → Translate (Skills 1 + 2)

```
# Step 1: Convert PDF to editable DOCX
/pdf-to-docx ./contract.pdf --output ./output

# Step 2: Translate the resulting workspace
/docx-translate-to-docx ./output/contract-docx-workspace --lang English
```

### Standalone DOCX Translate (Skill 3)

```
# Translate a DOCX file directly — no PDF or OCR needed
/another-pure-pure-docx-translate-to-docx ./report.docx --lang 繁體中文

# Translate a .doc file (auto-converts via LibreOffice)
/another-pure-pure-docx-translate-to-docx ./legacy.doc --lang "formal English"
```

### Image(s) → DOCX (Skill 4)

```
# Convert a single scanned image to editable DOCX
/image-to-docx ./scan.jpg

# Convert a directory of scanned pages to DOCX
/image-to-docx ./scans/ --output ./output

# Convert multiple images (processed in natural sort order)
/image-to-docx page1.jpg page2.jpg page3.jpg --output ./output
```
