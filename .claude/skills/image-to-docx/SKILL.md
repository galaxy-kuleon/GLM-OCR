# Skill: image-to-docx

Convert one or more scanned/photographed images (representing document pages) into an editable DOCX file. Maximizes reuse of the existing `pdf-to-docx` scripts; only adds image-specific bridge logic.

## Invocation

```
/image-to-docx <images> [--output <dir>]
```

- `<images>`: single file, directory (auto-finds jpg/jpeg/png/bmp/gif/webp), or space-separated multiple files
- `--output`: output directory, default `./output`

---

## Step 0: Parse Arguments & Validate Environment

1. **Collect image paths**: If a directory is given, find all files with extensions: jpg, jpeg, png, bmp, gif, webp. Natural-sort by filename (e.g., page-2 before page-10).
2. **OUTPUT_DIR**: use `--output` value, default `./output`.
3. **Validate** — stop and report if any check fails:
   ```bash
   command -v uv || { echo "ERROR: uv not found"; exit 1; }
   ```
   - At least 1 image found
   - All 7 fixed scripts exist (use `ls` to check each):
     - `.claude/skills/image-to-docx/scripts/prepare_images.py`
     - `.claude/skills/image-to-docx/scripts/consolidate_ocr_results.py`
     - `.claude/skills/pdf-to-docx/scripts/extract_styles.py`
     - `.claude/skills/pdf-to-docx/scripts/build_page_dsl.py`
     - `.claude/skills/pdf-to-docx/scripts/review_dsl.py`
     - `.claude/skills/pdf-to-docx/scripts/dsl_to_docx.py`
     - `.claude/skills/pdf-to-docx/scripts/verify_docx_visual.py`

4. **Set WORKSPACE** (use first image filename stem, without extension):
   ```bash
   FIRST_STEM=$(basename "$FIRST_IMAGE" | sed 's/\.[^.]*$//')
   WORKSPACE="$(pwd)/$OUTPUT_DIR/${FIRST_STEM}-img-docx-workspace"
   mkdir -p "$WORKSPACE"
   ```
5. Warn (but continue) if `OPENCODE_ZEN_API_KEY` is not set — style extraction and VLM review will fall back to defaults.

---

## Step 1: Prepare Images

Build a comma-separated absolute-path list of images in natural sort order, then run:

```bash
uv run --with Pillow \
  .claude/skills/image-to-docx/scripts/prepare_images.py \
  --images "/abs/path/img1.jpg,/abs/path/img2.jpg" \
  --workspace "$WORKSPACE"
```

**Read PAGE_COUNT** (use python3, not jq — jq may not be installed):
```bash
PAGE_COUNT=$(python3 -c "import json; print(json.load(open('$WORKSPACE/image-info.json'))['page_count'])")
```

**Verify** (must all exist before proceeding):
```bash
ls "$WORKSPACE/image-info.json"
ls "$WORKSPACE/input-pdf-rendered-pngs/page-1.png"
ls "$WORKSPACE/ocr-input/page-1/"
```

---

## Step 2: Per-Page OCR (loop N = 1 .. PAGE_COUNT)

**IMPORTANT — glmocr output path rule**: glmocr names its output subdirectory after the **stem of the input file**. Since OCR inputs are named `input.jpg` (or `input.png`, etc.), glmocr always writes to `<output-dir>/input/`.

For each page N (1-indexed):

```bash
# Get extension from image-info.json (use python3)
EXT=$(python3 -c "import json; print(json.load(open('$WORKSPACE/image-info.json'))['pages'][$((N-1))]['ext'])")

uv run glmocr parse \
  "$WORKSPACE/ocr-input/page-$N/input.$EXT" \
  --output "$WORKSPACE/ocr-output-pages/page-$N/"
```

**Expected output paths** (because input stem is "input"):
- `$WORKSPACE/ocr-output-pages/page-$N/input/input.json`
- `$WORKSPACE/ocr-output-pages/page-$N/input/input.md`
- `$WORKSPACE/ocr-output-pages/page-$N/input/imgs/` (may be empty if no image regions)

**Verify each page before proceeding:**
```bash
python3 -c "
import json, sys
data = json.load(open('$WORKSPACE/ocr-output-pages/page-$N/input/input.json'))
assert isinstance(data, list) and len(data) > 0, 'Bad structure'
print(f'Page $N: {len(data[0])} regions')
first = data[0][0] if data[0] else {}
print(f'  First region content preview: {repr(str(first.get(\"content\",\"\"))[:60])}')
"
```
If content looks like ` ```markdown\n\n``` ` (empty), the OCR model is broken — check your local VLM service.

---

## Step 3: Consolidate OCR Results

```bash
uv run \
  .claude/skills/image-to-docx/scripts/consolidate_ocr_results.py \
  --workspace "$WORKSPACE" \
  --pages "$PAGE_COUNT"
```

**Verify:**
```bash
python3 -c "
import json
data = json.load(open('$WORKSPACE/ocr-output/input/input.json'))
print(f'Pages in consolidated JSON: {len(data)} (expected $PAGE_COUNT)')
"
```

---

## Step 4: Style Extraction (reuses pdf-to-docx)

```bash
uv run --with requests,Pillow \
  .claude/skills/pdf-to-docx/scripts/extract_styles.py \
  --workspace "$WORKSPACE" \
  --pages "$PAGE_COUNT"
```

**Verify** each `style-page-N.json` exists:
```bash
ls "$WORKSPACE/ocr-output/input/style-page-1.json"
```

---

## Step 5: Build XML DSL (loop N = 1 .. PAGE_COUNT)

For each page N (use python3 to read dimensions):

```bash
PAGE_WIDTH=$(python3 -c "import json; print(json.load(open('$WORKSPACE/image-info.json'))['pages'][$((N-1))]['width_pts'])")
PAGE_HEIGHT=$(python3 -c "import json; print(json.load(open('$WORKSPACE/image-info.json'))['pages'][$((N-1))]['height_pts'])")

uv run --with lxml \
  .claude/skills/pdf-to-docx/scripts/build_page_dsl.py \
  --workspace "$WORKSPACE" \
  --page $N \
  --page-width-pts "$PAGE_WIDTH" \
  --page-height-pts "$PAGE_HEIGHT"
```

**Verify:**
```bash
ls "$WORKSPACE/dsl/page-$N.xml"
```

---

## Step 5.5: VLM Review of DSL (reuses pdf-to-docx)

```bash
uv run --with requests,lxml,Pillow \
  .claude/skills/pdf-to-docx/scripts/review_dsl.py \
  --workspace "$WORKSPACE" \
  --pages "$PAGE_COUNT"
```

Outputs `$WORKSPACE/dsl/review-page-{N}.json` for each page.

### Review JSON schema

Each item in `review-page-N.json` looks like:
```json
{
  "type": "font_mismatch",
  "region": 2,
  "field": "font-family",
  "expected": "sans",
  "actual": "serif",
  "description": "..."
}
```

- `region`: **1-based** index of the child element in `<page>` (1 = first child, 2 = second, etc.)
- `field`: which XML attribute to change
- `expected`: the value to set

**If the array is empty `[]`, no fixes needed — skip to Step 6.**

### How to apply fixes

For each item, find the `region`-th child element of `<page>` and apply the `field`/`expected` mapping:

| `field` | Apply to which XML element |
|---|---|
| `font-family` | attribute on `<heading>` or `<paragraph>` |
| `alignment` | attribute on `<heading>` or `<paragraph>` |
| `space-before-pt` | attribute on `<heading>` or `<paragraph>` |
| `space-after-pt` | attribute on `<paragraph>` |
| `line-spacing` | attribute on `<paragraph>` |
| `font-size-pt` | attribute on `<run>` inside the heading/paragraph |
| `bold` | attribute on `<run>` inside the heading/paragraph |
| `color-rgb` | attribute on `<run>` inside the heading/paragraph |

**Example**: item `{"region": 1, "field": "font-family", "expected": "sans"}` → change the `font-family` attribute on the **first** child of `<page>` to `"sans"`.

Apply ALL items. Then re-read each edited file to confirm changes are saved correctly.

---

## Step 6: DSL → DOCX (reuses pdf-to-docx)

```bash
uv run --with python-docx,lxml,Pillow \
  .claude/skills/pdf-to-docx/scripts/dsl_to_docx.py \
  --workspace "$WORKSPACE" \
  --output "$WORKSPACE/output.docx"
```

**Verify:**
```bash
python3 -c "
import os, sys
size = os.path.getsize('$WORKSPACE/output.docx')
print(f'output.docx size: {size} bytes')
assert size > 1000, 'File too small, something went wrong'
"
```

---

## Step 6.5: Visual Verification (optional, requires soffice)

Only run if `command -v soffice` succeeds:

```bash
uv run --with requests,Pillow \
  .claude/skills/pdf-to-docx/scripts/verify_docx_visual.py \
  --workspace "$WORKSPACE" \
  --pages "$PAGE_COUNT" \
  --docx "$WORKSPACE/output.docx"
```

Outputs `$WORKSPACE/dsl/visual-review-page-{N}.json` for each page.

### Visual-review JSON schema

Each item looks like:
```json
{
  "type": "font_difference",
  "element": "heading level='2' run '紅綠操盤法'",
  "field": "font-size-pt",
  "current_value": "36",
  "suggested_value": "48",
  "description": "..."
}
```

- `element`: description of which XML node (heading/paragraph and its run text)
- `field`: XML attribute to change
- `suggested_value`: value to apply (always apply this, not `current_value`)

**If any array is non-empty**, apply all fixes using the same attribute table from Step 5.5, then re-run Step 6 (DSL → DOCX) once more. Do NOT re-run verify_docx_visual.py again.

### How to locate the element from `element` description

- `"heading level='2' run '紅綠操盤法'"` → find `<heading level="2">` containing a `<run>` with that text
- `"paragraph run text"` → applies to ALL `<paragraph>` elements' `<run>` children
- Match by element type and run text content

---

## Step 7: Final Output

```bash
cp "$WORKSPACE/output.docx" "$WORKSPACE/final-output.docx"
ls -lh "$WORKSPACE/final-output.docx"
```

**Report to user:**
- Workspace path: `$WORKSPACE`
- Final output: `$WORKSPACE/final-output.docx`
- Pages processed: `$PAGE_COUNT`

---

## Workspace Structure Reference

```
<output>/<first-image-stem>-img-docx-workspace/
├── image-info.json                      # page count, per-page pts size, ext
├── input-pdf-rendered-pngs/
│   └── page-1.png                       # PNG for VLM reference
├── ocr-input/
│   └── page-1/input.jpg                 # copied original image (always named input.*)
├── ocr-output-pages/
│   └── page-1/
│       └── input/                       # ← glmocr names this dir after input stem ("input")
│           ├── input.json               # [[regions,...]] — single page
│           ├── input.md
│           └── imgs/                    # may be empty (no image regions)
├── ocr-output/input/                    # pdf-to-docx compatible (built by Step 3)
│   ├── input.json                       # [[page1],[page2],...] — all pages
│   ├── input.md
│   ├── style-page-1.json                # built by Step 4
│   └── imgs/                            # renamed crops: cropped_page{N-1}_idx*.jpg
├── dsl/
│   ├── page-1.xml                       # editable intermediate (Steps 5-6.5)
│   ├── review-page-1.json               # VLM review output (Step 5.5)
│   └── visual-review-page-1.json        # visual verify output (Step 6.5)
├── output.docx
└── final-output.docx
```

---

## DSL XML Quick Reference

Valid values for common attributes:

| Attribute | Valid values | Element |
|---|---|---|
| `font-family` | `"sans"`, `"serif"`, `"mono"` | `<heading>`, `<paragraph>` |
| `alignment` | `"left"`, `"center"`, `"right"`, `"justify"` | `<heading>`, `<paragraph>` |
| `font-size-pt` | number (e.g. `"12"`, `"14"`, `"24"`) | `<run>` |
| `bold` | `"true"`, `"false"` | `<run>` |
| `color-rgb` | `"R,G,B"` (e.g. `"0,0,0"`) | `<run>` |
| `space-before-pt` | number (e.g. `"0"`, `"12"`) | `<heading>`, `<paragraph>` |
| `space-after-pt` | number (e.g. `"0"`, `"6"`) | `<paragraph>` |
| `line-spacing` | number (e.g. `"1.0"`, `"1.5"`) | `<paragraph>` |

---

## Path Contracts (Hardcoded in Reused Scripts)

| Script | Reads | Writes |
|--------|-------|--------|
| `extract_styles.py` | `ocr-output/input/input.json`<br>`input-pdf-rendered-pngs/page-{N}.png` | `ocr-output/input/style-page-{N}.json` |
| `build_page_dsl.py` | `ocr-output/input/input.json`<br>`ocr-output/input/style-page-{N}.json` | `dsl/page-{N}.xml` |
| `review_dsl.py` | `dsl/page-{N}.xml`<br>`input-pdf-rendered-pngs/page-{N}.png` | `dsl/review-page-{N}.json` |
| `verify_docx_visual.py` | `input-pdf-rendered-pngs/page-{N}.png`<br>DOCX via soffice | `dsl/visual-review-page-{N}.json` |
| `dsl_to_docx.py` | `dsl/page-*.xml`<br>images: `$WORKSPACE/{src}` or `$WORKSPACE/ocr-output/input/{src}` | `--output` path |

**Page numbering**: scripts use 1-based filenames (`page-1.xml`); OCR JSON data array is 0-based (`data[0]` = page 1).
