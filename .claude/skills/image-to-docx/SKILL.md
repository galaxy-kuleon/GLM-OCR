---
name: image-to-docx
description: Converts one or more scanned or photographed document images into a high-fidelity editable DOCX. Use when the user wants image files or an image folder converted to Word while preserving layout.
argument-hint: <image-or-dir> [more-images...] [--output <dir>]
disable-model-invocation: false
allowed-tools:
  - Bash
  - Read
  - Write
  - Edit
  - Glob
  - Grep
---

# Skill: image-to-docx

Convert one or more scanned/photographed images (representing document pages) into an editable DOCX file. Maximizes reuse of the existing `pdf-to-docx` scripts; only adds image-specific bridge logic.

## Pipeline Checklist

Copy and track. Fill in values as you complete each step.

```
Pipeline Progress:
- [ ] Step 0: Validated, IMAGE_INPUTS=___
- [ ] Step 1: PAGE_COUNT=___ images prepared
- [ ] Step 2: OCR complete, pages=___ regions=___
- [ ] Step 3: Consolidated OCR → input.json
- [ ] Step 4: Style extraction complete
- [ ] Step 5: DSL files count=___
- [ ] Step 5.5: SKIP / DONE (optional)
- [ ] Step 6: output.docx size=___ bytes
- [ ] Step 6.5: SKIP / DONE (optional)
- [ ] Step 7: final-output.docx path=___
```

## Invocation

```
/image-to-docx <images> [--output <dir>]
```

- `<images>`: single file, directory (auto-finds jpg/jpeg/png/bmp/gif/webp), or space-separated multiple files
- `--output`: output directory, default `./output`

---

## Step 0: Parse Arguments & Validate Environment

1. Parse `$ARGUMENTS` into:
   - `IMAGE_INPUTS`: all positional arguments before `--output`
   - `OUTPUT_DIR`: value after `--output`, default `./output`
2. Input rules:
   - if there is exactly 1 positional path and it is a directory, load all supported images from that directory
   - otherwise, treat all positional paths as explicit image files in the exact order they were given
3. Supported extensions: `jpg`, `jpeg`, `png`, `bmp`, `gif`, `webp`
4. Build the final ordered absolute image path list.
5. Set `FIRST_IMAGE` to the first path in that final ordered list.
6. Validate with these exact checks and stop immediately if any fail:

```bash
mkdir -p "$OUTPUT_DIR"
command -v uv >/dev/null || { echo "ERROR: uv not found"; exit 1; }

test -f ".claude/skills/image-to-docx/scripts/prepare_images.py" || { echo "MISSING: .claude/skills/image-to-docx/scripts/prepare_images.py"; exit 1; }
test -f ".claude/skills/image-to-docx/scripts/consolidate_ocr_results.py" || { echo "MISSING: .claude/skills/image-to-docx/scripts/consolidate_ocr_results.py"; exit 1; }
test -f ".claude/skills/pdf-to-docx/scripts/extract_styles.py" || { echo "MISSING: .claude/skills/pdf-to-docx/scripts/extract_styles.py"; exit 1; }
test -f ".claude/skills/pdf-to-docx/scripts/build_page_dsl.py" || { echo "MISSING: .claude/skills/pdf-to-docx/scripts/build_page_dsl.py"; exit 1; }
test -f ".claude/skills/pdf-to-docx/scripts/review_dsl.py" || { echo "MISSING: .claude/skills/pdf-to-docx/scripts/review_dsl.py"; exit 1; }
test -f ".claude/skills/pdf-to-docx/scripts/dsl_to_docx.py" || { echo "MISSING: .claude/skills/pdf-to-docx/scripts/dsl_to_docx.py"; exit 1; }
test -f ".claude/skills/pdf-to-docx/scripts/verify_docx_visual.py" || { echo "MISSING: .claude/skills/pdf-to-docx/scripts/verify_docx_visual.py"; exit 1; }
test -n "$FIRST_IMAGE" || { echo "ERROR: no input images resolved"; exit 1; }
```

7. Derive the workspace with these exact commands:

```bash
FIRST_STEM=$(python3 -c "import os; print(os.path.splitext(os.path.basename('$FIRST_IMAGE'))[0])")
WORKSPACE="$(pwd)/$OUTPUT_DIR/${FIRST_STEM}-img-docx-workspace"
mkdir -p "$WORKSPACE"
echo "Using workspace: $WORKSPACE"
```

8. Export the final ordered absolute image path list as `IMAGE_PATHS_NL`, with one absolute path per line. Then write `$WORKSPACE/image-paths.json` with this exact block:

```bash
test -n "$IMAGE_PATHS_NL" || { echo "ERROR: IMAGE_PATHS_NL is empty"; exit 1; }
export WORKSPACE IMAGE_PATHS_NL
python3 << 'PYEOF'
import json, os

workspace = os.environ["WORKSPACE"]
paths = [line for line in os.environ["IMAGE_PATHS_NL"].splitlines() if line.strip()]
assert paths, "No image paths provided"

with open(os.path.join(workspace, "image-paths.json"), "w", encoding="utf-8") as f:
    json.dump(paths, f, ensure_ascii=False, indent=2)

print(f"Wrote {len(paths)} image paths")
PYEOF
```

Then derive `IMAGE_CSV` with this exact command:

```bash
test -f "$WORKSPACE/image-paths.json" || { echo "ERROR: missing $WORKSPACE/image-paths.json"; exit 1; }
IMAGE_CSV=$(python3 -c "import json; print(','.join(json.load(open('$WORKSPACE/image-paths.json', encoding='utf-8'))))")
```

9. If `OPENCODE_ZEN_API_KEY` is not set, continue. Style extraction and review scripts already have fallback behavior.

---

## Step 1: Prepare Images

Run:

```bash
uv run --with Pillow \
  .claude/skills/image-to-docx/scripts/prepare_images.py \
  --images "$IMAGE_CSV" \
  --workspace "$WORKSPACE"
```

**Read PAGE_COUNT** (use python3, not jq — jq may not be installed):
```bash
PAGE_COUNT=$(python3 -c "import json; print(json.load(open('$WORKSPACE/image-info.json'))['page_count'])")
```

**Verify** (must all exist before proceeding):
```bash
test -f "$WORKSPACE/image-info.json" || { echo "ERROR: missing $WORKSPACE/image-info.json"; exit 1; }
test -f "$WORKSPACE/input-pdf-rendered-pngs/page-1.png" || { echo "ERROR: missing $WORKSPACE/input-pdf-rendered-pngs/page-1.png"; exit 1; }
test -d "$WORKSPACE/ocr-input/page-1" || { echo "ERROR: missing $WORKSPACE/ocr-input/page-1"; exit 1; }
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
test -f "$WORKSPACE/ocr-output/input/style-page-1.json" || { echo "ERROR: missing $WORKSPACE/ocr-output/input/style-page-1.json"; exit 1; }
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
test -f "$WORKSPACE/dsl/page-$N.xml" || { echo "ERROR: missing $WORKSPACE/dsl/page-$N.xml"; exit 1; }
```

---

## Step 5.5: VLM Review of DSL (optional, default SKIP)

Default behavior: skip this step. Only run it if the user explicitly asked for review / extra layout polishing.

If the user did not explicitly ask for review, print `Step 5.5 SKIP: default skip unless user explicitly requested review` and continue to Step 6.

If the user explicitly asked for review, run:

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

Apply ALL items only when these conditions are true:

- 5 or fewer items on the page
- every item uses a known `field` from the table above
- the target region is unambiguous

If any condition fails, stop and report the page number instead of guessing.

After editing, re-read each edited file to confirm the changes are saved correctly.

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

## Step 6.5: Visual Verification (optional, default SKIP)

Default behavior: skip this step. Only run it if BOTH conditions are true:

- the user explicitly asked for visual QA
- `soffice` is available

If the user did not explicitly ask for visual QA, print `Step 6.5 SKIP: default skip unless user explicitly requested visual QA` and continue to Step 7.

If the user explicitly asked for visual QA, first check soffice:

```bash
command -v soffice >/dev/null && echo "Step 6.5 OK: soffice available" || echo "Step 6.5 SKIP: soffice not found"
```

Only run the next command if the check printed `Step 6.5 OK: soffice available`:

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

**If any array is non-empty**, only apply fixes when these conditions are true:

- 5 or fewer items on the page
- every item names a clear `field` and `suggested_value`
- the target element is unambiguous

If any condition fails, stop and report the page number instead of guessing.

After fixing XML, re-run Step 6 once more. Do NOT re-run verify_docx_visual.py again unless the user explicitly asks for another visual QA pass.

### How to locate the element from `element` description

- `"heading level='2' run '紅綠操盤法'"` → find `<heading level="2">` containing a `<run>` with that text
- `"paragraph run text"` → applies to ALL `<paragraph>` elements' `<run>` children
- Match by element type and run text content

---

## Step 7: Final Output

```bash
cp "$WORKSPACE/output.docx" "$WORKSPACE/final-output.docx"
python3 -c "import os; size=os.path.getsize('$WORKSPACE/final-output.docx'); print(f'final-output.docx size: {size} bytes'); assert size > 1000"
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
