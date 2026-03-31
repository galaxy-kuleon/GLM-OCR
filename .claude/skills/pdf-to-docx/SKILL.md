---
name: pdf-to-docx
description: Converts a PDF file to a high-fidelity editable DOCX. Use when the user wants to convert a PDF to Word, recreate a PDF as DOCX, or produce an editable document from a scanned/digital PDF. Automates OCR, layout detection, XML DSL generation, VLM review, and deterministic DOCX assembly.
argument-hint: <pdf-path> [--output <dir>]
disable-model-invocation: false
allowed-tools: Bash, Read, Write, Edit, Glob, Grep
---

# pdf-to-docx Skill

Convert a PDF to a high-fidelity editable DOCX through automated OCR, style extraction, XML DSL, and DOCX assembly.

**Architecture**: Each page produces a `page-{N}.xml` DSL file. Five fixed Python scripts handle the pipeline — no dynamically generated scripts.

## Pipeline Checklist

Copy and track. Fill in values as you complete each step.

```
Pipeline Progress:
- [ ] Step 0: PDF_PATH=___ WORKSPACE=___ PAGE_WIDTH=___ PAGE_HEIGHT=___
- [ ] Step 1: PAGE_COUNT=___
- [ ] Step 2: pdf-info.txt + pdf-fonts.txt + pdf-fulltext.txt
- [ ] Step 3: OCR pages=___ regions=___
- [ ] Step 4: Style extraction complete
- [ ] Step 5: DSL files count=___
- [ ] Step 5.5: SKIP / DONE (optional)
- [ ] Step 6: output.docx size=___ bytes
- [ ] Step 6.5: SKIP / DONE (optional)
- [ ] Step 7: final-output.docx path=___
```

## Step 0: Parse Arguments, Fail Fast, Create Workspace

### Parse `$ARGUMENTS`

- `PDF_PATH` (required): first positional argument — path to the input PDF
- `OUTPUT_DIR` (optional): value after `--output`, default `./output`

Derive:
- `PDF_STEM`: filename without extension
- `WORKSPACE`: `<OUTPUT_DIR>/<PDF_STEM>-docx-workspace`

### Validate input and prerequisites

Run this exact block. Abort immediately if any check fails.

```bash
test -f "$PDF_PATH" || { echo "ERROR: input PDF not found: $PDF_PATH"; exit 1; }

case "$PDF_PATH" in
  *.pdf|*.PDF) ;;
  *) echo "ERROR: input must be a .pdf file: $PDF_PATH"; exit 1 ;;
esac

mkdir -p "$OUTPUT_DIR"
PDF_STEM=$(basename "$PDF_PATH")
PDF_STEM="${PDF_STEM%.*}"
WORKSPACE="$OUTPUT_DIR/${PDF_STEM}-docx-workspace"
SKILL_DIR=".claude/skills/pdf-to-docx/scripts"

command -v pdftocairo >/dev/null || { echo "MISSING: pdftocairo (install poppler-utils)"; exit 1; }
command -v pdfinfo >/dev/null || { echo "MISSING: pdfinfo (install poppler-utils)"; exit 1; }
command -v pdffonts >/dev/null || { echo "MISSING: pdffonts (install poppler-utils)"; exit 1; }
command -v pdftotext >/dev/null || { echo "MISSING: pdftotext (install poppler-utils)"; exit 1; }
command -v uv >/dev/null || { echo "MISSING: uv"; exit 1; }

for script in extract_styles.py build_page_dsl.py review_dsl.py dsl_to_docx.py verify_docx_visual.py; do
  test -f "$SKILL_DIR/$script" || { echo "MISSING: $SKILL_DIR/$script"; exit 1; }
done

mkdir -p "$WORKSPACE"/{dsl,input-pdf-rendered-pngs}
cp "$PDF_PATH" "$WORKSPACE/input.pdf"
```

**IMPORTANT**: The input PDF is always copied as `input.pdf`. All OCR output paths use this stem: `ocr-output/input/input.json`, etc.

## Step 1: PDF to Reference PNGs

```bash
pdftocairo -png -r 200 "$WORKSPACE/input.pdf" "$WORKSPACE/input-pdf-rendered-pngs/page"
```

Normalize zero-padded filenames and count pages:

```bash
for f in "$WORKSPACE/input-pdf-rendered-pngs"/page-*.png; do
  base=$(basename "$f" .png)
  num=$(printf '%s' "$base" | sed 's/page-0*//')
  [ "$base" != "page-${num}" ] && mv "$f" "$WORKSPACE/input-pdf-rendered-pngs/page-${num}.png"
done

PAGE_COUNT=$(python3 -c "import glob; print(len(glob.glob('$WORKSPACE/input-pdf-rendered-pngs/page-*.png')))")
test "$PAGE_COUNT" -gt 0 || { echo "ERROR: no rendered page PNGs found"; exit 1; }
echo "PAGE_COUNT=$PAGE_COUNT"
```

## Step 2: PDF Metadata Extraction

```bash
pdfinfo "$WORKSPACE/input.pdf" > "$WORKSPACE/pdf-info.txt"
pdffonts "$WORKSPACE/input.pdf" > "$WORKSPACE/pdf-fonts.txt"
pdftotext "$WORKSPACE/input.pdf" "$WORKSPACE/pdf-fulltext.txt"
```

Extract page dimensions:

```bash
eval "$(python3 -c "import re; s=open('$WORKSPACE/pdf-info.txt', encoding='utf-8').read(); m=re.search(r'Page size:\\s*([0-9.]+) x ([0-9.]+) pts', s); assert m, 'Page size not found in pdf-info.txt'; print(f'PAGE_WIDTH={m.group(1)}'); print(f'PAGE_HEIGHT={m.group(2)}')")"
echo "PAGE_WIDTH=$PAGE_WIDTH PAGE_HEIGHT=$PAGE_HEIGHT"
```

## Step 3: OCR Parsing

```bash
uv run glmocr parse "$WORKSPACE/input.pdf" --output "$WORKSPACE/ocr-output/"
```

**IMPORTANT**: Do NOT run multiple `uv run glmocr parse` commands in parallel.

Output: `$WORKSPACE/ocr-output/input/` containing `input.json`, `input.md`, `imgs/`, `layout_vis/`.

### OCR JSON Structure

Each region in `input.json`:

```json
{
  "index": 0,
  "label": "text|table|formula|image",
  "native_label": "paragraph_title|text|table|image|doc_title|figure_title|vision_footnote|...",
  "content": "...",
  "bbox_2d": [x1, y1, x2, y2]
}
```

**CRITICAL**: `bbox_2d` values are normalized 0-1000 (not pixels).

### Cropped image naming

`cropped_page{N}_idx{M}.jpg` where N = page index (0-based), M = sequential image counter per page (NOT the region `index`).

### Native label categories

| native_label | label | Meaning |
|---|---|---|
| `doc_title` | text | Document title (H1), content prefixed with `# ` |
| `paragraph_title` | text | Section heading (H2), content prefixed with `## ` |
| `text` | text | Body paragraph |
| `figure_title` | text | Figure/image caption |
| `vision_footnote` | text | Footnote/endnote |
| `table` | table | Table (HTML in content) |
| `display_formula` | formula | Display math (LaTeX) |
| `image` | image | Image (content is null) |

## Step 4: Style Extraction (LMStudio)

```bash
uv run --with requests,Pillow \
  .claude/skills/pdf-to-docx/scripts/extract_styles.py \
  --workspace "$WORKSPACE" --pages "$PAGE_COUNT"
```

**What this does**: Sends page image + region summary to VLM, receives simplified style data (font size, bold, color, alignment). Results: `$WORKSPACE/ocr-output/input/style-page-{N}.json`.

**Fallback**: If API unavailable, auto-uses defaults by `native_label`. Pipeline continues.

### Style JSON output format

```json
[
  {
    "region_index": 0,
    "font_size_pt": 14,
    "bold": true,
    "italic": false,
    "underline": false,
    "color_rgb": [0, 0, 0],
    "alignment": "center"
  }
]
```

Additional optional fields: `tb` (text box), `bd` (border), `bg_rgb` (background color), `border_style`, `cell_overrides` (table cell-level colors with `col_colors`, `row_colors`, `cell_colors`, `keyword_styles`).

## Step 5: Build Per-Page XML DSL

```bash
for N in $(seq 1 $PAGE_COUNT); do
  uv run --with lxml \
    .claude/skills/pdf-to-docx/scripts/build_page_dsl.py \
    --workspace "$WORKSPACE" --page $N \
    --page-width-pts "$PAGE_WIDTH" --page-height-pts "$PAGE_HEIGHT"
done
```

Output: `$WORKSPACE/dsl/page-1.xml`, `page-2.xml`, ...

**Verify Step 5:**
```bash
python3 -c "import glob, os; paths=sorted(glob.glob('$WORKSPACE/dsl/page-*.xml')); print(f'DSL files: {len(paths)}'); assert len(paths) == int('$PAGE_COUNT'); assert all(os.path.getsize(p) > 0 for p in paths)"
```

### XML DSL Elements

| XML Element | python-docx Operation | Source |
|---|---|---|
| `<heading level="N">` | `doc.add_heading("", N)` + black color | doc_title->1, paragraph_title->2 |
| `<paragraph>` | `doc.add_paragraph(style=...)` | native_label mapping |
| `<run>` | `para.add_run()` + font/color/bold/italic | style JSON + markdown |
| `<table>` | `doc.add_table()` + occupancy grid merge | OCR HTML table |
| `<cell>` | cell text + shading + borders; may contain `<run>` children | style defaults + keyword_styles |
| `<image>` | `doc.add_picture()` + bbox scaling | sequential counter |
| `<text-frame>` | `w:framePr` + `w:pBdr` (TWIPS) | bbox + floating detection |
| `<side-by-side>` | invisible table `w:val="none"` | parallel layout detection |

### Run attributes (all optional)

| Attribute | Default | Description |
|---|---|---|
| `font-size-pt` | `11` | Font size |
| `bold` | `false` | Bold |
| `italic` | `false` | Italic |
| `underline` | `false` | Underline |
| `color-rgb` | `0,0,0` | Text color R,G,B |
| `superscript` | `false` | Superscript |
| `font-name` | (inherits page font) | Font name |

## Step 5.5: VLM Review — default SKIP

Print `Step 5.5 SKIP` and continue to Step 6. Only if the user explicitly asked for review, read [optional-steps-reference.md](optional-steps-reference.md) section "Step 5.5".

## Step 6: DSL to DOCX

```bash
uv run --with python-docx,lxml,Pillow \
  .claude/skills/pdf-to-docx/scripts/dsl_to_docx.py \
  --workspace "$WORKSPACE" --output "$WORKSPACE/output.docx"
```

**If the script fails:** Read the traceback. Check that all `<image src="...">` paths are valid and XML files are well-formed. Fix and re-run.

## Step 6.5: Visual Verification — default SKIP

Print `Step 6.5 SKIP` and continue to Step 7. Only if the user explicitly asked for visual QA, read [optional-steps-reference.md](optional-steps-reference.md) section "Step 6.5".

## Step 7: Validation and Final Output

### 7a. Image existence check

```bash
export WORKSPACE
python3 << 'PYEOF'
import glob, os, re, sys

workspace = os.environ["WORKSPACE"]
missing = []

for path in sorted(glob.glob(os.path.join(workspace, "dsl", "page-*.xml"))):
    text = open(path, encoding="utf-8").read()
    for src in re.findall(r'src="([^"]+)"', text):
        candidates = [
            os.path.join(workspace, src),
            os.path.join(workspace, "ocr-output", "input", src),
        ]
        if not any(os.path.isfile(candidate) for candidate in candidates):
            missing.append((path, src))

if missing:
    print("Missing image references:")
    for path, src in missing:
        print(f"  {path}: {src}")
    sys.exit(1)

print("Image existence check OK")
PYEOF
```

### 7b. Content completeness audit — default SKIP

Only run if the user explicitly asked. See [optional-steps-reference.md](optional-steps-reference.md) section "Step 7b".

### 7c. Fix and re-generate if needed

If Step 7a found issues: edit the XML files, re-run Step 6 once. If unclear, STOP and report.

### 7d. Final output

```bash
cp "$WORKSPACE/output.docx" "$WORKSPACE/final-output.docx"
```

Report: Final DOCX at `$WORKSPACE/final-output.docx`, page count, and any remaining issues.

## Important Notes

- The five scripts are at `.claude/skills/pdf-to-docx/scripts/` — NEVER regenerate them
- OCR `bbox_2d` is normalized 0-1000 — always convert before use
- Cropped image idx is a sequential counter, NOT the JSON region index
- Use `uv run --with` for all Python script execution
- VLM review compares XML text vs page image (not two images)
- Pipeline uses LMStudio (`http://localhost:1234/v1/chat/completions`) with model `qwen3.5-122b-a10b`
- Heading styles default to blue — `dsl_to_docx.py` overrides to black
- `w:framePr` uses TWIPS (1 pt = 20 twips), NOT EMU
- `soffice` is optional — only needed for Step 6.5

## Workspace Structure

```
$WORKSPACE/
├── input.pdf
├── input-pdf-rendered-pngs/page-{N}.png
├── pdf-info.txt, pdf-fonts.txt, pdf-fulltext.txt
├── ocr-output/input/
│   ├── input.json, input.md, style-page-{N}.json, imgs/, layout_vis/
├── dsl/page-{N}.xml, review-page-{N}.json, visual-review-page-{N}.json
├── docx-rendered-pngs/ (Step 6.5 only)
├── output.docx
└── final-output.docx
```

## Reference Files

- **[docx-assembly-guide.md](docx-assembly-guide.md)**: DOCX assembly details
- **[optional-steps-reference.md](optional-steps-reference.md)**: Steps 5.5, 6.5, 7b detailed procedures
- **[known-issues.md](known-issues.md)**: Known issues, lessons learned, edge cases
