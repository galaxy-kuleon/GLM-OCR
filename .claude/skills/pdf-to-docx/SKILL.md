---
name: pdf-to-docx
description: Converts a PDF file to a high-fidelity editable DOCX. Use when the user wants to convert a PDF to Word, recreate a PDF as DOCX, or produce an editable document from a scanned/digital PDF. Automates OCR, layout detection, XML DSL generation, VLM review, and deterministic DOCX assembly.
argument-hint: <pdf-path> [--output <dir>]
disable-model-invocation: false
allowed-tools: Bash, Read, Write, Edit, Glob, Grep
---

# pdf-to-docx Skill

Convert a PDF to a high-fidelity editable DOCX through automated OCR → style extraction → XML DSL → VLM review → DOCX assembly.

**Architecture**: Each page produces a `page-{N}.xml` DSL file. Five fixed Python scripts handle the pipeline — no dynamically generated assembly scripts.

## Step 0: Parse Arguments & Environment Check

### Parse `$ARGUMENTS`

Extract the following from `$ARGUMENTS`:

- `PDF_PATH` (required): first positional argument — path to the input PDF
- `OUTPUT_DIR` (optional): value after `--output`, default `./output`

Derive:

- `PDF_STEM`: filename without extension (e.g., `report` from `report.pdf`)
- `WORKSPACE`: `<OUTPUT_DIR>/<PDF_STEM>-docx-workspace`

### Validate prerequisites

Run these checks and **abort with a clear message** if any fail:

```bash
command -v pdftocairo || echo "MISSING: pdftocairo (install poppler-utils)"
command -v pdfinfo    || echo "MISSING: pdfinfo (install poppler-utils)"
command -v pdffonts   || echo "MISSING: pdffonts (install poppler-utils)"
command -v pdftotext  || echo "MISSING: pdftotext (install poppler-utils)"
command -v uv         || echo "MISSING: uv"
```

### Check Poe API key

```bash
# Check for POE_API_KEY in env or .env file
if [ -z "$POE_API_KEY" ]; then
  if [ -f ".env" ] && grep -q "POE_API_KEY" .env; then
    export POE_API_KEY=$(grep "POE_API_KEY" .env | cut -d'=' -f2 | tr -d '"' | tr -d "'")
  fi
fi
if [ -z "$POE_API_KEY" ]; then
  echo "WARNING: POE_API_KEY not found. Style extraction will use defaults only."
fi
```

### Check scripts exist

```bash
SKILL_DIR=".claude/skills/pdf-to-docx/scripts"
for script in extract_styles.py build_page_dsl.py review_dsl.py dsl_to_docx.py verify_docx_visual.py; do
  test -f "$SKILL_DIR/$script" || echo "MISSING: $SKILL_DIR/$script"
done
```

### Create workspace

```bash
mkdir -p "$WORKSPACE"/{dsl,input-pdf-rendered-pngs}
cp "$PDF_PATH" "$WORKSPACE/input.pdf"
```

**IMPORTANT**: The input PDF is always copied as `input.pdf`, so the OCR stem will always be `input`. All OCR output paths use this stem: `ocr-output/input/input.json`, etc.

---

## Step 1: PDF → Reference PNGs

Render each page of the input PDF to PNG at 200 DPI. These are the **ground truth** images used for VLM review.

```bash
pdftocairo -png -r 200 "$WORKSPACE/input.pdf" "$WORKSPACE/input-pdf-rendered-pngs/page"
```

Output: `$WORKSPACE/input-pdf-rendered-pngs/page-1.png`, `page-2.png`, ...

Note: for PDFs with ≥ 10 pages, pdftocairo zero-pads filenames (`page-01.png`, `page-02.png`, ...).

Count pages:

```bash
PAGE_COUNT=$(ls "$WORKSPACE/input-pdf-rendered-pngs"/page-*.png | wc -l | tr -d ' ')
```

---

## Step 2: PDF Metadata Extraction

Extract metadata needed for DOCX page setup and font mapping:

```bash
pdfinfo "$WORKSPACE/input.pdf" > "$WORKSPACE/pdf-info.txt"
pdffonts "$WORKSPACE/input.pdf" > "$WORKSPACE/pdf-fonts.txt"
pdftotext "$WORKSPACE/input.pdf" "$WORKSPACE/pdf-fulltext.txt"
```

Read `pdf-info.txt` to extract **page size** (in points). Parse the "Page size:" line:

```
Page size:      595.276 x 841.89 pts (A4)
```

Store as `PAGE_WIDTH` and `PAGE_HEIGHT` (in points).

---

## Step 3: OCR Parsing

Run glmocr to parse the PDF:

```bash
uv run glmocr parse "$WORKSPACE/input.pdf" --output "$WORKSPACE/ocr-output/"
```

**IMPORTANT**: Do NOT run multiple `uv run glmocr parse` commands in parallel — they will conflict.

Output directory: `$WORKSPACE/ocr-output/input/`

- `input.json` — structured OCR results (`List[List[Dict]]`, outer = pages 0-indexed, inner = regions)
- `input.md` — Markdown format
- `imgs/` — cropped region images
- `layout_vis/` — layout detection visualizations

### OCR JSON Structure

Each region:

```json
{
  "index": 0,
  "label": "text|table|formula|image",
  "native_label": "paragraph_title|text|table|image|doc_title|figure_title|vision_footnote|...",
  "content": "...",
  "bbox_2d": [x1, y1, x2, y2]
}
```

**CRITICAL**: `bbox_2d` values are **normalized 0–1000** (not pixels).

### Cropped image naming — CRITICAL

Image files are named `cropped_page{N}_idx{M}.jpg` where:

- `N` = page index (0-based)
- `M` = **sequential image counter** — NOT the region's `index` field

### Native label categories

| native_label        | label   | Meaning                                           |
| ------------------- | ------- | ------------------------------------------------- |
| `doc_title`         | text    | Document title (H1), content prefixed with `# `   |
| `paragraph_title`   | text    | Section heading (H2), content prefixed with `## ` |
| `text`              | text    | Body paragraph                                    |
| `figure_title`      | text    | Figure/image caption                              |
| `vision_footnote`   | text    | Footnote/endnote                                  |
| `table`             | table   | Table (HTML in content)                           |
| `display_formula`   | formula | Display math (LaTeX)                              |
| `image`             | image   | Image (content is null)                           |

---

## Step 4: Style Extraction (Poe AI)

Run the fixed script to extract styles per page:

```bash
uv run --with requests,Pillow \
  .claude/skills/pdf-to-docx/scripts/extract_styles.py \
  --workspace "$WORKSPACE" --pages "$PAGE_COUNT"
```

**What this does**: For each page, sends the page image + region summary to Poe AI (`gemini-3-flash`) and receives simplified style data (font size, bold, color, alignment). Results are saved as `$WORKSPACE/ocr-output/input/style-page-{N}.json`.

**Table cell-level style extraction**: For each table region, the script performs a second VLM call to detect non-default text colors and background colors at the column/row/cell level. This captures per-cell styling that would otherwise be lost when OCR treats the entire table as one region. Results are stored in the `cell_overrides` field of the style entry.

**Agent does NOT generate any script** — only executes the fixed script.

**Fallback**: If Poe API is unavailable, the script automatically uses default styles by `native_label`. The pipeline continues without interruption. Table cell-level extraction is best-effort — if it fails, the pipeline continues with region-level styles only.

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
  },
  {
    "region_index": 1,
    "font_size_pt": 9,
    "bold": false,
    "italic": false,
    "underline": false,
    "color_rgb": [0, 0, 0],
    "alignment": "left",
    "th": true,
    "cell_overrides": {
      "col_colors": [{"col": 5, "c": [204, 0, 0], "type": "text"}],
      "row_colors": [],
      "cell_colors": []
    }
  },
  {
    "region_index": 7,
    "font_size_pt": 10,
    "bold": false,
    "italic": false,
    "underline": false,
    "color_rgb": [0, 0, 0],
    "alignment": "left",
    "tb": true,
    "bd": true
  }
]
```

Additional optional fields:
- `tb` (boolean) — region is inside a text box or bordered frame
- `bd` (boolean) — text box has a visible border (only meaningful when `tb=true`)
- `bg_rgb` ([R,G,B]) — paragraph/heading background shading color (only if non-white)
- `border_style` ("single"/"double"/"none") — table border line style (only for table regions)
- `cell_overrides` (object) — table cell-level color overrides (only for table regions). Contains:
  - `col_colors`: columns with non-default text/background color
  - `row_colors`: rows with non-default text/background color
  - `cell_colors`: individual cells with non-default text/background color
  - Each entry has `c` (text color [R,G,B]), optional `bg` (background [R,G,B]), optional `text_bg` (text highlight [R,G,B]), and `type` ("text" or "bg")
  - `keyword_styles`: keyword-level styling within individual cells. Each entry has `row`, `col`, `keyword` (exact substring), optional `c` (text color), `bold`, `text_bg` (highlight color)

---

## Step 5: Build Per-Page XML DSL

For each page, convert OCR data + style data into an XML DSL file:

```bash
for N in $(seq 1 $PAGE_COUNT); do
  uv run --with lxml \
    .claude/skills/pdf-to-docx/scripts/build_page_dsl.py \
    --workspace "$WORKSPACE" --page $N \
    --page-width-pts "$PAGE_WIDTH" --page-height-pts "$PAGE_HEIGHT"
done
```

Output: `$WORKSPACE/dsl/page-1.xml`, `page-2.xml`, ...

**Agent does NOT generate any script** — only executes the fixed script.

**Table cell styling**: Header rows get bold + center alignment but NO hardcoded background color. Cell-level text colors and background colors are applied from `cell_overrides` in the style JSON (priority: cell-specific > column-level > row-level > region default). This preserves the original PDF's per-cell styling.

### XML DSL Schema

Each page produces a self-contained XML file. See the XML DSL reference below for element types:

- `<heading>` — document/section titles
- `<paragraph>` — body text, lists, captions, footnotes, formulas
- `<table>` — data tables with rows/cells
- `<image>` — cropped images from OCR
- `<text-frame>` — floating/positioned text
- `<side-by-side>` — parallel column layouts

### Quick sanity check

After building all DSL files, verify they exist and are non-empty:

```bash
ls -la "$WORKSPACE/dsl/"page-*.xml
```

---

## Step 5.5: VLM Review of XML DSL

Run the review script to compare each page's XML DSL against the original page image:

```bash
uv run --with requests,lxml,Pillow \
  .claude/skills/pdf-to-docx/scripts/review_dsl.py \
  --workspace "$WORKSPACE" --pages "$PAGE_COUNT"
```

**What this does**: Compares XML DSL against page images using Poe AI.

**Multi-page mode** (≤5 pages): All page PNGs + all XML DSLs are sent in a single API call. This enables cross-page consistency checking (font sizes, colors, column widths). The VLM responds with a JSON object keyed by page number. Timeout is 180 seconds for multi-image calls.

**Per-page fallback** (>5 pages or multi-page failure): Each page is reviewed independently in separate API calls with 120-second timeout.

Output: `$WORKSPACE/dsl/review-page-{N}.json` (per page, regardless of mode)

### Handle review results

For each page with non-empty issues:

1. **Read** `review-page-{N}.json`
2. **Analyze** each issue:
   - `missing_text` → add the missing content to `page-{N}.xml`
   - `wrong_style` → update the relevant `<run>` attributes
   - `wrong_order` → reorder elements in `page-{N}.xml`
   - `missing_image` → verify image path and add `<image>` element
   - `extra_content` → remove the extra element
   - `missing_textframe` → text with a visible border box in the image is rendered as plain paragraph; wrap it in `<text-frame has-border="true">`
   - `wrong_textframe` → text-frame exists but has wrong attributes (e.g. `has-border` should be `"true"` but is `"false"`)
   - `cross_page_inconsistency` → style differs between pages for the same type of content (e.g., font size 11pt on page 1 but 12pt on page 2); fix by aligning to the correct value
3. **Edit** `page-{N}.xml` directly using the Edit tool
4. **Optional**: Re-run review for pages with many fixes (maximum 1 re-review)

**Key advantage**: VLM reviews XML text (not comparing two images), which is much simpler for weak models.

---

## Step 6: DSL → DOCX

Convert all page XML DSL files into a single DOCX:

```bash
uv run --with python-docx,lxml,Pillow \
  .claude/skills/pdf-to-docx/scripts/dsl_to_docx.py \
  --workspace "$WORKSPACE" --output "$WORKSPACE/output.docx"
```

**Deterministic**: Same XML always produces the same DOCX. No dynamic code generation.

If the script fails, read the error and check:
- Are all `<image src="...">` paths valid?
- Are XML files well-formed?
- Fix any issues in the XML files and re-run.

---

## Step 6.5: DOCX Visual Verification

Render the DOCX to PNGs and compare against original PDF page images:

```bash
uv run --with requests,Pillow \
  .claude/skills/pdf-to-docx/scripts/verify_docx_visual.py \
  --workspace "$WORKSPACE" --pages "$PAGE_COUNT" \
  --docx "$WORKSPACE/output.docx"
```

**What this does**:
1. Converts output.docx → PDF via soffice
2. Renders PDF → PNGs at 200 DPI (docx-rendered-pngs/)
3. Detects page count mismatch (input PDF pages vs DOCX-rendered pages)
4. VLM compares pages with appropriate strategy based on page count match/mismatch
5. Reports visual differences

Output: `$WORKSPACE/dsl/visual-review-page-{N}.json` (N = input page number)

**Page count mismatch handling**: Input PDF may have 2 pages but DOCX may render as 5 pages due to font rendering, table reflow, line spacing differences, etc. When page counts differ, the script sends all images together and asks VLM to map content across pages. A `page_count_mismatch` issue is reported first.

### Handle visual review results

For each page with non-empty issues:
1. Read `visual-review-page-{N}.json`
2. If `page_count_mismatch` is present, investigate root cause (font size too large, table too wide, margins wrong) and fix in `page-{N}.xml`
3. Analyze other differences and edit `page-{N}.xml`
4. Re-run Step 6 (`dsl_to_docx.py`) to regenerate DOCX
5. Re-run Step 6.5 to verify fixes (maximum 1 re-verification)

**Prerequisite**: soffice must be available. If not, skip this step.

---

## Step 7: Structure Validation & Final Output

### 7a. Image existence check

Verify all `<image src="...">` paths in the DSL files exist:

```bash
for xml in "$WORKSPACE/dsl/"page-*.xml; do
  rg -o 'src="[^"]*"' "$xml" | while read -r match; do
    src=$(echo "$match" | sed 's/src="//;s/"//')
    test -f "$WORKSPACE/$src" || echo "MISSING: $WORKSPACE/$src"
  done
done
```

### 7b. Content completeness check

Compare key text fragments from `pdf-fulltext.txt` **and** `input.md` (OCR Markdown output) against the XML DSL content:

```bash
# Check against pdftotext output
head -20 "$WORKSPACE/pdf-fulltext.txt" | while IFS= read -r line; do
  line=$(echo "$line" | xargs)  # trim
  if [ -n "$line" ] && [ ${#line} -gt 5 ]; then
    if ! rg -q "$line" "$WORKSPACE/dsl/"; then
      echo "POSSIBLE MISSING (pdftotext): $line"
    fi
  fi
done
```

```bash
# Cross-check against OCR Markdown output (input.md) for more reliable text source
head -30 "$WORKSPACE/ocr-output/input/input.md" | while IFS= read -r line; do
  line=$(echo "$line" | xargs)  # trim
  # Skip markdown formatting lines (headers, separators, etc.)
  if [ -n "$line" ] && [ ${#line} -gt 5 ] && ! echo "$line" | rg -q '^[#|>-]'; then
    if ! rg -q "$line" "$WORKSPACE/dsl/"; then
      echo "POSSIBLE MISSING (OCR md): $line"
    fi
  fi
done
```

**Why two sources**: `pdf-fulltext.txt` (from `pdftotext`) captures raw text from the PDF, while `input.md` (from glmocr) captures OCR-recognized text including table content. Comparing against both provides better coverage for detecting missing content.

### 7c. Fix and re-generate if needed

If issues are found:
1. Edit the corresponding `page-{N}.xml` files
2. Re-run Step 6: `dsl_to_docx.py`

### 7d. Final output

```bash
cp "$WORKSPACE/output.docx" "$WORKSPACE/final-output.docx"
```

Report to the user:
- Final DOCX location: `$WORKSPACE/final-output.docx`
- Number of pages processed
- Summary of any remaining issues
- Workspace location for inspection

---

## Error Handling

- **glmocr parse fails**: Check if input PDF is valid; try re-running with `--log-level DEBUG`
- **extract_styles.py fails**: Script auto-falls back to defaults. Check error output.
- **build_page_dsl.py fails**: Check OCR JSON structure. Fix and retry.
- **review_dsl.py fails**: Script auto-writes empty reviews. VLM review is optional.
- **dsl_to_docx.py fails**: Read the Python traceback. Usually caused by malformed XML or missing images. Fix XML and retry.
- **verify_docx_visual.py fails**: Requires `soffice` (LibreOffice). If not available, step is automatically skipped with empty reviews.
- **soffice not available**: Not required for core pipeline. Only needed for Step 6.5 (DOCX Visual Verification).

## Important Notes

- All paths should use absolute paths internally
- The five scripts are at `.claude/skills/pdf-to-docx/scripts/` — NEVER regenerate them
- OCR `bbox_2d` is normalized 0–1000 — always convert before use
- Tables in OCR output are HTML strings — parsed by `lxml.html` in `build_page_dsl.py`
- Cropped image idx is a sequential counter, NOT the JSON region index
- Use `uv run --with` for all Python script execution — do NOT install packages globally
- XML DSL files are human-readable and can be manually edited for fine-tuning
- VLM review compares XML text vs page image (not two images) — works well with weak models
- Style extraction uses simplified prompts (5-6 fields) instead of complex nested JSON
- Pipeline does NOT require ollama — uses Poe AI (`https://api.poe.com/v1/chat/completions`)
- If `POE_API_KEY` is not set, pipeline still works with default styles
- Heading styles in python-docx default to blue — `dsl_to_docx.py` overrides to black
- `w:framePr` uses TWIPS (1 pt = 20 twips), NOT EMU
- `soffice` (LibreOffice) is optional — only needed for Step 6.5 (DOCX Visual Verification). If not available, step is skipped
- Paragraph/heading background shading (bg-color) and text-level highlight (text-bg-color) are supported in DSL and DOCX output
- Table border styles support `single`, `double`, and `none` — detected by VLM and applied in DOCX

## Workspace Structure

```
$WORKSPACE/
├── input.pdf
├── input-pdf-rendered-pngs/
│   ├── page-1.png, page-2.png, ...
├── pdf-info.txt, pdf-fonts.txt, pdf-fulltext.txt
├── ocr-output/input/
│   ├── input.json
│   ├── input.md
│   ├── style-page-1.json, style-page-2.json, ...
│   ├── imgs/
│   └── layout_vis/
├── dsl/
│   ├── page-1.xml, page-2.xml, ...
│   ├── review-page-1.json, review-page-2.json, ...
│   ├── visual-review-page-1.json, visual-review-page-2.json, ...
├── docx-rendered-pngs/          (Step 6.5, if soffice available)
│   ├── page-1.png, page-2.png, ...
├── output.docx
└── final-output.docx
```

## XML DSL Reference

### Page element

```xml
<page number="1" width-pts="595.276" height-pts="841.89"
      margin-top-cm="1.27" margin-bottom-cm="1.27"
      margin-left-cm="1.27" margin-right-cm="1.27"
      font-latin="Arial" font-cjk="SimSun">
  <!-- child elements -->
</page>
```

### Element types

| XML Element | python-docx Operation | Source |
|-------------|----------------------|--------|
| `<heading level="N">` | `doc.add_heading("", N)` + black color | doc_title→1, paragraph_title→2 |
| `<paragraph>` | `doc.add_paragraph(style=...)` | native_label mapping |
| `<run>` | `para.add_run()` + font/color/bold/italic | style JSON + markdown |
| `<table>` | `doc.add_table()` + occupancy grid merge | OCR HTML table |
| `<cell>` | cell text + shading + borders; may contain `<run>` children for keyword-level styling | style defaults + keyword_styles |
| `<image>` | `doc.add_picture()` + bbox scaling | sequential counter |
| `<text-frame>` | `w:framePr` + `w:pBdr` (TWIPS) | bbox + floating detection |
| `<side-by-side>` | invisible table `w:val="none"` | parallel layout detection |

### Run attributes (all optional)

| Attribute | Default | Description |
|-----------|---------|-------------|
| `font-size-pt` | `11` | Font size |
| `bold` | `false` | Bold |
| `italic` | `false` | Italic |
| `underline` | `false` | Underline |
| `color-rgb` | `0,0,0` | Text color R,G,B |
| `superscript` | `false` | Superscript |
| `font-name` | (inherits page font) | Font name |

## Known Issues & Lessons Learned

### Numbered/bulleted lists

- OCR content already contains list markers (e.g., `1. Item`, `- Item`)
- `build_page_dsl.py` does NOT apply Word `List Number` or `List Bullet` styles — this prevents double numbering
- List items are rendered as plain paragraphs with the marker text preserved as-is

### Table cell line breaks

- OCR HTML `<table>` content MAY contain `<br>` — `build_page_dsl.py` preserves them as `\n`
- Many cells show multi-line text due to column-width wrapping, not explicit breaks
- `dsl_to_docx.py` sets `autofit = False` + explicit column widths for natural wrapping

### Full-width CJK punctuation

- Assembly must NOT convert full-width to half-width (e.g., `（` → `(`)

### Heading colors

- python-docx defaults to blue — `dsl_to_docx.py` overrides all headings to black

### Text boxes

- **Use `w:framePr` + `w:pBdr`** for floating text with borders
- **NEVER use OOXML `wp:anchor`** — produces files Word cannot open
- For side-by-side layouts, use invisible tables (most reliable)

### Text box detection

- VLM style extraction reports `tb=true` and `bd=true` for regions inside bordered text boxes
- `build_page_dsl.py` groups spatially adjacent `tb=true` regions into `<text-frame>` elements
- Grouping criteria: X overlap > 50%, Y gap < 80 (normalized 0-1000 coordinates)
- Multiple stacked text regions in the same box produce one `<text-frame>` with multiple `<paragraph>` children
- Side-by-side detection remains as fallback for regions without VLM `tb` data
- `dsl_to_docx.py` already fully supports `<text-frame>` with `has-border` and multiple paragraphs

### Image regions and text-box detection

- VLM may incorrectly mark image regions as `tb=true` when the image has a visible border/frame
- `extract_styles.py` strips `tb`/`bd` from image regions (preventive)
- `build_page_dsl.py` filters `detect_textbox_regions_from_vlm()` to only consider `label="text"` (defensive)
- Without these guards, image regions get pulled out of document flow and appended at end of page

### Table cell-level styling

- OCR treats an entire table as one region, losing per-cell color information
- `extract_styles.py` performs a second VLM call per table to recover cell/column/row colors
- `cell_overrides` is optional and backward-compatible — if absent, region-level color applies uniformly
- Header rows no longer get hardcoded `bg-color="F0F0F0"` — background colors come from VLM detection only
- `dsl_to_docx.py` already supports `color-rgb` and `bg-color` on `<cell>` elements (no changes needed)

### Keyword-level cell styling

- VLM detects short words/substrings within table cells that have different styling from surrounding text (e.g., one word is red+bold)
- `keyword_styles` in `cell_overrides` maps `(row, col, keyword)` to per-keyword style overrides
- `build_page_dsl.py` splits cell text by keyword matches and creates `<run>` children inside `<cell>` with per-keyword attributes
- `dsl_to_docx.py` detects `<run>` children in `<cell>` and renders each run with its own formatting
- When no `keyword_styles` are present (or VLM fails to detect them), cells render with uniform styling as before
- Graceful degradation: if VLM misses keyword styling, the cell falls back to region-level color (same as previous behavior)

### Cropped image naming

- `cropped_page{N}_idx{M}.jpg` where M is sequential counter per page, NOT region index
