# Known Issues & Lessons Learned

## Contents

- [Numbered/bulleted lists](#numberedbulleted-lists)
- [Table cell line breaks](#table-cell-line-breaks)
- [CJK punctuation](#cjk-punctuation)
- [Heading colors](#heading-colors)
- [Text boxes](#text-boxes)
- [Text box detection](#text-box-detection)
- [Image regions and text-box detection](#image-regions-and-text-box-detection)
- [Table cell-level styling](#table-cell-level-styling)
- [Keyword-level cell styling](#keyword-level-cell-styling)
- [Cropped image naming](#cropped-image-naming)

---

## Numbered/bulleted lists

- OCR content already contains list markers (e.g., `1. Item`, `- Item`)
- `build_page_dsl.py` does NOT apply Word `List Number` or `List Bullet` styles — prevents double numbering
- List items are rendered as plain paragraphs with the marker text preserved as-is

## Table cell line breaks

- OCR HTML `<table>` content MAY contain `<br>` — `build_page_dsl.py` preserves them as `\n`
- Many cells show multi-line text due to column-width wrapping, not explicit breaks
- `dsl_to_docx.py` sets `autofit = False` + explicit column widths for natural wrapping

## CJK punctuation

- Assembly must NOT convert full-width to half-width (e.g., keep `（` not `(`)

## Heading colors

- python-docx defaults to blue — `dsl_to_docx.py` overrides all headings to black

## Text boxes

- **Use `w:framePr` + `w:pBdr`** for floating text with borders
- **NEVER use OOXML `wp:anchor`** — produces files Word cannot open
- For side-by-side layouts, use invisible tables (most reliable)

## Text box detection

- VLM style extraction reports `tb=true` and `bd=true` for regions inside bordered text boxes
- `build_page_dsl.py` groups spatially adjacent `tb=true` regions into `<text-frame>` elements
- Grouping criteria: X overlap > 50%, Y gap < 80 (normalized 0-1000 coordinates)
- Multiple stacked text regions in the same box produce one `<text-frame>` with multiple `<paragraph>` children
- Side-by-side detection remains as fallback for regions without VLM `tb` data
- `dsl_to_docx.py` already fully supports `<text-frame>` with `has-border` and multiple paragraphs

## Image regions and text-box detection

- VLM may incorrectly mark image regions as `tb=true` when the image has a visible border/frame
- `extract_styles.py` strips `tb`/`bd` from image regions (preventive)
- `build_page_dsl.py` filters `detect_textbox_regions_from_vlm()` to only consider `label="text"` (defensive)
- Without these guards, image regions get pulled out of document flow and appended at end of page

## Table cell-level styling

- OCR treats an entire table as one region, losing per-cell color information
- `extract_styles.py` performs a second VLM call per table to recover cell/column/row colors
- `cell_overrides` is optional and backward-compatible — if absent, region-level color applies uniformly
- Header rows no longer get hardcoded `bg-color="F0F0F0"` — background colors come from VLM detection only
- `dsl_to_docx.py` already supports `color-rgb` and `bg-color` on `<cell>` elements

## Keyword-level cell styling

- VLM detects short words/substrings within table cells that have different styling
- `keyword_styles` in `cell_overrides` maps `(row, col, keyword)` to per-keyword style overrides
- `build_page_dsl.py` splits cell text by keyword matches and creates `<run>` children inside `<cell>`
- `dsl_to_docx.py` detects `<run>` children in `<cell>` and renders each with own formatting
- When no `keyword_styles` are present, cells render with uniform styling as before
- Graceful degradation: if VLM misses keyword styling, falls back to region-level color

## Cropped image naming

- `cropped_page{N}_idx{M}.jpg` where M is sequential counter per page, NOT region index
