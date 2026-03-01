# glm-ocr Project Memory

## Skills Available
- `/pdf-to-docx` — PDF → editable DOCX (full pipeline)
- `/docx-translate-to-docx` — translates pdf-to-docx workspace
- `/another-pure-pure-docx-translate-to-docx` — translates any DOCX directly via OOXML
- `/image-to-docx` — images → editable DOCX (NEW, implemented 2026-03-02)

## image-to-docx Skill (implemented 2026-03-02)

### Files Created
- `.claude/skills/image-to-docx/SKILL.md` — complete step-by-step flow
- `.claude/skills/image-to-docx/scripts/prepare_images.py` — Pillow-based image prep
- `.claude/skills/image-to-docx/scripts/consolidate_ocr_results.py` — merge per-page OCR
- `.claude/settings.local.json` — added `Skill(image-to-docx)` permission
- `AGENTS.md` — added Skill 4 section

### Key Architecture
- Bridge pattern: prepares workspace compatible with existing pdf-to-docx scripts
- Workspace: `<output>/<first-stem>-img-docx-workspace/`
- Steps 4-7 fully reuse pdf-to-docx scripts unchanged
- DPI fallback: 200 DPI when EXIF missing

### Path Contracts (hardcoded in reused scripts)
- `extract_styles.py` reads: `ocr-output/input/input.json`, `input-pdf-rendered-pngs/page-{N}.png`
- `build_page_dsl.py` reads: `ocr-output/input/input.json`, `ocr-output/input/style-page-{N}.json`
- `review_dsl.py` reads: `dsl/page-{N}.xml`, `input-pdf-rendered-pngs/page-{N}.png`
- Page numbering: 1-based filenames, 0-based in OCR JSON data array

### consolidate_ocr_results.py Logic
- glmocr single-image output format: `[[regions,...]]` → takes `data[0]`
- Crop image renaming: `cropped_page0_idx{M}.*` → `cropped_page{N-1}_idx{M}.*`

### Verified Working (2026-03-02)
- Full pipeline tested with image_input/20260302_010005.jpg → final-output.docx OK
- Ollama 0.7.4 had a bug causing glm-ocr:bf16 to return empty `\`\`\`markdown\`\`\`` — fixed by upgrading Ollama
- glmocr output dir for images: `<output>/<stem>/` (not `<output>/input/`), files named `<stem>.json`, `<stem>.md`
  - For pdf-to-docx skill this is under `ocr-output/input/` but image-to-docx wraps it at `ocr-output-pages/page-N/input/`
- review_dsl.py defaults heading space-before-pt to 12 if not specified (no need to add manually)
- Visual review suggested corrections: font-family serif→sans, title 36→48pt bold, body 12→14pt, line-spacing 1.0→1.5
