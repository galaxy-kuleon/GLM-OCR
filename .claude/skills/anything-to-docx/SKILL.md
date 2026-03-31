---
name: anything-to-docx
description: Converts any document input (PDF, DOC/DOCX, images, Markdown) to a high-fidelity editable DOCX, with optional translation. Handles scanned PDFs, photographed pages, broken OCR metadata. Use when the user wants to convert any file to DOCX or translate a document while preserving layout.
argument-hint: <input-path> [--lang <target-language>] [--style <style-notes>] [--glossary <file.json>] [--term "A=B,C=D"] [--output <dir>]
disable-model-invocation: false
allowed-tools:
  - Bash
  - Read
  - Write
  - Edit
  - Glob
  - Grep
---

# anything-to-docx Skill

Convert any document to high-fidelity editable DOCX. Optional translation preserving layout/style.

**CRITICAL — Variable Persistence**: Shell variables are lost between separate bash calls. Step 0 saves all variables to `.atd-env.sh`. Every bash block after Step 0 MUST begin with `source .atd-env.sh` on its first line. If you combine multiple blocks into one bash call, one `source` at the top is enough.

**CRITICAL — No Skipping**: ALWAYS run every step in sequential order. Do NOT skip steps even if the workspace already contains files from a previous run. Do NOT jump to the final output step.

## Step 0: Parse Arguments, Detect Route, Fail Fast, Build Glossary

### 0.1 Parse `$ARGUMENTS`

```
INPUT_PATH  = first positional argument (required)
TARGET_LANG = --lang value or empty (optional)
STYLE_NOTES = --style value or empty (optional)
GLOSSARY_FILE = --glossary path or empty (optional)
TERM_PAIRS  = --term value or empty (optional)
OUTPUT_DIR  = --output value or ./output (optional)
```

### 0.2 Validate input path and detect route

Run this exact block. Do not modify.

```bash
test -e "$INPUT_PATH" || { echo "ERROR: input not found: $INPUT_PATH"; exit 1; }

if [ -d "$INPUT_PATH" ]; then
  ROUTE="B"
  INPUT_KIND="directory"
else
  EXT="${INPUT_PATH##*.}"
  EXT_LOWER=$(printf '%s' "$EXT" | tr '[:upper:]' '[:lower:]')
  case "$EXT_LOWER" in
    doc|docx|md|markdown)
      ROUTE="A"
      INPUT_KIND="$EXT_LOWER"
      ;;
    pdf|jpg|jpeg|png|bmp|gif|webp|tiff)
      ROUTE="B"
      INPUT_KIND="$EXT_LOWER"
      ;;
    *)
      echo "ERROR: unsupported input type: $INPUT_PATH"
      exit 1
      ;;
  esac
fi

echo "Detected route: $ROUTE ($INPUT_KIND)"
```

**After running the detection block above, continue to Step 0.3 and 0.4 below. Do NOT skip ahead to Route A or Route B yet.**

### 0.3 Derive workspace, clean stale data, save variables

```bash
mkdir -p "$OUTPUT_DIR"

if [ -d "$INPUT_PATH" ]; then
  STEM=$(basename "$INPUT_PATH")
else
  STEM=$(basename "$INPUT_PATH")
  STEM="${STEM%.*}"
fi

WORKSPACE="$OUTPUT_DIR/${STEM}-atd-workspace"

# Clean stale workspace from previous runs
if [ -d "$WORKSPACE" ]; then
  echo "Cleaning stale workspace: $WORKSPACE"
  rm -rf "$WORKSPACE"
fi

mkdir -p "$WORKSPACE"
SCRIPT_ROOT=".claude/skills"

# Save all variables for subsequent bash calls
cat > .atd-env.sh << ENVEOF
export INPUT_PATH="$INPUT_PATH"
export ROUTE="$ROUTE"
export INPUT_KIND="$INPUT_KIND"
export STEM="$STEM"
export WORKSPACE="$WORKSPACE"
export SCRIPT_ROOT="$SCRIPT_ROOT"
export OUTPUT_DIR="$OUTPUT_DIR"
export TARGET_LANG="${TARGET_LANG:-}"
export STYLE_NOTES="${STYLE_NOTES:-}"
ENVEOF
echo "Variables saved to .atd-env.sh"
```

Validate tools and fixed scripts needed for the detected route. Run ONLY the block for your detected ROUTE. Do NOT run the other route's block.

**If `ROUTE="A"`:**

```bash
source .atd-env.sh
command -v uv >/dev/null || { echo "MISSING: uv"; exit 1; }
test -f "$SCRIPT_ROOT/shared/verify_step.py" || { echo "MISSING: verify_step.py"; exit 1; }
test -f "$SCRIPT_ROOT/another-pure-pure-docx-translate-to-docx/scripts/extract_docx_texts.py" || { echo "MISSING: extract_docx_texts.py"; exit 1; }
test -f "$SCRIPT_ROOT/another-pure-pure-docx-translate-to-docx/scripts/apply_docx_translations.py" || { echo "MISSING: apply_docx_translations.py"; exit 1; }

if [ "$INPUT_KIND" = "doc" ]; then
  command -v soffice >/dev/null || { echo "MISSING: soffice"; exit 1; }
fi

if [ "$INPUT_KIND" = "md" ] || [ "$INPUT_KIND" = "markdown" ]; then
  command -v pandoc >/dev/null || { echo "MISSING: pandoc"; exit 1; }
fi
```

**If `ROUTE="B"`:**

```bash
source .atd-env.sh
command -v uv >/dev/null || { echo "MISSING: uv"; exit 1; }
for script in \
  "$SCRIPT_ROOT/shared/verify_step.py" \
  "$SCRIPT_ROOT/anything-to-docx/scripts/create_pdf_image_info.py" \
  "$SCRIPT_ROOT/anything-to-docx/scripts/resize_images.py" \
  "$SCRIPT_ROOT/anything-to-docx/scripts/vlm_generate_dsl.py" \
  "$SCRIPT_ROOT/anything-to-docx/scripts/vlm_merge_dsl.py" \
  "$SCRIPT_ROOT/anything-to-docx/scripts/deterministic_merge.py" \
  "$SCRIPT_ROOT/anything-to-docx/scripts/extract_dsl_texts.py" \
  "$SCRIPT_ROOT/anything-to-docx/scripts/apply_dsl_translations.py" \
  "$SCRIPT_ROOT/anything-to-docx/scripts/normalize_translations.py" \
  "$SCRIPT_ROOT/image-to-docx/scripts/consolidate_ocr_results.py" \
  "$SCRIPT_ROOT/pdf-to-docx/scripts/dsl_to_docx.py" \
  "$SCRIPT_ROOT/pdf-to-docx/scripts/verify_docx_visual.py"; do
  test -f "$script" || { echo "MISSING: $script"; exit 1; }
done

if [ "$INPUT_KIND" = "pdf" ]; then
  command -v pdftocairo >/dev/null || { echo "MISSING: pdftocairo (install poppler-utils)"; exit 1; }
  command -v pdfinfo >/dev/null || { echo "MISSING: pdfinfo (install poppler-utils)"; exit 1; }
  command -v pdftotext >/dev/null || { echo "MISSING: pdftotext (install poppler-utils)"; exit 1; }
fi
```

Abort immediately if any check fails.

### 0.4 Build glossary (only if `TARGET_LANG` is non-empty)

**If `TARGET_LANG` is empty, skip this step entirely.**

```bash
source .atd-env.sh
python3 << 'PYEOF'
import json, os, sys
glossary = {}
glossary_file = os.environ.get("GLOSSARY_FILE", "")
if glossary_file and os.path.exists(glossary_file):
    glossary.update(json.load(open(glossary_file)))
term_pairs = os.environ.get("TERM_PAIRS", "")
if term_pairs:
    for pair in term_pairs.split(","):
        parts = pair.strip().split("=", 1)
        if len(parts) == 2:
            glossary[parts[0].strip()] = parts[1].strip()
out_path = os.path.join(os.environ["WORKSPACE"], "glossary.json")
json.dump(glossary, open(out_path, "w"), ensure_ascii=False, indent=2)
print(f"Glossary: {len(glossary)} terms")
PYEOF
```

### State Summary

After completing Step 0, verify variables are saved:

```bash
source .atd-env.sh
echo "INPUT_PATH=$INPUT_PATH  ROUTE=$ROUTE  INPUT_KIND=$INPUT_KIND"
echo "WORKSPACE=$WORKSPACE  STEM=$STEM  SCRIPT_ROOT=$SCRIPT_ROOT"
echo "TARGET_LANG=${TARGET_LANG:-<empty>}"
```

**NOW go to your detected route: if ROUTE="A" go to "Route A" section. If ROUTE="B" go to "Route B" section. Do NOT read the other route's section.**

---

## Route A: DOC / DOCX / Markdown

Use this route for `.doc`, `.docx`, `.md`, `.markdown` inputs.

### Route A Checklist

Print this checklist now. After each step, reprint it with `[x]` and filled values:

```
Route A Progress:
- [ ] A1: Convert to DOCX -> DOCX_PATH=___
- [ ] A2: Translation (if --lang set, else SKIP)
- [ ] A3: Final output -> path=___
```

### A1: Convert to DOCX (if needed)

**If input is `.doc`:**
```bash
source .atd-env.sh
soffice --headless --convert-to docx --outdir "$WORKSPACE" "$INPUT_PATH"
DOCX_PATH="$WORKSPACE/$(basename "$INPUT_PATH" .doc).docx"
echo "export DOCX_PATH=\"$DOCX_PATH\"" >> .atd-env.sh
```

**If input is `.md` or `.markdown`:**
```bash
source .atd-env.sh
pandoc "$INPUT_PATH" -o "$WORKSPACE/${STEM}.docx"
DOCX_PATH="$WORKSPACE/${STEM}.docx"
echo "export DOCX_PATH=\"$DOCX_PATH\"" >> .atd-env.sh
```

**If input is `.docx`:**
```bash
source .atd-env.sh
DOCX_PATH="$INPUT_PATH"
echo "export DOCX_PATH=\"$DOCX_PATH\"" >> .atd-env.sh
```

**Verify A1:**
```bash
source .atd-env.sh
uv run .claude/skills/shared/verify_step.py --step A1 --workspace "$WORKSPACE" --docx-path "$DOCX_PATH"
```

### A2: Translation (if `TARGET_LANG` is set)

**If `TARGET_LANG` is empty, skip to A3.**

**A2a: Extract translatable text:**
```bash
source .atd-env.sh
uv run --with lxml \
  .claude/skills/another-pure-pure-docx-translate-to-docx/scripts/extract_docx_texts.py \
  --input "$DOCX_PATH" --output "$WORKSPACE/texts.json"
```

**Verify A2a:**
```bash
source .atd-env.sh
uv run .claude/skills/shared/verify_step.py --step A2a --workspace "$WORKSPACE"
```

**A2b: Translate** — go to [Translation Procedure](#translation-procedure) section below and follow T1-T5. When T5 passes, return HERE and continue to A2c.

**A2c: Apply translations:**
```bash
source .atd-env.sh
uv run --with lxml \
  .claude/skills/another-pure-pure-docx-translate-to-docx/scripts/apply_docx_translations.py \
  --input "$DOCX_PATH" --translations "$WORKSPACE/translations.json" \
  --output "$WORKSPACE/translated-output.docx"
```

**Verify A2c:**
```bash
source .atd-env.sh
uv run .claude/skills/shared/verify_step.py --step A2c --workspace "$WORKSPACE"
```

### A3: Final Output

```bash
source .atd-env.sh
if [ -f "$WORKSPACE/translated-output.docx" ]; then
  cp "$WORKSPACE/translated-output.docx" "$WORKSPACE/final-output.docx"
else
  cp "$DOCX_PATH" "$WORKSPACE/final-output.docx"
fi
ls -lh "$WORKSPACE/final-output.docx"
```

Report: `$WORKSPACE/final-output.docx`. **Route A done. STOP here.**

## Route B: PDF / Image -> VLM+OCR Pipeline

Use this route for `.pdf`, image files, or directories of images.

**ALWAYS run every step B0 through B8 in order. Do NOT skip steps. Do NOT reuse data from previous runs.**

### Route B Checklist

Print this checklist now. After each step, reprint it with `[x]` and filled values.

```
Route B Progress:
- [ ] B0: VLM_PROFILE=___
- [ ] B1: PAGE_COUNT=___
- [ ] B2: OCR pages=___ regions=___
- [ ] B3: VLM XML count=___
- [ ] B4: Merged XML count=___
- [ ] B5: output.docx size=___ bytes
- [ ] B6: SKIP / DONE
- [ ] B7: SKIP / DONE
- [ ] B8: final-output.docx path=___
```

### B0: Detect Model Profile

Run this exact block. Do not modify.

```bash
source .atd-env.sh
VLM_PROFILE="${VLM_MODEL_PROFILE:-weak}"
case "${VLM_MODEL:-}" in
  *gpt-4*|*claude*|*gemini*|*70b*|*70B*|*72b*|*72B*)
    VLM_PROFILE="strong"
    ;;
esac
export VLM_MODEL_PROFILE="$VLM_PROFILE"
echo "export VLM_PROFILE=\"$VLM_PROFILE\"" >> .atd-env.sh
echo "export VLM_MODEL_PROFILE=\"$VLM_PROFILE\"" >> .atd-env.sh
echo "Model profile: $VLM_PROFILE"
```

### B1: Prepare Input Images

**If input is a PDF:**

```bash
source .atd-env.sh
mkdir -p "$WORKSPACE/input-images"
pdftocairo -png -r 220 "$INPUT_PATH" "$WORKSPACE/input-images/page"

# Normalize zero-padded filenames (page-01.png -> page-1.png)
for f in "$WORKSPACE/input-images"/page-*.png; do
  base=$(basename "$f" .png)
  num=$(echo "$base" | sed 's/page-0*//')
  [ "$base" != "page-${num}" ] && mv "$f" "$WORKSPACE/input-images/page-${num}.png"
done

PAGE_COUNT=$(ls "$WORKSPACE/input-images"/page-*.png | wc -l | tr -d ' ')

pdfinfo "$INPUT_PATH" > "$WORKSPACE/pdf-info.txt"
pdftotext "$INPUT_PATH" "$WORKSPACE/pdf-fulltext.txt"

uv run .claude/skills/anything-to-docx/scripts/create_pdf_image_info.py \
  --workspace "$WORKSPACE" --pdf-info "$WORKSPACE/pdf-info.txt" --dpi 220

echo "export PAGE_COUNT=\"$PAGE_COUNT\"" >> .atd-env.sh
echo "Page count: $PAGE_COUNT"
```

**If input is image file(s) or a directory:**

```bash
source .atd-env.sh
python3 << 'PYEOF'
import json, os, re, sys

input_path = os.environ["INPUT_PATH"]
workspace = os.environ["WORKSPACE"]
exts = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".webp", ".tiff"}

def natural_key(path: str):
    name = os.path.basename(path)
    return [int(part) if part.isdigit() else part.lower() for part in re.split(r"(\d+)", name)]

if os.path.isdir(input_path):
    paths = []
    for name in os.listdir(input_path):
        path = os.path.join(input_path, name)
        if os.path.isfile(path) and os.path.splitext(name)[1].lower() in exts:
            paths.append(os.path.abspath(path))
    paths.sort(key=natural_key)
else:
    abs_path = os.path.abspath(input_path)
    if os.path.splitext(abs_path)[1].lower() not in exts:
        print(f"ERROR: unsupported image extension: {abs_path}", file=sys.stderr)
        sys.exit(1)
    paths = [abs_path]

if not paths:
    print(f"ERROR: no supported images found in {input_path}", file=sys.stderr)
    sys.exit(1)

out_path = os.path.join(workspace, "image-paths.json")
with open(out_path, "w", encoding="utf-8") as f:
    json.dump(paths, f, ensure_ascii=False, indent=2)

print(f"Collected {len(paths)} image paths")
PYEOF

IMAGES=$(python3 -c "import json; print(','.join(json.load(open('$WORKSPACE/image-paths.json', encoding='utf-8'))))")

uv run --with Pillow \
  .claude/skills/anything-to-docx/scripts/resize_images.py \
  --images "$IMAGES" --workspace "$WORKSPACE"

PAGE_COUNT=$(python3 -c "import json; print(json.load(open('$WORKSPACE/image-info.json'))['page_count'])")
echo "export PAGE_COUNT=\"$PAGE_COUNT\"" >> .atd-env.sh
echo "Page count: $PAGE_COUNT"
```

**Verify B1:**
```bash
source .atd-env.sh
uv run .claude/skills/shared/verify_step.py --step B1 --workspace "$WORKSPACE"
```

### B2: Run glm-ocr

**If input was a PDF:**
```bash
source .atd-env.sh
cp "$INPUT_PATH" "$WORKSPACE/input.pdf"
uv run glmocr parse "$WORKSPACE/input.pdf" --output "$WORKSPACE/ocr-output/"
```

**If input was image(s):**
```bash
source .atd-env.sh
for N in $(seq 1 $PAGE_COUNT); do
  mkdir -p "$WORKSPACE/ocr-input/page-$N"
  cp "$WORKSPACE/input-images/page-$N.png" "$WORKSPACE/ocr-input/page-$N/input.png"
  uv run glmocr parse "$WORKSPACE/ocr-input/page-$N/input.png" \
    --output "$WORKSPACE/ocr-output-pages/page-$N/"
done

uv run .claude/skills/image-to-docx/scripts/consolidate_ocr_results.py \
  --workspace "$WORKSPACE" --pages "$PAGE_COUNT"
```

**IMPORTANT**: Run glmocr commands ONE AT A TIME. Never in parallel.

**Verify B2:**
```bash
source .atd-env.sh
uv run .claude/skills/shared/verify_step.py --step B2 --workspace "$WORKSPACE"
```

### B3: VLM Generate XML DSL

Run this exact command. The script handles batch size and prompt selection based on `$VLM_PROFILE`.

```bash
source .atd-env.sh
uv run --with requests,Pillow,lxml \
  .claude/skills/anything-to-docx/scripts/vlm_generate_dsl.py \
  --workspace "$WORKSPACE" --pages "$PAGE_COUNT" --model-profile "$VLM_PROFILE"
```

**Verify B3:**
```bash
source .atd-env.sh
uv run .claude/skills/shared/verify_step.py --step B3 --workspace "$WORKSPACE"
```

**If verify fails:** Re-run the B3 command ONCE (maximum 1 retry). If the retry also fails, print `B3 FAILED -- VLM generation unsuccessful after 2 attempts` and STOP. Do NOT proceed to B4.

### B4: Merge

Run this exact block. The shell selects the right merge strategy based on `$VLM_PROFILE`.

```bash
source .atd-env.sh
if [ "$VLM_PROFILE" = "weak" ]; then
  uv run --with lxml \
    .claude/skills/anything-to-docx/scripts/deterministic_merge.py \
    --workspace "$WORKSPACE" --pages "$PAGE_COUNT"
else
  uv run --with requests,Pillow,lxml \
    .claude/skills/anything-to-docx/scripts/vlm_merge_dsl.py \
    --workspace "$WORKSPACE" --pages "$PAGE_COUNT"
fi
```

**Verify B4:**
```bash
source .atd-env.sh
uv run .claude/skills/shared/verify_step.py --step B4 --workspace "$WORKSPACE"
```

### B5: DSL to DOCX

```bash
source .atd-env.sh
uv run --with python-docx,lxml,Pillow \
  .claude/skills/pdf-to-docx/scripts/dsl_to_docx.py \
  --workspace "$WORKSPACE" --output "$WORKSPACE/output.docx"
```

**Verify B5:**
```bash
source .atd-env.sh
uv run .claude/skills/shared/verify_step.py --step B5 --workspace "$WORKSPACE"
```

### B6: Visual Verification -- default SKIP

Output `B6 SKIP` and go directly to B7. Do NOT read route-b-reference.md unless the user explicitly asked for visual QA.

### B7: Translation (if `TARGET_LANG` is set)

**If `TARGET_LANG` is empty, skip to B8.**

**B7a: Extract translatable text:**
```bash
source .atd-env.sh
uv run --with lxml \
  .claude/skills/anything-to-docx/scripts/extract_dsl_texts.py \
  --workspace "$WORKSPACE" --pages "$PAGE_COUNT" \
  --output "$WORKSPACE/texts.json"
```

**Verify B7a:**
```bash
source .atd-env.sh
uv run .claude/skills/shared/verify_step.py --step B7a --workspace "$WORKSPACE"
```

**B7b: Translate** -- go to [Translation Procedure](#translation-procedure) section below and follow T1-T5. When T5 passes, return HERE and continue to B7c.

**B7c: Apply translations to XML DSL:**
```bash
source .atd-env.sh
uv run --with lxml \
  .claude/skills/anything-to-docx/scripts/apply_dsl_translations.py \
  --workspace "$WORKSPACE" --pages "$PAGE_COUNT" \
  --translations "$WORKSPACE/translations.json" \
  --output-dir dsl-translated
```

**B7d: Generate translated DOCX:**
```bash
source .atd-env.sh
uv run --with python-docx,lxml,Pillow \
  .claude/skills/pdf-to-docx/scripts/dsl_to_docx.py \
  --workspace "$WORKSPACE" --dsl-dir "dsl-translated" \
  --output "$WORKSPACE/translated-output.docx"
```

**Verify B7d:**
```bash
source .atd-env.sh
uv run .claude/skills/shared/verify_step.py --step B7d --workspace "$WORKSPACE"
```

### B8: Final Output

```bash
source .atd-env.sh
if [ -f "$WORKSPACE/translated-output.docx" ]; then
  cp "$WORKSPACE/translated-output.docx" "$WORKSPACE/final-output.docx"
else
  cp "$WORKSPACE/output.docx" "$WORKSPACE/final-output.docx"
fi
ls -lh "$WORKSPACE/final-output.docx"
```

Report: output path, page count, translation status. **Route B done. STOP here.**

## Translation Procedure

**CRITICAL: YOU (the AI agent reading this) translate the text yourself. Do NOT create Python scripts that perform translation, do NOT call external APIs, and do NOT call LMStudio for translation.**

**CJK spacing rule**: never insert new whitespace between a CJK character and an adjacent Latin letter, digit, or ASCII punctuation mark unless the source text already contains that exact whitespace.

**Read and follow [translation-prompt.md](translation-prompt.md) steps T1 through T5 EXACTLY.**

**After T5 passes**: return to the calling step (A2c or B7c).

## Reference Files

- **[xml-dsl-reference.md](xml-dsl-reference.md)**: XML DSL element schema (all attributes and elements)
- **[translation-prompt.md](translation-prompt.md)**: Translation system/user prompt templates
- **[route-b-reference.md](route-b-reference.md)**: B6 visual verification, workspace structure, environment variables, error handling
