---
name: another-pure-pure-docx-translate-to-docx
description: Translates a DOCX (or DOC) file into a target language by directly parsing and modifying OOXML, preserving all formatting and structure. No dependency on PDF, OCR, or glm-ocr tools.
argument-hint: <input-file.docx> --lang <target-language> [--style <style-notes>] [--output <output-path>]
disable-model-invocation: false
allowed-tools:
  - Bash
  - Read
  - Write
  - Edit
  - Glob
  - Grep
---

# another-pure-pure-docx-translate-to-docx Skill

Translate a DOCX (or DOC) file directly into a target language by parsing and modifying OOXML. All formatting, structure, and embedded objects are preserved (zero-loss round-trip).

## Pipeline Checklist

Copy and track. Fill in values as you complete each step.

```
Translation Progress:
- [ ] Step 0: INPUT_PATH=___ TARGET_LANG=___
- [ ] Step 1: DOC→DOCX conversion (if needed): SKIP / DONE
- [ ] Step 2: Text extracted, segments=___
- [ ] Step 3: Translation complete, translations.json written
- [ ] Step 4: Translations applied → output.docx
- [ ] Step 5: Output verified, size=___ bytes
```

### Division of Responsibilities

- **Fixed script** `extract_docx_texts.py`: Extracts translatable text from DOCX → JSON
- **Claude agent (AI)**: Translates text segments in batches
- **Fixed script** `apply_docx_translations.py`: Applies translations back into DOCX

### Architecture

```
Input .doc/.docx
  → [soffice --headless] (if .doc → .docx conversion needed)
  → [extract_docx_texts.py] → texts.json
  → [Claude agent: translate in batches] → translations.json
  → [apply_docx_translations.py] → translated-output.docx
```

---

## Step 0: Parse Arguments & Validate Environment

### Parse `$ARGUMENTS`

Extract the following from `$ARGUMENTS`:

- `INPUT_PATH` (required): first positional argument — path to the DOCX or DOC file (e.g., `input/report.docx`)
- `TARGET_LANG` (required): value after `--lang` — target language and style (e.g., "English", "formal English", "繁體中文", "日本語")
- `STYLE_NOTES` (optional): value after `--style` — extra style/tone notes (e.g., "保留專業術語不翻譯", "使用台灣慣用譯法")
- `OUTPUT_PATH` (optional): value after `--output` — path for the output DOCX. Defaults to `<input-dir>/<input-name>.appdttd.docx` (placed next to the input file)

Examples of user invocations:

- `/another-pure-pure-docx-translate-to-docx report.docx --lang 繁體中文`
- `/another-pure-pure-docx-translate-to-docx contract.doc --lang "formal English" --style "keep legal terms in Latin"`
- `/another-pure-pure-docx-translate-to-docx input/manual.docx --lang 日本語 --output output/manual-ja.docx`

### Validate prerequisites

Run these checks and **abort with a clear message** if any fail:

```bash
command -v uv >/dev/null || { echo "MISSING: uv"; exit 1; }
```

If the input file has a `.doc` extension (not `.docx`), also check:

```bash
command -v soffice >/dev/null || { echo "MISSING: soffice (LibreOffice) — needed for .doc → .docx conversion"; exit 1; }
```

### Set up workspace

```bash
WORKSPACE_DIR="$(dirname "$INPUT_PATH")/workspace-$(date +%s)"
mkdir -p "$WORKSPACE_DIR"
```

---

## Step 1: Convert DOC to DOCX (if needed)

If the input file has a `.doc` extension:

```bash
soffice --headless --convert-to docx --outdir "$WORKSPACE_DIR" "$INPUT_PATH"
```

Update `DOCX_PATH` to point to the converted file. If already `.docx`, set `DOCX_PATH="$INPUT_PATH"`.

---

## Step 2: Extract Translatable Text

Run the extraction script:

```bash
uv run --with lxml \
  .claude/skills/another-pure-pure-docx-translate-to-docx/scripts/extract_docx_texts.py \
  --input "$DOCX_PATH" \
  --output "$WORKSPACE_DIR/texts.json"
```

Read and verify `$WORKSPACE_DIR/texts.json`:
- Check `total_segments` > 0
- Briefly review the segments to understand document structure

---

## Step 3: Translate Text Segments

**Claude agent performs translation directly** — no external API calls needed.

### 3.1 Read the translation prompt template

Read `.claude/skills/another-pure-pure-docx-translate-to-docx/translation-prompt.md` for the prompt format.

### 3.2 Batch and translate

Read `$WORKSPACE_DIR/texts.json` and process segments in batches of ~30:

**CRITICAL: YOU MUST TRANSLATE EVERY SINGLE BATCH. DO NOT STOP EARLY.**

Calculate total batches: `ceil(total_segments / 30)`. Track progress explicitly:

```
Total segments: N
Total batches:  ceil(N/30)
Batch 1/K: segments 0-29    → DONE ✓
Batch 2/K: segments 30-59   → DONE ✓
...
Batch K/K: segments ...      → DONE ✓
```

For each batch:

1. Prepare input JSON array with `id`, `type`, and `full_text` for each segment
2. Translate using the system/user prompt templates from `translation-prompt.md`
3. Collect translations with `id`, `type`, and `translated_text`
4. **Write each batch result to `$WORKSPACE_DIR/batch_N.json`** (one file per batch, for crash recovery)
5. Print progress: `Batch M/K complete: translated X segments`

**Translation approach:**
- Process ~30 segments per batch to maintain quality and prevent output truncation
- Use the element type (paragraph, header, table_cell, etc.) for context
- Apply the target language and style notes consistently
- Never insert new whitespace between a CJK character and an adjacent Latin letter, digit, or ASCII punctuation mark unless the source text already contains that exact whitespace
- Return translations in the exact format specified
- **ALL element types must be translated**: paragraphs, table_cell, textbox, header, footer, footnote, endnote, alt_text, chart — no type should be skipped

**STOP CONDITION: Only stop when ALL batches (1 through K) are complete. If you have translated fewer segments than total_segments, you are NOT done.**

### 3.3 Assemble and validate translations JSON

Merge all batch results into a single `translations.json`:

```json
{
  "translations": [
    {
      "id": "p[0]",
      "part": "word/document.xml",
      "type": "paragraph",
      "translated_text": "翻譯後的文字"
    }
  ]
}
```

**Important**: Each translation entry MUST include the `part` field (copied from the original segment in `texts.json`). The `apply_docx_translations.py` script uses `part` to locate the correct XML file within the DOCX.

Write the merged result to `$WORKSPACE_DIR/translations.json`.

### 3.4 Validate translation completeness

**This step is MANDATORY — do not skip it.**

After writing `translations.json`, verify completeness:

1. Count translated segments: `len(translations.json["translations"])`
2. Compare to `texts.json["total_segments"]`
3. If counts don't match:
   - Identify which segment IDs from `texts.json` are missing in `translations.json`
   - Report: `WARNING: {missing_count} segments not translated. Missing IDs: [...]`
   - Create a new batch containing ONLY the missing segments
   - Translate the missing batch
   - Merge into `translations.json` and re-validate
   - Repeat until all segments are translated or 3 retry attempts are exhausted
4. Print final status: `Translation complete: {translated}/{total} segments`

---

## Step 4: Apply Translations to DOCX

Run the application script:

```bash
uv run --with lxml \
  .claude/skills/another-pure-pure-docx-translate-to-docx/scripts/apply_docx_translations.py \
  --input "$DOCX_PATH" \
  --translations "$WORKSPACE_DIR/translations.json" \
  --output "$OUTPUT_PATH"
```

---

## Step 5: Verify Output & Report

1. **Confirm** `$OUTPUT_PATH` exists and has non-zero size
2. **Run the extraction script on the OUTPUT DOCX** to spot-check:
   ```bash
   uv run --with lxml \
     .claude/skills/another-pure-pure-docx-translate-to-docx/scripts/extract_docx_texts.py \
     --input "$OUTPUT_PATH" \
     --output "$WORKSPACE_DIR/output_check.json"
   ```
   - Read `output_check.json` and sample 5-10 segments from different types (paragraph, table_cell, header, etc.)
   - Verify they are in the target language, not the source language
   - If significant source-language text remains, report as a warning
3. **Report** translation results to the user:
   - Number of segments translated vs total extracted
   - Translation coverage percentage by type (paragraph, table_cell, header, etc.)
   - Target language used
   - Output file path
   - Any warnings about incomplete coverage

---

## Error Handling

- **Input file not found**: Abort with clear message
- **Not a valid DOCX/ZIP**: Abort — file may be corrupted
- **soffice not found** (for .doc input): Abort — suggest installing LibreOffice
- **0 segments extracted**: Abort — document may be image-only or password-protected
- **Translation errors**: If a batch fails, report the error and continue with remaining batches
- **apply script fails**: Check translations JSON format and retry
- **Output DOCX is 0 bytes**: Likely malformed translations — check script output

## Important Notes

- All paths should use absolute paths internally
- Use `uv run --with lxml` for all Python script execution — do NOT install packages globally
- The extraction and application scripts operate directly on ZIP/XML — no python-docx dependency
- Only `w:t` text nodes and attribute values (alt text) are modified — all formatting is preserved
- The `part` field in translations JSON is critical — it tells the apply script which XML file to modify
- Segments are identified by structural path IDs (e.g., `tbl[0]:tr[1]:tc[2]:p[0]`), not XPath
- Empty/whitespace-only segments are automatically skipped during extraction
- Textboxes from both `wps:txbx` and `v:textbox` are supported

## Workspace Structure (during execution)

```
$WORKSPACE_DIR/
├── texts.json                  ← extracted text segments
├── translations.json           ← translated segments
└── (converted.docx)            ← only if input was .doc

$OUTPUT_PATH                    ← final translated DOCX
```
