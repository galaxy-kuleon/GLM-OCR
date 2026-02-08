---
name: docx-translate-to-docx
description: Translates a DOCX workspace (produced by pdf-to-docx) into a target language using LM Studio or Poe AI, preserving all layout, formatting, and styling. Operates on XML DSL files for reliable translation.
argument-hint: <workspace-path> --lang <target-language-and-style> [--style <extra-style-notes>]
disable-model-invocation: false
allowed-tools:
  - Bash
  - Read
  - Write
  - Edit
  - Glob
  - Grep
---

# docx-translate-to-docx Skill

Translate a pdf-to-docx workspace into a target language by operating on XML DSL files, preserving all layout, formatting, and styling.

### Division of Responsibilities

- **Translation model** (LM Studio `qwen/qwen3-4b-2507` or Poe AI): Only responsible for translating text content (receives source JSON → returns translated JSON)
- **Claude agent**: Responsible for generating `translate_content.py`, flow control, and verification
- **Fixed script** `dsl_to_docx.py`: Converts translated XML DSL → DOCX (no dynamic code generation)

---

## Step 0: Parse Arguments & Validate Environment

### Parse `$ARGUMENTS`

Extract the following from `$ARGUMENTS`:

- `WORKSPACE` (required): first positional argument — path to the pdf-to-docx workspace directory (e.g., `output/small-docx-workspace`)
- `TARGET_LANG` (required): value after `--lang` — target language and style (e.g., "English", "formal English", "en-uk", "vivid 繁體中文"). This string is embedded directly into the translation prompt.
- `STYLE_NOTES` (optional): value after `--style` — extra style/tone notes (e.g., "保留專業術語不翻譯", "使用台灣慣用譯法")

Examples of user invocations:

- `/docx-translate-to-docx output/small-docx-workspace --lang English`
- `/docx-translate-to-docx output/contract_template-docx-workspace --lang "formal English"`
- `/docx-translate-to-docx output/small-docx-workspace --lang 日本語 --style "丁寧語（です・ます調）"`

### Validate prerequisites

Run these checks and **abort with a clear message** if any fail:

```bash
command -v uv || echo "MISSING: uv"
curl -sf http://localhost:1234/v1/models > /dev/null || echo "MISSING: LM Studio not running on localhost:1234"
```

### Validate workspace

Verify the workspace exists and contains required files:

- `$WORKSPACE/ocr-output/input/input.json` — OCR structured data
- `$WORKSPACE/ocr-output/input/input.md` — Markdown full text
- `$WORKSPACE/dsl/page-*.xml` — XML DSL files (from pdf-to-docx pipeline)

**Abort** if any of these are missing. The existence of `dsl/page-*.xml` confirms the pdf-to-docx XML DSL pipeline completed successfully.

### Create translation workspace

```bash
mkdir -p "$WORKSPACE/dsl-translated"
```

---

## Step 1: Read Source Data

Read these files:

1. **`$WORKSPACE/ocr-output/input/input.json`** — full OCR structure
2. **`$WORKSPACE/ocr-output/input/input.md`** — markdown full text (provides translation context)
3. **All `$WORKSPACE/dsl/page-*.xml`** — XML DSL files to translate

Understand the XML DSL structure:

- `<run>` elements contain translatable text (the `.text` content)
- `<cell>` elements contain translatable table cell text:
  - If `<cell>` has `<run>` children (keyword-level styling) → text is in each `<run>`, NOT in `cell.text`
  - If `<cell>` has no `<run>` children → text is in `cell.text` directly
- `<heading>` contains runs with heading text
- `<paragraph>` contains runs with paragraph text
- `<side-by-side>` → `<column>` → `<paragraph>` → `<run>` (side-by-side layout)
- `<text-frame>` → `<paragraph>` → `<run>` (floating text boxes)
- `<image>` may contain nested translatable text
- `<run is-math="true">` elements contain math formulas (LaTeX) — **do NOT translate**

---

## Step 2: Generate & Execute Translation Script

**Claude agent generates** `$WORKSPACE/translation/translate_content.py`. This Python script extracts text from XML DSL, sends to translation model, and writes translated XML DSL files.

### Script logic

The generated script must:

1. **Read each `page-{N}.xml`** and extract translatable text:
   - From `<run>` elements: `.text` content
   - **Skip** `<run>` elements with `is-math="true"` attribute (math formulas must not be translated)
   - From `<cell>` elements — use branching logic:
     - If `<cell>` has `<run>` children → extract each `<run>`'s text (xpath: `table[0]/row[0]/cell[0]/run[0]`)
     - If `<cell>` has no `<run>` children → extract `cell.text` (xpath: `table[0]/row[0]/cell[0]`)
   - Traverse `<side-by-side>` → `<column>` → `<paragraph>` → `<run>` (xpath: `side-by-side[0]/column[0]/paragraph[0]/run[0]`)
   - Skip non-text attributes (font-size, color, bold, etc.)

2. **Build translation payload per page**: Package extracted text into JSON, combine with translation prompt, and send to LM Studio:
   - Endpoint: `http://localhost:1234/v1/chat/completions`
   - Model: `qwen/qwen3-4b-2507`
   - **One API call per page**
   - System prompt includes: target language & style instructions, full document markdown as reference
   - User prompt includes: current page's text content as JSON array

   Input format per page:
   ```json
   [
     {"xpath": "heading[1]/run[1]", "text": "文件標題"},
     {"xpath": "paragraph[1]/run[1]", "text": "一般文字"},
     {"xpath": "table[1]/row[0]/cell[0]", "text": "表頭1"},
     {"xpath": "table[1]/row[1]/cell[2]/run[0]", "text": "關鍵字"},
     {"xpath": "side-by-side[0]/column[0]/paragraph[0]/run[0]", "text": "左欄文字"}
   ]
   ```

   Note: `<run is-math="true">` elements are excluded — math/formula content is never sent for translation.

   Expected output:
   ```json
   {
     "translations": [
       {"xpath": "heading[1]/run[1]", "translated_text": "Document Title"},
       {"xpath": "paragraph[1]/run[1]", "translated_text": "Normal text"},
       {"xpath": "table[1]/row[0]/cell[0]", "translated_text": "Header 1"},
       {"xpath": "table[1]/row[1]/cell[2]/run[0]", "translated_text": "Keyword"},
       {"xpath": "side-by-side[0]/column[0]/paragraph[0]/run[0]", "translated_text": "Left column text"}
     ]
   }
   ```

3. **Write translated XML**: For each page:
   - Deep copy the original `page-{N}.xml`
   - Replace text in matched `<run>` and `<cell>` elements using xpath lookup
   - For `<cell>` with `<run>` children: replace each `run.text`, NOT `cell.text`
   - For `<cell>` without `<run>` children: replace `cell.text` directly
   - Do NOT modify `<run is-math="true">` elements (math formulas are untouched)
   - Preserve ALL attributes (font-size, color, bold, alignment, etc.)
   - Save to `$WORKSPACE/dsl-translated/page-{N}.xml`

4. **Error handling**:
   - Timeout: 120 seconds per API call
   - Retry: up to 2 retries per page on failure
   - Use `response_format` with `json_schema` for structured output if model supports it
   - If structured output fails, fall back to manual JSON parsing

### Prompt template

Read the prompt template from `translation-prompt.md` (located in the same skill directory as this file) and fill in the template variables:
- `{target_language}` → `TARGET_LANG`
- `{style_notes}` → `STYLE_NOTES` or empty
- `{full_document_markdown}` → content of `input.md`

### Execute

```bash
uv run --with requests,lxml "$WORKSPACE/translation/translate_content.py"
```

If it fails, read the error, fix the script, and retry (up to 3 times).

---

## Step 3: Generate Translated DOCX

Use the **fixed `dsl_to_docx.py` script** with the translated DSL directory:

```bash
uv run --with python-docx,lxml,Pillow \
  .claude/skills/pdf-to-docx/scripts/dsl_to_docx.py \
  --workspace "$WORKSPACE" \
  --dsl-dir "dsl-translated" \
  --output "$WORKSPACE/translation/translated-output.docx"
```

**No dynamic assembly script generation needed** — the fixed `dsl_to_docx.py` reads any DSL directory.

---

## Step 4: Verify Output

1. **Confirm** `$WORKSPACE/translation/translated-output.docx` exists and has non-zero size
2. **Report** translation results to the user:
   - Number of pages translated
   - Number of text elements translated
   - Target language used
   - Output file path: `$WORKSPACE/translation/translated-output.docx`
   - Remind user to check translation quality and open the DOCX to verify layout

---

## Error Handling

- **LM Studio not running**: Abort with message to start LM Studio with `qwen/qwen3-4b-2507` model
- **LM Studio API error / timeout**: Retry up to 2 times per page with 10s delay
- **JSON parse error from LM**: Strip markdown code fences, try `json.loads` again; if still fails, retry
- **XML DSL files missing**: Abort — user needs to run pdf-to-docx first with the new XML DSL pipeline
- **translate_content.py execution fails**: Read traceback, fix script, retry up to 3 times
- **dsl_to_docx.py execution fails**: Check translated XML is well-formed. Fix and retry.
- **Output DOCX is 0 bytes**: Likely malformed translated XML — check traceback

## Important Notes

- All paths should use absolute paths internally
- Translation operates on XML DSL text nodes — much simpler than modifying Python assembly scripts
- Style JSON files are NOT translated — visual styles are language-independent
- XML attributes (font-size, color, bold, alignment) are preserved exactly as-is
- Only `<run>` text content and `<cell>` text content are translated (except `<run is-math="true">` which is skipped)
- The `dsl_to_docx.py` script's `--dsl-dir` option allows pointing to `dsl-translated/`
- Use `uv run --with` for all Python script execution — do NOT install packages globally
- OCR `bbox_2d` values are normalized 0–1000, not pixels
- Cropped image idx is a sequential counter per page, not the JSON region index

## Workspace Structure (after translation)

```
$WORKSPACE/
├── dsl/                          ← original XML DSL
│   ├── page-1.xml, page-2.xml, ...
├── dsl-translated/               ← translated XML DSL
│   ├── page-1.xml, page-2.xml, ...
├── translation/
│   ├── translate_content.py      ← generated by agent
│   └── translated-output.docx   ← final translated DOCX
└── ...
```

## Advantages Over Previous Approach

| Previous | New |
|----------|-----|
| Find and modify `assemble_docx.py` | Operate on XML text nodes |
| Remove hardcoded text checks | No hardcoded checks in fixed scripts |
| Generate new assembly script | Use fixed `dsl_to_docx.py` with `--dsl-dir` |
| Fragile: depends on assembly script structure | Robust: XML text replacement is deterministic |
