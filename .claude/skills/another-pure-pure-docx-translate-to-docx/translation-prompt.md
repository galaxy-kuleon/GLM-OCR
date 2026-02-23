# Translation Prompt Template

Prompt template used for AI translation of DOCX document content.

## Usage

When translating text extracted from a DOCX file, use this template with the following variables:
- `{target_language}` — target language and style (from `--lang` argument)
- `{style_notes}` — extra style/tone notes (from `--style` argument, may be empty)
- `{batch_number}` — current batch number (1-based)
- `{total_batches}` — total number of batches
- `{texts_json}` — current batch's translatable text segments (see format below)

## Input Format (texts JSON)

Text is extracted at the paragraph level from DOCX XML. Each item includes an `id` for identification, a `type` describing the element kind, and the `full_text` content to translate.

```json
[
  {
    "id": "p[3]",
    "type": "paragraph",
    "full_text": "This is a paragraph with bold and normal text."
  },
  {
    "id": "tbl[0]:tr[1]:tc[2]:p[0]",
    "type": "table_cell",
    "full_text": "Table cell content"
  },
  {
    "id": "txbx[0]:p[0]",
    "type": "textbox",
    "full_text": "Text box content"
  },
  {
    "id": "header1:p[0]",
    "type": "header",
    "full_text": "Company Name"
  },
  {
    "id": "docPr[0]:descr",
    "type": "alt_text",
    "full_text": "Photo of the building entrance"
  },
  {
    "id": "chart:t[0]",
    "type": "chart",
    "full_text": "Revenue"
  }
]
```

### Element types

| Type | Description |
|------|-------------|
| `paragraph` | Body text, headings |
| `table_cell` | Table cell content |
| `textbox` | Text box / floating text frame |
| `header` | Page header |
| `footer` | Page footer |
| `footnote` | Footnote |
| `endnote` | Endnote |
| `alt_text` | Image alt text / title |
| `chart` | Chart text (axis labels, titles, data labels) |

## Output Format (AI returns)

The output is a JSON **object** with a `translations` array:

```json
{
  "translations": [
    {
      "id": "p[3]",
      "type": "paragraph",
      "translated_text": "這是一個包含粗體和普通文字的段落。"
    },
    {
      "id": "tbl[0]:tr[1]:tc[2]:p[0]",
      "type": "table_cell",
      "translated_text": "表格儲存格內容"
    },
    {
      "id": "txbx[0]:p[0]",
      "type": "textbox",
      "translated_text": "文字方塊內容"
    },
    {
      "id": "header1:p[0]",
      "type": "header",
      "translated_text": "公司名稱"
    }
  ]
}
```

The translation logic extracts the `translations` array from this wrapper object via `response["translations"]`, then maps each `id` back to the corresponding segment for downstream application.

## System Prompt Template

```
You are a professional document translator. Your task is to translate document content into {target_language}.

{style_notes}

## Rules

1. **Only translate text content** — do not modify id identifiers or type fields.
2. **Preserve proper nouns** — keep brand names, product names, and technical identifiers in their original form unless there is a widely accepted translation.
3. **Handle mixed languages** — if the source contains mixed languages, translate all text to the target language but keep proper nouns and technical terms as-is.
4. **Match the count** — your output `translations` array MUST have exactly the same number of items as the input, with matching ids.
5. **Output format** — respond with a JSON object containing a `translations` array. Each item has `id` (string, unchanged), `type` (string, unchanged), and `translated_text` (string).
6. **Do NOT translate empty strings** — if the text is empty, keep it empty.
7. **Do NOT translate math/formula content** — items containing LaTeX expressions or mathematical formulas should be returned as-is without translation.
8. **Preserve numbers and units** — keep numeric values, dates, units of measurement, and currency symbols in their conventional form for the target language.
9. **Context awareness** — consider the element type (paragraph, header, table_cell, etc.) to produce contextually appropriate translations.
```

## User Prompt Template

```
Translate the following text items (batch {batch_number}/{total_batches}) into {target_language}.

Input:
{texts_json}

Output a JSON object with a "translations" array. Each item has "id" (string, unchanged), "type" (string, unchanged), and "translated_text" (string).
```

## Batching Strategy

- Each batch contains approximately **50 segments** (adjustable based on total count)
- Segments are grouped by their `part` (XML source file) to maintain context locality
- All batches for a document share the same system prompt for consistency

## Key Design Decisions

| Decision | Rationale |
|----------|-----------|
| id-based segment identification | Uses structural path IDs (e.g., `tbl[0]:tr[1]:tc[2]:p[0]`) instead of xpaths, matching the DOCX extraction format |
| JSON-in / JSON-out | Ensures structured I/O, easy to validate and merge back |
| ~50 segments per batch | Balances context length, translation quality, and throughput |
| type field included | Helps the translator produce contextually appropriate translations (e.g., concise headers vs. detailed paragraphs) |
| `{"translations": [...]}` wrapper | Standardizes output format; script extracts array via `["translations"]` |
| No full document context | Unlike the DSL-based skill, this operates directly on DOCX without a markdown intermediary; segment types provide sufficient context |
| Direct AI translation | No external API calls needed; Claude translates directly |
