# Translation Prompt Template

Prompt template used by `translate_content.py` to send translation requests to LM Studio (`qwen/qwen3-4b-2507`).

## Usage

In the generated `translate_content.py` script, read this template and replace the template variables:
- `{target_language}` — target language and style (from `--lang` argument)
- `{style_notes}` — extra style/tone notes (from `--style` argument, may be empty)
- `{full_document_markdown}` — full content of `input.md` (provides terminology/context reference)
- `{page_num}` — current page number (1-based)
- `{texts_json}` — current page's translatable text elements extracted from XML DSL (see format below)

## Input Format (texts JSON sent to model)

Text is extracted at the element level from `<run>` and `<cell>` elements in the XML DSL. Each item includes an `xpath` for identification and the `text` content to translate.

```json
[
  {
    "xpath": "heading[0]/run[0]",
    "text": "This is Title"
  },
  {
    "xpath": "paragraph[0]/run[0]",
    "text": "This is a body text."
  },
  {
    "xpath": "table[0]/row[0]/cell[1]",
    "text": "Table header col 1"
  },
  {
    "xpath": "table[0]/row[1]/cell[2]/run[0]",
    "text": "Keyword-styled cell text"
  },
  {
    "xpath": "side-by-side[0]/column[0]/paragraph[0]/run[0]",
    "text": "Left column text"
  }
]
```

- `<run>` elements: xpath is built from parent hierarchy, e.g. `heading[0]/run[0]`, `text-frame[0]/paragraph[1]/run[0]`, `side-by-side[0]/column[0]/paragraph[0]/run[0]`
- `<cell>` elements (no `<run>` children): xpath uses table/row/col indices, e.g. `table[0]/row[2]/cell[3]`
- `<cell>` elements (with `<run>` children): xpath includes run index, e.g. `table[0]/row[1]/cell[2]/run[0]`
- `<run is-math="true">` elements are **excluded** — math/formula content is never translated
- Empty text elements are excluded from the input.

## Output Format (model returns)

The output is a JSON **object** with a `translations` array (required by structured output — top-level must be object):

```json
{
  "translations": [
    {
      "xpath": "heading[0]/run[0]",
      "translated_text": "이것은 제목입니다"
    },
    {
      "xpath": "paragraph[0]/run[0]",
      "translated_text": "이것은 본문 텍스트입니다."
    },
    {
      "xpath": "table[0]/row[0]/cell[1]",
      "translated_text": "표 헤더 열 1"
    }
  ]
}
```

The `translate_content.py` script extracts the `translations` array from this wrapper object via `response["translations"]`, then maps each `xpath` back to the corresponding XML element to replace its text.

## System Prompt Template

```
You are a professional document translator. Your task is to translate document content into {target_language}.

{style_notes}

## Rules

1. **Only translate text content** — do not modify xpath identifiers.
2. **Preserve proper nouns** — keep brand names, product names, and technical identifiers in their original form unless there is a widely accepted translation.
3. **Handle mixed languages** — if the source contains mixed languages, translate all text to the target language but keep proper nouns and technical terms as-is.
4. **Match the count** — your output `translations` array MUST have exactly the same number of items as the input, with matching xpaths.
5. **Output format** — respond with a JSON object containing a `translations` array. Each item has `xpath` (string, unchanged) and `translated_text` (string).
6. **Do NOT translate empty strings** — if the text is empty, keep it empty.
7. **Do NOT translate math/formula content** — items containing LaTeX expressions or mathematical formulas should be returned as-is without translation. (Note: math content is normally excluded before sending, but if any slips through, preserve it unchanged.)

## Full Document Context (for reference only — do not include in output)

{full_document_markdown}
```

## User Prompt Template

```
Translate the following text items from page {page_num} into {target_language}.

Input:
{texts_json}

Output a JSON object with a "translations" array. Each item has "xpath" (string, unchanged) and "translated_text" (string).
```

## Key Design Decisions

| Decision | Rationale |
|----------|-----------|
| xpath-based element-level extraction | Operates directly on XML DSL `<run>` and `<cell>` elements, enabling precise text replacement without affecting structure or attributes |
| JSON-in / JSON-out | Ensures structured I/O, easy to validate and merge back into XML |
| Full document as system context | 256k context is sufficient; helps with terminology consistency |
| One page per API call | Balances context length and translation consistency |
| Structured output (`response_format`) | LM Studio enforces valid JSON via grammar-based sampling, eliminating code fence cleanup and JSON parse retries |
| `{"translations": [...]}` wrapper | OpenAI structured output spec requires top-level object; script extracts array via `["translations"]` |
| XML attributes preserved as-is | font-size, color, bold, italic, alignment etc. are language-independent and never modified |
| Cell/run branching | `<cell>` with `<run>` children uses per-run xpaths; without uses cell-level xpath — matches keyword-styling DSL output |
| Math exclusion | `<run is-math="true">` excluded at extraction time; rule 7 as safety net if any slip through |
