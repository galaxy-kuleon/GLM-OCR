# DOCX Translator

Translate DOCX content while preserving ALL formatting (fonts, sizes, colors, bold, italic, alignment, etc.).

## Features

- **Structure preservation**: Handles paragraphs, tables, headers, footers, text boxes
- **Formatting preservation**: All run-level formatting (font, size, color, bold, italic, underline, etc.) is preserved
- **Style control**: Formal, legal, casual, technical, or neutral translation styles
- **Batch translation**: Efficient batched LLM calls for large documents
- **Glossary support**: Custom term translations for domain-specific vocabulary
- **Progress tracking**: Shows translation progress per batch

## Usage

### Basic translation

```bash
python3 translator/translator.py input.docx -o output.docx --target-lang zh
```

### With style control

```bash
python3 translator/translator.py input.docx -o output.docx \
    --target-lang zh \
    --style legal
```

### Custom LLM endpoint

```bash
python3 translator/translator.py input.docx -o output.docx \
    --target-lang en \
    --api-base http://localhost:11434/v1 \
    --model qwen3.6-27b-k-xl
```

### With glossary

Create a `glossary.json`:
```json
{
    "Force Majeure": "不可抗力",
    "Indemnification": "赔偿",
    "Arbitration": "仲裁"
}
```

```bash
python3 translator/translator.py input.docx -o output.docx \
    --target-lang zh \
    --style legal \
    --glossary glossary.json
```

## Options

| Option | Default | Description |
|--------|---------|-------------|
| `--target-lang`, `-l` | `zh` | Target language code |
| `--source-lang`, `-s` | auto | Source language (auto-detect if not specified) |
| `--style` | `neutral` | Translation style: formal, legal, casual, technical, neutral |
| `--api-base` | `http://localhost:11234/v1` | LLM API base URL |
| `--api-key` | `change-me-local-key` | LLM API key |
| `--model` | `qwen3.6-27b-k-xl` | LLM model name |
| `--batch-size` | `20` | Text segments per batch |
| `--glossary` | none | Path to glossary JSON file |

## How It Works

1. **Extract**: Walk through all document elements (paragraphs, tables, headers, footers) and extract text runs with their locations
2. **Batch**: Group text segments into batches for efficient LLM calls
3. **Translate**: Send batches to LLM with style instructions and glossary
4. **Apply**: Replace original text with translations while preserving all formatting (font, size, color, bold, italic, etc.)
5. **Save**: Write the translated DOCX

## Architecture

```
DOCX → Extract Text Segments → Batch → LLM Translation → Apply to DOCX → Output
         (preserve locations)        (with style/glossary)  (preserve formatting)
```

### Text Segment Structure

Each text segment tracks:
- Original text
- Location (paragraph index, run index)
- Context (in table? which row/col? header/footer?)
- Translated text (after translation)

### Formatting Preservation

The translator only modifies `run.text` - all other run properties (font, size, color, bold, italic, underline, spacing, etc.) remain untouched. This ensures perfect formatting preservation.

## Integration with Pipeline

Can be used as a post-processing step after DOCX generation:

```bash
# Generate DOCX from PDF
python3 run_pipeline.py input.pdf -o output.docx --digital --positioned

# Translate the output
python3 translator/translator.py output.docx -o output_zh.docx --target-lang zh --style formal
```
