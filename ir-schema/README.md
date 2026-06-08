# DocIR Pipeline

**Document Intermediate Representation** — A pipeline for converting PDF documents to editable DOCX while preserving layout, typography, and structure.

## Overview

```
PDF → GLM-OCR → DocIR XML → Styled DocIR XML → Editable DOCX
         │            │              │
         │            │              └── VLM typography extraction
         │            └── Intermediate representation (XML)
         └── Layout detection + OCR
```

The pipeline consists of:
1. **GLM-OCR** — Layout detection and OCR (via Ollama)
2. **IR Builder** — Converts GLM-OCR output to DocIR XML
3. **Semantic Validator** — Validates DocIR XML against XSD schema
4. **Style Extractor** — VLM-based typography extraction
5. **DOCX Generator** — Converts DocIR XML to editable DOCX

## Quick Start

```bash
# Full pipeline: PDF → DOCX
python run_pipeline.py input.pdf -o output.docx

# With absolute positioning (preserves layout)
python run_pipeline.py input.pdf -o output.docx --positioned

# With table style extraction
python run_pipeline.py input.pdf -o output.docx --table-styles

# Skip OCR (use existing GLM-OCR output)
python run_pipeline.py input.pdf -o output.docx --skip-ocr --glmocr-output ./output-dir/

# Skip style extraction (faster, no VLM needed)
python run_pipeline.py input.pdf -o output.docx --skip-style
```

## Requirements

- Python 3.10+
- GLM-OCR (via Ollama with `glm-ocr:latest` model)
- VLM service (qwen3.6-35b-a3b-q7 @ localhost:11234) — for style extraction
- LibreOffice — for visual comparison tests

### Python Packages

```bash
pip install pymupdf lxml requests python-docx Pillow
```

## Architecture

### DocIR XML Schema

The intermediate representation uses XML with namespace `urn:docir:v0.1`.

**Key elements:**
- `<docir:document>` — Root element with metadata
- `<docir:page>` — Page with regions
- `<docir:region>` — Layout region (text, table, image, formula)
- `<docir:bbox>` — Bounding box in PDF points (origin at bottom-left)
- `<docir:text_content>` — Paragraphs and runs with style attributes
- `<docir:table_content>` — Structured table with rows, cells, spans
- `<docir:image_content>` — Image reference with visual features
- `<docir:assets>` — Asset registry (extracted images)

**Coordinate system:**
- Units: PDF points (1/72 inch)
- Origin: bottom-left of page
- Conversion: `pt_x = (norm_x / 1000) * page_width_pt`

### Components

#### 1. IR Builder (`builder/ir_builder.py`)

Converts GLM-OCR model JSON to DocIR XML.

```bash
python builder/ir_builder.py model.json source.pdf -o output.docir.xml --title "Document"
```

**Features:**
- Coordinate conversion (normalized 0-1000 → PDF points)
- Region type mapping (text, table, image, formula)
- Table parsing (markdown tables, tab-separated)
- Image asset registration with dimensions
- Content cleaning (removes OCR artifacts)

#### 2. Semantic Validator (`validator/semantic_validator.py`)

Validates DocIR XML against XSD schema + semantic rules.

```bash
python validator/semantic_validator.py input.docir.xml
```

**Validation rules:**
1. XSD schema compliance
2. Bbox within page bounds
3. Non-empty text content
4. Valid region types
5. Asset references exist
6. Coordinate consistency
7. Table dimensions match cells
8. Page count matches metadata

#### 3. Style Extractor (`style_extractor/style_extractor.py`)

VLM-based typography extraction from cropped region images.

```bash
# Text styles only
python style_extractor/style_extractor.py input.docir.xml source.pdf -o styled.xml

# Text + table styles
python style_extractor/style_extractor.py input.docir.xml source.pdf -o styled.xml --table-styles
```

**Extracted styles:**
- Font name, size, bold, italic, underline
- Text color
- Evidence: pixel height, color sample, confidence
- Font size calibration (pixel height → points using DPI)

**Table styles:**
- Border visibility and color
- Header row detection
- Cell background colors

**Features:**
- Retry logic with exponential backoff
- Empty region skipping
- Small region filtering
- Font size calibration using pixel measurements

#### 4. DOCX Generator (`docx_generator/docx_generator.py`)

Converts DocIR XML to editable DOCX using python-docx.

```bash
# Flow mode (linear document)
python docx_generator/docx_generator.py input.docir.xml -o output.docx

# Positioned mode (preserves layout)
python docx_generator/docx_generator.py input.docir.xml -o output.docx --positioned
```

**Modes:**
- **Flow** — Regions become sequential paragraphs (default)
- **Positioned** — Elements placed at PDF coordinates (zero margins, space_before/left_indent)

**Supported content:**
- Text with typography styles
- Tables with borders, headers, cell spans
- Images with dimensions from bbox
- Page breaks for multi-page documents

#### 5. Table Parser (`builder/table_parser.py`)

Parses markdown-like table content from GLM-OCR.

**Supported formats:**
- Markdown tables (`| col1 | col2 |`)
- Tab-separated values
- Empty/single-cell content

#### 6. Visual Comparison (`tests/visual_comparison.py`)

Renders DOCX back to PDF and compares with source using VLM judge.

```bash
python tests/visual_comparison.py source.pdf generated.docx
```

**Pipeline:**
1. Convert DOCX → PDF (LibreOffice headless)
2. Render both PDFs to page images (pymupdf)
3. Create side-by-side comparison images
4. VLM scores: text/layout/style fidelity (0-10 each)

## Testing

```bash
# Quick test (positioned DOCX, no VLM needed)
python tests/test_integration.py --positioned-only

# Full pipeline test (requires VLM)
python tests/test_integration.py --full

# Visual comparison
python tests/visual_comparison.py source.pdf generated.docx
```

## File Structure

```
ir-schema/
├── docir-v0.1.0.xsd          # XSD schema
├── run_pipeline.py            # End-to-end pipeline script
├── builder/
│   ├── ir_builder.py          # GLM-OCR → DocIR XML
│   └── table_parser.py        # Markdown table parser
├── validator/
│   └── semantic_validator.py  # XML validation
├── style_extractor/
│   ├── style_extractor.py     # VLM typography extraction
│   └── debug_crops.py         # Visual inspection tool
├── docx_generator/
│   └── docx_generator.py      # DocIR XML → DOCX
├── tests/
│   ├── test_integration.py    # Integration tests
│   └── visual_comparison.py   # Visual fidelity testing
├── samples/
│   ├── small-pdf-built.xml    # Sample built XML
│   └── small-pdf-styled.xml   # Sample styled XML
└── PROGRESS.json              # Task tracking
```

## Current Status

**Completed features (15/20 tasks):**
- ✅ Table markdown parser
- ✅ DOCX Generator v1-v4 (text, tables, images, positioning)
- ✅ End-to-end pipeline script
- ✅ Content cleaning (OCR artifacts)
- ✅ Integration tests
- ✅ Image dimension extraction
- ✅ VLM style extraction improvements
- ✅ Font size calibration
- ✅ Table style extraction
- ✅ Visual comparison test
- ✅ Confidence extraction (investigated - not available in GLM-OCR)
- ✅ DOCX visual fidelity scoring

**Remaining tasks:**
- ⬜ Cross-page merge hints
- ⬜ Performance optimization (batch VLM, parallel processing)
- ⬜ Text box detection
- ⬜ Multi-page test

## Known Limitations

1. **Confidence scores** — GLM-OCR model JSON doesn't include layout detector confidence. Uses placeholder 0.85.
2. **Empty regions** — Some text regions have no OCR content (VLM sees blank crops).
3. **Table content** — GLM-OCR sometimes returns placeholder text instead of actual table data.
4. **Image extraction** — Depends on GLM-OCR correctly cropping and saving images.
5. **Positioning accuracy** — Positioned mode approximates layout using paragraph spacing/indentation.

## Development

### Adding New Region Types

1. Update XSD schema (`docir-v0.1.0.xsd`) — add to `Region` type enumeration
2. Update IR Builder (`builder/ir_builder.py`) — add to `determine_region_type()` mapping
3. Update DOCX Generator — add processing function for new type

### VLM Configuration

Default VLM settings in `style_extractor/style_extractor.py`:
```python
VLMConfig(
    api_base="http://localhost:11234/v1",
    api_key="change-me-local-key",
    model="qwen3.6-35b-a3b-q7",
    max_tokens=4096,
    temperature=0.1,
    timeout=120
)
```

## License

Internal project — zai-org/GLM-OCR pipeline extension.
