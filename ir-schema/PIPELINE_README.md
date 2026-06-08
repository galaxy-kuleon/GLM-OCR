# DocIR Pipeline - Complete Documentation

## Overview

DocIR (Document Intermediate Representation) is a commercial-grade pipeline for converting PDF documents to high-fidelity editable DOCX files. The pipeline uses GLM-OCR for layout detection and OCR, VLM (Visual Language Models) for style extraction, and produces an XML-based intermediate representation that maps directly to OOXML (DOCX format).

## Architecture

```
PDF → GLM-OCR Pipeline → DocIR XML → Style Extraction → Styled DocIR XML → DOCX Generator → Editable DOCX
```

### Pipeline Stages

1. **GLM-OCR Layout Detection & OCR**
   - Detects document layout (text, tables, images, formulas)
   - Performs OCR on text regions
   - Outputs structured JSON with regions, bounding boxes, and content

2. **IR Builder**
   - Converts GLM-OCR JSON to DocIR XML
   - Transforms coordinates from normalized (0-1000) to PDF points
   - Cleans OCR content (removes artifacts, markdown fences)
   - Extracts image dimensions
   - Registers image assets

3. **Style Extraction**
   - Crops text regions from PDF at 200 DPI
   - Sends to VLM (qwen3.6-35b-a3b-q7) for style analysis
   - Extracts: font name, font size, bold, italic, color
   - Calibrates font sizes using pixel height and DPI
   - Handles edge cases (small regions, empty content)

4. **Table Style Extraction** (Optional)
   - Crops table regions from PDF
   - Extracts: border visibility, border color, header row detection
   - Updates table_style elements in DocIR XML

5. **DOCX Generator**
   - Converts DocIR XML to editable DOCX
   - Two modes:
     - **Flow mode** (default): Linear document flow
     - **Positioned mode**: Absolute positioning using page coordinates
   - Preserves typography styles (font, size, bold, italic, color)
   - Renders tables with borders and formatting
   - Embeds images with correct dimensions

6. **Visual Comparison** (Optional)
   - Converts generated DOCX back to PDF
   - Renders both PDFs to images
   - Calculates similarity score
   - Generates diff images highlighting differences

## Installation

### Prerequisites

- Python 3.10+
- Ollama with glm-ocr model
- VLM service (qwen3.6-35b-a3b-q7 @ localhost:11234)
- LibreOffice (for DOCX → PDF conversion in visual comparison)

### Setup

```bash
# Navigate to GLM-OCR repository
cd ~/Works/glm-ocr

# Install dependencies
uv pip install pymupdf lxml requests python-docx Pillow numpy

# Verify Ollama is running
curl http://localhost:11434/api/tags

# Verify VLM service is running
curl http://localhost:11234/v1/models -H "Authorization: Bearer change-me-local-key"
```

## Usage

### Quick Start

```bash
# Run the full pipeline
.venv/bin/python3 ir-schema/run_pipeline.py \
  input.pdf \
  -o output.docx

# Skip OCR (use existing GLM-OCR output)
.venv/bin/python3 ir-schema/run_pipeline.py \
  input.pdf \
  -o output.docx \
  --skip-ocr \
  --glmocr-output output-dir

# Skip style extraction
.venv/bin/python3 ir-schema/run_pipeline.py \
  input.pdf \
  -o output.docx \
  --skip-style
```

### Individual Components

#### 1. GLM-OCR Pipeline

```bash
uv run --extra layout glmocr parse input.pdf \
  --mode selfhosted \
  --layout-device cpu \
  --set pipeline.ocr_api.api_mode ollama_generate \
  --set pipeline.ocr_api.api_path /api/generate \
  --set pipeline.ocr_api.api_host localhost \
  --set pipeline.ocr_api.api_port 11434 \
  --set pipeline.ocr_api.model glm-ocr:latest \
  --set pipeline.layout.device cpu \
  --output ./output
```

#### 2. IR Builder

```bash
.venv/bin/python3 ir-schema/builder/ir_builder.py \
  output/input_model.json \
  input.pdf \
  -o input.docir.xml \
  --title "Document Title"
```

#### 3. Style Extraction

```bash
.venv/bin/python3 ir-schema/style_extractor/style_extractor.py \
  input.docir.xml \
  input.pdf \
  -o input-styled.docir.xml \
  --dpi 200
```

#### 4. Table Style Extraction (Optional)

```bash
.venv/bin/python3 ir-schema/style_extractor/table_style_extractor.py \
  input.docir.xml \
  input.pdf \
  -o input-table-styled.docir.xml
```

#### 5. DOCX Generator

```bash
# Flow mode (default)
.venv/bin/python3 ir-schema/docx_generator/docx_generator.py \
  input-styled.docir.xml \
  -o output.docx

# Positioned mode
.venv/bin/python3 ir-schema/docx_generator/docx_generator.py \
  input-styled.docir.xml \
  -o output.docx \
  --positioned
```

#### 6. Visual Comparison

```bash
.venv/bin/python3 ir-schema/tests/visual_comparison.py \
  input.pdf \
  output.docx \
  -o comparison-output/
```

## DocIR XML Schema

### Root Element

```xml
<docir:document version="0.1.0" source_pdf="input.pdf" generated_at="2026-06-09T00:00:00" generator="ir-builder-v0.1.0">
  <docir:metadata>...</docir:metadata>
  <docir:pages>...</docir:pages>
  <docir:assets>...</docir:assets>
  <docir:cross_page_hints>...</docir:cross_page_hints>
</docir:document>
```

### Key Elements

#### Region

```xml
<docir:region id="r0_p0" type="text" native_label="paragraph_title" order="0">
  <docir:bbox x="217.88" y="772.01" width="157.76" height="27.78"/>
  <docir:polygon>...</docir:polygon>
  <docir:provenance>
    <docir:source>glm-ocr-pipeline</docir:source>
    <docir:confidence>0.85</docir:confidence>
    <docir:detection_model>PP-DocLayoutV3</docir:detection_model>
    <docir:style_extractor>qwen3.6-35b-a3b-q7</docir:style_extractor>
  </docir:provenance>
  <docir:text_content>
    <docir:paragraph>
      <docir:run font_name="Arial" font_size_pt="36.0" bold="true" color="#000000"
                 evidence_pixel_height="48.0" evidence_confidence="0.95">
        This is Title
      </docir:run>
    </docir:paragraph>
  </docir:text_content>
</docir:region>
```

#### Table

```xml
<docir:region id="r5_p0" type="table" order="5">
  <docir:bbox x="55.96" y="571.64" width="481.60" height="141.44"/>
  <docir:table_content rows="4" cols="3">
    <docir:table_style border_visible="true" border_color="#000000" header_row="true"/>
    <docir:row_group type="header">
      <docir:row>
        <docir:cell>
          <docir:text_content>
            <docir:paragraph>
              <docir:run>Header 1</docir:run>
            </docir:paragraph>
          </docir:text_content>
        </docir:cell>
        ...
      </docir:row>
    </docir:row_group>
    <docir:row_group type="body">
      ...
    </docir:row_group>
  </docir:table_content>
</docir:region>
```

#### Image Asset

```xml
<docir:assets>
  <docir:asset id="img_page0_region6" mime_type="image/jpeg" width_px="641" height_px="634">
    <docir:file_path>imgs/cropped_page0_idx0.jpg</docir:file_path>
    <docir:extraction_source>pdf_page_0_region_6</docir:extraction_source>
  </docir:asset>
</docir:assets>
```

## Schema Validation

```bash
# Validate DocIR XML
.venv/bin/python3 ir-schema/validator/semantic_validator.py \
  input.docir.xml \
  --xsd ir-schema/docir-v0.1.0.xsd \
  -v
```

### Validation Rules

1. **XSD Schema Validation** - Structure and data types
2. **Bounding Box Validation** - Coordinates within page boundaries
3. **Table Structure Validation** - Row/column counts match cells
4. **Region Order Validation** - Monotonically increasing order
5. **Image Reference Validation** - All image references resolve to assets
6. **Confidence Range Validation** - Values in [0, 1]
7. **Cross-Page Hint Validation** - References are valid
8. **Provenance Completeness** - All regions have provenance data

## Coordinate Systems

### GLM-OCR Output
- **Normalized coordinates**: 0-1000 scale
- **Origin**: Top-left corner
- **Format**: `[x1, y1, x2, y2]`

### DocIR XML
- **PDF points**: 1/72 inch
- **Origin**: Bottom-left corner (PDF standard)
- **Format**: `x, y, width, height`

### Conversion Formula

```python
# Normalized (0-1000) to PDF points
pt_x = (norm_x / 1000) * page_width_pt
pt_y = page_height_pt - (norm_y / 1000) * page_height_pt
```

### DOCX Output
- **EMU (English Metric Units)**: 1 inch = 914400 EMU
- **Conversion**: 1 pt = 12700 EMU

## Font Size Calibration

The style extractor calibrates font sizes using pixel height and DPI:

```python
# At 200 DPI, 1 inch = 200 pixels, 1 inch = 72 points
# So: points = pixels * 72 / DPI
calibrated_size = pixel_height * 72.0 / dpi
```

The final font size is a weighted average of VLM estimate and calibration:
- High confidence (>0.7): 70% VLM, 30% calibration
- Low confidence (<0.7): 30% VLM, 70% calibration

## Performance

### Typical Processing Times

| Stage | Time per Page | Notes |
|-------|---------------|-------|
| GLM-OCR Layout | 30-60s | CPU-based, can be slow for complex layouts |
| GLM-OCR OCR | 10-30s per region | Depends on region count and size |
| IR Builder | <1s | Fast, CPU-only |
| Style Extraction | 10-45s per region | VLM inference, bottleneck |
| Table Style Extraction | 15-50s per table | VLM inference |
| DOCX Generator | <1s | Fast, CPU-only |

### Optimization Tips

1. **Use GPU for layout detection** if available
2. **Skip style extraction** for draft conversions
3. **Use flow mode** instead of positioned mode for faster generation
4. **Batch process** multiple PDFs in parallel

## Testing

### Integration Test

```bash
.venv/bin/python3 ir-schema/tests/test_integration.py
```

Tests the full pipeline: build → style → DOCX → validate

### Visual Comparison Test

```bash
.venv/bin/python3 ir-schema/tests/visual_comparison.py \
  input.pdf \
  output.docx \
  -o comparison-output/
```

Generates similarity score and diff images.

### Multi-Page Test

```bash
# Create test PDF
.venv/bin/python3 ir-schema/tests/create_multipage_test_pdf.py \
  /tmp/multipage-test.pdf \
  3

# Run pipeline
.venv/bin/python3 ir-schema/run_pipeline.py \
  /tmp/multipage-test.pdf \
  -o /tmp/multipage-test.docx
```

## Troubleshooting

### GLM-OCR Timeout

**Problem**: GLM-OCR hangs or times out on large PDFs

**Solution**:
- Use `--layout-device cpu` for CPU-only processing
- Reduce `--set pipeline.max_workers 4` to limit concurrency
- Process PDFs page-by-page if needed

### VLM Style Extraction Fails

**Problem**: VLM returns empty response or JSON parse error

**Solution**:
- Check VLM service is running: `curl http://localhost:11234/v1/models`
- Verify API key: `change-me-local-key`
- Increase timeout: `--timeout 300`
- Check VLM logs for errors

### DOCX Generation Errors

**Problem**: python-docx throws errors during generation

**Solution**:
- Validate DocIR XML first: `semantic_validator.py`
- Check image paths are correct
- Ensure all required elements are present

### Low Visual Fidelity Score

**Problem**: Visual comparison shows <80% similarity

**Solution**:
- Use positioned mode: `--positioned`
- Check style extraction completed successfully
- Verify table styles were extracted
- Check image dimensions are correct

## File Structure

```
ir-schema/
├── docir-v0.1.0.xsd              # XSD schema
├── DESIGN.md                      # Design document
├── README.md                      # This file
├── PROGRESS.json                  # Progress tracking
├── run_pipeline.py                # End-to-end pipeline script
├── builder/
│   ├── __init__.py
│   ├── ir_builder.py              # GLM-OCR → DocIR converter
│   ├── table_parser.py            # Markdown table parser
│   └── README.md
├── style_extractor/
│   ├── style_extractor.py         # Text style extraction
│   ├── table_style_extractor.py   # Table style extraction
│   ├── debug_crops.py             # Debug utility
│   └── README.md
├── docx_generator/
│   ├── docx_generator.py          # DocIR → DOCX converter
│   └── README.md
├── validator/
│   └── semantic_validator.py      # Validation rules
├── samples/
│   ├── small-pdf-sample.xml       # Hand-crafted example
│   └── small-pdf-built.xml        # Auto-generated example
└── tests/
    ├── test_integration.py        # Integration test
    ├── visual_comparison.py       # Visual comparison
    └── create_multipage_test_pdf.py
```

## Design Principles

1. **One-way pipeline**: VLM output → DOCX input (no feedback loops)
2. **PDF-native coordinates**: PDF points (1/72 inch)
3. **Region-based structure**: Layout detection regions + inline text
4. **Dual-layer styles**: Computed values + visual evidence
5. **XML serialization**: Direct OOXML mapping
6. **Multi-tier validation**: XSD + semantic checks
7. **Provenance tracking**: Source, confidence, model, timing

## Future Work

- [ ] Cross-page merge hints for table/paragraph continuations
- [ ] Batch VLM requests for better performance
- [ ] Text box detection and preservation
- [ ] Advanced table style extraction (cell colors, merged cells)
- [ ] Formula recognition and MathML export
- [ ] Handwriting recognition support

## License

Apache 2.0

## Credits

- **GLM-OCR**: https://github.com/zai-org/GLM-OCR
- **python-docx**: https://python-docx.readthedocs.io/
- **PyMuPDF**: https://pymupdf.readthedocs.io/
- **VLM Models**: qwen3.6-35b-a3b-q7 via LM Studio

## Contact

For issues and questions, please open an issue on the GitHub repository.
