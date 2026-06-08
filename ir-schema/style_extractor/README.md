# DocIR Style Extractor

Extracts typography styles from cropped region images using VLM (Visual Language Model).

## Overview

The Style Extractor is a component of the DocIR pipeline that analyzes cropped text regions from PDF pages and extracts detailed style information including:

- **Font name** (e.g., Arial, Times New Roman, Yu Gothic)
- **Font size** in PDF points
- **Style attributes** (bold, italic, underline)
- **Text color** in hex format
- **Visual evidence** (pixel height, color sample, confidence score)

## Architecture

```
DocIR XML + Source PDF
         ↓
    Crop text regions (pymupdf @ 200 DPI)
         ↓
    VLM analysis (qwen3.6-35b-a3b-q7 @ localhost:11234)
         ↓
    Parse JSON response
         ↓
    Update DocIR XML with style attributes
```

## Usage

### Command Line

```bash
# Basic usage
python ir-schema/style_extractor/style_extractor.py \
  input.docir.xml \
  source.pdf \
  -o output.docir.xml

# Custom model and API settings
python ir-schema/style_extractor/style_extractor.py \
  input.docir.xml \
  source.pdf \
  -o output.docir.xml \
  --model qwen3.6-35b-a3b-q7 \
  --api-base http://localhost:11234/v1 \
  --api-key change-me-local-key \
  --dpi 200

# Process specific region types
python ir-schema/style_extractor/style_extractor.py \
  input.docir.xml \
  source.pdf \
  --region-types text table
```

### Python API

```python
from pathlib import Path
from style_extractor.style_extractor import extract_styles_from_docir, VLMConfig

vlm_config = VLMConfig(
    api_base="http://localhost:11234/v1",
    api_key="change-me-local-key",
    model="qwen3.6-35b-a3b-q7"
)

extract_styles_from_docir(
    docir_path=Path("input.docir.xml"),
    pdf_path=Path("source.pdf"),
    output_path=Path("output.docir.xml"),
    vlm_config=vlm_config,
    dpi=200,
    region_types=["text"]
)
```

## VLM Prompt

The style extractor uses a carefully designed prompt that instructs the VLM to:

1. Analyze the cropped region image
2. Extract precise typography information
3. Return structured JSON with computed styles and visual evidence
4. Rate confidence in the detection

Example response:
```json
{
  "font_name": "Arial",
  "font_size_pt": 36.0,
  "bold": true,
  "italic": false,
  "underline": false,
  "color": "#000000",
  "text_alignment": "center",
  "evidence": {
    "pixel_height": 48.0,
    "color_sample": "#000000",
    "confidence": 0.95,
    "notes": "Large bold title text"
  }
}
```

## Output Format

The style extractor updates the DocIR XML by adding attributes to `<docir:run>` elements:

```xml
<docir:run 
  font_name="Arial" 
  font_size_pt="36.0" 
  bold="true" 
  italic="false" 
  underline="false" 
  color="#000000"
  evidence_pixel_height="48.0" 
  evidence_color_sample="#000000" 
  evidence_confidence="0.95">
  This is Title
</docir:run>
```

It also updates the `<docir:provenance>` section:

```xml
<docir:provenance>
  <docir:source>glm-ocr-pipeline</docir:source>
  <docir:confidence>0.85</docir:confidence>
  <docir:detection_model>PP-DocLayoutV3</docir:detection_model>
  <docir:style_extractor>qwen3.6-35b-a3b-q7</docir:style_extractor>
  <docir:style_extraction_time_ms>45627</docir:style_extraction_time_ms>
</docir:provenance>
```

## Coordinate System

The style extractor handles coordinate transformation between:

- **DocIR/PDF coordinates**: Origin at bottom-left, y increases upward
- **pymupdf coordinates**: Origin at top-left, y increases downward

The conversion formula:
```python
pymupdf_y = page_height - pdf_y
```

## Performance

Typical processing times (on M4 Max 128GB):

- **VLM inference**: 10-45 seconds per region (depends on image complexity)
- **PDF cropping**: < 1 second per region
- **Total for 7 regions**: ~3-5 minutes

## Debugging

Use the debug script to save cropped images for visual inspection:

```bash
python ir-schema/style_extractor/debug_crops.py \
  input.docir.xml \
  source.pdf \
  /tmp/crops
```

This saves each cropped region as a JPEG file for manual verification.

## Limitations

1. **VLM accuracy**: Style detection depends on VLM capability. Complex fonts or unusual styling may have lower confidence.
2. **Resolution**: Higher DPI improves accuracy but increases processing time.
3. **Small text**: Very small text (< 8pt) may be difficult for VLM to analyze accurately.
4. **Mixed styles**: If a region contains multiple styles, the VLM may only detect the dominant style.

## Integration

The style extractor fits into the DocIR pipeline as follows:

```
1. GLM-OCR pipeline → Model JSON
2. IR Builder → DocIR XML (basic structure)
3. Style Extractor → DocIR XML (with styles) ← THIS STEP
4. Validator → Validated DocIR XML
5. DOCX Generator → Editable DOCX (future)
```

## Testing

Run the integration test to verify the full pipeline:

```bash
python ir-schema/tests/test_integration.py
```

This test:
1. Builds DocIR from GLM-OCR output
2. Extracts styles using VLM
3. Validates the final XML
4. Cleans up test artifacts

## Requirements

- Python 3.10+
- pymupdf (PyMuPDF)
- lxml
- requests
- VLM service (qwen3.6-35b-a3b-q7 @ localhost:11234)

Install dependencies:
```bash
cd ~/Works/glm-ocr
uv pip install pymupdf lxml requests
```

## Example Results

From `small.pdf` test case:

| Region | Text | Font | Size | Style | Color | Confidence |
|--------|------|------|------|-------|-------|------------|
| r0_p0 | This is Title | Arial | 36pt | bold | #000000 | 0.95 |
| r1_p0 | This is a body text. | Times New Roman | 14pt | - | #000000 | 0.95 |
| r2_p0 | This is a body text with color and italic. | Times New Roman | 14pt | italic | #FF0000 | 0.95 |
| r3_p0 | This is heading 1 | Arial | 36pt | bold | #000000 | 0.95 |
| r4_p0 | This is heading 2 | Arial | 28pt | bold | #000000 | 0.95 |
| r7_p0 | This is a text box, 文字方塊 (YuGothic) ! | Yu Gothic | 28pt | bold | #000000 | 0.95 |
| r8_p0 | This is a body text with color and italic in a text box!! 耶！ | Times New Roman | 24pt | italic | #FF0000 | 0.95 |

All styles extracted successfully with high confidence (0.95).
