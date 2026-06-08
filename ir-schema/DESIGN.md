# DocIR — Document Intermediate Representation

## Design Document v0.1.0

### Vision

DocIR is a **commercial-grade intermediate representation** for converting PDF documents into high-fidelity editable DOCX files. It sits between a VLM/OCR pipeline (GLM-OCR + style extraction) and a DOCX generator (OOXML assembler).

```
PDF ──→ GLM-OCR Pipeline ──→ DocIR (XML) ──→ DOCX Generator ──→ Editable DOCX
         │                      ↑
         ├── Layout Detection    │
         ├── Region Cropping     │ Dual-layer:
         ├── OCR (Ollama)        │ computed style + evidence
         └── VLM Style Extract ──┘
```

### Design Principles

| # | Principle | Decision | Rationale |
|---|-----------|----------|-----------|
| 1 | IR 定位 | VLM → DOCX one-way pipeline | 簡單直接，不需要 multi-consumer 的複雜性 |
| 2 | 座標系統 | PDF pt 為主，export 轉 EMU | PDF-native，轉換公式確定 (1pt = 12700 EMU) |
| 3 | Element 結構 | Region tree + inline structure | 匹配 layout detection 輸出，text 內部有語義結構 |
| 4 | Style 資訊 | Dual: computed + evidence | Downstream 可以驗證 VLM 推斷是否正確 |
| 5 | Inline structure | Semantic paragraphs + positional runs | 兼顧可讀性和精確性 |
| 6 | Table | Tree structure (nested support) | 完整支持 merged cells + nested tables |
| 7 | 跨頁 | Per-page snapshot + merge_hints | Pipeline 分兩階段，簡單但需要 post-processing |
| 8 | Image/Chart | Reference + VLM visual features | 不存 raw data，VLM 提取特徵供 downstream 使用 |
| 9 | Error handling | Multi-tier inclusion | confidence-based filtering，DOCX generator 可選擇性 include |
| 10 | Serialization | **XML** | Tag-based，直接映射 OOXML，DOCX generator 零轉換 |
| 11 | Validation | Schema (XSD) + semantic checks | Structure + semantics 雙重驗證 |
| 12 | Extensibility | Evolve as needed | Pragmatic，不預先 over-engineer |
| 13 | Token efficiency | Don't optimize | VLM context window 夠大 (128k+) |
| 14 | Debugging | All: logging + checkpoints + provenance + visual | 完整 traceability |

---

### Schema Overview

```xml
<document version="0.1.0" source_pdf="..." generated_at="..." generator="...">
  <metadata>
    <title/> <author/> <page_count/>
    <default_page_size width_pt="..." height_pt="..."/>
    <pipeline_info>
      <layout_detector/> <ocr_engine/> <style_extractor/>
    </pipeline_info>
  </metadata>
  
  <pages>
    <page index="0">
      <page_size width_pt="..." height_pt="..."/>
      <regions>
        <region id="r0_p0" type="text|table|image|formula" order="0">
          <bbox x="..." y="..." width="..." height="..."/>
          <polygon>...</polygon>
          <provenance>
            <source/> <confidence/> <processing_time_ms/>
            <detection_model/> <notes/>
          </provenance>
          <text_content|table_content|image_content|formula_content>
            ...
          </text_content|table_content|image_content|formula_content>
          <merge_hint type="..." linked_region="..."/>
        </region>
      </regions>
    </page>
  </pages>
  
  <assets>
    <asset id="..." mime_type="...">
      <file_path/> <extraction_source/>
    </asset>
  </assets>
  
  <cross_page_hints>
    <hint type="table_continuation|paragraph_continuation"
          from_region="..." to_region="..."/>
  </cross_page_hints>
</document>
```

---

### Key Design Decisions Explained

#### 1. XML over JSON

**Why XML?**
- OOXML (DOCX) is XML-based → direct mapping with zero transformation
- XSD schema provides strong validation
- XPath/XQuery for programmatic access
- Namespaces for extensibility
- VLM can output XML (tested with qwen3.6-35b-a3b-q7, 100% JSON validity)

**Trade-off:** XML is more verbose than JSON, but we chose not to optimize for token efficiency (Q13=E).

#### 2. Dual-Layer Style (Computed + Evidence)

Every style attribute has two representations:

```xml
<docir:run font_size_pt="14.0" bold="true" color="#FF0000"
           evidence_pixel_height="16" evidence_color_sample="#CC3333"
           evidence_confidence="0.85">
  Heading text
</docir:run>
```

- **Computed** (`font_size_pt`, `color`): VLM's best guess, used by DOCX generator
- **Evidence** (`evidence_pixel_height`, `evidence_color_sample`): Raw visual observations, used for:
  - Verification (does computed match evidence?)
  - Debugging (why did VLM choose this font size?)
  - Re-inference (if confidence is low, re-run with different prompt)

#### 3. Region-Based Top-Level Structure

We use GLM-OCR's native region output as the top-level structure:

```
page → regions[] → region (type=text|table|image|formula)
                         ↓
                    content (paragraphs | table tree | image ref | formula)
```

This matches the layout detection output directly, avoiding an unnecessary transformation step.

#### 4. Table Tree Structure

Tables use a tree structure (row_group → row → cell) instead of flat grid:

```xml
<table_content rows="7" cols="5">
  <row_group type="header">
    <row>
      <cell row_span="1" col_span="2">
        <text_content>...</text_content>
      </cell>
      ...
    </row>
  </row_group>
  <row_group type="body">
    <row>
      <cell>
        <!-- Nested table support -->
        <table_content rows="2" cols="2">...</table_content>
      </cell>
    </row>
  </row_group>
</table_content>
```

This directly maps to OOXML's `w:tbl > w:tr > w:tc` structure.

#### 5. Cross-Page Handling

Each page is self-contained. Cross-page elements use `merge_hint`:

```xml
<!-- Page 0, last region -->
<region id="r5_p0" type="table">
  <merge_hint type="continues_on_next_page" linked_region="r0_p1"/>
</region>

<!-- Page 1, first region -->
<region id="r0_p1" type="table">
  <merge_hint type="continued_from_previous" linked_region="r5_p0"/>
</region>
```

Plus a global `cross_page_hints` section for the post-processor:

```xml
<cross_page_hints>
  <hint type="table_continuation" from_region="r5_p0" to_region="r0_p1" confidence="0.95"/>
</cross_page_hints>
```

#### 6. Multi-Tier Error Handling

Confidence-based inclusion tiers:

| Confidence | Tier | Action |
|-----------|------|--------|
| ≥ 0.8 | Include | Direct use in DOCX |
| 0.5 – 0.8 | Review | Include with `needs_review="true"` attribute |
| < 0.5 | Placeholder | Replace with `[UNREADABLE]` marker |

The DOCX generator decides whether to include `needs_review` regions based on user preference.

---

### Pipeline Integration

#### Stage 1: GLM-OCR Layout + OCR

```
Input: PDF pages
Output: JSON regions (bbox, type, confidence, raw text)
```

#### Stage 2: Coordinate Conversion

```
Input: Normalized bbox (0-1000)
Output: PDF pt coordinates
Formula: pt_x = (norm_x / 1000) * page_width_pt
```

#### Stage 3: VLM Style Extraction

```
Input: Cropped region images + raw text
Model: qwen3.6-35b-a3b-q7 @ localhost:11234
Output: font-name, font-size, color, bold, italic + evidence
```

#### Stage 4: IR Assembly

```
Input: Regions + styles + coordinates
Output: DocIR XML document
```

#### Stage 5: Semantic Validation

```
Input: DocIR XML
Checks:
  - bbox within page boundaries
  - table rows × cols matches actual cell count
  - region order is monotonically increasing
  - image references resolve to existing assets
  - confidence values in valid range [0, 1]
```

#### Stage 6: DOCX Generation

```
Input: Validated DocIR XML
Output: Editable DOCX file
Mapping:
  - region(text) → w:p (paragraph)
  - region(table) → w:tbl
  - region(image) → w:drawing
  - run → w:r with w:rPr
  - bbox → positioning (EMU conversion)
```

---

### VLM Model Recommendation

Based on benchmark testing (2026-06-08):

| Model | Speed | Accuracy | JSON Validity | Verdict |
|-------|-------|----------|---------------|---------|
| **qwen3.6-35b-a3b-q7** | 38s avg | 🏆 Best | 100% | **Recommended** |
| qwen3.6-27b-k-xl | 71s avg | Good depth | 67% (scale errors) | Too slow |
| qwen3.5-35b-a3b-gf | 8s avg | Overly optimistic | 100% | Too fast, misses issues |

**Winner: `qwen3.6-35b-a3b-q7`** — best balance of speed, accuracy, and instruction following.

---

### File Structure

```
ir-schema/
├── docir-v0.1.0.xsd          # XSD schema definition
├── DESIGN.md                  # This document
├── samples/
│   ├── small-pdf-sample.xml   # Hand-crafted sample IR from small.pdf
│   └── small-pdf-built.xml    # Auto-generated by IR Builder
├── builder/
│   ├── __init__.py
│   ├── ir_builder.py          # GLM-OCR JSON → DocIR XML converter
│   └── README.md              # Builder usage documentation
├── validator/
│   └── semantic_validator.py  # Semantic checks beyond XSD
└── tests/
    └── test_integration.py    # End-to-end pipeline test
```

---

### Implementation Status

#### ✅ Completed (v0.1.0)

1. **XSD Schema** — Complete DocIR v0.1.0 schema with all element types
2. **Semantic Validator** — 8 validation rules (bbox, table, region order, image refs, confidence, cross-page, provenance, style evidence)
3. **IR Builder** — GLM-OCR model JSON → DocIR XML converter
   - Coordinate conversion (normalized 0-1000 → PDF pt)
   - Region type mapping (text/table/image/formula)
   - Content cleaning (remove OCR repetition artifacts)
   - Image asset registration
   - Multi-page support
4. **Style Extractor** — VLM-based typography extraction
   - PDF region cropping (pymupdf @ 200 DPI)
   - VLM analysis (qwen3.6-35b-a3b-q7 @ localhost:11234)
   - Dual-layer style output (computed + evidence)
   - Provenance tracking (model, timing, confidence)
5. **Integration Test** — Full pipeline test (build → style → validate)
6. **VLM Model Benchmark** — qwen3.6-35b-a3b-q7 selected for style extraction

#### 🚧 In Progress

- Table markdown parser (extract rows/cols from OCR content)
- Confidence score extraction from layout detector

####  Planned

1. **DOCX generator** — DocIR XML → python-docx / direct OOXML
2. **Cross-page merge post-processor** — Handle table/paragraph continuation
3. **End-to-end visual test** — small.pdf → DocIR → DOCX → visual comparison

---

### Quick Start

```bash
# 1. Run GLM-OCR pipeline (produces model JSON)
cd ~/Works/glm-ocr
uv run --extra layout glmocr parse small.pdf \
  --mode selfhosted --layout-device cpu \
  --set pipeline.ocr_api.api_mode ollama_generate \
  --set pipeline.ocr_api.api_path /api/generate \
  --output ./output-small

# 2. Build DocIR XML
.venv/bin/python3 ir-schema/builder/ir_builder.py \
  output-small/small_model.json \
  small.pdf \
  -o small.docir.xml

# 3. Validate
.venv/bin/python3 ir-schema/validator/semantic_validator.py \
  small.docir.xml \
  --xsd ir-schema/docir-v0.1.0.xsd -v

# 4. Run integration test
.venv/bin/python3 ir-schema/tests/test_integration.py
```
