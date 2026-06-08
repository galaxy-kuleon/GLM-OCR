# DocIR Builder

將 GLM-OCR pipeline 輸出轉換為 DocIR XML 格式。

## 功能

- 讀取 GLM-OCR 的 model JSON（包含 regions、bbox、labels、content）
- 從 PDF 取得頁面尺寸
- 將 normalized bbox (0-1000) 轉換為 PDF points
- 產生符合 DocIR v0.1.0 schema 的 XML

## 座標轉換

GLM-OCR 使用 normalized coordinates (0-1000)，原點在左上角。  
DocIR 使用 PDF points (1/72 inch)，原點在左下角。

轉換公式：
```
pt_x = (norm_x / 1000) * page_width_pt
pt_y = page_height_pt - (norm_y / 1000) * page_height_pt
```

## 使用方式

### Command Line

```bash
# 基本用法
python ir-schema/builder/ir_builder.py \
  output/doc_model.json \
  source.pdf \
  -o output.docir.xml

# 指定 title
python ir-schema/builder/ir_builder.py \
  output/doc_model.json \
  source.pdf \
  -o output.docir.xml \
  --title "My Document"
```

### Python API

```python
from pathlib import Path
from ir_schema.builder.ir_builder import build_docir

build_docir(
    model_json_path=Path("output/doc_model.json"),
    pdf_path=Path("source.pdf"),
    output_path=Path("output.docir.xml"),
    pipeline_info={
        "title": "My Document",
        "source": "glm-ocr-pipeline",
        "detection_model": "PP-DocLayoutV3",
        "ocr_engine": "Ollama glm-ocr:latest",
        "style_extractor": "qwen3.6-35b-a3b-q7"
    }
)
```

## 輸入格式

### GLM-OCR Model JSON

```json
[
  [
    {
      "index": 0,
      "label": "text",
      "bbox_2d": [366, 83, 631, 116],
      "polygon": [[366, 83], [366, 115], [630, 115], [630, 83]],
      "content": "This is text content"
    },
    {
      "index": 1,
      "label": "table",
      "bbox_2d": [94, 321, 903, 489],
      "polygon": [[94, 321], [94, 488], [903, 488], [903, 321]],
      "content": "| Col1 | Col2 |\n|------|------|\n| A    | B    |"
    },
    {
      "index": 2,
      "label": "image",
      "bbox_2d": [55, 528, 442, 799],
      "polygon": [[55, 528], [55, 799], [441, 799], [441, 528]],
      "content": null,
      "image_path": "imgs/cropped_page0_idx2.jpg"
    }
  ]
]
```

## 輸出格式

DocIR XML v0.1.0，包含：

- `<docir:metadata>` — 文件資訊、pipeline 設定
- `<docir:pages>` — 頁面列表，每頁包含 regions
- `<docir:regions>` — 區域列表（text/table/image/formula）
- `<docir:assets>` — 圖片資源註冊
- `<docir:cross_page_hints>` — 跨頁元素提示

每個 region 包含：
- `id`, `type`, `native_label`, `order`
- `<docir:bbox>` — PDF points 座標
- `<docir:polygon>` — 多邊形頂點
- `<docir:provenance>` — 來源、信心度、模型
- `<docir:text_content>` / `<docir:table_content>` / `<docir:image_content>` — 內容

## 驗證

使用 semantic validator 檢查輸出的正確性：

```bash
python ir-schema/validator/semantic_validator.py \
  output.docir.xml \
  --xsd ir-schema/docir-v0.1.0.xsd \
  -v
```

檢查項目：
- XSD schema validation
- bbox 是否在頁面範圍內
- table rows/cols 是否匹配實際 cell 數量
- region order 是否遞增
- image references 是否已註冊
- confidence 是否在 [0, 1] 範圍內

## 限制

目前版本（v0.1.0）的限制：

1. **Confidence scores** — model JSON 沒有 confidence，使用 placeholder (0.85)
2. **Table parsing** — 只產生 minimal table structure，未解析 markdown table
3. **Style extraction** — 未整合 VLM style extraction（font-size, color, bold）
4. **Cross-page merge** — 未實作跨頁 table/paragraph merge
5. **Image assets** — 只註冊路徑，未提取 dimensions 或 visual features

## 下一步

- [ ] 整合 GLM-OCR layout detector 的 confidence scores
- [ ] 實作 table markdown parser
- [ ] 整合 VLM style extraction（qwen3.6-35b-a3b-q7）
- [ ] 實作 cross-page merge hints
- [ ] 提取 image dimensions 和 visual features

## 相關文件

- [DESIGN.md](../DESIGN.md) — DocIR 設計文件
- [docir-v0.1.0.xsd](../docir-v0.1.0.xsd) — XSD schema
- [semantic_validator.py](../validator/semantic_validator.py) — 驗證器
