#!/usr/bin/env python3
"""
End-to-end test for stream table detection and DOCX rendering.

Tests:
1. Create PDF with borderless table
2. Run IR builder (marks as text)
3. Run stream table detector (converts to stream_table)
4. Run DOCX generator (renders borderless table)
5. Verify DOCX output
"""

import sys
import tempfile
from pathlib import Path

# Add parent to path
sys.path.insert(0, str(Path(__file__).parent.parent))

import pymupdf
from builder.ir_builder import build_docir
from builder.stream_table_detector import detect_stream_tables
from docx_generator.docx_generator import generate_docx
from lxml import etree


DOCIR_NS = "urn:docir:v0.1"


def create_test_pdf_with_stream_table(output_path: Path):
    """Create a PDF with a borderless table."""
    doc = pymupdf.open()
    page = doc.new_page(width=612, height=792)
    
    # Title
    page.insert_text((72, 50), "Stream Table Test", fontsize=16, fontname="helv")
    
    # Borderless table (3 columns, 5 rows)
    data = [
        ["Product", "Price", "Quantity"],
        ["Widget A", "$10.00", "100"],
        ["Widget B", "$25.50", "50"],
        ["Widget C", "$5.75", "200"],
        ["Widget D", "$15.00", "75"],
    ]
    
    col_positions = [72, 220, 340]
    for row_idx, row in enumerate(data):
        y = 100 + row_idx * 20
        for col_idx, cell in enumerate(row):
            page.insert_text((col_positions[col_idx], y), cell, fontsize=11, fontname="helv")
    
    # Some text after the table
    page.insert_text((72, 250), "End of table.", fontsize=11, fontname="helv")
    
    doc.save(str(output_path))
    doc.close()
    print(f"✓ Created test PDF: {output_path}")


def create_mock_model_json(pdf_path: Path, output_path: Path):
    """Create a mock model.json from the PDF (simulating GLM-OCR output)."""
    import json
    
    doc = pymupdf.open(str(pdf_path))
    page = doc[0]
    
    # GLM-OCR model.json format: list of pages, each page is list of regions
    # Region keys: index, label, content, bbox_2d (normalized 0-1000), polygon
    # Page is 612x792, so normalize: x_norm = x/612*1000, y_norm = y/792*1000
    page_w, page_h = 612, 792
    model_data = [
        [  # Page 0
            {
                "index": 0,
                "label": "text",
                "content": "Product Price Quantity\nWidget A $10.00 100\nWidget B $25.50 50\nWidget C $5.75 200\nWidget D $15.00 75",
                "bbox_2d": [60/page_w*1000, 85/page_h*1000, 450/page_w*1000, 210/page_h*1000],
                "polygon": [[60/page_w*1000, 85/page_h*1000], [450/page_w*1000, 85/page_h*1000], [450/page_w*1000, 210/page_h*1000], [60/page_w*1000, 210/page_h*1000]]
            },
            {
                "index": 1,
                "label": "text",
                "content": "End of table.",
                "bbox_2d": [60/page_w*1000, 240/page_h*1000, 300/page_w*1000, 260/page_h*1000],
                "polygon": [[60/page_w*1000, 240/page_h*1000], [300/page_w*1000, 240/page_h*1000], [300/page_w*1000, 260/page_h*1000], [60/page_w*1000, 260/page_h*1000]]
            }
        ]
    ]
    
    doc.close()
    
    with open(output_path, "w") as f:
        json.dump(model_data, f, indent=2)
    
    print(f"✓ Created mock model.json: {output_path}")


def test_stream_table_detection():
    """Test full stream table pipeline."""
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)
        
        # Step 1: Create test PDF
        pdf_path = tmpdir / "test.pdf"
        create_test_pdf_with_stream_table(pdf_path)
        
        # Step 2: Create mock model.json
        model_json = tmpdir / "model.json"
        create_mock_model_json(pdf_path, model_json)
        
        # Step 3: Build DocIR XML
        docir_path = tmpdir / "test.docir.xml"
        build_docir(model_json, pdf_path, docir_path)
        print(f"✓ Built DocIR XML")
        
        # Check that region is initially text
        tree = etree.parse(str(docir_path))
        root = tree.getroot()
        regions = root.findall(f".//{{{DOCIR_NS}}}region")
        text_regions = [r for r in regions if r.get("type") == "text"]
        print(f"  Initial text regions: {len(text_regions)}")
        
        # Step 4: Run stream table detection
        detect_stream_tables(docir_path, pdf_path)
        print(f"✓ Ran stream table detection")
        
        # Check that one region is now stream_table
        tree = etree.parse(str(docir_path))
        root = tree.getroot()
        regions = root.findall(f".//{{{DOCIR_NS}}}region")
        stream_tables = [r for r in regions if r.get("type") == "stream_table"]
        text_regions = [r for r in regions if r.get("type") == "text"]
        
        print(f"  Stream tables: {len(stream_tables)}")
        print(f"  Remaining text regions: {len(text_regions)}")
        
        if len(stream_tables) == 0:
            print("✗ FAIL: No stream tables detected")
            return False
        
        # Verify stream table structure
        st = stream_tables[0]
        table_content = st.find(f".//{{{DOCIR_NS}}}table_content")
        if table_content is None:
            print("✗ FAIL: No table_content in stream_table")
            return False
        
        num_rows = int(table_content.get("rows", "0"))
        num_cols = int(table_content.get("cols", "0"))
        table_type = table_content.get("type", "")
        
        print(f"  Table: {num_rows} rows x {num_cols} cols, type={table_type}")
        
        if num_rows != 5 or num_cols != 3:
            print(f"✗ FAIL: Expected 5x3, got {num_rows}x{num_cols}")
            return False
        
        if table_type != "stream":
            print(f"✗ FAIL: Expected type='stream', got '{table_type}'")
            return False
        
        # Verify cells
        cells = table_content.findall(f".//{{{DOCIR_NS}}}cell")
        print(f"  Cells: {len(cells)}")
        
        # Check first row (header)
        first_row_cells = [c for c in cells if c.get("row") == "0"]
        header_texts = [c.text for c in first_row_cells if c.text]
        print(f"  Header: {header_texts}")
        
        if "Product" not in header_texts or "Price" not in header_texts:
            print("✗ FAIL: Header cells not found")
            return False
        
        # Step 5: Generate DOCX
        docx_path = tmpdir / "test.docx"
        generate_docx(docir_path, docx_path)
        print(f"✓ Generated DOCX: {docx_path}")
        
        # Verify DOCX
        from docx import Document
        doc = Document(str(docx_path))
        
        tables = doc.tables
        print(f"  DOCX tables: {len(tables)}")
        
        if len(tables) == 0:
            print("✗ FAIL: No tables in DOCX")
            return False
        
        table = tables[0]
        print(f"  Table: {len(table.rows)} rows x {len(table.columns)} cols")
        
        if len(table.rows) != 5 or len(table.columns) != 3:
            print(f"✗ FAIL: Expected 5x3 table in DOCX, got {len(table.rows)}x{len(table.columns)}")
            return False
        
        # Check that borders are NOT visible (stream table)
        from docx.oxml.ns import qn
        tbl = table._tbl
        tblPr = tbl.tblPr
        if tblPr is not None:
            borders = tblPr.find(qn('w:tblBorders'))
            if borders is not None:
                print("✗ FAIL: Stream table should not have visible borders")
                return False
        
        print()
        print("=" * 60)
        print("✓ ALL STREAM TABLE TESTS PASSED")
        print("=" * 60)
        return True


if __name__ == "__main__":
    success = test_stream_table_detection()
    sys.exit(0 if success else 1)
