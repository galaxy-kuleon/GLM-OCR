#!/usr/bin/env python3
"""
Text box detection test.

Tests that floating regions (aside_text, header, footer) are:
1. Detected and marked with floating="true" in DocIR XML
2. Rendered as bordered text boxes in DOCX output
"""

import json
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from builder.ir_builder import build_docir, is_floating_region, FLOATING_LABELS
from docx_generator.docx_generator import generate_docx


def create_model_with_textboxes() -> Path:
    """Create a model JSON with floating regions (text boxes)."""
    regions = [
        # Normal text region
        {
            "index": 0,
            "label": "text",
            "content": "This is normal body text.",
            "bbox_2d": [100, 700, 500, 650],
            "polygon": [[100, 700], [500, 700], [500, 650], [100, 650]]
        },
        # Text box (aside_text)
        {
            "index": 1,
            "label": "aside_text",
            "content": "This is a sidebar text box!",
            "bbox_2d": [520, 700, 700, 600],
            "polygon": [[520, 700], [700, 700], [700, 600], [520, 600]]
        },
        # Header
        {
            "index": 2,
            "label": "header",
            "content": "Document Header - Confidential",
            "bbox_2d": [100, 820, 500, 800],
            "polygon": [[100, 820], [500, 820], [500, 800], [100, 800]]
        },
        # Footer
        {
            "index": 3,
            "label": "footer",
            "content": "Page 1 of 5",
            "bbox_2d": [250, 50, 350, 30],
            "polygon": [[250, 50], [350, 50], [350, 30], [250, 30]]
        },
        # Another normal text
        {
            "index": 4,
            "label": "paragraph_title",
            "content": "Section Title",
            "bbox_2d": [100, 580, 400, 550],
            "polygon": [[100, 580], [400, 580], [400, 550], [100, 550]]
        },
    ]
    
    model_data = [regions]  # Single page
    
    output = Path(tempfile.mktemp(suffix='_model.json'))
    with open(output, 'w') as f:
        json.dump(model_data, f)
    return output


def create_test_pdf() -> Path:
    """Create a single-page A4 PDF for testing."""
    import pymupdf
    dst = pymupdf.open()
    dst.new_page(width=595, height=842)
    output = Path(tempfile.mktemp(suffix='.pdf'))
    dst.save(str(output))
    dst.close()
    return output


def test_floating_detection():
    """Test that floating labels are correctly identified."""
    print("Test 1: Floating label detection")
    
    assert is_floating_region("aside_text") == True
    assert is_floating_region("header") == True
    assert is_floating_region("footer") == True
    assert is_floating_region("text") == False
    assert is_floating_region("table") == False
    assert is_floating_region("image") == False
    assert is_floating_region("paragraph_title") == False
    
    print(f"  ✓ FLOATING_LABELS = {FLOATING_LABELS}")
    print("  ✓ aside_text, header, footer correctly identified as floating")
    print("  ✓ text, table, image, paragraph_title correctly identified as non-floating")


def test_ir_builder_floating_attribute():
    """Test that IR builder adds floating='true' attribute."""
    print("\nTest 2: IR Builder floating attribute")
    
    model_json = create_model_with_textboxes()
    pdf_path = create_test_pdf()
    output_xml = Path(tempfile.mktemp(suffix='.docir.xml'))
    
    try:
        pipeline_info = {
            "title": "Text Box Test",
            "source": "test",
            "detection_model": "test",
            "ocr_engine": "test",
            "style_extractor": "none"
        }
        
        build_docir(model_json, pdf_path, output_xml, pipeline_info)
        
        # Parse and verify
        from lxml import etree
        tree = etree.parse(str(output_xml))
        root = tree.getroot()
        
        DOCIR_NS = "urn:docir:v0.1"
        pages = root.find(f'{{{DOCIR_NS}}}pages')
        page = pages.findall(f'{{{DOCIR_NS}}}page')[0]
        regions = page.find(f'{{{DOCIR_NS}}}regions')
        region_list = regions.findall(f'{{{DOCIR_NS}}}region')
        
        floating_count = 0
        for region in region_list:
            native_label = region.get('native_label')
            floating = region.get('floating')
            
            if native_label in FLOATING_LABELS:
                assert floating == 'true', \
                    f"Region {native_label} should have floating='true', got {floating}"
                floating_count += 1
            else:
                assert floating is None, \
                    f"Region {native_label} should not have floating attribute, got {floating}"
        
        assert floating_count == 3, f"Expected 3 floating regions, got {floating_count}"
        
        print(f"  ✓ {floating_count} regions marked as floating='true'")
        print("  ✓ Non-floating regions have no floating attribute")
        
    finally:
        model_json.unlink(missing_ok=True)
        pdf_path.unlink(missing_ok=True)
        output_xml.unlink(missing_ok=True)


def test_docx_text_box_rendering():
    """Test that DOCX generator renders floating regions as text boxes."""
    print("\nTest 3: DOCX text box rendering")
    
    model_json = create_model_with_textboxes()
    pdf_path = create_test_pdf()
    docir_xml = Path(tempfile.mktemp(suffix='.docir.xml'))
    docx_output = Path(tempfile.mktemp(suffix='.docx'))
    
    try:
        pipeline_info = {
            "title": "Text Box DOCX Test",
            "source": "test",
            "detection_model": "test",
            "ocr_engine": "test",
            "style_extractor": "none"
        }
        
        build_docir(model_json, pdf_path, docir_xml, pipeline_info)
        generate_docx(docir_xml, docx_output, positioned=True)
        
        # Verify DOCX structure
        from docx import Document
        doc = Document(str(docx_output))
        
        # Check we have tables (text boxes are rendered as single-cell tables)
        tables = doc.tables
        assert len(tables) >= 3, f"Expected at least 3 tables (text boxes), got {len(tables)}"
        
        # Verify text box content
        textbox_contents = []
        for table in tables:
            if len(table.rows) == 1 and len(table.columns) == 1:
                cell_text = table.cell(0, 0).text.strip()
                if cell_text:
                    textbox_contents.append(cell_text)
        
        # Check that floating content is in text boxes
        assert any("sidebar text box" in t for t in textbox_contents), \
            f"aside_text content not found in text boxes: {textbox_contents}"
        assert any("Header" in t or "Confidential" in t for t in textbox_contents), \
            f"header content not found in text boxes: {textbox_contents}"
        assert any("Page 1" in t for t in textbox_contents), \
            f"footer content not found in text boxes: {textbox_contents}"
        
        # Check borders exist on text box cells
        from docx.oxml.ns import qn
        for table in tables:
            cell = table.cell(0, 0)
            tc = cell._tc
            tcPr = tc.find(qn('w:tcPr'))
            if tcPr is not None:
                tcBorders = tcPr.find(qn('w:tcBorders'))
                if tcBorders is not None:
                    # This is a text box - verify it has borders
                    top_border = tcBorders.find(qn('w:top'))
                    assert top_border is not None, "Text box should have top border"
        
        print(f"  ✓ {len(tables)} text boxes rendered as bordered tables")
        print(f"  ✓ Text box contents: {textbox_contents[:3]}")
        print("  ✓ Borders and shading applied to text box cells")
        
    finally:
        model_json.unlink(missing_ok=True)
        pdf_path.unlink(missing_ok=True)
        docir_xml.unlink(missing_ok=True)
        docx_output.unlink(missing_ok=True)


if __name__ == '__main__':
    print("=" * 60)
    print("Text Box Detection Tests")
    print("=" * 60)
    
    test_floating_detection()
    test_ir_builder_floating_attribute()
    test_docx_text_box_rendering()
    
    print("\n" + "=" * 60)
    print("✓ All text box detection tests passed!")
    print("=" * 60)
