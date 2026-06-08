#!/usr/bin/env python3
"""
Stream Table Detector - Detect borderless tables from text alignment patterns.

Algorithm (inspired by pdf2docx):
1. For each text region, collect all line x0 positions
2. Cluster x-positions into columns (within threshold)
3. If 2+ columns with consistent vertical alignment, mark as stream table
4. Output: table_content with rows/cols derived from text alignment

This handles tables without visible borders but with aligned columns.
"""

import pymupdf
from collections import defaultdict
from pathlib import Path
from typing import Optional, Dict, List, Tuple
from lxml import etree


DOCIR_NS = "urn:docir:v0.1"


def detect_stream_tables(
    docir_path: Path,
    pdf_path: Path,
    output_path: Optional[Path] = None,
    col_threshold: float = 10.0,
    min_rows: int = 2,
    min_cols: int = 2,
    min_confidence: float = 0.6
) -> Path:
    """
    Detect stream tables (borderless) in DocIR and update region types.
    
    Args:
        docir_path: Path to input DocIR XML
        pdf_path: Path to source PDF
        output_path: Path to output DocIR XML (default: overwrite input)
        col_threshold: Max x-distance to consider same column (points)
        min_rows: Minimum rows to qualify as table
        min_cols: Minimum columns to qualify as table
        min_confidence: Minimum confidence to mark as table
    
    Returns:
        Path to output DocIR XML
    """
    if output_path is None:
        output_path = docir_path
    
    # Load DocIR XML
    tree = etree.parse(str(docir_path))
    root = tree.getroot()
    
    # Open PDF
    doc = pymupdf.open(str(pdf_path))
    
    tables_detected = 0
    
    # Process each page
    for page_elem in root.findall(f".//{{{DOCIR_NS}}}page"):
        page_idx = int(page_elem.get("index", 0))
        if page_idx >= len(doc):
            continue
        
        page = doc[page_idx]
        text_dict = page.get_text("dict")
        
        # Process each text region
        for region in page_elem.findall(f".//{{{DOCIR_NS}}}region"):
            region_type = region.get("type")
            
            # Only process text regions (not already tables)
            if region_type not in ["text", "paragraph"]:
                continue
            
            # Get region bbox
            bbox_elem = region.find(f".//{{{DOCIR_NS}}}bbox")
            if bbox_elem is None:
                continue
            
            x = float(bbox_elem.get("x", 0))
            y = float(bbox_elem.get("y", 0))  # y is TOP of region in PDF coords (from bottom)
            w = float(bbox_elem.get("width", 0))
            h = float(bbox_elem.get("height", 0))
            
            # DocIR bbox is in PDF coords (y from bottom), but PyMuPDF text dict
            # uses y from TOP. Convert:
            # - PDF y_top (from bottom) -> screen y_top (from top) = page_height - y
            # - PDF y_bottom (from bottom) = y - h -> screen y_bottom = page_height - (y - h) = page_height - y + h
            page_height = page.rect.height
            screen_y0 = page_height - y         # top of region (from top of page)
            screen_y1 = page_height - y + h     # bottom of region (from top of page)
            region_rect = pymupdf.Rect(x, screen_y0, x + w, screen_y1)
            
            # Detect stream table
            result = detect_stream_table(
                text_dict, region_rect, 
                col_threshold=col_threshold
            )
            
            if result and result["num_rows"] >= min_rows and result["num_cols"] >= min_cols:
                if result["confidence"] >= min_confidence:
                    # Convert to stream table
                    _convert_to_stream_table(region, result, page, region_rect)
                    tables_detected += 1
    
    doc.close()
    
    # Write output
    tree.write(
        str(output_path),
        xml_declaration=True,
        encoding="UTF-8",
        pretty_print=True
    )
    
    print(f"✓ Stream table detection complete")
    print(f"  Tables detected: {tables_detected}")
    print(f"  Output: {output_path}")
    
    return output_path


def detect_stream_table(
    text_dict: dict,
    region_rect: pymupdf.Rect,
    col_threshold: float = 10.0
) -> Optional[Dict]:
    """
    Detect borderless table from text alignment patterns.
    
    Args:
        text_dict: PyMuPDF text dictionary
        region_rect: Region bounding box
        col_threshold: Max x-distance for same column
    
    Returns:
        Dict with columns, num_rows, num_cols, confidence, or None
    """
    lines_by_y = defaultdict(list)  # y_center -> [(x0, x1, text)]
    
    for block in text_dict.get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            bbox = line.get("bbox")
            if not bbox:
                continue
            x0, y0, x1, y1 = bbox
            line_rect = pymupdf.Rect(x0, y0, x1, y1)
            if not region_rect.intersects(line_rect):
                continue
            
            # Get text content
            text = " ".join(span.get("text", "") for span in line.get("spans", []))
            if not text.strip():
                continue
            
            # Use y-center for row grouping
            y_center = (y0 + y1) / 2
            lines_by_y[y_center].append((x0, x1, text))
    
    if len(lines_by_y) < 2:
        return None  # Need at least 2 rows
    
    # Sort by y position
    sorted_rows = sorted(lines_by_y.items(), key=lambda x: x[0])
    
    # Collect all x0 positions
    all_x0 = []
    for y, items in sorted_rows:
        for x0, x1, text in items:
            all_x0.append(x0)
    
    if not all_x0:
        return None
    
    # Cluster x0 positions into columns
    all_x0.sort()
    columns = []
    current_col = [all_x0[0]]
    
    for x0 in all_x0[1:]:
        if x0 - current_col[-1] < col_threshold:
            current_col.append(x0)
        else:
            columns.append(sum(current_col) / len(current_col))
            current_col = [x0]
    columns.append(sum(current_col) / len(current_col))
    
    if len(columns) < 2:
        return None  # Need at least 2 columns
    
    # Check column consistency: most rows should have text in most columns
    row_col_counts = []
    for y, items in sorted_rows:
        cols_in_row = set()
        for x0, x1, text in items:
            # Find which column this belongs to
            for i, col_x in enumerate(columns):
                if abs(x0 - col_x) < col_threshold:
                    cols_in_row.add(i)
                    break
        row_col_counts.append(len(cols_in_row))
    
    # If most rows have 2+ columns, it's a table
    avg_cols = sum(row_col_counts) / len(row_col_counts) if row_col_counts else 0
    if avg_cols >= 2:
        return {
            "columns": columns,
            "num_rows": len(sorted_rows),
            "num_cols": len(columns),
            "confidence": min(1.0, avg_cols / len(columns)),
            "rows_data": sorted_rows,
            "column_boundaries": columns
        }
    
    return None


def _convert_to_stream_table(
    region: etree.Element,
    detection: Dict,
    page: pymupdf.Page,
    region_rect: pymupdf.Rect
):
    """Convert a text region to a stream table region."""
    region_id = region.get("id", "?")
    
    # Change region type
    region.set("type", "stream_table")
    
    # Remove text_content
    text_content = region.find(f".//{{{DOCIR_NS}}}text_content")
    if text_content is not None:
        region.remove(text_content)
    
    # Create table_content
    table_content = etree.SubElement(region, f"{{{DOCIR_NS}}}table_content")
    table_content.set("rows", str(detection["num_rows"]))
    table_content.set("cols", str(detection["num_cols"]))
    table_content.set("type", "stream")  # Mark as borderless/stream
    
    # Add column boundaries
    for i, col_x in enumerate(detection["columns"]):
        col_elem = etree.SubElement(table_content, f"{{{DOCIR_NS}}}column")
        col_elem.set("index", str(i))
        col_elem.set("x", f"{col_x:.2f}")
        # Calculate column width (distance to next column or region edge)
        if i < len(detection["columns"]) - 1:
            width = detection["columns"][i + 1] - col_x
        else:
            width = region_rect.x1 - col_x
        col_elem.set("width", f"{width:.2f}")
    
    # Extract cell data from rows
    rows_data = detection["rows_data"]
    columns = detection["columns"]
    col_threshold = 10.0
    
    for row_idx, (y_center, items) in enumerate(rows_data):
        row_elem = etree.SubElement(table_content, f"{{{DOCIR_NS}}}row")
        row_elem.set("index", str(row_idx))
        
        # Group items by column
        cells = defaultdict(str)
        for x0, x1, text in items:
            # Find which column
            for col_idx, col_x in enumerate(columns):
                if abs(x0 - col_x) < col_threshold:
                    cells[col_idx] = text.strip()
                    break
        
        # Create cells
        for col_idx in range(len(columns)):
            cell_elem = etree.SubElement(row_elem, f"{{{DOCIR_NS}}}cell")
            cell_elem.set("row", str(row_idx))
            cell_elem.set("col", str(col_idx))
            cell_text = cells.get(col_idx, "")
            if cell_text:
                cell_elem.text = cell_text
    
    print(f"  ✓ Region {region_id}: stream table {detection['num_rows']}x{detection['num_cols']} (conf={detection['confidence']:.2f})")


if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description="Detect stream tables in DocIR")
    parser.add_argument("docir", type=Path, help="Input DocIR XML")
    parser.add_argument("pdf", type=Path, help="Source PDF")
    parser.add_argument("-o", "--output", type=Path, help="Output DocIR XML")
    parser.add_argument("--col-threshold", type=float, default=10.0,
                       help="Column detection threshold (points)")
    parser.add_argument("--min-rows", type=int, default=2,
                       help="Minimum rows for table")
    parser.add_argument("--min-cols", type=int, default=2,
                       help="Minimum columns for table")
    parser.add_argument("--min-confidence", type=float, default=0.6,
                       help="Minimum confidence threshold")
    
    args = parser.parse_args()
    
    detect_stream_tables(
        args.docir,
        args.pdf,
        args.output,
        col_threshold=args.col_threshold,
        min_rows=args.min_rows,
        min_cols=args.min_cols,
        min_confidence=args.min_confidence
    )
