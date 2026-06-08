#!/usr/bin/env python3
"""
DocIR Builder v0.1.0

Converts GLM-OCR pipeline output to DocIR XML format.

Input:
  - GLM-OCR model JSON (with regions, bbox, labels, content)
  - Source PDF (for page dimensions)

Output:
  - DocIR XML document (validated against docir-v0.1.0.xsd)

Coordinate conversion:
  - GLM-OCR bbox: normalized 0-1000
  - DocIR bbox: PDF points (1/72 inch)
  - Formula: pt_x = (norm_x / 1000) * page_width_pt
"""

import json
import sys
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Any, Optional, Tuple
from lxml import etree
import pymupdf
from PIL import Image


# DocIR namespace
DOCIR_NS = "urn:docir:v0.1"
NSMAP = {"docir": DOCIR_NS}


def docir_tag(name: str) -> str:
    """Create a fully qualified DocIR tag name."""
    return f"{{{DOCIR_NS}}}{name}"


def get_pdf_page_dimensions(pdf_path: Path) -> List[Tuple[float, float]]:
    """
    Extract page dimensions from PDF.
    
    Returns:
        List of (width_pt, height_pt) tuples, one per page.
    """
    doc = pymupdf.open(str(pdf_path))
    dimensions = []
    for page in doc:
        dimensions.append((page.rect.width, page.rect.height))
    doc.close()
    return dimensions


def normalize_bbox_to_pt(
    bbox_norm: List[float],
    page_width_pt: float,
    page_height_pt: float
) -> Tuple[float, float, float, float]:
    """
    Convert normalized bbox (0-1000) to PDF points.
    
    GLM-OCR bbox format: [x1, y1, x2, y2] where coordinates are normalized 0-1000.
    Origin is top-left in GLM-OCR, but PDF origin is bottom-left.
    
    Returns:
        (x, y, width, height) in PDF points, with y adjusted for PDF coordinate system.
    """
    x1_norm, y1_norm, x2_norm, y2_norm = bbox_norm
    
    # Convert to PDF points
    x1_pt = (x1_norm / 1000.0) * page_width_pt
    x2_pt = (x2_norm / 1000.0) * page_width_pt
    
    # GLM-OCR y is top-down, PDF y is bottom-up
    y1_pt = page_height_pt - (y1_norm / 1000.0) * page_height_pt
    y2_pt = page_height_pt - (y2_norm / 1000.0) * page_height_pt
    
    # Calculate width and height
    width_pt = abs(x2_pt - x1_pt)
    height_pt = abs(y2_pt - y1_pt)
    
    # Use top-left corner as (x, y) for DocIR
    x_pt = min(x1_pt, x2_pt)
    y_pt = max(y1_pt, y2_pt)  # y is bottom-left in PDF, so use the larger value
    
    return (x_pt, y_pt, width_pt, height_pt)


def polygon_to_pt(
    polygon_norm: List[List[float]],
    page_width_pt: float,
    page_height_pt: float
) -> List[Tuple[float, float]]:
    """
    Convert normalized polygon coordinates to PDF points.
    
    Returns:
        List of (x, y) tuples in PDF points.
    """
    points_pt = []
    for x_norm, y_norm in polygon_norm:
        x_pt = (x_norm / 1000.0) * page_width_pt
        y_pt = page_height_pt - (y_norm / 1000.0) * page_height_pt
        points_pt.append((x_pt, y_pt))
    return points_pt


def clean_ocr_content(content: str) -> str:
    """
    Clean OCR content by removing repetition artifacts and markdown fences.
    
    GLM-OCR sometimes:
    - Repeats content with "Text Recognition:" prefix
    - Wraps content in markdown code fences (```markdown ... ```)
    - Produces empty blocks
    
    This function extracts the first meaningful line and removes artifacts.
    """
    if not content:
        return ""
    
    lines = content.strip().split('\n')
    if not lines:
        return ""
    
    # Filter out empty lines and markdown fences
    clean_lines = []
    in_code_fence = False
    
    for line in lines:
        stripped = line.strip()
        
        # Skip empty lines
        if not stripped:
            continue
        
        # Skip markdown code fences
        if stripped.startswith('```'):
            in_code_fence = not in_code_fence
            continue
        
        # Skip lines inside code fences
        if in_code_fence:
            continue
        
        # Remove "Text Recognition:" prefix if present
        if stripped.startswith("Text Recognition:"):
            stripped = stripped[len("Text Recognition:"):].strip()
        
        # Skip if line is now empty after removing prefix
        if not stripped:
            continue
        
        clean_lines.append(stripped)
    
    if not clean_lines:
        return ""
    
    # Return the first clean line (most reliable)
    # TODO: Could be smarter about combining multiple lines
    return clean_lines[0]


def determine_region_type(label: str) -> str:
    """
    Map GLM-OCR label to DocIR region type.
    
    GLM-OCR labels: text, table, image, formula, chart, seal, etc.
    DocIR types: text, table, image, formula, chart, seal
    """
    # Direct mapping for most labels
    type_map = {
        "text": "text",
        "table": "table",
        "image": "image",
        "formula": "formula",
        "chart": "chart",
        "seal": "seal",
        "paragraph_title": "text",  # Titles are text regions
        # Additional PP-DocLayoutV3 labels
        "doc_title": "text",
        "figure_title": "text",
        "aside_text": "text",  # Text boxes / floating text
        "header": "text",
        "footer": "text",
        "footnote": "text",
        "reference": "text",
        "reference_content": "text",
        "abstract": "text",
        "content": "text",
        "algorithm": "text",
        "number": "text",
        "formula_number": "text",
        "vision_footnote": "text",
    }
    
    return type_map.get(label, "text")


# Labels that represent floating/positioned elements (text boxes)
FLOATING_LABELS = {"aside_text", "header", "footer"}


def is_floating_region(label: str) -> bool:
    """
    Check if a region label represents a floating/text box element.
    
    Floating regions are positioned independently from the main text flow:
    - aside_text: sidebars, callout boxes, text boxes
    - header: page headers
    - footer: page footers
    """
    return label in FLOATING_LABELS


def build_text_content(content: str) -> etree.Element:
    """
    Build DocIR text_content element from OCR text.
    
    Structure:
      <text_content>
        <paragraph>
          <run>text</run>
        </paragraph>
      </text_content>
    """
    text_content = etree.Element(f"{{{DOCIR_NS}}}text_content")
    
    # Split content into paragraphs (by newlines)
    paragraphs = [p.strip() for p in content.split('\n') if p.strip()]
    
    if not paragraphs:
        # Empty content - create paragraph with empty run (XSD requires at least one run)
        para = etree.SubElement(text_content, f"{{{DOCIR_NS}}}paragraph")
        run = etree.SubElement(para, f"{{{DOCIR_NS}}}run")
        run.text = ""
        return text_content
    
    for para_text in paragraphs:
        para = etree.SubElement(text_content, f"{{{DOCIR_NS}}}paragraph")
        run = etree.SubElement(para, f"{{{DOCIR_NS}}}run")
        run.text = para_text
    
    return text_content


def build_table_content(content: str) -> etree.Element:
    """
    Build DocIR table_content element from OCR content.
    
    Uses table_parser to extract structured table data.
    """
    import sys
    from pathlib import Path
    
    # Add parent directory to path for import
    builder_dir = Path(__file__).parent
    if str(builder_dir) not in sys.path:
        sys.path.insert(0, str(builder_dir))
    
    from table_parser import parse_table_content, TableData
    
    table_data = parse_table_content(content)
    
    # Build XML element
    table_content = etree.Element(f"{{{DOCIR_NS}}}table_content")
    table_content.set("rows", str(table_data.num_rows))
    table_content.set("cols", str(table_data.num_cols))
    
    # Table style
    table_style = etree.SubElement(table_content, f"{{{DOCIR_NS}}}table_style")
    table_style.set("border_visible", "true")
    table_style.set("border_color", "#000000")
    if table_data.has_header:
        table_style.set("header_row", "true")
    
    # Row groups
    if table_data.has_header:
        # Header row group
        header_group = etree.SubElement(table_content, f"{{{DOCIR_NS}}}row_group")
        header_group.set("type", "header")
        header_row_elem = etree.SubElement(header_group, f"{{{DOCIR_NS}}}row")
        
        for cell in table_data.rows[0]:
            cell_elem = etree.SubElement(header_row_elem, f"{{{DOCIR_NS}}}cell")
            if cell.col_span > 1:
                cell_elem.set("col_span", str(cell.col_span))
            if cell.row_span > 1:
                cell_elem.set("row_span", str(cell.row_span))
            
            text_content = etree.SubElement(cell_elem, f"{{{DOCIR_NS}}}text_content")
            para = etree.SubElement(text_content, f"{{{DOCIR_NS}}}paragraph")
            run = etree.SubElement(para, f"{{{DOCIR_NS}}}run")
            run.text = cell.content
        
        # Body row group
        if len(table_data.rows) > 1:
            body_group = etree.SubElement(table_content, f"{{{DOCIR_NS}}}row_group")
            body_group.set("type", "body")
            
            for row in table_data.rows[1:]:
                row_elem = etree.SubElement(body_group, f"{{{DOCIR_NS}}}row")
                for cell in row:
                    cell_elem = etree.SubElement(row_elem, f"{{{DOCIR_NS}}}cell")
                    if cell.col_span > 1:
                        cell_elem.set("col_span", str(cell.col_span))
                    if cell.row_span > 1:
                        cell_elem.set("row_span", str(cell.row_span))
                    
                    text_content = etree.SubElement(cell_elem, f"{{{DOCIR_NS}}}text_content")
                    para = etree.SubElement(text_content, f"{{{DOCIR_NS}}}paragraph")
                    run = etree.SubElement(para, f"{{{DOCIR_NS}}}run")
                    run.text = cell.content
    else:
        # All body rows
        body_group = etree.SubElement(table_content, f"{{{DOCIR_NS}}}row_group")
        body_group.set("type", "body")
        
        for row in table_data.rows:
            row_elem = etree.SubElement(body_group, f"{{{DOCIR_NS}}}row")
            for cell in row:
                cell_elem = etree.SubElement(row_elem, f"{{{DOCIR_NS}}}cell")
                if cell.col_span > 1:
                    cell_elem.set("col_span", str(cell.col_span))
                if cell.row_span > 1:
                    cell_elem.set("row_span", str(cell.row_span))
                
                text_content = etree.SubElement(cell_elem, f"{{{DOCIR_NS}}}text_content")
                para = etree.SubElement(text_content, f"{{{DOCIR_NS}}}paragraph")
                run = etree.SubElement(para, f"{{{DOCIR_NS}}}run")
                run.text = cell.content
    
    return table_content


def build_image_content(region: Dict[str, Any], page_idx: int = 0) -> etree.Element:
    """
    Build DocIR image_content element.
    
    Structure:
      <image_content>
        <image_reference asset_id="..." mime_type="image/jpeg"/>
        <visual_features>
          <contains_text>false</contains_text>
        </visual_features>
      </image_content>
    """
    image_content = etree.Element(f"{{{DOCIR_NS}}}image_content")
    
    # Image reference
    image_ref = etree.SubElement(image_content, f"{{{DOCIR_NS}}}image_reference")
    asset_id = f"img_page{page_idx}_region{region.get('index', 0)}"
    image_ref.set("asset_id", asset_id)
    image_ref.set("mime_type", "image/jpeg")
    
    # Visual features (minimal for now)
    visual_features = etree.SubElement(image_content, f"{{{DOCIR_NS}}}visual_features")
    contains_text = etree.SubElement(visual_features, f"{{{DOCIR_NS}}}contains_text")
    contains_text.text = "false"
    
    return image_content


def build_provenance(region: Dict[str, Any], pipeline_info: Dict[str, str]) -> etree.Element:
    """
    Build DocIR provenance element.
    
    Structure:
      <provenance>
        <source>glm-ocr-pipeline</source>
        <confidence>0.85</confidence>
        <detection_model>PP-DocLayoutV3</detection_model>
      </provenance>
    """
    provenance = etree.Element(f"{{{DOCIR_NS}}}provenance")
    
    source = etree.SubElement(provenance, f"{{{DOCIR_NS}}}source")
    source.text = pipeline_info.get("source", "glm-ocr-pipeline")
    
    # Confidence is not in the model JSON, use placeholder
    # TODO: Extract confidence from layout detector output
    confidence = etree.SubElement(provenance, f"{{{DOCIR_NS}}}confidence")
    confidence.text = "0.85"  # Placeholder
    
    detection_model = etree.SubElement(provenance, f"{{{DOCIR_NS}}}detection_model")
    detection_model.text = pipeline_info.get("detection_model", "PP-DocLayoutV3")
    
    return provenance


def build_region(
    region: Dict[str, Any],
    page_width_pt: float,
    page_height_pt: float,
    pipeline_info: Dict[str, str],
    page_idx: int = 0
) -> etree.Element:
    """
    Build DocIR region element from GLM-OCR region.
    """
    region_elem = etree.Element(f"{{{DOCIR_NS}}}region")
    
    # Attributes
    region_id = f"r{region['index']}_p{page_idx}"
    region_elem.set("id", region_id)
    region_elem.set("type", determine_region_type(region["label"]))
    region_elem.set("native_label", region["label"])
    region_elem.set("order", str(region["index"]))
    
    # Mark floating/text box regions
    if is_floating_region(region["label"]):
        region_elem.set("floating", "true")
    
    # Bounding box (convert to PDF pt)
    bbox_pt = normalize_bbox_to_pt(region["bbox_2d"], page_width_pt, page_height_pt)
    bbox_elem = etree.SubElement(region_elem, f"{{{DOCIR_NS}}}bbox")
    bbox_elem.set("x", f"{bbox_pt[0]:.2f}")
    bbox_elem.set("y", f"{bbox_pt[1]:.2f}")
    bbox_elem.set("width", f"{bbox_pt[2]:.2f}")
    bbox_elem.set("height", f"{bbox_pt[3]:.2f}")
    
    # Polygon (if present)
    if "polygon" in region and region["polygon"]:
        polygon_pt = polygon_to_pt(region["polygon"], page_width_pt, page_height_pt)
        polygon_elem = etree.SubElement(region_elem, f"{{{DOCIR_NS}}}polygon")
        for x, y in polygon_pt:
            point = etree.SubElement(polygon_elem, f"{{{DOCIR_NS}}}point")
            point.set("x", f"{x:.2f}")
            point.set("y", f"{y:.2f}")
    
    # Provenance
    provenance = build_provenance(region, pipeline_info)
    region_elem.append(provenance)
    
    # Content (based on region type)
    region_type = determine_region_type(region["label"])
    content = clean_ocr_content(region.get("content", ""))
    
    if region_type == "text":
        text_content = build_text_content(content)
        region_elem.append(text_content)
    elif region_type == "table":
        table_content = build_table_content(content)
        region_elem.append(table_content)
    elif region_type == "image":
        image_content = build_image_content(region, page_idx)
        region_elem.append(image_content)
    else:
        # For other types, use text content as fallback
        text_content = build_text_content(content)
        region_elem.append(text_content)
    
    return region_elem


def detect_cross_page_continuations(
    model_data: List[List[Dict[str, Any]]],
    pages_elem,
    ns: str
) -> List[Dict[str, Any]]:
    """
    Detect cross-page continuations (tables/paragraphs split across pages).
    
    Heuristics:
    - Table continuation: last table on page N AND first table on page N+1
    - Paragraph continuation: last text on page N doesn't end with sentence punctuation
      AND first text on page N+1 starts with lowercase
    
    Returns list of merge hint dicts.
    """
    hints = []
    
    if len(model_data) < 2:
        return hints
    
    # Sentence-ending punctuation (both Western and CJK)
    sentence_endings = {'.', '!', '?', '。', '！', '？', ';', '；', ':', '：'}
    
    for page_idx in range(len(model_data) - 1):
        current_page = model_data[page_idx]
        next_page = model_data[page_idx + 1]
        
        if not current_page or not next_page:
            continue
        
        # Find last/first regions by type on each page
        # Tables: last table on current page, first table on next page
        tables_current = [r for r in current_page if r.get("label") == "table"]
        tables_next = [r for r in next_page if r.get("label") == "table"]
        
        if tables_current and tables_next:
            last_table = tables_current[-1]
            first_table = tables_next[0]
            
            last_id = f"r{last_table['index']}_p{page_idx}"
            first_id = f"r{first_table['index']}_p{page_idx + 1}"
            
            hints.append({
                "type": "table_continuation",
                "from_region": last_id,
                "to_region": first_id,
                "from_page": page_idx,
                "to_page": page_idx + 1,
                "confidence": 0.7  # Moderate confidence - could be separate tables
            })
        
        # Text: last text on current page, first text on next page
        text_labels = {"text", "paragraph_title"}
        text_current = [r for r in current_page if r.get("label") in text_labels]
        text_next = [r for r in next_page if r.get("label") in text_labels]
        
        if text_current and text_next:
            last_text = text_current[-1]
            first_text = text_next[0]
            
            last_content = (last_text.get("content") or "").strip()
            first_content = (first_text.get("content") or "").strip()
            
            if last_content and first_content:
                last_char = last_content[-1] if last_content else ""
                first_char = first_content[0] if first_content else ""
                
                # Continuation signal: no sentence ending + lowercase start
                no_sentence_end = last_char not in sentence_endings
                lowercase_start = first_char.isalpha() and first_char.islower()
                
                if no_sentence_end and lowercase_start:
                    last_id = f"r{last_text['index']}_p{page_idx}"
                    first_id = f"r{first_text['index']}_p{page_idx + 1}"
                    
                    hints.append({
                        "type": "paragraph_continuation",
                        "from_region": last_id,
                        "to_region": first_id,
                        "from_page": page_idx,
                        "to_page": page_idx + 1,
                        "confidence": 0.8  # Higher confidence for text continuation
                    })
    
    return hints


def build_docir(
    model_json_path: Path,
    pdf_path: Path,
    output_path: Path,
    pipeline_info: Optional[Dict[str, str]] = None
) -> None:
    """
    Build DocIR XML from GLM-OCR output.
    
    Args:
        model_json_path: Path to GLM-OCR model JSON file
        pdf_path: Path to source PDF file
        output_path: Path to output DocIR XML file
        pipeline_info: Optional metadata about the pipeline
    """
    if pipeline_info is None:
        pipeline_info = {}
    
    # Default pipeline info
    pipeline_info.setdefault("source", "glm-ocr-pipeline")
    pipeline_info.setdefault("detection_model", "PP-DocLayoutV3")
    pipeline_info.setdefault("ocr_engine", "Ollama glm-ocr:latest")
    pipeline_info.setdefault("style_extractor", "pending")
    
    # Load GLM-OCR model JSON
    with open(model_json_path, 'r', encoding='utf-8') as f:
        model_data = json.load(f)
    
    # Get PDF page dimensions
    page_dimensions = get_pdf_page_dimensions(pdf_path)
    
    # Collect image assets from all regions
    # GLM-OCR convention: imgs/cropped_page{N}_idx{M}.jpg where M is sequential per page
    image_assets = []
    image_counter_per_page = {}  # Track sequential image index per page
    for page_idx, page_regions in enumerate(model_data):
        image_counter_per_page[page_idx] = 0
        for region in page_regions:
            if region.get("label") == "image":
                asset_id = f"img_page{page_idx}_region{region['index']}"
                # Use image_path if available, otherwise generate from convention
                # GLM-OCR uses sequential counter for images, not region index
                if "image_path" in region and region["image_path"]:
                    image_path = region["image_path"]
                else:
                    img_idx = image_counter_per_page[page_idx]
                    image_path = f"imgs/cropped_page{page_idx}_idx{img_idx}.jpg"
                    image_counter_per_page[page_idx] += 1
                
                image_assets.append({
                    "id": asset_id,
                    "path": image_path,
                    "mime_type": "image/jpeg",
                    "region_index": region["index"],
                    "page_index": page_idx
                })
    
    # Build DocIR XML
    document = etree.Element(docir_tag("document"), nsmap=NSMAP)
    document.set("version", "0.1.0")
    document.set("source_pdf", pdf_path.name)
    document.set("generated_at", datetime.now().isoformat())
    document.set("generator", "ir-builder-v0.1.0")
    
    # Metadata
    metadata = etree.SubElement(document, f"{{{DOCIR_NS}}}metadata")
    
    title = etree.SubElement(metadata, f"{{{DOCIR_NS}}}title")
    title.text = pipeline_info.get("title", pdf_path.stem)
    
    page_count = etree.SubElement(metadata, f"{{{DOCIR_NS}}}page_count")
    page_count.text = str(len(page_dimensions))
    
    default_page_size = etree.SubElement(metadata, f"{{{DOCIR_NS}}}default_page_size")
    default_page_size.set("width_pt", f"{page_dimensions[0][0]:.2f}")
    default_page_size.set("height_pt", f"{page_dimensions[0][1]:.2f}")
    
    pipeline_info_elem = etree.SubElement(metadata, f"{{{DOCIR_NS}}}pipeline_info")
    
    layout_detector = etree.SubElement(pipeline_info_elem, f"{{{DOCIR_NS}}}layout_detector")
    layout_detector.text = pipeline_info["detection_model"]
    
    ocr_engine = etree.SubElement(pipeline_info_elem, f"{{{DOCIR_NS}}}ocr_engine")
    ocr_engine.text = pipeline_info["ocr_engine"]
    
    style_extractor = etree.SubElement(pipeline_info_elem, f"{{{DOCIR_NS}}}style_extractor")
    style_extractor.text = pipeline_info["style_extractor"]
    
    # Pages
    pages_elem = etree.SubElement(document, f"{{{DOCIR_NS}}}pages")
    
    for page_idx, page_regions in enumerate(model_data):
        page_width_pt, page_height_pt = page_dimensions[page_idx]
        
        page_elem = etree.SubElement(pages_elem, f"{{{DOCIR_NS}}}page")
        page_elem.set("index", str(page_idx))
        
        page_size = etree.SubElement(page_elem, f"{{{DOCIR_NS}}}page_size")
        page_size.set("width_pt", f"{page_width_pt:.2f}")
        page_size.set("height_pt", f"{page_height_pt:.2f}")
        
        regions_elem = etree.SubElement(page_elem, f"{{{DOCIR_NS}}}regions")
        
        for region in page_regions:
            region_elem = build_region(region, page_width_pt, page_height_pt, pipeline_info, page_idx)
            regions_elem.append(region_elem)
    
    # Assets (register image assets)
    assets_elem = etree.SubElement(document, f"{{{DOCIR_NS}}}assets")
    for asset in image_assets:
        asset_elem = etree.SubElement(assets_elem, f"{{{DOCIR_NS}}}asset")
        asset_elem.set("id", asset["id"])
        asset_elem.set("mime_type", asset["mime_type"])
        
        # Extract image dimensions if file exists
        # Try relative to model JSON directory first, then relative to cwd
        image_path = Path(asset["path"])
        if not image_path.is_absolute():
            # Try relative to model JSON directory
            model_dir = model_json_path.parent
            relative_to_model = model_dir / image_path
            if relative_to_model.exists():
                image_path = relative_to_model
        
        if image_path.exists():
            try:
                with Image.open(image_path) as img:
                    width_px, height_px = img.size
                    asset_elem.set("width_px", str(width_px))
                    asset_elem.set("height_px", str(height_px))
            except Exception as e:
                # If we can't read dimensions, just skip those attributes
                pass
        
        file_path = etree.SubElement(asset_elem, f"{{{DOCIR_NS}}}file_path")
        file_path.text = asset["path"]
        
        extraction_source = etree.SubElement(asset_elem, f"{{{DOCIR_NS}}}extraction_source")
        extraction_source.text = f"pdf_page_{asset['page_index']}_region_{asset['region_index']}"
    
    # Cross-page merge hints
    cross_page_hints = etree.SubElement(document, f"{{{DOCIR_NS}}}cross_page_hints")
    merge_hints = detect_cross_page_continuations(model_data, pages_elem, DOCIR_NS)
    for hint in merge_hints:
        hint_elem = etree.SubElement(cross_page_hints, f"{{{DOCIR_NS}}}merge_hint")
        hint_elem.set("type", hint["type"])
        hint_elem.set("from_region", hint["from_region"])
        hint_elem.set("to_region", hint["to_region"])
        hint_elem.set("from_page", str(hint["from_page"]))
        hint_elem.set("to_page", str(hint["to_page"]))
        hint_elem.set("confidence", f"{hint['confidence']:.2f}")
    
    # Write XML
    tree = etree.ElementTree(document)
    tree.write(
        str(output_path),
        encoding="UTF-8",
        xml_declaration=True,
        pretty_print=True
    )
    
    print(f"✓ DocIR XML written to: {output_path}")
    print(f"  Pages: {len(page_dimensions)}")
    print(f"  Regions: {sum(len(page) for page in model_data)}")
    print(f"  Assets: {len(image_assets)}")


def main():
    """CLI entry point."""
    import argparse
    
    parser = argparse.ArgumentParser(
        description="Build DocIR XML from GLM-OCR output"
    )
    parser.add_argument(
        "model_json",
        type=Path,
        help="Path to GLM-OCR model JSON file"
    )
    parser.add_argument(
        "pdf",
        type=Path,
        help="Path to source PDF file"
    )
    parser.add_argument(
        "-o", "--output",
        type=Path,
        help="Output DocIR XML path (default: <pdf-stem>.docir.xml)"
    )
    parser.add_argument(
        "--title",
        type=str,
        help="Document title (default: PDF filename)"
    )
    
    args = parser.parse_args()
    
    # Default output path
    if args.output is None:
        args.output = args.pdf.with_suffix(".docir.xml")
    
    # Pipeline info
    pipeline_info = {}
    if args.title:
        pipeline_info["title"] = args.title
    
    # Build DocIR
    build_docir(args.model_json, args.pdf, args.output, pipeline_info)


if __name__ == "__main__":
    main()
