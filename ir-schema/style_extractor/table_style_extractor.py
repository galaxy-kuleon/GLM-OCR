#!/usr/bin/env python3
"""
DocIR Table Style Extractor

Extracts table styles (borders, colors, headers) from cropped table images using VLM.
Updates DocIR XML with table style information.

Usage:
    python table_style_extractor.py input.docir.xml source.pdf -o output.docir.xml
"""

import base64
import json
import sys
import time
from pathlib import Path
from typing import Dict, Optional
from dataclasses import dataclass

import pymupdf
import requests
from lxml import etree


# DocIR namespace
DOCIR_NS = "urn:docir:v0.1"
NSMAP = {"docir": DOCIR_NS}


@dataclass
class TableStyle:
    """Extracted table style information."""
    border_visible: Optional[bool] = None
    border_color: Optional[str] = None  # hex like "#000000"
    header_row: Optional[bool] = None
    cell_background_color: Optional[str] = None  # hex like "#FFFFFF"
    header_background_color: Optional[str] = None  # hex like "#CCCCCC"
    
    # Evidence
    evidence_confidence: Optional[float] = None
    evidence_notes: Optional[str] = None
    
    # Metadata
    vlm_model: str = "qwen3.6-35b-a3b-q7"
    extraction_time_ms: Optional[int] = None


@dataclass
class VLMConfig:
    """VLM API configuration."""
    api_base: str = "http://localhost:11234/v1"
    api_key: str = "change-me-local-key"
    model: str = "qwen3.6-35b-a3b-q7"
    max_tokens: int = 4096
    temperature: float = 0.1
    timeout: int = 120


# VLM Table Style Extraction Prompt
TABLE_STYLE_EXTRACTION_PROMPT = """You are a commercial-grade table style analyzer. Analyze this cropped table region from a document and extract precise style information.

IMPORTANT: This is a table region crop. Focus on the table's visual styling.

Respond in JSON only (no markdown fences, no explanation):
{
  "border_visible": <true|false - are borders visible around cells?>,
  "border_color": "<hex color like #000000 or #CCCCCC, or null if no borders>",
  "header_row": <true|false - is there a distinct header row with different styling?>,
  "cell_background_color": "<hex color for regular cells like #FFFFFF, or null if transparent/white>",
  "header_background_color": "<hex color for header row like #CCCCCC or #E0E0E0, or null if same as cells>",
  "evidence": {
    "confidence": <0.0-1.0 confidence in style detection>,
    "notes": "any observations about the table styling"
  }
}

Guidelines:
- border_visible: true if you can see clear borders around cells, false if no borders or very faint
- border_color: Most common border color. Use #000000 for black, #CCCCCC for light gray, etc.
- header_row: true if the first row has distinct styling (bold, different background, etc.)
- cell_background_color: Background color of regular data cells. Use #FFFFFF for white
- header_background_color: Background color of header row if different from cells
- confidence scoring:
  * 0.9-1.0: Very clear table structure, confident in all attributes
  * 0.7-0.9: Clear table but some uncertainty (e.g., border color)
  * 0.5-0.7: Somewhat unclear table structure
  * <0.5: Very unclear, low confidence
- If no table is visible, set all fields to null and confidence to 0.0

Respond ONLY with valid JSON."""


def crop_table_region_from_pdf(
    pdf_path: Path,
    page_index: int,
    bbox_pt: tuple,
    dpi: int = 200
) -> bytes:
    """
    Crop a table region from a PDF page and return as JPEG bytes.
    
    Args:
        pdf_path: Path to source PDF
        page_index: Page number (0-indexed)
        bbox_pt: Bounding box in PDF points (x, y, width, height)
        dpi: Resolution for rendering
    
    Returns:
        JPEG image bytes
    """
    doc = pymupdf.open(str(pdf_path))
    page = doc[page_index]
    page_height = page.rect.height
    
    x, y, w, h = bbox_pt
    
    # Convert from DocIR coords (origin at bottom-left) to pymupdf coords (origin at top-left)
    pdf_y_bottom = y - h
    pdf_y_top = y
    
    pymupdf_y_top = page_height - pdf_y_top
    pymupdf_y_bottom = page_height - pdf_y_bottom
    
    clip = pymupdf.Rect(x, pymupdf_y_top, x + w, pymupdf_y_bottom)
    
    zoom = dpi / 72.0
    mat = pymupdf.Matrix(zoom, zoom)
    
    pix = page.get_pixmap(matrix=mat, clip=clip)
    img_data = pix.tobytes("jpeg")
    doc.close()
    
    return img_data


def extract_table_style_from_image(
    image_bytes: bytes,
    vlm_config: VLMConfig
) -> Optional[TableStyle]:
    """
    Send cropped table image to VLM and extract style information.
    
    Args:
        image_bytes: JPEG image bytes
        vlm_config: VLM API configuration
    
    Returns:
        TableStyle with extracted information, or None on failure
    """
    img_b64 = base64.b64encode(image_bytes).decode()
    
    payload = {
        "model": vlm_config.model,
        "messages": [
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": TABLE_STYLE_EXTRACTION_PROMPT},
                    {
                        "type": "image_url",
                        "image_url": {"url": f"data:image/jpeg;base64,{img_b64}"}
                    }
                ]
            }
        ],
        "max_tokens": vlm_config.max_tokens,
        "temperature": vlm_config.temperature
    }
    
    headers = {
        "Authorization": f"Bearer {vlm_config.api_key}",
        "Content-Type": "application/json"
    }
    
    start_time = time.time()
    
    try:
        resp = requests.post(
            f"{vlm_config.api_base}/chat/completions",
            json=payload,
            headers=headers,
            timeout=vlm_config.timeout
        )
        
        elapsed_ms = int((time.time() - start_time) * 1000)
        
        if resp.status_code != 200:
            print(f"  ✗ VLM API error: HTTP {resp.status_code}")
            return None
        
        data = resp.json()
        content = data["choices"][0]["message"].get("content", "")
        
        if not content:
            print(f"  ✗ Empty VLM response")
            return None
        
        # Parse JSON response
        clean = content.strip()
        if clean.startswith("```"):
            lines = clean.split("\n")
            clean = "\n".join(lines[1:-1]) if lines[-1].strip() == "```" else "\n".join(lines[1:])
        
        try:
            parsed = json.loads(clean)
        except json.JSONDecodeError as e:
            print(f"  ✗ JSON parse error: {e}")
            print(f"    Content: {content[:200]}")
            return None
        
        # Extract table style information
        style = TableStyle()
        style.vlm_model = vlm_config.model
        style.extraction_time_ms = elapsed_ms
        
        style.border_visible = parsed.get("border_visible")
        style.border_color = parsed.get("border_color")
        style.header_row = parsed.get("header_row")
        style.cell_background_color = parsed.get("cell_background_color")
        style.header_background_color = parsed.get("header_background_color")
        
        # Evidence
        evidence = parsed.get("evidence", {})
        style.evidence_confidence = evidence.get("confidence")
        style.evidence_notes = evidence.get("notes")
        
        return style
        
    except requests.exceptions.Timeout:
        print(f"  ✗ VLM request timeout ({vlm_config.timeout}s)")
        return None
    except Exception as e:
        print(f"  ✗ VLM request error: {e}")
        return None


def extract_table_styles_from_docir(
    docir_path: Path,
    pdf_path: Path,
    output_path: Optional[Path] = None,
    vlm_config: Optional[VLMConfig] = None,
    dpi: int = 200
) -> Path:
    """
    Extract table styles for all table regions in a DocIR document.
    
    Args:
        docir_path: Path to input DocIR XML
        pdf_path: Path to source PDF
        output_path: Path to output DocIR XML (default: overwrite input)
        vlm_config: VLM API configuration
        dpi: Resolution for cropping
    
    Returns:
        Path to output DocIR XML
    """
    if vlm_config is None:
        vlm_config = VLMConfig()
    
    if output_path is None:
        output_path = docir_path
    
    # Load DocIR XML
    tree = etree.parse(str(docir_path))
    root = tree.getroot()
    
    # Find all table regions
    tables_to_process = []
    for page in root.findall(f".//{{{DOCIR_NS}}}page"):
        page_idx = int(page.get("index", 0))
        
        for region in page.findall(f".//{{{DOCIR_NS}}}region"):
            region_type = region.get("type")
            region_id = region.get("id", "?")
            
            if region_type != "table":
                continue
            
            # Get bbox
            bbox_elem = region.find(f"{{{DOCIR_NS}}}bbox")
            if bbox_elem is None:
                continue
            
            x = float(bbox_elem.get("x", 0))
            y = float(bbox_elem.get("y", 0))
            w = float(bbox_elem.get("width", 0))
            h = float(bbox_elem.get("height", 0))
            
            # Skip very small tables
            if h < 20 or w < 50:
                continue
            
            tables_to_process.append({
                "element": region,
                "page_idx": page_idx,
                "region_id": region_id,
                "bbox_pt": (x, y, w, h)
            })
    
    print(f"Found {len(tables_to_process)} table regions to process")
    print(f"Using VLM: {vlm_config.model} @ {vlm_config.api_base}")
    print()
    
    # Process each table
    for i, info in enumerate(tables_to_process):
        region_elem = info["element"]
        page_idx = info["page_idx"]
        region_id = info["region_id"]
        bbox_pt = info["bbox_pt"]
        
        print(f"[{i+1}/{len(tables_to_process)}] Table {region_id} (page {page_idx})...")
        
        # Crop table from PDF
        try:
            img_bytes = crop_table_region_from_pdf(pdf_path, page_idx, bbox_pt, dpi=dpi)
        except Exception as e:
            print(f"  ✗ Crop error: {e}")
            continue
        
        # Extract table style via VLM
        style = extract_table_style_from_image(img_bytes, vlm_config)
        
        if style is None:
            print(f"  ✗ Table style extraction failed")
            continue
        
        # Update DocIR XML
        table_content = region_elem.find(f"{{{DOCIR_NS}}}table_content")
        if table_content is None:
            print(f"  ⚠ No table_content element")
            continue
        
        # Find or create table_style element
        table_style_elem = table_content.find(f"{{{DOCIR_NS}}}table_style")
        if table_style_elem is None:
            table_style_elem = etree.SubElement(table_content, f"{{{DOCIR_NS}}}table_style")
        
        # Set style attributes
        if style.border_visible is not None:
            table_style_elem.set("border_visible", str(style.border_visible).lower())
        if style.border_color:
            table_style_elem.set("border_color", style.border_color)
        if style.header_row is not None:
            table_style_elem.set("header_row", str(style.header_row).lower())
        
        # Add evidence as notes
        if style.evidence_notes:
            notes_elem = table_style_elem.find(f"{{{DOCIR_NS}}}notes")
            if notes_elem is None:
                notes_elem = etree.SubElement(table_style_elem, f"{{{DOCIR_NS}}}notes")
            notes_elem.text = style.evidence_notes
        
        # Log success
        style_info = []
        if style.border_visible:
            style_info.append(f"borders={style.border_color or 'visible'}")
        if style.header_row:
            style_info.append("header=yes")
        
        conf = f"conf={style.evidence_confidence:.2f}" if style.evidence_confidence else ""
        print(f"  ✓ {', '.join(style_info) if style_info else 'no distinct style'} {conf} ({style.extraction_time_ms}ms)")
    
    # Save updated XML
    tree.write(
        str(output_path),
        encoding="UTF-8",
        xml_declaration=True,
        pretty_print=True
    )
    
    print(f"\n✓ Updated DocIR saved to: {output_path}")
    return output_path


def main():
    """CLI entry point."""
    import argparse
    
    parser = argparse.ArgumentParser(
        description="Extract table styles from DocIR regions using VLM"
    )
    parser.add_argument(
        "docir_xml",
        type=Path,
        help="Path to DocIR XML file"
    )
    parser.add_argument(
        "pdf",
        type=Path,
        help="Path to source PDF file"
    )
    parser.add_argument(
        "-o", "--output",
        type=Path,
        help="Output DocIR XML path (default: overwrite input)"
    )
    parser.add_argument(
        "--model",
        type=str,
        default="qwen3.6-35b-a3b-q7",
        help="VLM model name (default: qwen3.6-35b-a3b-q7)"
    )
    parser.add_argument(
        "--api-base",
        type=str,
        default="http://localhost:11234/v1",
        help="VLM API base URL"
    )
    parser.add_argument(
        "--api-key",
        type=str,
        default="change-me-local-key",
        help="VLM API key"
    )
    parser.add_argument(
        "--dpi",
        type=int,
        default=200,
        help="DPI for region cropping (default: 200)"
    )
    
    args = parser.parse_args()
    
    vlm_config = VLMConfig(
        api_base=args.api_base,
        api_key=args.api_key,
        model=args.model
    )
    
    extract_table_styles_from_docir(
        docir_path=args.docir_xml,
        pdf_path=args.pdf,
        output_path=args.output,
        vlm_config=vlm_config,
        dpi=args.dpi
    )


if __name__ == "__main__":
    main()
