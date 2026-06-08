#!/usr/bin/env python3
"""
Debug script to crop and save region images for visual inspection.
"""

import sys
from pathlib import Path

# Add parent directory to path to import style_extractor module
script_dir = Path(__file__).parent
sys.path.insert(0, str(script_dir))

from style_extractor import crop_region_from_pdf
from lxml import etree

DOCIR_NS = "urn:docir:v0.1"

def debug_crops(docir_path: Path, pdf_path: Path, output_dir: Path):
    """Crop all text regions and save as images for inspection."""
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # Load DocIR XML
    tree = etree.parse(str(docir_path))
    root = tree.getroot()
    
    # Get page dimensions
    page_size = root.find(f".//{{{DOCIR_NS}}}default_page_size")
    if page_size is not None:
        page_w = float(page_size.get("width_pt", 0))
        page_h = float(page_size.get("height_pt", 0))
        print(f"Page size: {page_w:.2f} x {page_h:.2f} pt")
    
    # Process each region
    for page in root.findall(f".//{{{DOCIR_NS}}}page"):
        page_idx = int(page.get("index", 0))
        
        for region in page.findall(f".//{{{DOCIR_NS}}}region"):
            region_type = region.get("type")
            region_id = region.get("id", "?")
            
            if region_type != "text":
                continue
            
            # Get bbox
            bbox_elem = region.find(f"{{{DOCIR_NS}}}bbox")
            if bbox_elem is None:
                continue
            
            x = float(bbox_elem.get("x", 0))
            y = float(bbox_elem.get("y", 0))
            w = float(bbox_elem.get("width", 0))
            h = float(bbox_elem.get("height", 0))
            
            print(f"\n{region_id}: bbox=({x:.2f}, {y:.2f}, {w:.2f}, {h:.2f})")
            print(f"  Type: {region_type}")
            
            # Get text content
            text_content = region.find(f".//{{{DOCIR_NS}}}run")
            if text_content is not None and text_content.text:
                print(f"  Text: {text_content.text[:50]}...")
            
            # Crop and save
            try:
                img_bytes = crop_region_from_pdf(pdf_path, page_idx, (x, y, w, h), dpi=200)
                output_path = output_dir / f"{region_id}.jpg"
                with open(output_path, "wb") as f:
                    f.write(img_bytes)
                print(f"  ✓ Saved: {output_path}")
            except Exception as e:
                print(f"  ✗ Error: {e}")

if __name__ == "__main__":
    if len(sys.argv) < 4:
        print("Usage: debug_crops.py <docir.xml> <pdf> <output_dir>")
        sys.exit(1)
    
    docir_path = Path(sys.argv[1])
    pdf_path = Path(sys.argv[2])
    output_dir = Path(sys.argv[3])
    
    debug_crops(docir_path, pdf_path, output_dir)
