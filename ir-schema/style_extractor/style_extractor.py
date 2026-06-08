#!/usr/bin/env python3
"""
DocIR Style Extractor v0.1.0

Extracts typography styles from cropped region images using VLM.
Updates DocIR XML with computed styles and visual evidence.

Pipeline:
  DocIR XML + Source PDF → Crop text regions → VLM analysis → Updated DocIR XML

Model: qwen3.6-35b-a3b-q7 @ localhost:11234
"""

import base64
import json
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from io import BytesIO
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from dataclasses import dataclass

import pymupdf
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# Shared HTTP session with connection pooling for parallel VLM requests
_http_session = None

def get_http_session() -> requests.Session:
    """Get or create a shared HTTP session with connection pooling."""
    global _http_session
    if _http_session is None:
        _http_session = requests.Session()
        # Configure retry strategy with backoff
        retry_strategy = Retry(
            total=3,
            backoff_factor=0.5,
            status_forcelist=[429, 500, 502, 503, 504],
        )
        adapter = HTTPAdapter(
            max_retries=retry_strategy,
            pool_connections=10,
            pool_maxsize=10,
        )
        _http_session.mount("http://", adapter)
        _http_session.mount("https://", adapter)
    return _http_session
from lxml import etree


# DocIR namespace
DOCIR_NS = "urn:docir:v0.1"
NSMAP = {"docir": DOCIR_NS}


@dataclass
class RegionStyle:
    """Extracted style information for a region."""
    font_name: Optional[str] = None
    font_size_pt: Optional[float] = None
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    underline: Optional[bool] = None
    color: Optional[str] = None  # hex like "#FF0000"
    
    # Evidence
    evidence_pixel_height: Optional[float] = None
    evidence_color_sample: Optional[str] = None
    evidence_confidence: Optional[float] = None
    
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


# VLM Style Extraction Prompt
STYLE_EXTRACTION_PROMPT = """You are a commercial-grade typography analyzer. Analyze this cropped region from a document and extract precise style information.

IMPORTANT: This is a single text region crop. Focus on the ACTUAL visible text styling.

Respond in JSON only (no markdown fences, no explanation):
{
  "font_name": "detected font family name or null if uncertain",
  "font_size_pt": <estimated font size in PDF points>,
  "bold": <true|false|null>,
  "italic": <true|false|null>,
  "underline": <true|false|null>,
  "color": "<hex color like #000000 or #FF0000, or null if standard black>",
  "text_alignment": "left|center|right|justify",
  "evidence": {
    "pixel_height": <measured character height in pixels>,
    "color_sample": "<sampled dominant text color hex>",
    "confidence": <0.0-1.0 confidence in style detection>,
    "notes": "any observations about the styling"
  }
}

Guidelines:
- font_size_pt: Estimate based on character height relative to typical document scaling
  - Typical body text: 10-12pt
  - Headings: 14-24pt
  - Titles: 24-36pt
- color: Use #000000 for black text, only specify color if clearly non-black
- confidence: Rate your certainty about the style detection
  - 0.9-1.0: Very clear, high confidence
  - 0.7-0.9: Reasonably clear
  - 0.5-0.7: Uncertain, blurry or small text
  - <0.5: Very uncertain, consider skipping
- If text is too small or blurry, set confidence lower and note uncertainty
- font_name: Common fonts include Arial, Times New Roman, Calibri, Helvetica, Georgia
- If you cannot determine a style attribute with reasonable confidence, set it to null

Respond ONLY with valid JSON."""


# Table Style Extraction Prompt
TABLE_STYLE_EXTRACTION_PROMPT = """You are a table structure analyzer. Analyze this cropped table region from a document and extract table styling information.

Respond in JSON only (no markdown fences, no explanation):
{
  "border_visible": <true|false>,
  "border_color": "<hex color like #000000 or #CCCCCC, or null if no borders>",
  "header_row": <true|false - is the first row visually distinct as a header?>,
  "header_style": {
    "bold": <true|false|null>,
    "background_color": "<hex color or null>"
  },
  "cell_background_colors": ["<hex color of first non-header cell or null>", "..."],
  "confidence": <0.0-1.0 confidence in detection>,
  "notes": "any observations about table styling"
}

Guidelines:
- border_visible: true if the table has visible grid lines/borders
- border_color: the color of the border lines (usually #000000 for black, #CCCCCC for gray)
- header_row: true if the first row appears to be a header (bold text, different background, etc.)
- header_style: styling of the header row cells
- cell_background_colors: list of background colors for body cells (usually null/white)
- confidence: rate your certainty (0.9+ for clear tables, lower for ambiguous)

Respond ONLY with valid JSON."""


@dataclass
class TableStyle:
    """Extracted table style information."""
    border_visible: bool = True
    border_color: Optional[str] = "#000000"
    header_row: bool = False
    header_bold: Optional[bool] = None
    header_background_color: Optional[str] = None
    confidence: Optional[float] = None
    notes: Optional[str] = None


def extract_table_style_from_image(
    image_bytes: bytes,
    vlm_config: VLMConfig,
    max_retries: int = 2
) -> Optional[TableStyle]:
    """
    Send cropped table image to VLM and extract table style information.
    
    Args:
        image_bytes: JPEG image bytes of the table
        vlm_config: VLM API configuration
        max_retries: Maximum number of retry attempts
    
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
    
    for attempt in range(max_retries + 1):
        start_time = time.time()
        
        try:
            resp = get_http_session().post(
                f"{vlm_config.api_base}/chat/completions",
                json=payload,
                headers=headers,
                timeout=vlm_config.timeout
            )
            
            elapsed_ms = int((time.time() - start_time) * 1000)
            
            if resp.status_code != 200:
                if attempt < max_retries:
                    print(f"  ⚠ VLM API error: HTTP {resp.status_code}, retrying ({attempt+1}/{max_retries})...")
                    time.sleep(1 * (attempt + 1))
                    continue
                print(f"  ✗ VLM API error: HTTP {resp.status_code} (after {max_retries} retries)")
                return None
            
            data = resp.json()
            content = data["choices"][0]["message"].get("content", "")
            
            if not content:
                if attempt < max_retries:
                    print(f"  ⚠ Empty VLM response, retrying...")
                    time.sleep(1 * (attempt + 1))
                    continue
                return None
            
            # Parse JSON response
            clean = content.strip()
            if clean.startswith("```"):
                lines = clean.split("\n")
                clean = "\n".join(lines[1:-1]) if lines[-1].strip() == "```" else "\n".join(lines[1:])
            
            try:
                parsed = json.loads(clean)
            except json.JSONDecodeError as e:
                if attempt < max_retries:
                    print(f"  ⚠ JSON parse error, retrying...")
                    time.sleep(1 * (attempt + 1))
                    continue
                print(f"  ✗ JSON parse error: {e}")
                return None
            
            # Extract table style
            style = TableStyle()
            style.border_visible = parsed.get("border_visible", True)
            style.border_color = parsed.get("border_color", "#000000")
            style.header_row = parsed.get("header_row", False)
            style.confidence = parsed.get("confidence")
            style.notes = parsed.get("notes")
            
            header_style = parsed.get("header_style", {})
            style.header_bold = header_style.get("bold")
            style.header_background_color = header_style.get("background_color")
            
            return style
            
        except requests.exceptions.Timeout:
            if attempt < max_retries:
                print(f"  ⚠ Timeout, retrying...")
                time.sleep(1 * (attempt + 1))
                continue
            print(f"  ✗ VLM request timeout")
            return None
        except Exception as e:
            if attempt < max_retries:
                print(f"  ⚠ Error: {e}, retrying...")
                time.sleep(1 * (attempt + 1))
                continue
            print(f"  ✗ VLM request error: {e}")
            return None
    
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
    table_regions = []
    for page in root.findall(f".//{{{DOCIR_NS}}}page"):
        page_idx = int(page.get("index", 0))
        
        for region in page.findall(f".//{{{DOCIR_NS}}}region"):
            if region.get("type") != "table":
                continue
            
            region_id = region.get("id", "?")
            bbox_elem = region.find(f"{{{DOCIR_NS}}}bbox")
            if bbox_elem is None:
                continue
            
            x = float(bbox_elem.get("x", 0))
            y = float(bbox_elem.get("y", 0))
            w = float(bbox_elem.get("width", 0))
            h = float(bbox_elem.get("height", 0))
            
            table_regions.append({
                "element": region,
                "page_idx": page_idx,
                "region_id": region_id,
                "bbox_pt": (x, y, w, h)
            })
    
    print(f"Found {len(table_regions)} table regions to process")
    print(f"Using VLM: {vlm_config.model} @ {vlm_config.api_base}")
    print()
    
    # Process each table region
    for i, info in enumerate(table_regions):
        region_elem = info["element"]
        page_idx = info["page_idx"]
        region_id = info["region_id"]
        bbox_pt = info["bbox_pt"]
        
        print(f"[{i+1}/{len(table_regions)}] Table {region_id} (page {page_idx})...")
        
        # Crop table region from PDF
        try:
            img_bytes = crop_region_from_pdf(pdf_path, page_idx, bbox_pt, dpi=dpi)
        except Exception as e:
            print(f"  ✗ Crop error: {e}")
            continue
        
        # Extract table style via VLM
        style = extract_table_style_from_image(img_bytes, vlm_config)
        
        if style is None:
            print(f"  ✗ Table style extraction failed")
            continue
        
        # Update DocIR XML - find or create table_style element
        table_content = region_elem.find(f"{{{DOCIR_NS}}}table_content")
        if table_content is None:
            print(f"  ⚠ No table_content element")
            continue
        
        table_style_elem = table_content.find(f"{{{DOCIR_NS}}}table_style")
        if table_style_elem is None:
            table_style_elem = etree.SubElement(table_content, f"{{{DOCIR_NS}}}table_style")
        
        # Update table style attributes
        table_style_elem.set("border_visible", str(style.border_visible).lower())
        if style.border_color:
            table_style_elem.set("border_color", style.border_color)
        if style.header_row:
            table_style_elem.set("header_row", "true")
        
        # Log success
        style_info = []
        if style.border_visible:
            style_info.append(f"borders={style.border_color}")
        else:
            style_info.append("no borders")
        if style.header_row:
            style_info.append("has header")
        if style.header_bold:
            style_info.append("header bold")
        
        conf = f"conf={style.confidence:.2f}" if style.confidence else ""
        print(f"  ✓ {', '.join(style_info)} {conf}")
    
    # Save updated XML
    tree.write(
        str(output_path),
        encoding="UTF-8",
        xml_declaration=True,
        pretty_print=True
    )
    
    print(f"\n✓ Updated DocIR saved to: {output_path}")
    return output_path


def crop_region_from_pdf(
    pdf_path: Path,
    page_index: int,
    bbox_pt: Tuple[float, float, float, float],
    dpi: int = 200
) -> bytes:
    """
    Crop a region from a PDF page and return as JPEG bytes.
    
    Args:
        pdf_path: Path to source PDF
        page_index: Page number (0-indexed)
        bbox_pt: Bounding box in PDF points (x, y, width, height)
                 Note: DocIR uses PDF coordinates (origin at bottom-left, y increases upward)
        dpi: Resolution for rendering
    
    Returns:
        JPEG image bytes
    """
    doc = pymupdf.open(str(pdf_path))
    page = doc[page_index]
    page_height = page.rect.height
    
    x, y, w, h = bbox_pt
    
    # In DocIR/PDF coordinates:
    # - y is distance from BOTTOM of page
    # - y represents the TOP edge of the region
    # - The region extends from (y - h) at the bottom to y at the top
    
    pdf_y_bottom = y - h  # Bottom edge (lower y value, from bottom)
    pdf_y_top = y          # Top edge (higher y value, from bottom)
    
    # Convert to pymupdf coordinates (origin at top-left, y increases downward)
    pymupdf_y_top = page_height - pdf_y_top
    pymupdf_y_bottom = page_height - pdf_y_bottom
    
    # Create clip rectangle (x0, y0, x1, y1) where y0 < y1 in pymupdf coords
    clip = pymupdf.Rect(x, pymupdf_y_top, x + w, pymupdf_y_bottom)
    
    # Render at specified DPI
    zoom = dpi / 72.0
    mat = pymupdf.Matrix(zoom, zoom)
    
    # Render the clipped region
    pix = page.get_pixmap(matrix=mat, clip=clip)
    
    # Convert to JPEG
    img_data = pix.tobytes("jpeg")
    doc.close()
    
    return img_data


def extract_style_from_image(
    image_bytes: bytes,
    vlm_config: VLMConfig,
    dpi: int = 200,
    max_retries: int = 2
) -> Optional[RegionStyle]:
    """
    Send cropped image to VLM and extract style information.
    
    Args:
        image_bytes: JPEG image bytes
        vlm_config: VLM API configuration
        dpi: DPI used for cropping (for font size calibration)
        max_retries: Maximum number of retry attempts on failure
    
    Returns:
        RegionStyle with extracted information, or None on failure
    """
    # Encode image to base64
    img_b64 = base64.b64encode(image_bytes).decode()
    
    # Build API request
    payload = {
        "model": vlm_config.model,
        "messages": [
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": STYLE_EXTRACTION_PROMPT},
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
    
    for attempt in range(max_retries + 1):
        start_time = time.time()
        
        try:
            resp = get_http_session().post(
                f"{vlm_config.api_base}/chat/completions",
                json=payload,
                headers=headers,
                timeout=vlm_config.timeout
            )
            
            elapsed_ms = int((time.time() - start_time) * 1000)
            
            if resp.status_code != 200:
                if attempt < max_retries:
                    print(f"  ⚠ VLM API error: HTTP {resp.status_code}, retrying ({attempt+1}/{max_retries})...")
                    time.sleep(1 * (attempt + 1))  # Exponential backoff
                    continue
                print(f"  ✗ VLM API error: HTTP {resp.status_code} (after {max_retries} retries)")
                return None
            
            data = resp.json()
            content = data["choices"][0]["message"].get("content", "")
            
            if not content:
                if attempt < max_retries:
                    print(f"  ⚠ Empty VLM response, retrying ({attempt+1}/{max_retries})...")
                    time.sleep(1 * (attempt + 1))
                    continue
                print(f"  ✗ Empty VLM response (after {max_retries} retries)")
                return None
            
            # Parse JSON response
            clean = content.strip()
            if clean.startswith("```"):
                lines = clean.split("\n")
                clean = "\n".join(lines[1:-1]) if lines[-1].strip() == "```" else "\n".join(lines[1:])
            
            try:
                parsed = json.loads(clean)
            except json.JSONDecodeError as e:
                if attempt < max_retries:
                    print(f"  ⚠ JSON parse error: {e}, retrying ({attempt+1}/{max_retries})...")
                    time.sleep(1 * (attempt + 1))
                    continue
                print(f"  ✗ JSON parse error: {e}")
                print(f"    Content: {content[:200]}")
                return None
            
            # Extract style information
            style = RegionStyle()
            style.vlm_model = vlm_config.model
            style.extraction_time_ms = elapsed_ms
            
            style.font_name = parsed.get("font_name")
            style.font_size_pt = parsed.get("font_size_pt")
            style.bold = parsed.get("bold")
            style.italic = parsed.get("italic")
            style.underline = parsed.get("underline")
            style.color = parsed.get("color")
            
            # Evidence
            evidence = parsed.get("evidence", {})
            style.evidence_pixel_height = evidence.get("pixel_height")
            style.evidence_color_sample = evidence.get("color_sample")
            style.evidence_confidence = evidence.get("confidence")
            
            # Calibrate font size using pixel height and DPI if available
            if style.evidence_pixel_height and dpi:
                # Convert pixel height to points using DPI
                # 1 inch = 72 points, 1 inch = DPI pixels
                # So: points = pixels * 72 / DPI
                calibrated_size = style.evidence_pixel_height * 72.0 / dpi
                
                # Use calibrated size if VLM didn't provide one, or if calibration is more reliable
                if not style.font_size_pt:
                    style.font_size_pt = calibrated_size
                else:
                    # Average VLM estimate and calibration for better accuracy
                    # Weight calibration more if VLM confidence is low
                    vlm_confidence = style.evidence_confidence or 0.5
                    if vlm_confidence < 0.7:
                        # Low confidence, trust calibration more
                        style.font_size_pt = calibrated_size * 0.7 + style.font_size_pt * 0.3
                    else:
                        # High confidence, trust VLM more
                        style.font_size_pt = style.font_size_pt * 0.7 + calibrated_size * 0.3
            
            return style
            
        except requests.exceptions.Timeout:
            if attempt < max_retries:
                print(f"  ⚠ VLM request timeout ({vlm_config.timeout}s), retrying ({attempt+1}/{max_retries})...")
                time.sleep(1 * (attempt + 1))
                continue
            print(f"  ✗ VLM request timeout ({vlm_config.timeout}s) after {max_retries} retries")
            return None
        except Exception as e:
            if attempt < max_retries:
                print(f"  ⚠ VLM request error: {e}, retrying ({attempt+1}/{max_retries})...")
                time.sleep(1 * (attempt + 1))
                continue
            print(f"  ✗ VLM request error: {e} (after {max_retries} retries)")
            return None
    
    return None


def extract_styles_from_docir(
    docir_path: Path,
    pdf_path: Path,
    output_path: Optional[Path] = None,
    vlm_config: Optional[VLMConfig] = None,
    dpi: int = 200,
    region_types: List[str] = None,
    parallel: int = 1
) -> Path:
    """
    Extract styles for all text regions in a DocIR document.
    
    Args:
        docir_path: Path to input DocIR XML
        pdf_path: Path to source PDF
        output_path: Path to output DocIR XML (default: overwrite input)
        vlm_config: VLM API configuration
        dpi: Resolution for cropping
        region_types: Which region types to process (default: ["text"])
    
    Returns:
        Path to output DocIR XML
    """
    if vlm_config is None:
        vlm_config = VLMConfig()
    
    if output_path is None:
        output_path = docir_path
    
    if region_types is None:
        region_types = ["text"]
    
    # Load DocIR XML
    tree = etree.parse(str(docir_path))
    root = tree.getroot()
    
    # Find all regions to process
    regions_to_process = []
    skipped_regions = []
    for page in root.findall(f".//{{{DOCIR_NS}}}page"):
        page_idx = int(page.get("index", 0))
        
        for region in page.findall(f".//{{{DOCIR_NS}}}region"):
            region_type = region.get("type")
            region_id = region.get("id", "?")
            
            if region_type not in region_types:
                continue
            
            # Get bbox
            bbox_elem = region.find(f"{{{DOCIR_NS}}}bbox")
            if bbox_elem is None:
                continue
            
            x = float(bbox_elem.get("x", 0))
            y = float(bbox_elem.get("y", 0))
            w = float(bbox_elem.get("width", 0))
            h = float(bbox_elem.get("height", 0))
            
            # Skip very small regions (likely unreadable)
            if h < 10 or w < 20:
                skipped_regions.append((region_id, f"too small ({w:.1f}x{h:.1f}pt)"))
                continue
            
            # Skip regions with no text content (empty runs)
            if region_type == "text":
                text_content = region.find(f".//{{{DOCIR_NS}}}text_content")
                if text_content is not None:
                    has_text = False
                    for run in text_content.findall(f".//{{{DOCIR_NS}}}run"):
                        if run.text and run.text.strip():
                            has_text = True
                            break
                    if not has_text:
                        skipped_regions.append((region_id, "empty text content"))
                        continue
            
            regions_to_process.append({
                "element": region,
                "page_idx": page_idx,
                "region_id": region_id,
                "bbox_pt": (x, y, w, h)
            })
    
    print(f"Found {len(regions_to_process)} regions to process")
    if skipped_regions:
        print(f"Skipped {len(skipped_regions)} regions (too small)")
        for region_id, reason in skipped_regions[:5]:  # Show first 5
            print(f"  - {region_id}: {reason}")
        if len(skipped_regions) > 5:
            print(f"  ... and {len(skipped_regions) - 5} more")
    print(f"Using VLM: {vlm_config.model} @ {vlm_config.api_base}")
    print()
    
    # Worker function for parallel processing
    def process_region(info):
        region_elem = info["element"]
        page_idx = info["page_idx"]
        region_id = info["region_id"]
        bbox_pt = info["bbox_pt"]
        
        # Crop region from PDF
        try:
            img_bytes = crop_region_from_pdf(pdf_path, page_idx, bbox_pt, dpi=dpi)
        except Exception as e:
            return {"region_id": region_id, "success": False, "error": f"Crop error: {e}"}
        
        # Extract style via VLM
        style = extract_style_from_image(img_bytes, vlm_config, dpi=dpi)
        
        if style is None:
            return {"region_id": region_id, "success": False, "error": "Style extraction failed"}
        
        return {
            "region_id": region_id,
            "success": True,
            "style": style,
            "element": region_elem
        }
    
    # Process regions in parallel
    results = []
    max_workers = min(parallel, len(regions_to_process)) if parallel > 1 else 1
    
    if max_workers > 1:
        print(f"Processing with {max_workers} parallel workers")
    
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {executor.submit(process_region, info): info for info in regions_to_process}
        
        for i, future in enumerate(as_completed(futures), 1):
            info = futures[future]
            region_id = info["region_id"]
            page_idx = info["page_idx"]
            
            print(f"[{i}/{len(regions_to_process)}] Region {region_id} (page {page_idx})...")
            
            try:
                result = future.result()
                results.append(result)
                
                if result["success"]:
                    style = result["style"]
                    style_info = []
                    if style.font_size_pt:
                        style_info.append(f"{style.font_size_pt:.1f}pt")
                    if style.bold:
                        style_info.append("bold")
                    if style.italic:
                        style_info.append("italic")
                    if style.color and style.color != "#000000":
                        style_info.append(style.color)
                    if style.font_name:
                        style_info.append(style.font_name)
                    
                    conf = f"conf={style.evidence_confidence:.2f}" if style.evidence_confidence else ""
                    print(f"  ✓ {', '.join(style_info)} {conf} ({style.extraction_time_ms}ms)")
                else:
                    print(f"  ✗ {result['error']}")
            except Exception as e:
                print(f"  ✗ Unexpected error: {e}")
                results.append({"region_id": region_id, "success": False, "error": str(e)})
    
    # Update DocIR XML with results (single-threaded)
    for result in results:
        if not result["success"]:
            continue
        
        region_elem = result["element"]
        style = result["style"]
        
        # Find or create text_content > paragraph > run elements
        text_content = region_elem.find(f"{{{DOCIR_NS}}}text_content")
        if text_content is None:
            continue
        
        for para in text_content.findall(f"{{{DOCIR_NS}}}paragraph"):
            for run in para.findall(f"{{{DOCIR_NS}}}run"):
                # Set computed style attributes
                if style.font_name:
                    run.set("font_name", style.font_name)
                if style.font_size_pt:
                    run.set("font_size_pt", f"{style.font_size_pt:.1f}")
                if style.bold is not None:
                    run.set("bold", str(style.bold).lower())
                if style.italic is not None:
                    run.set("italic", str(style.italic).lower())
                if style.underline is not None:
                    run.set("underline", str(style.underline).lower())
                if style.color:
                    run.set("color", style.color)
                
                # Set evidence attributes
                if style.evidence_pixel_height:
                    run.set("evidence_pixel_height", f"{style.evidence_pixel_height:.1f}")
                if style.evidence_color_sample:
                    run.set("evidence_color_sample", style.evidence_color_sample)
                if style.evidence_confidence is not None:
                    run.set("evidence_confidence", f"{style.evidence_confidence:.2f}")
        
        # Update provenance with style extraction info
        provenance = region_elem.find(f"{{{DOCIR_NS}}}provenance")
        if provenance is not None:
            style_extractor = provenance.find(f"{{{DOCIR_NS}}}style_extractor")
            if style_extractor is None:
                style_extractor = etree.SubElement(provenance, f"{{{DOCIR_NS}}}style_extractor")
            style_extractor.text = style.vlm_model
            
            style_time = provenance.find(f"{{{DOCIR_NS}}}style_extraction_time_ms")
            if style_time is None:
                style_time = etree.SubElement(provenance, f"{{{DOCIR_NS}}}style_extraction_time_ms")
            style_time.text = str(style.extraction_time_ms)
    
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
        description="Extract typography styles from DocIR regions using VLM"
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
    parser.add_argument(
        "--region-types",
        type=str,
        nargs="+",
        default=["text"],
        help="Region types to process (default: text)"
    )
    parser.add_argument(
        "--table-styles",
        action="store_true",
        help="Also extract table styles (borders, headers, cell colors)"
    )
    parser.add_argument(
        "--parallel",
        type=int,
        default=1,
        metavar="N",
        help="Number of parallel VLM requests (default: 1, sequential)"
    )
    
    args = parser.parse_args()
    
    vlm_config = VLMConfig(
        api_base=args.api_base,
        api_key=args.api_key,
        model=args.model
    )
    
    # Extract text styles
    output = extract_styles_from_docir(
        docir_path=args.docir_xml,
        pdf_path=args.pdf,
        output_path=args.output,
        vlm_config=vlm_config,
        dpi=args.dpi,
        region_types=args.region_types,
        parallel=args.parallel
    )
    
    # Extract table styles if requested
    if args.table_styles:
        print(f"\n{'='*60}")
        print("Extracting table styles...")
        print(f"{'='*60}")
        extract_table_styles_from_docir(
            docir_path=output,
            pdf_path=args.pdf,
            output_path=output,
            vlm_config=vlm_config,
            dpi=args.dpi
        )


if __name__ == "__main__":
    main()
