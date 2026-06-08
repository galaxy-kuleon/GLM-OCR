#!/usr/bin/env python3
"""
Visual Comparison Test for DocIR Pipeline.

Renders DOCX back to PDF (via LibreOffice), then compares with source PDF.
Uses VLM judge to score visual fidelity.

Pipeline:
  1. Convert DOCX → PDF (LibreOffice headless)
  2. Render both PDFs to page images (pymupdf)
  3. Send side-by-side comparison to VLM for scoring
"""

import base64
import json
import subprocess
import sys
import tempfile
import time
from io import BytesIO
from pathlib import Path
from typing import Optional, Tuple, List
from dataclasses import dataclass

import pymupdf
import requests


@dataclass
class VLMConfig:
    """VLM API configuration."""
    api_base: str = "http://localhost:11234/v1"
    api_key: str = "change-me-local-key"
    model: str = "qwen3.6-35b-a3b-q7"
    max_tokens: int = 4096
    temperature: float = 0.1
    timeout: int = 180


@dataclass
class ComparisonResult:
    """Result of visual comparison."""
    overall_score: float  # 0-10
    text_fidelity: float  # 0-10
    layout_fidelity: float  # 0-10
    style_fidelity: float  # 0-10
    notes: str
    page_count_source: int
    page_count_generated: int


# VLM Comparison Prompt
COMPARISON_PROMPT = """You are a document fidelity judge. Compare the SOURCE document (left image) with the GENERATED document (right image) and score how well the generated version reproduces the original.

The GENERATED document was created by: PDF → OCR → DocIR XML → DOCX → PDF roundtrip.

Score each dimension from 0-10:
1. **text_fidelity**: How accurately is the text content reproduced? (spelling, completeness, reading order)
2. **layout_fidelity**: How well does the layout match? (positions, spacing, column structure, tables)
3. **style_fidelity**: How well are visual styles preserved? (fonts, sizes, bold/italic, colors, alignment)
4. **overall_score**: Overall visual similarity score (0=completely different, 10=identical)

Respond in JSON only (no markdown fences, no explanation):
{
  "text_fidelity": <0-10>,
  "layout_fidelity": <0-10>,
  "style_fidelity": <0-10>,
  "overall_score": <0-10>,
  "notes": "<brief description of what matches and what doesn't>"
}

Guidelines:
- 9-10: Nearly identical, minor differences only
- 7-8: Good reproduction, some layout or style differences
- 5-6: Recognizable as same document, significant differences
- 3-4: Same general content but poor reproduction
- 0-2: Very different or unrecognizable

Consider:
- Text content accuracy (OCR may have errors)
- Positioning and spacing (may not be pixel-perfect)
- Font appearance (may differ but should be similar weight/size)
- Table structure (rows/cols should match)
- Images (should be present, may differ in exact rendering)

Respond ONLY with valid JSON."""


def convert_docx_to_pdf(docx_path: Path, output_dir: Path) -> Optional[Path]:
    """
    Convert DOCX to PDF using LibreOffice headless.
    
    Args:
        docx_path: Path to input DOCX file
        output_dir: Directory for output PDF
    
    Returns:
        Path to generated PDF, or None on failure
    """
    output_dir.mkdir(parents=True, exist_ok=True)
    
    cmd = [
        "/opt/homebrew/bin/soffice",
        "--headless",
        "--convert-to", "pdf",
        "--outdir", str(output_dir),
        str(docx_path)
    ]
    
    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=60
        )
        
        if result.returncode != 0:
            print(f"✗ LibreOffice error: {result.stderr}")
            return None
        
        # Find the output PDF
        pdf_name = docx_path.stem + ".pdf"
        pdf_path = output_dir / pdf_name
        
        if pdf_path.exists():
            return pdf_path
        else:
            print(f"✗ Output PDF not found: {pdf_path}")
            return None
            
    except subprocess.TimeoutExpired:
        print("✗ LibreOffice conversion timed out")
        return None
    except Exception as e:
        print(f"✗ Conversion error: {e}")
        return None


def render_pdf_page_to_image(
    pdf_path: Path,
    page_index: int = 0,
    dpi: int = 150
) -> bytes:
    """
    Render a PDF page to JPEG image bytes.
    
    Args:
        pdf_path: Path to PDF file
        page_index: Page number (0-indexed)
        dpi: Resolution for rendering
    
    Returns:
        JPEG image bytes
    """
    doc = pymupdf.open(str(pdf_path))
    
    if page_index >= len(doc):
        doc.close()
        raise ValueError(f"Page {page_index} not found (PDF has {len(doc)} pages)")
    
    page = doc[page_index]
    zoom = dpi / 72.0
    mat = pymupdf.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat)
    img_data = pix.tobytes("jpeg")
    doc.close()
    
    return img_data


def create_side_by_side_image(
    source_img: bytes,
    generated_img: bytes,
    max_width: int = 800
) -> bytes:
    """
    Create a side-by-side comparison image.
    
    Args:
        source_img: Source PDF page image (JPEG bytes)
        generated_img: Generated PDF page image (JPEG bytes)
        max_width: Maximum width for each image
    
    Returns:
        Combined JPEG image bytes
    """
    from PIL import Image
    
    # Load images
    source = Image.open(BytesIO(source_img))
    generated = Image.open(BytesIO(generated_img))
    
    # Resize to fit
    def resize_to_width(img, target_width):
        ratio = target_width / img.width
        new_height = int(img.height * ratio)
        return img.resize((target_width, new_height), Image.LANCZOS)
    
    source_resized = resize_to_width(source, max_width)
    generated_resized = resize_to_width(generated, max_width)
    
    # Create combined image
    total_width = max_width * 2 + 20  # 20px gap
    total_height = max(source_resized.height, generated_resized.height) + 40  # 40px for labels
    
    combined = Image.new('RGB', (total_width, total_height), (255, 255, 255))
    
    # Paste images
    combined.paste(source_resized, (0, 40))
    combined.paste(generated_resized, (max_width + 20, 40))
    
    # Add labels
    from PIL import ImageDraw
    draw = ImageDraw.Draw(combined)
    draw.text((10, 10), "SOURCE", fill=(0, 0, 0))
    draw.text((max_width + 30, 10), "GENERATED", fill=(0, 0, 0))
    
    # Convert to JPEG
    output = BytesIO()
    combined.save(output, format="JPEG", quality=85)
    return output.getvalue()


def compare_with_vlm(
    comparison_image: bytes,
    vlm_config: VLMConfig
) -> Optional[ComparisonResult]:
    """
    Send comparison image to VLM for fidelity scoring.
    
    Args:
        comparison_image: Side-by-side JPEG image bytes
        vlm_config: VLM configuration
    
    Returns:
        ComparisonResult or None on failure
    """
    img_b64 = base64.b64encode(comparison_image).decode()
    
    payload = {
        "model": vlm_config.model,
        "messages": [
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": COMPARISON_PROMPT},
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
            print(f"✗ VLM API error: HTTP {resp.status_code}")
            return None
        
        data = resp.json()
        content = data["choices"][0]["message"].get("content", "")
        
        if not content:
            print("✗ Empty VLM response")
            return None
        
        # Parse JSON response
        clean = content.strip()
        if clean.startswith("```"):
            lines = clean.split("\n")
            clean = "\n".join(lines[1:-1]) if lines[-1].strip() == "```" else "\n".join(lines[1:])
        
        try:
            parsed = json.loads(clean)
        except json.JSONDecodeError as e:
            print(f"✗ JSON parse error: {e}")
            print(f"  Content: {content[:300]}")
            return None
        
        result = ComparisonResult(
            overall_score=float(parsed.get("overall_score", 0)),
            text_fidelity=float(parsed.get("text_fidelity", 0)),
            layout_fidelity=float(parsed.get("layout_fidelity", 0)),
            style_fidelity=float(parsed.get("style_fidelity", 0)),
            notes=parsed.get("notes", ""),
            page_count_source=0,
            page_count_generated=0
        )
        
        print(f"✓ VLM comparison completed in {elapsed_ms}ms")
        return result
        
    except requests.exceptions.Timeout:
        print(f"✗ VLM request timeout ({vlm_config.timeout}s)")
        return None
    except Exception as e:
        print(f"✗ VLM request error: {e}")
        return None


def run_visual_comparison(
    source_pdf: Path,
    generated_docx: Path,
    vlm_config: Optional[VLMConfig] = None,
    dpi: int = 150,
    output_dir: Optional[Path] = None
) -> Optional[ComparisonResult]:
    """
    Run full visual comparison pipeline.
    
    Args:
        source_pdf: Path to source PDF
        generated_docx: Path to generated DOCX
        vlm_config: VLM configuration
        dpi: Resolution for rendering
        output_dir: Directory for intermediate files
    
    Returns:
        ComparisonResult or None on failure
    """
    if vlm_config is None:
        vlm_config = VLMConfig()
    
    if output_dir is None:
        output_dir = Path(tempfile.mkdtemp(prefix="docir-comparison-"))
    
    print(f"{'='*60}")
    print("Visual Comparison Test")
    print(f"{'='*60}")
    print(f"Source PDF:    {source_pdf}")
    print(f"Generated DOCX: {generated_docx}")
    print(f"Output dir:    {output_dir}")
    print()
    
    # Step 1: Convert DOCX to PDF
    print("→ Step 1: Converting DOCX to PDF...")
    generated_pdf = convert_docx_to_pdf(generated_docx, output_dir)
    
    if generated_pdf is None:
        print("✗ Failed to convert DOCX to PDF")
        return None
    
    print(f"✓ Generated PDF: {generated_pdf}")
    
    # Step 2: Get page counts
    source_doc = pymupdf.open(str(source_pdf))
    generated_doc = pymupdf.open(str(generated_pdf))
    
    source_pages = len(source_doc)
    generated_pages = len(generated_doc)
    
    source_doc.close()
    generated_doc.close()
    
    print(f"  Source pages: {source_pages}")
    print(f"  Generated pages: {generated_pages}")
    
    # Step 3: Compare each page
    pages_to_compare = min(source_pages, generated_pages)
    
    if pages_to_compare == 0:
        print("✗ No pages to compare")
        return None
    
    print(f"\n→ Step 2: Comparing {pages_to_compare} page(s)...")
    
    all_results = []
    
    for page_idx in range(pages_to_compare):
        print(f"\n  Page {page_idx + 1}/{pages_to_compare}:")
        
        # Render pages
        try:
            source_img = render_pdf_page_to_image(source_pdf, page_idx, dpi=dpi)
            generated_img = render_pdf_page_to_image(generated_pdf, page_idx, dpi=dpi)
        except Exception as e:
            print(f"  ✗ Render error: {e}")
            continue
        
        # Create side-by-side comparison
        comparison_img = create_side_by_side_image(source_img, generated_img)
        
        # Save comparison image for debugging
        comparison_path = output_dir / f"comparison_page{page_idx}.jpg"
        with open(comparison_path, "wb") as f:
            f.write(comparison_img)
        
        # Send to VLM for scoring
        result = compare_with_vlm(comparison_img, vlm_config)
        
        if result:
            result.page_count_source = source_pages
            result.page_count_generated = generated_pages
            all_results.append(result)
            
            print(f"  ✓ Overall: {result.overall_score}/10")
            print(f"    Text: {result.text_fidelity}/10, Layout: {result.layout_fidelity}/10, Style: {result.style_fidelity}/10")
            print(f"    Notes: {result.notes[:100]}...")
    
    if not all_results:
        print("\n✗ No successful comparisons")
        return None
    
    # Aggregate results (average across pages)
    avg_result = ComparisonResult(
        overall_score=sum(r.overall_score for r in all_results) / len(all_results),
        text_fidelity=sum(r.text_fidelity for r in all_results) / len(all_results),
        layout_fidelity=sum(r.layout_fidelity for r in all_results) / len(all_results),
        style_fidelity=sum(r.style_fidelity for r in all_results) / len(all_results),
        notes=" | ".join(r.notes for r in all_results),
        page_count_source=source_pages,
        page_count_generated=generated_pages
    )
    
    print(f"\n{'='*60}")
    print("Visual Comparison Summary")
    print(f"{'='*60}")
    print(f"Pages compared: {len(all_results)}/{source_pages}")
    print(f"Overall score:  {avg_result.overall_score:.1f}/10")
    print(f"Text fidelity:  {avg_result.text_fidelity:.1f}/10")
    print(f"Layout fidelity: {avg_result.layout_fidelity:.1f}/10")
    print(f"Style fidelity: {avg_result.style_fidelity:.1f}/10")
    print(f"Notes: {avg_result.notes[:200]}")
    print(f"{'='*60}")
    
    return avg_result


def main():
    """CLI entry point."""
    import argparse
    
    parser = argparse.ArgumentParser(
        description="Visual comparison test: DOCX → PDF → compare with source"
    )
    parser.add_argument(
        "source_pdf",
        type=Path,
        help="Path to source PDF"
    )
    parser.add_argument(
        "generated_docx",
        type=Path,
        help="Path to generated DOCX"
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        help="Directory for intermediate files"
    )
    parser.add_argument(
        "--dpi",
        type=int,
        default=150,
        help="DPI for rendering (default: 150)"
    )
    parser.add_argument(
        "--model",
        type=str,
        default="qwen3.6-35b-a3b-q7",
        help="VLM model name"
    )
    
    args = parser.parse_args()
    
    if not args.source_pdf.exists():
        print(f"Error: Source PDF not found: {args.source_pdf}", file=sys.stderr)
        sys.exit(1)
    
    if not args.generated_docx.exists():
        print(f"Error: Generated DOCX not found: {args.generated_docx}", file=sys.stderr)
        sys.exit(1)
    
    vlm_config = VLMConfig(model=args.model)
    
    result = run_visual_comparison(
        source_pdf=args.source_pdf,
        generated_docx=args.generated_docx,
        vlm_config=vlm_config,
        dpi=args.dpi,
        output_dir=args.output_dir
    )
    
    if result is None:
        sys.exit(1)


if __name__ == "__main__":
    main()
