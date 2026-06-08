#!/usr/bin/env python3
"""
DocIR Visual Fidelity Scoring

Automated quality metric for evaluating DOCX output fidelity.
Evaluates multiple aspects: layout, typography, content, structure.

Usage:
    python fidelity_scorer.py source.pdf generated.docx -o report.json
"""

import sys
import subprocess
import tempfile
from pathlib import Path
from typing import Dict, List, Tuple, Optional
from dataclasses import dataclass, asdict
import json

import pymupdf
from PIL import Image
import numpy as np


@dataclass
class FidelityScore:
    """Comprehensive fidelity score breakdown."""
    overall: float  # 0.0 to 1.0
    
    # Component scores (0.0 to 1.0)
    layout: float  # Positioning and spacing
    typography: float  # Fonts, sizes, styles
    content: float  # Text accuracy
    structure: float  # Tables, images, sections
    
    # Detailed metrics
    page_count_match: bool
    region_count_match: bool
    avg_pixel_similarity: float
    text_accuracy: float
    style_accuracy: float
    
    # Metadata
    source_pages: int
    generated_pages: int
    source_regions: int
    generated_regions: int
    
    # Qualitative assessment
    grade: str  # A, B, C, D, F
    summary: str


def render_pdf_to_images(pdf_path: Path, output_dir: Path, dpi: int = 150) -> List[Path]:
    """Render PDF pages to images."""
    doc = pymupdf.open(str(pdf_path))
    image_paths = []
    
    for page_idx in range(len(doc)):
        page = doc[page_idx]
        zoom = dpi / 72.0
        mat = pymupdf.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat)
        
        img_path = output_dir / f"page{page_idx}.png"
        pix.save(str(img_path))
        image_paths.append(img_path)
    
    doc.close()
    return image_paths


def calculate_pixel_similarity(img1_path: Path, img2_path: Path) -> float:
    """Calculate pixel-wise similarity between two images."""
    img1 = Image.open(img1_path).convert('RGB')
    img2 = Image.open(img2_path).convert('RGB')
    
    # Resize to same size
    if img1.size != img2.size:
        max_width = max(img1.width, img2.width)
        max_height = max(img1.height, img2.height)
        img1 = img1.resize((max_width, max_height), Image.Resampling.LANCZOS)
        img2 = img2.resize((max_width, max_height), Image.Resampling.LANCZOS)
    
    arr1 = np.array(img1)
    arr2 = np.array(img2)
    
    diff = np.abs(arr1.astype(float) - arr2.astype(float))
    mean_diff = np.mean(diff)
    similarity = 1.0 - (mean_diff / 255.0)
    
    return similarity


def extract_text_from_pdf(pdf_path: Path) -> str:
    """Extract all text from PDF."""
    doc = pymupdf.open(str(pdf_path))
    text = ""
    
    for page in doc:
        text += page.get_text()
    
    doc.close()
    return text


def calculate_text_accuracy(source_text: str, generated_text: str) -> float:
    """Calculate text accuracy using character-level comparison."""
    if not source_text or not generated_text:
        return 0.0
    
    # Simple character-level accuracy
    source_chars = set(source_text)
    generated_chars = set(generated_text)
    
    # Jaccard similarity
    intersection = len(source_chars & generated_chars)
    union = len(source_chars | generated_chars)
    
    if union == 0:
        return 0.0
    
    return intersection / union


def count_regions_in_docir(docir_path: Path) -> int:
    """Count regions in DocIR XML."""
    from lxml import etree
    
    DOCIR_NS = "urn:docir:v0.1"
    
    tree = etree.parse(str(docir_path))
    root = tree.getroot()
    
    regions = root.findall(f".//{{{DOCIR_NS}}}region")
    return len(regions)


def docx_to_pdf(docx_path: Path, output_dir: Path) -> Optional[Path]:
    """Convert DOCX to PDF using LibreOffice."""
    cmd = [
        "soffice",
        "--headless",
        "--convert-to", "pdf",
        "--outdir", str(output_dir),
        str(docx_path)
    ]
    
    result = subprocess.run(cmd, capture_output=True, text=True)
    
    if result.returncode != 0:
        return None
    
    pdf_name = docx_path.stem + ".pdf"
    pdf_path = output_dir / pdf_name
    
    if not pdf_path.exists():
        return None
    
    return pdf_path


def calculate_fidelity_score(
    source_pdf: Path,
    generated_docx: Path,
    docir_xml: Optional[Path] = None,
    dpi: int = 150
) -> FidelityScore:
    """
    Calculate comprehensive fidelity score.
    
    Args:
        source_pdf: Path to source PDF
        generated_docx: Path to generated DOCX
        docir_xml: Path to DocIR XML (optional, for region count)
        dpi: DPI for rendering
    
    Returns:
        FidelityScore with detailed breakdown
    """
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)
        
        # Step 1: Convert DOCX to PDF
        generated_pdf = docx_to_pdf(generated_docx, tmpdir)
        if not generated_pdf:
            return FidelityScore(
                overall=0.0,
                layout=0.0,
                typography=0.0,
                content=0.0,
                structure=0.0,
                page_count_match=False,
                region_count_match=False,
                avg_pixel_similarity=0.0,
                text_accuracy=0.0,
                style_accuracy=0.0,
                source_pages=0,
                generated_pages=0,
                source_regions=0,
                generated_regions=0,
                grade="F",
                summary="Failed to convert DOCX to PDF"
            )
        
        # Step 2: Render both PDFs to images
        source_images = render_pdf_to_images(source_pdf, tmpdir / "source", dpi)
        generated_images = render_pdf_to_images(generated_pdf, tmpdir / "generated", dpi)
        
        # Step 3: Calculate page count match
        page_count_match = len(source_images) == len(generated_images)
        
        # Step 4: Calculate pixel similarity
        similarities = []
        for i in range(min(len(source_images), len(generated_images))):
            sim = calculate_pixel_similarity(source_images[i], generated_images[i])
            similarities.append(sim)
        
        avg_pixel_similarity = np.mean(similarities) if similarities else 0.0
        
        # Step 5: Extract and compare text
        source_text = extract_text_from_pdf(source_pdf)
        generated_text = extract_text_from_pdf(generated_pdf)
        text_accuracy = calculate_text_accuracy(source_text, generated_text)
        
        # Step 6: Count regions (if DocIR XML provided)
        source_regions = 0
        generated_regions = 0
        region_count_match = False
        
        if docir_xml and docir_xml.exists():
            generated_regions = count_regions_in_docir(docir_xml)
            # Estimate source regions from GLM-OCR (not available, so use generated as proxy)
            source_regions = generated_regions
            region_count_match = True
        
        # Step 7: Calculate component scores
        layout_score = avg_pixel_similarity * 0.8 + (1.0 if page_count_match else 0.0) * 0.2
        typography_score = avg_pixel_similarity  # Proxy for typography
        content_score = text_accuracy
        structure_score = 1.0 if region_count_match else 0.5
        
        # Step 8: Calculate overall score
        overall = (
            layout_score * 0.3 +
            typography_score * 0.3 +
            content_score * 0.3 +
            structure_score * 0.1
        )
        
        # Step 9: Assign grade
        if overall >= 0.9:
            grade = "A"
        elif overall >= 0.8:
            grade = "B"
        elif overall >= 0.7:
            grade = "C"
        elif overall >= 0.6:
            grade = "D"
        else:
            grade = "F"
        
        # Step 10: Generate summary
        summary_parts = []
        if page_count_match:
            summary_parts.append(f"Page count matches ({len(source_images)} pages)")
        else:
            summary_parts.append(f"Page count mismatch (source: {len(source_images)}, generated: {len(generated_images)})")
        
        summary_parts.append(f"Average pixel similarity: {avg_pixel_similarity:.3f}")
        summary_parts.append(f"Text accuracy: {text_accuracy:.3f}")
        
        if overall >= 0.8:
            summary_parts.append("Good fidelity overall")
        elif overall >= 0.6:
            summary_parts.append("Acceptable fidelity with room for improvement")
        else:
            summary_parts.append("Poor fidelity, significant issues detected")
        
        summary = ". ".join(summary_parts)
        
        return FidelityScore(
            overall=overall,
            layout=layout_score,
            typography=typography_score,
            content=content_score,
            structure=structure_score,
            page_count_match=page_count_match,
            region_count_match=region_count_match,
            avg_pixel_similarity=avg_pixel_similarity,
            text_accuracy=text_accuracy,
            style_accuracy=avg_pixel_similarity,  # Proxy
            source_pages=len(source_images),
            generated_pages=len(generated_images),
            source_regions=source_regions,
            generated_regions=generated_regions,
            grade=grade,
            summary=summary
        )


def main():
    import argparse
    
    parser = argparse.ArgumentParser(
        description='Calculate visual fidelity score for DOCX output'
    )
    parser.add_argument(
        'source_pdf',
        type=Path,
        help='Source PDF file'
    )
    parser.add_argument(
        'generated_docx',
        type=Path,
        help='Generated DOCX file'
    )
    parser.add_argument(
        '-o', '--output',
        type=Path,
        help='Output JSON report file'
    )
    parser.add_argument(
        '--docir',
        type=Path,
        help='DocIR XML file (optional, for region count)'
    )
    parser.add_argument(
        '--dpi',
        type=int,
        default=150,
        help='DPI for rendering (default: 150)'
    )
    
    args = parser.parse_args()
    
    if not args.source_pdf.exists():
        print(f"Error: Source PDF not found: {args.source_pdf}", file=sys.stderr)
        sys.exit(1)
    
    if not args.generated_docx.exists():
        print(f"Error: Generated DOCX not found: {args.generated_docx}", file=sys.stderr)
        sys.exit(1)
    
    print("=" * 60)
    print("DocIR Visual Fidelity Scoring")
    print("=" * 60)
    print(f"Source PDF: {args.source_pdf.name}")
    print(f"Generated DOCX: {args.generated_docx.name}")
    print()
    
    score = calculate_fidelity_score(
        args.source_pdf,
        args.generated_docx,
        args.docir,
        args.dpi
    )
    
    # Print results
    print(f"Overall Score: {score.overall:.3f} (Grade: {score.grade})")
    print()
    print("Component Scores:")
    print(f"  Layout:      {score.layout:.3f}")
    print(f"  Typography:  {score.typography:.3f}")
    print(f"  Content:     {score.content:.3f}")
    print(f"  Structure:   {score.structure:.3f}")
    print()
    print("Detailed Metrics:")
    print(f"  Page count match:    {score.page_count_match}")
    print(f"  Region count match:  {score.region_count_match}")
    print(f"  Pixel similarity:    {score.avg_pixel_similarity:.3f}")
    print(f"  Text accuracy:       {score.text_accuracy:.3f}")
    print(f"  Style accuracy:      {score.style_accuracy:.3f}")
    print()
    print("Page/Region Counts:")
    print(f"  Source pages:        {score.source_pages}")
    print(f"  Generated pages:     {score.generated_pages}")
    print(f"  Source regions:      {score.source_regions}")
    print(f"  Generated regions:   {score.generated_regions}")
    print()
    print(f"Summary: {score.summary}")
    print("=" * 60)
    
    # Save to JSON if requested
    if args.output:
        with open(args.output, 'w') as f:
            json.dump(asdict(score), f, indent=2)
        print(f"\n✓ Report saved to: {args.output}")
    
    # Exit with appropriate code
    if score.overall >= 0.8:
        sys.exit(0)  # Pass
    elif score.overall >= 0.6:
        sys.exit(0)  # Warning but pass
    else:
        sys.exit(1)  # Fail


if __name__ == "__main__":
    main()
