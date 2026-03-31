#!/usr/bin/env python3
"""Create image-info.json for the PDF route.

Parses pdfinfo output to extract page dimensions, lists rendered PNGs,
and writes a structured JSON manifest.

Usage:
    uv run .claude/skills/anything-to-docx/scripts/create_pdf_image_info.py \
        --workspace /path/to/workspace \
        --pdf-info /path/to/workspace/pdf-info.txt \
        --dpi 220
"""

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Optional

# A4 paper: 210mm × 297mm = 595.276 × 841.89 points (ISO 216 standard)
# Used as fallback when pdfinfo cannot determine page size
DEFAULT_WIDTH_PTS = 595.276
DEFAULT_HEIGHT_PTS = 841.89


def parse_page_size(text: str) -> tuple[float, float]:
    """Extract page width/height in pts from pdfinfo text.

    Handles lines like:
        Page size:      595.276 x 841.89 pts (A4)
        Page size:      612 x 792 pts

    Returns (width_pts, height_pts). Falls back to A4 if not found.
    """
    pattern = r"Page\s+size:\s+([\d.]+)\s+x\s+([\d.]+)\s+pts"
    match = re.search(pattern, text)
    if match:
        return float(match.group(1)), float(match.group(2))
    return DEFAULT_WIDTH_PTS, DEFAULT_HEIGHT_PTS


def parse_page_count(text: str) -> Optional[int]:
    """Extract page count from pdfinfo text.

    Handles lines like:
        Pages:          3

    Returns None if not found.
    """
    match = re.search(r"Pages:\s+(\d+)", text)
    if match:
        return int(match.group(1))
    return None


def count_pngs(input_images_dir: Path) -> int:
    """Count page-*.png files in the input-images directory."""
    if not input_images_dir.is_dir():
        return 0
    return len(sorted(input_images_dir.glob("page-*.png")))


def pts_to_px(pts: float, dpi: int) -> int:
    """Convert points to pixels at given DPI. 1 pt = 1/72 inch."""
    return round(pts * dpi / 72)


def read_pdf_info(pdf_info_path: Path) -> str:
    """Read pdf-info.txt, returning empty string on failure."""
    try:
        return pdf_info_path.read_text(encoding="utf-8")
    except (OSError, UnicodeDecodeError):
        return ""


def build_image_info(
    workspace: Path,
    pdf_info_text: str,
    dpi: int,
) -> dict:
    """Build the image-info dict from parsed pdfinfo and actual PNGs.

    Pure function: takes data in, returns data out.
    """
    width_pts, height_pts = parse_page_size(pdf_info_text)
    width_px = pts_to_px(width_pts, dpi)
    height_px = pts_to_px(height_pts, dpi)

    input_images_dir = workspace / "input-images"

    # Determine page count: prefer pdfinfo, fall back to counting PNGs
    parsed_count = parse_page_count(pdf_info_text)
    actual_png_count = count_pngs(input_images_dir)
    page_count = parsed_count if parsed_count is not None else actual_png_count

    # Build page entries only for PNGs that actually exist
    pages = []
    for i in range(1, page_count + 1):
        png_name = f"page-{i}.png"
        png_path = input_images_dir / png_name
        if png_path.exists():
            pages.append(
                {
                    "index": i,
                    "width_px": width_px,
                    "height_px": height_px,
                    "width_pts": width_pts,
                    "height_pts": height_pts,
                    "ext": "png",
                    "original": f"input-images/{png_name}",
                }
            )

    return {
        "page_count": len(pages),
        "dpi": dpi,
        "pages": pages,
    }


def write_image_info(workspace: Path, info: dict) -> Path:
    """Write image-info.json to workspace. Returns the output path."""
    output_path = workspace / "image-info.json"
    output_path.write_text(
        json.dumps(info, indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )
    return output_path


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Create image-info.json for PDF route"
    )
    parser.add_argument(
        "--workspace",
        type=Path,
        required=True,
        help="Path to the workspace directory",
    )
    parser.add_argument(
        "--pdf-info",
        type=Path,
        required=True,
        help="Path to pdf-info.txt from pdfinfo",
    )
    parser.add_argument(
        "--dpi",
        type=int,
        required=True,
        help="DPI used for rendering PNGs",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> None:
    args = parse_args(argv)
    pdf_info_text = read_pdf_info(args.pdf_info)
    info = build_image_info(args.workspace, pdf_info_text, args.dpi)
    output_path = write_image_info(args.workspace, info)
    print(f"Wrote {output_path} ({info['page_count']} pages, {info['dpi']} DPI)")


if __name__ == "__main__":
    main()
