#!/usr/bin/env python3
"""Resize images for VLM processing.

Caps width at 1820px (Lanczos downscale, no upscaling), converts to PNG,
and writes page metadata as JSON.

Usage:
    uv run --with Pillow .claude/skills/anything-to-docx/scripts/resize_images.py \
        --images "/abs/path/img1.jpg,/abs/path/img2.jpg" \
        --workspace /path/to/workspace
"""

import argparse
import json
import os
import sys

from PIL import Image

# A4 at 220 DPI = 1819px wide — fits within Qwen3.5 VLM's 2560px patch grid upper bound
# Going wider wastes tokens (model downscales internally); going narrower loses OCR detail
MAX_WIDTH = 1820

# 220 DPI balances OCR quality vs VLM token cost: ~5980 tokens/page for A4
# (200 DPI = ~5040 tokens, 300 DPI = ~11700 tokens — diminishing returns above 220)
DPI = 220

PX_TO_PT = 72 / DPI  # 1 point = 1/72 inch (standard PostScript/PDF definition)


def px_to_pts(px):
    return px * PX_TO_PT


def process_image(src_path, index, output_dir):
    """Open, optionally resize, save as PNG. Returns page info dict or None on error."""
    try:
        img = Image.open(src_path)
        img.load()  # force decode to catch corrupt files early
    except Exception as e:
        print(f"ERROR: Cannot open '{src_path}': {e}", file=sys.stderr)
        return None

    orig_w, orig_h = img.size

    if orig_w > MAX_WIDTH:
        ratio = MAX_WIDTH / orig_w
        new_w = MAX_WIDTH
        new_h = round(orig_h * ratio)
        img = img.resize((new_w, new_h), Image.LANCZOS)
        print(f"Resized page {index}: {orig_w}x{orig_h} -> {new_w}x{new_h}")
    else:
        new_w, new_h = orig_w, orig_h
        print(f"Page {index}: {new_w}x{new_h} (no resize needed)")

    # Convert to RGB if needed (handles RGBA, palette, etc.)
    if img.mode not in ("RGB", "L"):
        img = img.convert("RGB")

    dest = os.path.join(output_dir, f"page-{index}.png")
    img.save(dest, "PNG")

    return {
        "index": index,
        "width_px": new_w,
        "height_px": new_h,
        "width_pts": round(px_to_pts(new_w), 3),
        "height_pts": round(px_to_pts(new_h), 3),
        "original": src_path,
    }


def main():
    parser = argparse.ArgumentParser(description="Resize images for VLM processing")
    parser.add_argument("--images", required=True, help="Comma-separated absolute image paths")
    parser.add_argument("--workspace", required=True, help="Workspace directory path")
    args = parser.parse_args()

    image_paths = [p.strip() for p in args.images.split(",") if p.strip()]
    output_dir = os.path.join(args.workspace, "input-images")
    os.makedirs(output_dir, exist_ok=True)

    pages = []
    for i, src in enumerate(image_paths, start=1):
        info = process_image(src, i, output_dir)
        if info is not None:
            pages.append(info)

    if not pages:
        print("ERROR: No valid images processed.", file=sys.stderr)
        sys.exit(1)

    manifest = {"page_count": len(pages), "pages": pages}
    manifest_path = os.path.join(args.workspace, "image-info.json")
    with open(manifest_path, "w") as f:
        json.dump(manifest, f, indent=2)

    print(f"Wrote {manifest_path} ({len(pages)} pages)")


if __name__ == "__main__":
    main()
