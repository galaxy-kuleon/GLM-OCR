#!/usr/bin/env python3
"""Build DocIR from raster page images.

This is the PNG-first non-digital survival path:
  PDF/PNG(s) -> page PNG assets -> full-page image DocIR -> DOCX

It intentionally does not pretend raster pages are editable text. The goal is a
high-fidelity, non-timeout visual baseline for scanned/image-only input. OCR/VLM
editable overlays can be added on top of this image-backed truth layer later.
"""
from __future__ import annotations

import argparse
import shutil
import struct
from datetime import datetime
from pathlib import Path
from typing import Iterable

import lxml.etree as etree
import pymupdf

DOCIR_NS = "urn:docir:v0.1"
NSMAP = {"docir": DOCIR_NS}
IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg"}
MIN_DPI = 36
MAX_DPI = 600
MAX_PAGES = 500
MAX_RENDER_PIXELS_PER_PAGE = 80_000_000


def tag(name: str) -> str:
    return f"{{{DOCIR_NS}}}{name}"


def png_size(path: Path) -> tuple[int, int]:
    with path.open("rb") as f:
        header = f.read(24)
    if len(header) < 24 or not header.startswith(b"\x89PNG\r\n\x1a\n"):
        raise ValueError(f"Not a PNG file: {path}")
    return struct.unpack(">II", header[16:24])


def jpeg_size(path: Path) -> tuple[int, int]:
    # Small dependency-free JPEG SOF parser; stream instead of reading the full
    # file so large page images do not create avoidable memory spikes.
    with path.open("rb") as f:
        if f.read(2) != b"\xff\xd8":
            raise ValueError(f"Not a JPEG file: {path}")
        while True:
            marker = b""
            marker_prefix = f.read(1)
            if not marker_prefix:
                break
            while marker_prefix == b"\xff":
                marker = f.read(1)
                if marker != b"\xff":
                    break
            if not marker:
                break
            marker_int = marker[0]
            if marker_int in {0xD8, 0xD9, 0x01} or 0xD0 <= marker_int <= 0xD7:
                continue
            length_bytes = f.read(2)
            if len(length_bytes) != 2:
                break
            length = int.from_bytes(length_bytes, "big")
            if length < 2:
                break
            if marker_int in {0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF}:
                segment = f.read(length - 2)
                if len(segment) < 5:
                    break
                height = int.from_bytes(segment[1:3], "big")
                width = int.from_bytes(segment[3:5], "big")
                return width, height
            f.seek(length - 2, 1)
    raise ValueError(f"Could not read JPEG dimensions: {path}")


def image_size(path: Path) -> tuple[int, int]:
    suffix = path.suffix.lower()
    if suffix == ".png":
        return png_size(path)
    if suffix in {".jpg", ".jpeg"}:
        return jpeg_size(path)
    raise ValueError(f"Unsupported image type: {path}")


def iter_input_images(input_path: Path) -> Iterable[Path]:
    if input_path.is_dir():
        yield from sorted(p for p in input_path.iterdir() if p.suffix.lower() in IMAGE_SUFFIXES)
    elif input_path.suffix.lower() in IMAGE_SUFFIXES:
        yield input_path
    else:
        return


def render_pdf_to_pngs(pdf_path: Path, assets_dir: Path, dpi: int) -> list[tuple[Path, float, float, int, int]]:
    """Render PDF pages to PNGs and return (path, width_pt, height_pt, width_px, height_px)."""
    doc = pymupdf.open(str(pdf_path))
    try:
        if len(doc) > MAX_PAGES:
            raise ValueError(f"PDF has {len(doc)} pages; raster fallback limit is {MAX_PAGES}")
        zoom = dpi / 72.0
        matrix = pymupdf.Matrix(zoom, zoom)
        pages: list[tuple[Path, float, float, int, int]] = []
        for page_idx in range(len(doc)):
            page = doc[page_idx]
            expected_pixels = int(page.rect.width * zoom) * int(page.rect.height * zoom)
            if expected_pixels > MAX_RENDER_PIXELS_PER_PAGE:
                raise ValueError(
                    f"Page {page_idx + 1} would render to ~{expected_pixels:,} pixels; "
                    f"limit is {MAX_RENDER_PIXELS_PER_PAGE:,}. Lower --dpi."
                )
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            rel = assets_dir / f"page_{page_idx + 1:04d}.png"
            pix.save(str(rel))
            pages.append((rel, float(page.rect.width), float(page.rect.height), int(pix.width), int(pix.height)))
        return pages
    finally:
        doc.close()


def collect_images(input_path: Path, assets_dir: Path, dpi: int) -> list[tuple[Path, float, float, int, int]]:
    """Collect/render input pages as PNG/JPEG assets.

    Returns (asset_path, width_pt, height_pt, width_px, height_px).
    """
    assets_dir.mkdir(parents=True, exist_ok=True)
    if input_path.suffix.lower() == ".pdf":
        return render_pdf_to_pngs(input_path, assets_dir, dpi)

    image_paths = list(iter_input_images(input_path))
    if len(image_paths) > MAX_PAGES:
        raise ValueError(f"Image input has {len(image_paths)} pages; raster fallback limit is {MAX_PAGES}")

    pages = []
    for idx, src in enumerate(image_paths, start=1):
        width_px, height_px = image_size(src)
        pixels = width_px * height_px
        if pixels > MAX_RENDER_PIXELS_PER_PAGE:
            raise ValueError(
                f"Image page {idx} has {pixels:,} pixels; "
                f"limit is {MAX_RENDER_PIXELS_PER_PAGE:,}. Downsample the input."
            )
        width_pt = width_px * 72.0 / dpi
        height_pt = height_px * 72.0 / dpi
        dest = assets_dir / f"page_{idx:04d}{src.suffix.lower()}"
        if src.resolve() != dest.resolve():
            shutil.copy2(src, dest)
        pages.append((dest, width_pt, height_pt, width_px, height_px))
    if not pages:
        raise ValueError(f"Input is not a PDF, PNG/JPEG, or image directory: {input_path}")
    return pages


def add_text(parent: etree._Element, name: str, value: str) -> etree._Element:
    elem = etree.SubElement(parent, tag(name))
    elem.text = value
    return elem


def build_docir(input_path: Path, output_path: Path, title: str | None = None, dpi: int = 200) -> Path:
    if dpi < MIN_DPI or dpi > MAX_DPI:
        raise ValueError(f"dpi must be between {MIN_DPI} and {MAX_DPI}; got {dpi}")
    assets_dir = output_path.parent / "assets"
    pages = collect_images(input_path, assets_dir, dpi)
    if not pages:
        raise ValueError(f"No renderable pages found in input: {input_path}")

    document = etree.Element(tag("document"), nsmap=NSMAP)
    document.set("version", "0.1.0")
    document.set("source_pdf", input_path.name)
    document.set("generated_at", datetime.now().isoformat())
    document.set("generator", "raster-ir-builder-v0.1.0")

    metadata = etree.SubElement(document, tag("metadata"))
    add_text(metadata, "title", title or input_path.stem)
    add_text(metadata, "page_count", str(len(pages)))
    default_size = etree.SubElement(metadata, tag("default_page_size"))
    default_size.set("width_pt", f"{pages[0][1]:.2f}")
    default_size.set("height_pt", f"{pages[0][2]:.2f}")
    pipe = etree.SubElement(metadata, tag("pipeline_info"))
    add_text(pipe, "layout_detector", "Raster page image fallback")
    add_text(pipe, "ocr_engine", "none: image-backed non-digital baseline")
    add_text(pipe, "style_extractor", "none")

    pages_el = etree.SubElement(document, tag("pages"))
    assets_el = etree.SubElement(document, tag("assets"))
    for page_idx, (asset_path, width_pt, height_pt, width_px, height_px) in enumerate(pages):
        asset_id = f"page_image_{page_idx + 1:04d}"
        asset_el = etree.SubElement(assets_el, tag("asset"))
        asset_el.set("id", asset_id)
        asset_el.set("mime_type", "image/png" if asset_path.suffix.lower() == ".png" else "image/jpeg")
        asset_el.set("width_px", str(width_px))
        asset_el.set("height_px", str(height_px))
        add_text(asset_el, "file_path", str(asset_path.relative_to(output_path.parent)))
        add_text(asset_el, "extraction_source", "raster-page-render")

        page_el = etree.SubElement(pages_el, tag("page"))
        page_el.set("index", str(page_idx))
        size_el = etree.SubElement(page_el, tag("page_size"))
        size_el.set("width_pt", f"{width_pt:.2f}")
        size_el.set("height_pt", f"{height_pt:.2f}")
        regions_el = etree.SubElement(page_el, tag("regions"))
        region = etree.SubElement(regions_el, tag("region"))
        region.set("id", f"r0_p{page_idx}")
        region.set("type", "image")
        region.set("native_label", "raster_page_image")
        region.set("order", "0")
        bbox = etree.SubElement(region, tag("bbox"))
        bbox.set("x", "0.00")
        # Schema/semantic-validator-friendly bbox origin. The DOCX generator
        # special-cases native_label=raster_page_image so this still anchors at
        # page top, while y + height remains inside page bounds.
        bbox.set("y", "0.00")
        bbox.set("width", f"{width_pt:.2f}")
        bbox.set("height", f"{height_pt:.2f}")
        prov = etree.SubElement(region, tag("provenance"))
        add_text(prov, "source", "raster-ir-builder")
        add_text(prov, "confidence", "1.00")
        add_text(prov, "detection_model", "pdftocairo-compatible page image")
        image_content = etree.SubElement(region, tag("image_content"))
        image_ref = etree.SubElement(image_content, tag("image_reference"))
        image_ref.set("asset_id", asset_id)
        image_ref.set("mime_type", asset_el.get("mime_type") or "image/png")
        vf = etree.SubElement(image_content, tag("visual_features"))
        add_text(vf, "contains_text", "true")
        add_text(vf, "description", "raster_fallback=true; full-page image-backed non-digital baseline")

    etree.SubElement(document, tag("cross_page_hints"))
    output_path.parent.mkdir(parents=True, exist_ok=True)
    etree.ElementTree(document).write(str(output_path), encoding="UTF-8", xml_declaration=True, pretty_print=True)
    print(f"✓ Raster DocIR XML written to: {output_path}")
    print(f"  Pages: {len(pages)}")
    print(f"  Assets: {len(pages)}")
    return output_path


def main() -> None:
    ap = argparse.ArgumentParser(description="Build image-backed DocIR from PDF/PNG pages")
    ap.add_argument("input", type=Path, help="Input PDF, PNG/JPEG, or directory of page images")
    ap.add_argument("-o", "--output", type=Path, required=True)
    ap.add_argument("--title")
    ap.add_argument("--dpi", type=int, default=200, help="DPI for PDF render or PNG pixel-to-point mapping")
    args = ap.parse_args()
    build_docir(args.input, args.output, args.title, args.dpi)


if __name__ == "__main__":
    main()
