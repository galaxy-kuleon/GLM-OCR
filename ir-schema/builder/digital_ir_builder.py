#!/usr/bin/env python3
"""Build DocIR directly from born-digital PDFs using PyMuPDF.

This is the true digital fast path: PDF → PyMuPDF blocks/spans → DocIR.
It avoids GLM-OCR/Ollama OCR entirely for PDFs with extractable text.
"""
from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path
from typing import Any

import lxml.etree as etree
import pymupdf

DOCIR_NS = "urn:docir:v0.1"
NSMAP = {"docir": DOCIR_NS}


def tag(name: str) -> str:
    return f"{{{DOCIR_NS}}}{name}"


def color_to_hex(color: int | None) -> str | None:
    if color is None:
        return None
    try:
        return f"#{int(color) & 0xFFFFFF:06x}"
    except Exception:
        return None


def is_bold(font_name: str, flags: int = 0) -> bool:
    s = (font_name or "").lower()
    return "bold" in s or "black" in s or bool(flags & 16)


def is_italic(font_name: str, flags: int = 0) -> bool:
    s = (font_name or "").lower()
    return "italic" in s or "oblique" in s or bool(flags & 2)


def normalize_font_name(font_name: str) -> str:
    import re
    if not font_name:
        return ""
    name = re.sub(r"^[A-Z]{6}\+", "", font_name)
    name = name.replace("MT", "") if name.endswith("MT") else name
    name = name.replace("PS", "")
    name = re.sub(r"[-,]?(Bold|Italic|Oblique|Regular|Light|Medium|Black).*$", "", name, flags=re.I)
    name = re.sub(r"([a-z])([A-Z])", r"\1 \2", name)
    return name.replace("-", " ").strip() or font_name


def add_provenance(region: etree._Element, source: str = "pymupdf-digital-direct") -> None:
    prov = etree.SubElement(region, tag("provenance"))
    etree.SubElement(prov, tag("source")).text = source
    etree.SubElement(prov, tag("confidence")).text = "1.00"
    etree.SubElement(prov, tag("detection_model")).text = "PyMuPDF text blocks"


def add_bbox(parent: etree._Element, rect: tuple[float, float, float, float], page_h: float) -> None:
    x0, y0, x1, y1 = rect
    # PyMuPDF y is from page top; DocIR y is PDF coordinate from bottom, at visual top.
    y_docir = page_h - y0
    bbox = etree.SubElement(parent, tag("bbox"))
    bbox.set("x", f"{x0:.2f}")
    bbox.set("y", f"{y_docir:.2f}")
    bbox.set("width", f"{max(0, x1 - x0):.2f}")
    bbox.set("height", f"{max(0, y1 - y0):.2f}")


def build_text_region(block: dict[str, Any], page_idx: int, region_idx: int, page_h: float) -> etree._Element | None:
    lines = block.get("lines") or []
    if not lines:
        return None

    text = "".join(
        span.get("text", "")
        for line in lines
        for span in (line.get("spans") or [])
    ).strip()
    if not text:
        return None

    region = etree.Element(tag("region"))
    region.set("id", f"r{region_idx}_p{page_idx}")
    region.set("type", "text")
    region.set("native_label", "digital_text_block")
    region.set("order", str(region_idx))

    bbox = block.get("bbox")
    if not bbox:
        return None
    add_bbox(region, tuple(bbox), page_h)
    add_provenance(region)

    text_content = etree.SubElement(region, tag("text_content"))
    for line in lines:
        para = etree.SubElement(text_content, tag("paragraph"))
        # Preserve line-level runs/spans for formatting.
        spans = line.get("spans") or []
        if not spans:
            etree.SubElement(para, tag("run")).text = ""
            continue
        for span in spans:
            span_text = span.get("text", "")
            if not span_text:
                continue
            run = etree.SubElement(para, tag("run"))
            run.text = span_text
            font = span.get("font") or ""
            size = span.get("size")
            flags = int(span.get("flags") or 0)
            color = color_to_hex(span.get("color"))
            if font:
                run.set("font_name", normalize_font_name(font))
            if size:
                run.set("font_size_pt", f"{float(size):.2f}")
            if is_bold(font, flags):
                run.set("bold", "true")
            if is_italic(font, flags):
                run.set("italic", "true")
            if color:
                run.set("color", color)

    return region


def build_text_box_region(
    box_rect: tuple[float, float, float, float],
    blocks: list[dict[str, Any]],
    page_idx: int,
    region_idx: int,
    page_h: float,
) -> etree._Element | None:
    """Build one floating text-box region from text blocks inside a drawn rectangle."""
    inner_blocks = []
    for block in blocks:
        if block.get("type") != 0 or not block.get("bbox"):
            continue
        if rect_overlap_ratio(tuple(block["bbox"]), box_rect) > 0.5:
            inner_blocks.append(block)
    if not inner_blocks:
        return None
    inner_blocks.sort(key=lambda b: (b["bbox"][1], b["bbox"][0]))

    region = etree.Element(tag("region"))
    region.set("id", f"r{region_idx}_p{page_idx}")
    region.set("type", "text")
    region.set("native_label", "pymupdf_text_box")
    region.set("floating", "true")
    region.set("order", str(region_idx))
    add_bbox(region, box_rect, page_h)
    add_provenance(region, "pymupdf-drawing-rect")

    text_content = etree.SubElement(region, tag("text_content"))
    for block in inner_blocks:
        for line in block.get("lines", []) or []:
            para = etree.SubElement(text_content, tag("paragraph"))
            for span in line.get("spans", []) or []:
                span_text = span.get("text", "")
                if not span_text:
                    continue
                run = etree.SubElement(para, tag("run"))
                run.text = span_text
                font = span.get("font") or ""
                size = span.get("size")
                flags = int(span.get("flags") or 0)
                color = color_to_hex(span.get("color"))
                if font:
                    run.set("font_name", normalize_font_name(font))
                if size:
                    run.set("font_size_pt", f"{float(size):.2f}")
                if is_bold(font, flags):
                    run.set("bold", "true")
                if is_italic(font, flags):
                    run.set("italic", "true")
                if color:
                    run.set("color", color)
    if not text_content.findall(tag("paragraph")):
        return None
    return region


def build_table_region(table: Any, page_idx: int, region_idx: int, page_h: float) -> etree._Element:
    """Build a DocIR table region from PyMuPDF Table."""
    region = etree.Element(tag("region"))
    region.set("id", f"r{region_idx}_p{page_idx}")
    region.set("type", "table")
    region.set("native_label", "pymupdf_table")
    region.set("order", str(region_idx))
    add_bbox(region, tuple(table.bbox), page_h)
    add_provenance(region, "pymupdf-find_tables")

    data = table.extract() or []
    rows = len(data)
    cols = max((len(r) for r in data), default=0)
    table_content = etree.SubElement(region, tag("table_content"))
    table_content.set("rows", str(rows))
    table_content.set("cols", str(cols))
    table_style = etree.SubElement(table_content, tag("table_style"))
    table_style.set("border_visible", "true")
    table_style.set("border_color", "#000000")
    if rows > 1:
        table_style.set("header_row", "true")

    if rows:
        header = etree.SubElement(table_content, tag("row_group"))
        header.set("type", "header")
        row_el = etree.SubElement(header, tag("row"))
        for cell_text in data[0]:
            cell_el = etree.SubElement(row_el, tag("cell"))
            tc = etree.SubElement(cell_el, tag("text_content"))
            para = etree.SubElement(tc, tag("paragraph"))
            run = etree.SubElement(para, tag("run"))
            run.text = cell_text or ""
    if rows > 1:
        body = etree.SubElement(table_content, tag("row_group"))
        body.set("type", "body")
        for row in data[1:]:
            row_el = etree.SubElement(body, tag("row"))
            for i in range(cols):
                cell_el = etree.SubElement(row_el, tag("cell"))
                tc = etree.SubElement(cell_el, tag("text_content"))
                para = etree.SubElement(tc, tag("paragraph"))
                run = etree.SubElement(para, tag("run"))
                run.text = (row[i] if i < len(row) else "") or ""

    return region


def build_image_region(block: dict[str, Any], page_idx: int, region_idx: int, page_h: float, asset_id: str) -> etree._Element | None:
    """Build a DocIR image region from a PyMuPDF image block."""
    bbox = block.get("bbox")
    if not bbox:
        return None
    region = etree.Element(tag("region"))
    region.set("id", f"r{region_idx}_p{page_idx}")
    region.set("type", "image")
    region.set("native_label", "pymupdf_image_block")
    region.set("order", str(region_idx))
    add_bbox(region, tuple(bbox), page_h)
    add_provenance(region, "pymupdf-image-block")

    image_content = etree.SubElement(region, tag("image_content"))
    image_ref = etree.SubElement(image_content, tag("image_reference"))
    image_ref.set("asset_id", asset_id)
    image_ref.set("mime_type", f"image/{block.get('ext', 'png')}")
    vf = etree.SubElement(image_content, tag("visual_features"))
    etree.SubElement(vf, tag("contains_text")).text = "false"
    return region


def rect_overlap_ratio(a: tuple[float, float, float, float], b: tuple[float, float, float, float]) -> float:
    """Overlap area / area(a) for screen-coordinate rect tuples."""
    ax0, ay0, ax1, ay1 = a
    bx0, by0, bx1, by1 = b
    ix0, iy0 = max(ax0, bx0), max(ay0, by0)
    ix1, iy1 = min(ax1, bx1), min(ay1, by1)
    if ix1 <= ix0 or iy1 <= iy0:
        return 0.0
    inter = (ix1 - ix0) * (iy1 - iy0)
    area = max(1.0, (ax1 - ax0) * (ay1 - ay0))
    return inter / area


def build_docir(pdf_path: Path, output_path: Path, title: str | None = None) -> Path:
    doc = pymupdf.open(str(pdf_path))
    document = etree.Element(tag("document"), nsmap=NSMAP)
    document.set("version", "0.1.0")
    document.set("source_pdf", pdf_path.name)
    document.set("generated_at", datetime.now().isoformat())
    document.set("generator", "digital-ir-builder-v0.1.0")

    metadata = etree.SubElement(document, tag("metadata"))
    etree.SubElement(metadata, tag("title")).text = title or pdf_path.stem
    etree.SubElement(metadata, tag("page_count")).text = str(len(doc))
    default_size = etree.SubElement(metadata, tag("default_page_size"))
    if len(doc):
        default_size.set("width_pt", f"{doc[0].rect.width:.2f}")
        default_size.set("height_pt", f"{doc[0].rect.height:.2f}")
    pipe = etree.SubElement(metadata, tag("pipeline_info"))
    etree.SubElement(pipe, tag("layout_detector")).text = "PyMuPDF text blocks"
    etree.SubElement(pipe, tag("ocr_engine")).text = "PyMuPDF digital text extraction"
    etree.SubElement(pipe, tag("style_extractor")).text = "PyMuPDF span attributes"

    pages_el = etree.SubElement(document, tag("pages"))
    total_regions = 0
    assets: list[dict[str, str]] = []
    page_count = len(doc)
    for page_idx in range(page_count):
        page = doc[page_idx]
        page_el = etree.SubElement(pages_el, tag("page"))
        page_el.set("index", str(page_idx))
        size_el = etree.SubElement(page_el, tag("page_size"))
        size_el.set("width_pt", f"{page.rect.width:.2f}")
        size_el.set("height_pt", f"{page.rect.height:.2f}")
        regions_el = etree.SubElement(page_el, tag("regions"))

        text_dict: Any = page.get_text("dict")
        blocks = text_dict.get("blocks", []) if isinstance(text_dict, dict) else []
        region_items: list[tuple[tuple[float, float, float, float], etree._Element]] = []

        # Native PyMuPDF table detection catches lattice tables and some stream tables.
        table_bboxes: list[tuple[float, float, float, float]] = []
        try:
            table_finder = page.find_tables()
            tables = getattr(table_finder, "tables", None) or []
        except Exception:
            tables = []
        for table in tables:
            try:
                bbox = tuple(table.bbox)
                table_bboxes.append(bbox)
                region = build_table_region(table, page_idx, len(region_items), page.rect.height)
                region_items.append((bbox, region))
            except Exception:
                continue

        # Drawn text boxes: large stroked rectangles with text inside (not table grid lines).
        text_box_bboxes: list[tuple[float, float, float, float]] = []
        try:
            for drawing in page.get_drawings():
                rect = drawing.get("rect")
                if not rect:
                    continue
                if rect.width > 80 and rect.height > 30 and float(drawing.get("width") or 0) >= 1.0:
                    text_box_bboxes.append((rect.x0, rect.y0, rect.x1, rect.y1))
        except Exception:
            pass
        for box in text_box_bboxes:
            region = build_text_box_region(box, blocks, page_idx, len(region_items), page.rect.height)
            if region is not None:
                region_items.append((box, region))

        # Text and image blocks. Skip text blocks mostly covered by detected tables/text boxes.
        image_count = 0
        for block in blocks:
            btype = block.get("type")
            bbox = block.get("bbox")
            if not bbox:
                continue
            bbox_t = tuple(bbox)
            if btype == 0:
                if any(rect_overlap_ratio(bbox_t, tb) > 0.6 for tb in table_bboxes):
                    continue
                if any(rect_overlap_ratio(bbox_t, box) > 0.5 for box in text_box_bboxes):
                    continue
                region = build_text_region(block, page_idx, len(region_items), page.rect.height)
                if region is not None:
                    region_items.append((bbox_t, region))
            elif btype == 1:
                ext = block.get("ext", "png")
                asset_id = f"img_page{page_idx}_{image_count}"
                rel_path = f"assets/{asset_id}.{ext}"
                asset_path = output_path.parent / rel_path
                asset_path.parent.mkdir(parents=True, exist_ok=True)
                img_bytes = block.get("image")
                if img_bytes:
                    asset_path.write_bytes(img_bytes)
                    assets.append({
                        "id": asset_id,
                        "path": rel_path,
                        "mime_type": f"image/{ext}",
                        "width_px": str(block.get("width", "")),
                        "height_px": str(block.get("height", "")),
                    })
                    region = build_image_region(block, page_idx, len(region_items), page.rect.height, asset_id)
                    if region is not None:
                        region_items.append((bbox_t, region))
                        image_count += 1

        # Sort visually top-to-bottom, left-to-right and rewrite order/id sequence.
        region_items.sort(key=lambda item: (item[0][1], item[0][0]))
        for region_idx, (_bbox, region) in enumerate(region_items):
            region.set("id", f"r{region_idx}_p{page_idx}")
            region.set("order", str(region_idx))
            regions_el.append(region)
            total_regions += 1

    assets_el = etree.SubElement(document, tag("assets"))
    for asset in assets:
        asset_el = etree.SubElement(assets_el, tag("asset"))
        asset_el.set("id", asset["id"])
        asset_el.set("mime_type", asset["mime_type"])
        if asset.get("width_px"):
            asset_el.set("width_px", asset["width_px"])
        if asset.get("height_px"):
            asset_el.set("height_px", asset["height_px"])
        etree.SubElement(asset_el, tag("file_path")).text = asset["path"]
        etree.SubElement(asset_el, tag("extraction_source")).text = "pymupdf-image-block"
    etree.SubElement(document, tag("cross_page_hints"))
    output_path.parent.mkdir(parents=True, exist_ok=True)
    etree.ElementTree(document).write(str(output_path), encoding="UTF-8", xml_declaration=True, pretty_print=True)
    page_count = len(doc)
    doc.close()
    print(f"✓ Digital DocIR XML written to: {output_path}")
    print(f"  Pages: {page_count}")
    print(f"  Regions: {total_regions}")
    return output_path


def main() -> None:
    ap = argparse.ArgumentParser(description="Build DocIR directly from digital PDF text blocks")
    ap.add_argument("pdf", type=Path)
    ap.add_argument("-o", "--output", type=Path, required=True)
    ap.add_argument("--title")
    args = ap.parse_args()
    build_docir(args.pdf, args.output, args.title)


if __name__ == "__main__":
    main()
