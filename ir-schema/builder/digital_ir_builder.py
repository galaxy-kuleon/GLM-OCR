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


def span_is_visible(span: dict[str, Any]) -> bool:
    """Return False for invisible/transparent PDF text spans.

    Born-digital PDFs often contain hidden OCR/decoy layers. PyMuPDF exposes
    fully transparent text with alpha=0; rendering that into DOCX creates
    human-visible garbage even though it was invisible in the source PDF.
    """
    alpha = span.get("alpha")
    if alpha is not None:
        try:
            # PyMuPDF alpha is 0..255. Treat fully transparent and near-zero
            # antialiasing artifacts as invisible.
            if float(alpha) <= 1:
                return False
        except (TypeError, ValueError):
            pass
    return True


def visible_spans(line: dict[str, Any]) -> list[dict[str, Any]]:
    return [span for span in (line.get("spans") or []) if span_is_visible(span)]


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
        for span in visible_spans(line)
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
        # Preserve line-level runs/spans for formatting.
        spans = visible_spans(line)
        if not spans:
            continue
        para = etree.SubElement(text_content, tag("paragraph"))
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
            spans = visible_spans(line)
            if not spans:
                continue
            para = etree.SubElement(text_content, tag("paragraph"))
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
    if not text_content.findall(tag("paragraph")):
        return None
    return region


def best_span_for_cell(page: Any, cell_bbox: tuple[float, float, float, float], cell_text: str) -> dict[str, Any] | None:
    """Find the most relevant PyMuPDF text span for a table cell.

    PyMuPDF's table.extract() gives clean cell text but drops font/color/style.
    We recover a best-effort style hint from spans that geometrically overlap
    the source cell, preserving obvious human-visible styling like red italic
    table text.
    """
    if page is None or not cell_text:
        return None
    try:
        text_dict = page.get_text("dict")
    except Exception:
        return None
    best: tuple[float, dict[str, Any]] | None = None
    for block in text_dict.get("blocks", []) if isinstance(text_dict, dict) else []:
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []) or []:
            for span in visible_spans(line):
                span_text = (span.get("text") or "").strip()
                if not span_text:
                    continue
                span_bbox = span.get("bbox")
                if not span_bbox or len(span_bbox) != 4:
                    continue
                span_rect = (float(span_bbox[0]), float(span_bbox[1]), float(span_bbox[2]), float(span_bbox[3]))
                overlap = rect_overlap_ratio(span_rect, cell_bbox)
                # Header cells can be represented as one long span crossing
                # multiple cells; allow textual overlap as a weak signal.
                if overlap <= 0 and span_text not in cell_text and cell_text not in span_text:
                    continue
                score = overlap
                if span_text in cell_text or cell_text in span_text:
                    score += 0.25
                if best is None or score > best[0]:
                    best = (score, span)
    return best[1] if best is not None else None


def apply_span_style(run: etree._Element, span: dict[str, Any] | None) -> None:
    if not span:
        return
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


def add_cell_text(cell_el: etree._Element, cell_text: str, style_span: dict[str, Any] | None = None) -> None:
    tc = etree.SubElement(cell_el, tag("text_content"))
    para = etree.SubElement(tc, tag("paragraph"))
    run = etree.SubElement(para, tag("run"))
    # Keep embedded newlines inside one paragraph. python-docx converts them
    # into line breaks, avoiding the extra paragraph spacing that otherwise
    # clips multi-line table cells when row heights match the source PDF.
    run.text = str(cell_text or "")
    apply_span_style(run, style_span)


def table_row_heights(table: Any, rows: int) -> list[float]:
    cells = [c for c in (getattr(table, "cells", None) or []) if c]
    if not cells or rows <= 0:
        x0, y0, x1, y1 = tuple(table.bbox)
        return [(y1 - y0) / max(1, rows)] * max(0, rows)
    y_edges = sorted({round(float(c[1]), 2) for c in cells} | {round(float(c[3]), 2) for c in cells})
    heights = [max(1.0, y_edges[i + 1] - y_edges[i]) for i in range(len(y_edges) - 1)]
    if len(heights) != rows:
        x0, y0, x1, y1 = tuple(table.bbox)
        return [(y1 - y0) / max(1, rows)] * rows
    return heights


def cell_bbox_at(table: Any, row_idx: int, col_idx: int, rows: int, cols: int) -> tuple[float, float, float, float] | None:
    cells = getattr(table, "cells", None) or []
    # PyMuPDF currently exposes cells column-major for this table API.
    idx = col_idx * rows + row_idx
    if 0 <= idx < len(cells) and cells[idx]:
        return tuple(cells[idx])
    return None


def build_table_region(table: Any, page_idx: int, region_idx: int, page_h: float, page: Any | None = None) -> etree._Element:
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
    row_heights = table_row_heights(table, rows)
    table_content = etree.SubElement(region, tag("table_content"))
    table_content.set("rows", str(rows))
    table_content.set("cols", str(cols))
    table_style = etree.SubElement(table_content, tag("table_style"))
    table_style.set("border_visible", "true")
    table_style.set("border_color", "#000000")
    if rows > 1:
        table_style.set("header_row", "true")

    def add_row(parent: etree._Element, row_data: list[str], row_idx: int) -> None:
        row_el = etree.SubElement(parent, tag("row"))
        if row_idx < len(row_heights):
            row_el.set("height_pt", f"{row_heights[row_idx]:.2f}")
        for col_idx in range(cols):
            cell_text = (row_data[col_idx] if col_idx < len(row_data) else "") or ""
            cell_el = etree.SubElement(row_el, tag("cell"))
            cell_bbox = cell_bbox_at(table, row_idx, col_idx, rows, cols)
            if cell_bbox is not None:
                x0, y0, x1, y1 = cell_bbox
                cell_el.set("width_pt", f"{max(1.0, x1 - x0):.2f}")
                cell_el.set("height_pt", f"{max(1.0, y1 - y0):.2f}")
            style_span = best_span_for_cell(page, cell_bbox, cell_text) if cell_bbox is not None else None
            add_cell_text(cell_el, cell_text, style_span)

    if rows:
        header = etree.SubElement(table_content, tag("row_group"))
        header.set("type", "header")
        add_row(header, data[0], 0)
    if rows > 1:
        body = etree.SubElement(table_content, tag("row_group"))
        body.set("type", "body")
        for row_idx, row in enumerate(data[1:], start=1):
            add_row(body, row, row_idx)

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
                region = build_table_region(table, page_idx, len(region_items), page.rect.height, page=page)
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
