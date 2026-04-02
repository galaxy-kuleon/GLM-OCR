#!/usr/bin/env python3
"""Generate XML DSL pages directly from parse_digital_pdf() output.

Bypasses VLM entirely for digital PDFs. Transforms the plain-data output
of parse_digital_pdf into XML DSL files consumable by dsl_to_docx.py.

Pure functions throughout — no classes, no mutation, no side effects
except the final write_dsl_files() at the system boundary.
"""

import argparse
import glob as _glob
import os
import sys
import xml.etree.ElementTree as ET
from pathlib import Path


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

# Font family mapping: pdftohtml family names → standard DOCX font names
# Keys are lowercase for case-insensitive lookup.
FONT_FAMILY_TO_LATIN: dict[str, str] = {
    "liberationsans": "Arial",
    "liberation sans": "Arial",
    "arial": "Arial",
    "helvetica": "Arial",
    "timesnewroman": "Times New Roman",
    "times": "Times New Roman",
    "liberationserif": "Times New Roman",
    "liberation serif": "Times New Roman",
    "couriernew": "Courier New",
    "courier": "Courier New",
    "liberationmono": "Courier New",
    "liberation mono": "Courier New",
}

FONT_FAMILY_TO_CJK: dict[str, str] = {
    "notosanscjksc": "SimHei",
    "notosanscjktc": "SimHei",
    "notosanscjk": "SimHei",
    "noto sans cjk sc": "SimHei",
    "microsoftyahei": "SimHei",
    "simhei": "SimHei",
    "simsun": "SimSun",
    "notoserif": "SimSun",
    "notoserif cjk sc": "SimSun",
    "songti sc": "SimSun",
    "stfangsong": "SimSun",
}

DEFAULT_FONT_LATIN = "Arial"
DEFAULT_FONT_CJK = "SimSun"
DEFAULT_MARGIN_CM = "1.27"


# ---------------------------------------------------------------------------
# Font detection — pure functions
# ---------------------------------------------------------------------------

def detect_page_fonts(fontspecs: list[dict]) -> tuple[str, str]:
    """Detect the dominant Latin and CJK fonts from a page's fontspec list.

    Returns (latin_font, cjk_font) as standard DOCX font names.
    Pure function.
    """
    if not fontspecs:
        return DEFAULT_FONT_LATIN, DEFAULT_FONT_CJK

    # Collect all families
    families = [fs.get("family", "") for fs in fontspecs]

    latin_font = _detect_font_from_families(families, FONT_FAMILY_TO_LATIN, DEFAULT_FONT_LATIN)
    cjk_font = _detect_font_from_families(families, FONT_FAMILY_TO_CJK, DEFAULT_FONT_CJK)

    return latin_font, cjk_font


def _detect_font_from_families(
    families: list[str],
    mapping: dict[str, str],
    default: str,
) -> str:
    """Find the first matching font family in the mapping.

    Tries exact lowercase match first. Pure function.
    """
    for family in families:
        if family is None:
            continue
        key = family.lower().replace(" ", "")
        if key in mapping:
            return mapping[key]
        # Also try with spaces preserved
        key_spaced = family.lower()
        if key_spaced in mapping:
            return mapping[key_spaced]
    return default


# ---------------------------------------------------------------------------
# Color, font resolution, and font size — pure functions
# ---------------------------------------------------------------------------

def hex_color_to_rgb(hex_color: str) -> str | None:
    """Convert '#RRGGBB' hex color to 'R,G,B' string for DSL color-rgb attribute.

    Returns None for black (#000000) — the default, no need to emit.
    Returns None for white (#ffffff) — typically invisible text (table backgrounds).
    Returns None for empty/invalid input.
    Pure function.
    """
    if not hex_color or not hex_color.startswith("#") or len(hex_color) != 7:
        return None

    hex_color_lower = hex_color.lower()

    # Skip black (default) and white (invisible text)
    if hex_color_lower in ("#000000", "#ffffff"):
        return None

    try:
        r = int(hex_color[1:3], 16)
        g = int(hex_color[3:5], 16)
        b = int(hex_color[5:7], 16)
        return f"{r},{g},{b}"
    except ValueError:
        return None


def resolve_font_name(font_family: str, page_latin: str, page_cjk: str) -> str | None:
    """Resolve paragraph font family to a DOCX font name override.

    Returns a font name string if it differs from the page default,
    or None if it matches (meaning no override needed).
    Pure function.
    """
    if not font_family:
        return None

    # Try Latin mapping first
    key = font_family.lower().replace(" ", "")
    mapped = FONT_FAMILY_TO_LATIN.get(key)
    if mapped is None:
        # Try with spaces preserved
        mapped = FONT_FAMILY_TO_LATIN.get(font_family.lower())
    if mapped is None:
        # Try CJK mapping
        mapped = FONT_FAMILY_TO_CJK.get(key)
    if mapped is None:
        mapped = FONT_FAMILY_TO_CJK.get(font_family.lower())

    if mapped is None:
        return None  # Unknown font — use page default

    # Only emit override if different from page default
    if mapped == page_latin or mapped == page_cjk:
        return None

    return mapped


def format_font_size(size_pts: float) -> str:
    """Format font size for XML attribute.

    Always rounds to nearest integer.
    e.g., 11.0 → "11", 11.3 → "11", 9.7 → "10", 22.0 → "22"

    The existing DSL uses integer font sizes (see page-1.xml reference).
    """
    rounded = round(size_pts)
    return str(rounded)


# ---------------------------------------------------------------------------
# PDF image map — pure functions (adapted from deterministic_merge.py)
# ---------------------------------------------------------------------------

def build_pdf_images_map(workspace: str, total_pages: int) -> dict[int, list[tuple[str, int, int]]]:
    """Build a map of PDF-extracted images, filtering smasks and repeating images.

    Scans workspace/pdf-images/ directory. Filters out:
    1. Smask (alpha mask) files (grayscale mode 'L')
    2. Repeating images appearing on 3+ pages (headers, footers, watermarks)
       detected by file-size + dimensions fingerprint.

    Returns: {page_num: [(relative_path, pixel_width, pixel_height), ...]}
    Relative paths are relative to workspace (e.g., "pdf-images/img-001-002.png").

    Pure function — reads filesystem but produces no side effects.
    """
    pdf_dir = Path(workspace) / "pdf-images"
    if not pdf_dir.exists():
        return {}

    from PIL import Image as _PILImage

    # Collect all non-smask images grouped by page
    all_images: dict[int, list[tuple[str, int, int, str]]] = {}
    for page_num in range(1, total_pages + 1):
        pattern = str(pdf_dir / f"img-{page_num:03d}-*.png")
        candidates = sorted(_glob.glob(pattern))
        page_imgs: list[tuple[str, int, int, str]] = []
        for path in candidates:
            try:
                with _PILImage.open(path) as img:
                    if img.mode != "L":  # skip smask (alpha channel)
                        fsize = os.path.getsize(path)
                        fhash = f"{fsize}_{img.width}x{img.height}"
                        rel_path = f"pdf-images/{Path(path).name}"
                        page_imgs.append((rel_path, img.width, img.height, fhash))
            except Exception:
                continue
        all_images[page_num] = page_imgs

    # Count how many pages each image fingerprint appears on
    hash_page_count: dict[str, set[int]] = {}
    for page_num, imgs in all_images.items():
        for _, _, _, fhash in imgs:
            if fhash not in hash_page_count:
                hash_page_count[fhash] = set()
            hash_page_count[fhash].add(page_num)

    # Filter: exclude images appearing on 3+ pages (repeating headers/footers)
    repeat_threshold = min(3, max(2, total_pages // 2))
    repeating_hashes = {h for h, pages in hash_page_count.items()
                        if len(pages) >= repeat_threshold}

    result: dict[int, list[tuple[str, int, int]]] = {}
    for page_num, imgs in all_images.items():
        filtered = [(p, w, h) for p, w, h, fhash in imgs
                     if fhash not in repeating_hashes]
        if filtered:
            result[page_num] = filtered

    return result


def bbox_pts_to_normalized(
    bbox_pts: dict, page_width_pts: float, page_height_pts: float
) -> str:
    """Convert bbox from pts {top, left, width, height} to normalized 'left,top,right,bottom' string.

    The DSL bbox format is normalized 0-1000 (left,top,right,bottom),
    matching dsl_to_docx.py's process_image which divides by 1000.0.
    Pure function.
    """
    # Guard against zero page dimensions (T2 evaluator finding)
    if page_width_pts <= 0 or page_height_pts <= 0:
        return "0,0,0,0"

    left_pts = bbox_pts.get("left", 0)
    top_pts = bbox_pts.get("top", 0)
    width_pts = bbox_pts.get("width", 0)
    height_pts = bbox_pts.get("height", 0)

    # Normalize to 0-1000 range
    left = round(left_pts / page_width_pts * 1000)
    top = round(top_pts / page_height_pts * 1000)
    right = round((left_pts + width_pts) / page_width_pts * 1000)
    bottom = round((top_pts + height_pts) / page_height_pts * 1000)

    return f"{left},{top},{right},{bottom}"


def match_images_to_pdf_images(
    pdftohtml_images: list[dict],
    pdf_images: list[tuple[str, int, int]],
) -> list[dict]:
    """Match pdftohtml images to pdf-images by aspect ratio.

    For each pdf-image, finds the best matching pdftohtml image by aspect
    ratio similarity. Returns a list of dicts with merged info:
    {src, bbox_pts, top} for each matched image.

    Unmatched pdf-images (those with no close aspect-ratio match) are dropped.
    Pure function.
    """
    if not pdf_images:
        return []

    # Build pdftohtml image aspect ratios
    pdftohtml_aspects = []
    for img in pdftohtml_images:
        w = max(img.get("width", 1), 1)
        h = max(img.get("height", 1), 1)
        pdftohtml_aspects.append(w / h)

    used_pdftohtml: set[int] = set()
    matched: list[dict] = []

    for pdf_path, pdf_w, pdf_h in pdf_images:
        if pdf_h == 0:
            continue
        pdf_aspect = pdf_w / pdf_h

        # Find best matching pdftohtml image by aspect ratio
        best_idx = -1
        best_score = float("inf")
        for i, pt_aspect in enumerate(pdftohtml_aspects):
            if i in used_pdftohtml:
                continue
            score = abs(pdf_aspect - pt_aspect) / max(pdf_aspect, pt_aspect, 0.01)
            if score < best_score:
                best_score = score
                best_idx = i

        if best_idx >= 0 and best_score < 0.5:
            used_pdftohtml.add(best_idx)
            pt_img = pdftohtml_images[best_idx]
            matched.append({
                "src": pdf_path,
                "bbox_pts": pt_img.get("bbox_pts", {}),
                "top": pt_img.get("top", 0),
            })

    return matched


def image_to_xml(image_data: dict, page_width_pts: float, page_height_pts: float) -> ET.Element:
    """Convert matched image data to an <image> XML element.

    image_data has: src, bbox_pts, top.
    Pure function.
    """
    elem = ET.Element("image")
    elem.set("src", image_data.get("src", ""))

    bbox_pts = image_data.get("bbox_pts", {})
    if bbox_pts:
        bbox_str = bbox_pts_to_normalized(bbox_pts, page_width_pts, page_height_pts)
        elem.set("bbox", bbox_str)

    elem.set("page-width-pts", str(int(page_width_pts)))

    return elem


# ---------------------------------------------------------------------------
# Table → XML element — pure functions
# ---------------------------------------------------------------------------

def col_boundaries_to_ratios(col_boundaries_px: list[int]) -> list[float]:
    """Convert pixel column boundaries to proportional width ratios.

    col_boundaries_px has N+1 entries: [left_0, left_1, ..., left_N-1, right_edge].
    Returns N ratios that sum to ~1.0.

    Example: [100, 200, 300, 500] → 3 columns, widths [100, 100, 200], total 400
             → ratios [0.25, 0.25, 0.5]

    Pure function.
    """
    if len(col_boundaries_px) < 2:
        return []

    total_width = col_boundaries_px[-1] - col_boundaries_px[0]
    if total_width <= 0:
        # Degenerate case: all boundaries at same position
        n = len(col_boundaries_px) - 1
        ratios = [round(1.0 / n, 4)] * n
    else:
        ratios = []
        for i in range(len(col_boundaries_px) - 1):
            col_width = col_boundaries_px[i + 1] - col_boundaries_px[i]
            ratios.append(round(col_width / total_width, 4))

    # Normalize: adjust last element so sum(ratios) == 1.0 exactly.
    # Independent rounding can drift by up to ~0.0003 per column.
    if ratios:
        drift = 1.0 - sum(ratios)
        ratios[-1] = round(ratios[-1] + drift, 4)

    return ratios


def table_to_xml(
    table: dict, page_width_pts: float, page_height_pts: float
) -> ET.Element:
    """Convert a parsed table dict to a <table> XML element.

    The table dict (from parse_digital_pdf) has:
      rows: list[list[dict]] — each cell has text, bold, font_size_pts, col
      col_count, row_count: int
      bbox_pts: dict with top, left, width, height
      col_boundaries_px: list[int] — N+1 boundary positions

    Output matches the DSL format consumed by dsl_to_docx.py process_table.
    Pure function.
    """
    row_count = table.get("row_count", 0)
    col_count = table.get("col_count", 0)
    rows = table.get("rows", [])

    elem = ET.Element("table")
    elem.set("rows", str(row_count))
    elem.set("cols", str(col_count))

    # bbox — normalized 0-1000 "left,top,right,bottom"
    bbox_pts = table.get("bbox_pts", {})
    if bbox_pts:
        bbox_str = bbox_pts_to_normalized(bbox_pts, page_width_pts, page_height_pts)
        elem.set("bbox", bbox_str)

    elem.set("page-width-pts", str(int(page_width_pts)))
    elem.set("border-style", "full")

    # col-widths — proportional ratios from pixel boundaries
    col_boundaries = table.get("col_boundaries_px", [])
    ratios = col_boundaries_to_ratios(col_boundaries)
    if ratios:
        col_widths_elem = ET.SubElement(elem, "col-widths")
        col_widths_elem.text = ",".join(str(r) for r in ratios)

    # Rows and cells
    for row_idx, row_cells in enumerate(rows):
        row_elem = ET.SubElement(elem, "row")
        row_elem.set("index", str(row_idx))

        # If first row and majority of column positions are bold, mark as header.
        # Guards against vacuous truth (empty row) and sparse rows (1 bold out of 5 cols).
        if row_idx == 0 and row_cells:
            bold_count = sum(1 for c in row_cells if c.get("bold", False))
            min_bold_needed = max(1, (col_count + 1) // 2)  # majority of columns
            if bold_count >= min_bold_needed:
                row_elem.set("is-header", "true")

        for cell in row_cells:
            cell_elem = ET.SubElement(row_elem, "cell")
            cell_elem.set("row", str(row_idx))
            cell_elem.set("col", str(cell.get("col", 0)))

            font_size = cell.get("font_size_pts", 9)
            cell_elem.set("font-size-pt", format_font_size(font_size))

            if cell.get("bold", False):
                cell_elem.set("bold", "true")
            if cell.get("italic", False):
                cell_elem.set("italic", "true")

            cell_elem.text = cell.get("text", "").strip()

    return elem


# ---------------------------------------------------------------------------
# Paragraph → XML element — pure functions
# ---------------------------------------------------------------------------

def paragraph_to_xml(para: dict) -> ET.Element | None:
    """Convert a single parsed paragraph dict to an XML Element.

    Returns None if the paragraph should be skipped (e.g., table cells, empty).
    Pure function — creates a new Element, does not mutate para.
    """
    para_type = para.get("type", "paragraph")

    # Skip table cell paragraphs — handled by table generation (T3).
    # Exception: footer/header content at table-row positions should still
    # render (the is_table_cell flag is a heuristic that can fire on footer
    # lines when multiple footer elements share the same vertical position).
    if para.get("is_table_cell", False) and para_type not in ("footer", "header"):
        return None

    text = para.get("text", "").strip()

    # Skip empty paragraphs
    if not text:
        return None

    # Map type to XML element
    if para_type == "heading":
        return _heading_to_xml(para)
    elif para_type in ("footer", "header"):
        return _footer_header_to_xml(para)
    else:
        # "paragraph" and anything else → <paragraph>
        return _paragraph_to_xml(para)


def _heading_to_xml(para: dict) -> ET.Element:
    """Convert a heading paragraph to <heading> XML element. Pure function."""
    level = para.get("heading_level", 2)
    alignment = para.get("alignment", "left")

    elem = ET.Element("heading")
    elem.set("level", str(level))
    elem.set("alignment", alignment)

    run = _make_run(para)
    elem.append(run)

    return elem


def _paragraph_to_xml(para: dict) -> ET.Element:
    """Convert a body paragraph to <paragraph> XML element. Pure function."""
    alignment = para.get("alignment", "left")

    elem = ET.Element("paragraph")
    elem.set("alignment", alignment)

    run = _make_run(para)
    elem.append(run)

    return elem


def _footer_header_to_xml(para: dict) -> ET.Element:
    """Convert footer/header to <paragraph> XML element.

    Footers and headers are rendered as regular paragraphs with their
    (typically small) font size preserved. Pure function.
    """
    alignment = para.get("alignment", "left")

    elem = ET.Element("paragraph")
    elem.set("alignment", alignment)

    run = _make_run(para)
    elem.append(run)

    return elem


def _make_run(para: dict) -> ET.Element:
    """Create a <run> element from paragraph data. Pure function."""
    run = ET.Element("run")

    # Font size
    font_size = para.get("font_size_pts", 11)
    run.set("font-size-pt", format_font_size(font_size))

    # Bold / italic — only set if true (matches existing DSL convention)
    if para.get("bold", False):
        run.set("bold", "true")
    if para.get("italic", False):
        run.set("italic", "true")

    # Color — skip black (default), skip white (invisible)
    color_rgb = hex_color_to_rgb(para.get("color", ""))
    if color_rgb:
        run.set("color-rgb", color_rgb)

    # Font name — only set if different from page default (caller passes context)
    font_name_override = para.get("_font_name_override")
    if font_name_override:
        run.set("font-name", font_name_override)

    # Text content
    run.text = para.get("text", "").strip()

    return run


# ---------------------------------------------------------------------------
# Page → XML — pure function
# ---------------------------------------------------------------------------

def generate_page_dsl(
    page_data: dict,
    page_images: list[tuple[str, int, int]] | None = None,
) -> str:
    """Generate XML DSL string for one page.

    Takes a page dict (from parse_digital_pdf output) and an optional list
    of filtered pdf-images for this page. Returns a complete XML string.

    Elements (paragraphs, tables, images) are interleaved by vertical position
    (top coordinate) to preserve document reading order. Table-cell paragraphs
    are filtered out when tables are present (rendered as table cells instead).

    Pure function — no side effects.
    """
    page_num = page_data.get("number", 1)
    width_pts = page_data.get("width_pts", 612.0)
    height_pts = page_data.get("height_pts", 792.0)
    fontspecs = page_data.get("fontspecs", [])
    paragraphs = page_data.get("paragraphs", [])
    tables = page_data.get("tables", [])
    pdftohtml_images = page_data.get("images", [])

    # Detect fonts from fontspecs
    latin_font, cjk_font = detect_page_fonts(fontspecs)

    # Create page root element
    page = ET.Element("page")
    page.set("number", str(page_num))
    page.set("width-pts", str(int(width_pts)) if width_pts == int(width_pts) else str(width_pts))
    page.set("height-pts", str(int(height_pts)) if height_pts == int(height_pts) else str(height_pts))
    page.set("margin-top-cm", DEFAULT_MARGIN_CM)
    page.set("margin-bottom-cm", DEFAULT_MARGIN_CM)
    page.set("margin-left-cm", DEFAULT_MARGIN_CM)
    page.set("margin-right-cm", DEFAULT_MARGIN_CM)
    page.set("font-latin", latin_font)
    page.set("font-cjk", cjk_font)

    # --- Build sortable elements: list of (top_position, xml_element) ---
    elements: list[tuple[float, ET.Element]] = []

    # Paragraphs → XML, enriched with font-name override
    for para in paragraphs:
        font_override = resolve_font_name(
            para.get("font_family", ""), latin_font, cjk_font
        )
        enriched_para = {**para, "_font_name_override": font_override} if font_override else para

        xml_elem = paragraph_to_xml(enriched_para)
        if xml_elem is not None:
            # Use paragraph top position (in px) for ordering
            top_pos = float(para.get("top", 0))
            elements.append((top_pos, xml_elem))

    # Tables → XML, positioned by bbox_px top coordinate
    for tbl in tables:
        xml_elem = table_to_xml(tbl, width_pts, height_pts)
        # Use bbox_px top (in px) for ordering — consistent with paragraph/image px coords.
        # Fallback: bbox_pts top if bbox_px is missing (defensive guard).
        bbox_px = tbl.get("bbox_px") or {}
        if "top" in bbox_px:
            top_pos = float(bbox_px["top"])
        else:
            # Fallback: bbox_pts top if bbox_px is missing or lacks "top"
            top_pos = float(tbl.get("bbox_pts", {}).get("top", 0))
        elements.append((top_pos, xml_elem))

    # Images → XML (matched to pdf-images, filtered for repeating)
    if page_images:
        matched = match_images_to_pdf_images(pdftohtml_images, page_images)
        for img_data in matched:
            xml_elem = image_to_xml(img_data, width_pts, height_pts)
            # Use pdftohtml top position (in px) for ordering
            top_pos = float(img_data.get("top", 0))
            elements.append((top_pos, xml_elem))

    # Sort all elements by vertical position (stable sort preserves order for same top)
    elements.sort(key=lambda pair: pair[0])

    for _, xml_elem in elements:
        page.append(xml_elem)

    # Pretty-print with indentation
    ET.indent(page, space="  ")

    return ET.tostring(page, encoding="unicode")


# ---------------------------------------------------------------------------
# All pages → disk — system boundary (side effects confined here)
# ---------------------------------------------------------------------------

def generate_all_dsl(parsed_data: dict, workspace: str) -> list[str]:
    """Generate DSL XML files for all pages and write to disk.

    Builds a pdf-images map from workspace/pdf-images/ to include
    non-repeating images in the DSL output.

    Side effects: creates directory, writes files, reads pdf-images/.
    Returns list of written file paths.
    """
    dsl_dir = Path(workspace) / "dsl"
    dsl_dir.mkdir(parents=True, exist_ok=True)

    pages = parsed_data.get("pages", [])
    total_pages = len(pages)

    # Build pdf-images map (filtered: no smasks, no repeating headers)
    pdf_images_map = build_pdf_images_map(workspace, total_pages)

    written_paths = []

    for page_data in pages:
        page_num = page_data.get("number", 1)
        page_images = pdf_images_map.get(page_num)
        xml_str = generate_page_dsl(page_data, page_images=page_images)
        out_path = dsl_dir / f"page-{page_num}.xml"
        out_path.write_text(xml_str, encoding="utf-8")
        written_paths.append(str(out_path))

    return written_paths


# ---------------------------------------------------------------------------
# CLI — system boundary
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Generate XML DSL from digital PDF structure (bypasses VLM)."
    )
    parser.add_argument("--pdf", required=True, help="Path to digital PDF file")
    parser.add_argument("--workspace", required=True, help="Output workspace directory")
    parser.add_argument("--dry-run", action="store_true",
                        help="Print DSL to stdout instead of writing files")
    args = parser.parse_args()

    pdf_path = args.pdf
    if not Path(pdf_path).exists():
        print(f"Error: {pdf_path} not found", file=sys.stderr)
        sys.exit(1)

    # Import here to keep module importable without parse_pdf_structure on sys.path
    sys.path.insert(0, str(Path(__file__).parent))
    from parse_pdf_structure import parse_digital_pdf

    print(f"Parsing {pdf_path}...", file=sys.stderr)
    parsed_data = parse_digital_pdf(pdf_path)

    page_count = len(parsed_data.get("pages", []))
    print(f"Parsed {page_count} pages", file=sys.stderr)

    if args.dry_run:
        # Build pdf-images map for dry-run too (workspace must exist for images)
        pdf_images_map = build_pdf_images_map(args.workspace, page_count)
        for page_data in parsed_data.get("pages", []):
            page_num = page_data.get("number", 1)
            page_images = pdf_images_map.get(page_num)
            xml_str = generate_page_dsl(page_data, page_images=page_images)
            print(xml_str)
            print()
    else:
        written = generate_all_dsl(parsed_data, args.workspace)
        for path in written:
            print(f"  Written: {path}", file=sys.stderr)
        print(f"Generated {len(written)} DSL files", file=sys.stderr)


if __name__ == "__main__":
    main()
