#!/usr/bin/env python3
"""Format parsed PDF structure into VLM-prompt-ready reference text.

Reads output of parse_digital_pdf() and produces per-page structured text blocks
suitable for injection into VLM prompts alongside page images.

Design: Pure functions, data in → text out. No classes, no mutation, no side effects.
"""

import argparse
import json
import sys
from pathlib import Path

# Import shared constant — single source of truth
sys.path.insert(0, str(Path(__file__).parent))
from parse_pdf_structure import PDFTOHTML_SCALE  # noqa: E402


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _px_to_pts(px: int | float) -> float:
    """Convert pdftohtml pixel units to PDF points."""
    return round(px / PDFTOHTML_SCALE, 1)


def _build_fontspec_lookup(fontspecs: list[dict]) -> dict[str, dict]:
    """Build fontspec lookup by id. Pure function."""
    return {fs["id"]: fs for fs in fontspecs}


# ---------------------------------------------------------------------------
# FONTS section
# ---------------------------------------------------------------------------

def format_fonts_section(
    paragraphs: list[dict],
    fontspec_by_id: dict[str, dict],
) -> str:
    """Format the FONTS section listing all fonts used on this page.

    Only includes fonts actually referenced by paragraphs (not phantom fontspecs).
    Each font shows: id, size in pts, color, and style (bold detected from paragraphs).

    Uses the cumulative fontspec lookup to resolve fonts declared on earlier pages.
    """
    if not paragraphs or not fontspec_by_id:
        return ""

    # Collect font_ids actually used by paragraphs
    used_font_ids = set()
    # Also track which font_ids appear in bold context
    bold_font_ids = set()
    for para in paragraphs:
        used_font_ids.add(para.get("font_id", ""))
        if para.get("bold"):
            bold_font_ids.add(para.get("font_id", ""))
        # Also check individual lines for mixed-style paragraphs
        for line in para.get("lines", []):
            used_font_ids.add(line.get("font_id", ""))
            if line.get("bold"):
                bold_font_ids.add(line.get("font_id", ""))

    # Sort by font id (numeric sort) for stable output
    sorted_font_ids = sorted(used_font_ids, key=lambda fid: int(fid) if fid.isdigit() else 999)

    lines = ["FONTS:"]
    for fid in sorted_font_ids:
        fs = fontspec_by_id.get(fid)
        if fs is None:
            continue
        size_pts = fs.get("size_pts", 0)
        color = fs.get("color", "#000000")
        style = "bold" if fid in bold_font_ids else "regular"
        lines.append(f"  id={fid}: {size_pts}pt {color} ({style})")

    return "\n".join(lines) if len(lines) > 1 else ""


# ---------------------------------------------------------------------------
# TEXT section
# ---------------------------------------------------------------------------

def _para_overlaps_any_table(para: dict, tables: list[dict]) -> bool:
    """Check if a paragraph's bounding box overlaps any table's bbox (in px).

    Uses px coordinates since paragraphs store positions in px.
    A paragraph overlaps a table if both vertical AND horizontal extents intersect.
    """
    para_top = para.get("top", 0)
    para_bottom = para_top + para.get("height", 0)
    para_left = para.get("left", 0)
    para_right = para_left + para.get("width", 0)

    for table in tables:
        bbox = table.get("bbox_px", {})
        table_top = bbox.get("top", 0)
        table_bottom = table_top + bbox.get("height", 0)
        table_left = bbox.get("left", 0)
        table_right = table_left + bbox.get("width", 0)
        # Both vertical AND horizontal overlap required (with small tolerance)
        vertical_overlap = para_top < table_bottom + 2 and para_bottom > table_top - 2
        horizontal_overlap = para_left < table_right + 2 and para_right > table_left - 2
        if vertical_overlap and horizontal_overlap:
            return True
    return False


def format_text_section(
    paragraphs: list[dict], fontspec_by_id: dict, tables: list[dict]
) -> tuple[list[str], list[str]]:
    """Format the TEXT and FOOTER sections with per-line position and content.

    Uses paragraph lines for positional accuracy, with paragraph-level type
    annotations (heading level, footer/header) for context.
    Positions are in pts (converted from px).

    Paragraphs that overlap detected table bboxes are excluded from TEXT
    (they appear in the TABLES section instead).

    Returns (text_lines, footer_lines) — each a list of formatted strings.
    """
    if not paragraphs:
        return [], []

    # Separate footer/header first — they always appear regardless of table overlap
    footer_paras = [p for p in paragraphs if p.get("type") in ("footer", "header")]

    # For body text: exclude table-cell paragraphs AND paragraphs overlapping table bboxes
    body_paras = [
        p for p in paragraphs
        if p.get("type") not in ("footer", "header")
        and not p.get("is_table_cell", False)
        and not _para_overlaps_any_table(p, tables)
    ]

    # Build content first, only prepend header if non-empty
    # (paragraphs may yield no lines after rstrip filtering)
    text_content = []
    for para in body_paras:
        text_content.extend(_format_paragraph_lines(para, fontspec_by_id))
    text_lines = (["TEXT (top_pts, left_pts → content):"] + text_content) if text_content else []

    footer_content = []
    for para in footer_paras:
        footer_content.extend(_format_paragraph_lines(para, fontspec_by_id))
    footer_lines = (["FOOTER:"] + footer_content) if footer_content else []

    return text_lines, footer_lines


def _format_paragraph_lines(para: dict, fontspec_by_id: dict) -> list[str]:
    """Format individual lines of a paragraph with position and style info.

    Returns list of formatted line strings.
    """
    result = []
    lines = para.get("lines", [])
    para_type = para.get("type", "paragraph")

    for line in lines:
        text = line.get("text", "").rstrip()
        # Skip empty text elements (pdftohtml artifact)
        if not text:
            continue

        top_pts = _px_to_pts(line.get("top", 0))
        left_pts = _px_to_pts(line.get("left", 0))

        # Get font info from the line's font_id
        fid = line.get("font_id", "")
        fs = fontspec_by_id.get(fid)
        size_pts = fs.get("size_pts", 0) if fs else para.get("font_size_pts", 0)
        color = fs.get("color", "#000000") if fs else para.get("color", "#000000")
        is_bold = line.get("bold", False)

        # Build style annotations
        style_parts = [f"{size_pts}pt"]
        if is_bold:
            style_parts.append("BOLD")
        if color and color != "#000000":
            style_parts.append(color)

        style_str = " ".join(style_parts)

        # Add heading annotation if applicable
        prefix = ""
        if para_type == "heading":
            level = para.get("heading_level", "?")
            prefix = f"[H{level}] "

        result.append(f'  ({top_pts},{left_pts}) {prefix}{style_str}: "{text}"')

    return result


# ---------------------------------------------------------------------------
# TABLES section
# ---------------------------------------------------------------------------

def format_tables_section(tables: list[dict]) -> str:
    """Format the TABLES section with position, dimensions, and content.

    Each table shows header + data rows with pipe-separated values.
    """
    if not tables:
        return ""

    lines = ["TABLES (detected from text grid analysis):"]

    for table in tables:
        bbox = table.get("bbox_pts", {})
        top = bbox.get("top", 0)
        left = bbox.get("left", 0)
        bottom = round(top + bbox.get("height", 0), 1)
        right = round(left + bbox.get("width", 0), 1)
        rows = table.get("rows", [])
        row_count = table.get("row_count", len(rows))
        col_count = table.get("col_count", 0)

        lines.append(
            f"  Table at ({top},{left})-({bottom},{right}): "
            f"~{row_count} rows × {col_count} cols"
        )

        for i, row in enumerate(rows):
            # Sort cells by column index
            sorted_cells = sorted(row, key=lambda c: c.get("col", 0))
            cell_texts = [c.get("text", "").strip() for c in sorted_cells]
            row_str = " | ".join(cell_texts)

            if i == 0:
                lines.append(f"    Header: {row_str}")
            else:
                lines.append(f"    Row {i}: {row_str}")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# IMAGES section
# ---------------------------------------------------------------------------

def format_document_order(
    paragraphs: list[dict],
    tables: list[dict],
    images: list[dict],
    fontspec_by_id: dict[str, dict],
) -> str:
    """Format the DOCUMENT ORDER section — explicit element ordering for VLM.

    Builds a numbered list of ALL elements on the page sorted by vertical
    position (top_pts), giving the VLM an unambiguous sequence to follow.

    This is the most important section for weak VLMs — it tells them exactly
    what order to output elements in.
    """
    # Collect all elements with their type, label, and vertical position
    items: list[tuple[float, float, str]] = []  # (top_pts, left_pts, description)

    for para in paragraphs:
        if para.get("is_table_cell", False):
            continue
        top_pts = _px_to_pts(para.get("top", 0))
        left_pts = _px_to_pts(para.get("left", 0))
        para_type = para.get("type", "paragraph")
        text = para.get("text", "").strip()
        if not text:
            continue

        # Truncate long text for readability (full text is in TEXT section)
        display_text = text[:80] + "..." if len(text) > 80 else text

        if para_type == "heading":
            level = para.get("heading_level", "?")
            items.append((top_pts, left_pts, f'[heading-{level}] "{display_text}"'))
        elif para_type in ("footer", "header"):
            items.append((top_pts, left_pts, f'[{para_type}] "{display_text}"'))
        else:
            # Get font info for context
            fid = para.get("font_id", "")
            fs = fontspec_by_id.get(fid)
            size = fs.get("size_pts", 0) if fs else para.get("font_size_pts", 0)
            bold = "bold " if para.get("bold") else ""
            items.append((top_pts, left_pts, f'[paragraph {bold}{size}pt] "{display_text}"'))

    for table in tables:
        bbox = table.get("bbox_pts", {})
        top_pts = bbox.get("top", 0)
        left_pts = bbox.get("left", 0)
        row_count = table.get("row_count", 0)
        col_count = table.get("col_count", 0)
        items.append((top_pts, left_pts, f"[table] {row_count} rows × {col_count} cols"))

    for img in images:
        bbox = img.get("bbox_pts", {})
        top_pts = bbox.get("top", 0)
        left_pts = bbox.get("left", 0)
        w = bbox.get("width", 0)
        h = bbox.get("height", 0)
        items.append((top_pts, left_pts, f"[image] {w}×{h} pts"))

    if not items:
        return ""

    # Sort by vertical position (top), then horizontal (left)
    items.sort(key=lambda x: (x[0], x[1]))

    lines = ["DOCUMENT ORDER (output elements in this EXACT sequence):"]
    for i, (top, left, desc) in enumerate(items, 1):
        lines.append(f"  {i}. {desc}")

    return "\n".join(lines)


def format_images_section(images: list[dict]) -> str:
    """Format the IMAGES section with position and dimensions.

    Each image shows its top-left position and size in pts.
    """
    if not images:
        return ""

    lines = ["IMAGES:"]
    for img in images:
        bbox = img.get("bbox_pts", {})
        top = bbox.get("top", 0)
        left = bbox.get("left", 0)
        width = bbox.get("width", 0)
        height = bbox.get("height", 0)
        lines.append(f"  ({top},{left}) {width}×{height} pts — embedded image")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Page-level composition
# ---------------------------------------------------------------------------

def format_page_reference(
    page_data: dict,
    cumulative_fontspec_by_id: dict[str, dict] | None = None,
) -> str:
    """Format complete reference text for one page.

    Composes: fonts + text + tables + images + footer sections.
    Returns a single multi-line string ready for VLM prompt injection.

    cumulative_fontspec_by_id: optional pre-built lookup covering all pages.
    If None, builds from this page's fontspecs only (works for single-page use).
    """
    page_num = page_data.get("number", "?")
    width_pts = page_data.get("width_pts", 0)
    height_pts = page_data.get("height_pts", 0)

    fontspecs = page_data.get("fontspecs", [])
    paragraphs = page_data.get("paragraphs", [])
    tables = page_data.get("tables", [])
    images = page_data.get("images", [])

    fontspec_by_id = cumulative_fontspec_by_id or _build_fontspec_lookup(fontspecs)

    # Build sections
    header = f"[EXACT PDF STRUCTURE — Page {page_num}]"
    page_size = f"Page: {width_pts} × {height_pts} pts"
    fonts_section = format_fonts_section(paragraphs, fontspec_by_id)
    order_section = format_document_order(paragraphs, tables, images, fontspec_by_id)
    text_lines, footer_lines = format_text_section(paragraphs, fontspec_by_id, tables)
    tables_section = format_tables_section(tables)
    images_section = format_images_section(images)

    # Compose all non-empty sections — DOCUMENT ORDER first for VLM emphasis
    sections = [header, page_size]

    if order_section:
        sections.append(order_section)

    if fonts_section:
        sections.append(fonts_section)

    if text_lines:
        sections.append("\n".join(text_lines))

    if tables_section:
        sections.append(tables_section)

    if images_section:
        sections.append(images_section)

    if footer_lines:
        sections.append("\n".join(footer_lines))

    return "\n\n".join(sections)


# ---------------------------------------------------------------------------
# All-pages composition
# ---------------------------------------------------------------------------

def format_all_pages(parsed_data: dict) -> dict[int, str]:
    """Format reference text for all pages.

    Takes the output of parse_digital_pdf() and returns a dict
    mapping page number to formatted reference text.

    Builds a cumulative fontspec lookup across all pages so that font references
    declared on page 1 but used on page 3 still resolve correctly.
    """
    # Build cumulative fontspec lookup (fontspecs may be declared on any page)
    cumulative_fontspecs: dict[str, dict] = {}
    for page in parsed_data.get("pages", []):
        for fs in page.get("fontspecs", []):
            cumulative_fontspecs[fs["id"]] = fs

    result = {}
    for page in parsed_data.get("pages", []):
        page_num = page.get("number", 0)
        result[page_num] = format_page_reference(page, cumulative_fontspecs)

    return result


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Format parsed PDF structure into VLM-prompt-ready reference text."
    )
    source = parser.add_mutually_exclusive_group(required=True)
    source.add_argument("--pdf-structure", type=str,
                        help="Path to JSON file from parse_pdf_structure.py")
    source.add_argument("--pdf", type=str,
                        help="Path to PDF file (will run parse_digital_pdf internally)")
    parser.add_argument("--page", type=int,
                        help="Output only this page number (default: all pages)")
    parser.add_argument("--output", type=str,
                        help="Output file path (default: stdout)")
    args = parser.parse_args()

    # Load parsed data
    if args.pdf_structure:
        json_path = Path(args.pdf_structure)
        if not json_path.exists():
            print(f"Error: {json_path} not found", file=sys.stderr)
            sys.exit(1)
        parsed_data = json.loads(json_path.read_text(encoding="utf-8"))
    else:
        pdf_path = args.pdf
        if not Path(pdf_path).exists():
            print(f"Error: {pdf_path} not found", file=sys.stderr)
            sys.exit(1)
        # Import and run parser
        script_dir = Path(__file__).parent
        sys.path.insert(0, str(script_dir))
        from parse_pdf_structure import parse_digital_pdf
        parsed_data = parse_digital_pdf(pdf_path)

    # Format
    all_pages = format_all_pages(parsed_data)

    # Output
    if args.page is not None:
        if args.page not in all_pages:
            print(f"Error: page {args.page} not found (available: {sorted(all_pages.keys())})",
                  file=sys.stderr)
            sys.exit(1)
        output_text = all_pages[args.page]
    else:
        output_text = "\n\n".join(all_pages[k] for k in sorted(all_pages.keys()))

    if args.output:
        Path(args.output).write_text(output_text, encoding="utf-8")
        print(f"Written to {args.output}", file=sys.stderr)
    else:
        print(output_text)


if __name__ == "__main__":
    main()
