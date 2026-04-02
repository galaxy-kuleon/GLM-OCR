#!/usr/bin/env python3
"""Parse digital PDF structure using poppler tools (pdftohtml, pdffonts, pdftotext).

Returns plain data (dicts/lists) — no classes, no side effects beyond subprocess calls.
Designed for composition: each function takes input and returns data.
"""

import argparse
import html
import json
import re
import subprocess
import sys
from pathlib import Path


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

PDFTOHTML_SCALE = 1.5  # pdftohtml outputs at 1.5x PDF pts
SUBSET_PREFIX_RE = re.compile(r"^[A-Z]{6}\+")
DIGITAL_MIN_CHARS_PER_PAGE = 50  # non-whitespace chars; language-agnostic (CJK-safe)

# Paragraph grouping thresholds
PARA_VERTICAL_GAP_FACTOR = 1.3  # max top-diff / line-height ratio for same paragraph
PARA_LEFT_TOLERANCE_PX = 30     # max left-margin difference for same-paragraph lines
TABLE_ROW_MIN_ELEMENTS = 3      # min elements at same top to be considered a table row
TABLE_ROW_MAX_ELEMENTS = 8      # max elements — beyond this, likely CJK character fragments
                                # KNOWN LIMITATION: tables with 9+ columns will not be detected.
                                # This ceiling is required to reject CJK fragment false positives.
TABLE_ROW_MEDIAN_WIDTH_MIN_PX = 30  # min median element width — CJK fragments are ~17px
TABLE_ROW_COVERAGE_MAX = 0.85   # max width coverage ratio — paragraph text tiles ~1.0, table rows ~0.3-0.7
TABLE_ROW_TOP_TOLERANCE_PX = 5  # max top difference to be considered same row
LIST_ITEM_RE = re.compile(r"^\s*(\d+[\.\)]\s|[-•–—]\s)")  # numbered or bulleted list
ALIGNMENT_TOLERANCE_PX = 15     # tolerance for alignment detection

# Heading classification thresholds (in pts)
HEADING_LEVEL_1_MIN_PTS = 18    # size_pts >= 18
HEADING_LEVEL_2_MIN_PTS = 14    # size_pts >= 14
HEADING_LEVEL_3_MIN_PTS = 12    # size_pts >= 12
HEADING_LEVEL_4_MIN_PTS = 11    # size_pts >= 11 (bold only)

# Heading content filter — rejects trivially short text fragments (e.g., "月", "___")
# that pdftohtml produces for date fields and form placeholders at heading font sizes.
# After stripping whitespace and underscores, text must have >= this many characters.
HEADING_MIN_CONTENT_CHARS = 2

# Header/footer detection thresholds
FOOTER_TOP_RATIO = 0.92         # top_pts > page_height_pts * 0.92
HEADER_BOTTOM_RATIO = 0.08      # top_pts < page_height_pts * 0.08
HEADER_FOOTER_MAX_FONT_PTS = 10 # font_size_pts <= 10


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _run_cmd(cmd: list[str]) -> subprocess.CompletedProcess:
    """Run a subprocess, returning the CompletedProcess. Caller decides what to do with it."""
    return subprocess.run(cmd, capture_output=True, text=True, timeout=30)


def _clean_font_family(raw_family: str) -> str:
    """Strip subset prefix (e.g., 'BAAAAA+LiberationSans' -> 'LiberationSans').
    Also strip style suffix for the family name (e.g., '-Bold', '-Italic', '-BoldItalic')."""
    family = SUBSET_PREFIX_RE.sub("", raw_family)
    # Remove style suffixes to get the base family
    for suffix in ("-BoldItalic", "-BoldOblique", "-Bold", "-Italic", "-Oblique", "-Regular"):
        if family.endswith(suffix):
            return family[: -len(suffix)]
    return family


def _px_to_pts(px: int | float) -> float:
    """Convert pdftohtml pixel units to PDF points."""
    return round(px / PDFTOHTML_SCALE, 1)


# ---------------------------------------------------------------------------
# parse_pdffonts
# ---------------------------------------------------------------------------

def parse_pdffonts(pdf_path: str) -> list[dict]:
    """Run pdffonts and parse the tabular output into a list of font dicts.

    Returns list of:
      {"name": str, "type": str, "encoding": str,
       "embedded": bool, "subset": bool, "unicode": bool}
    """
    result = _run_cmd(["pdffonts", pdf_path])
    if result.returncode != 0:
        return []

    lines = result.stdout.strip().splitlines()
    if len(lines) < 2:
        return []

    # The second line is dashes — use it to determine column boundaries
    dash_line = lines[1]
    # Find column boundaries: each column is a run of dashes separated by spaces
    col_ranges = []
    in_dash = False
    start = 0
    for i, ch in enumerate(dash_line):
        if ch == "-" and not in_dash:
            start = i
            in_dash = True
        elif ch == " " and in_dash:
            col_ranges.append((start, i))
            in_dash = False
    if in_dash:
        col_ranges.append((start, len(dash_line)))

    fonts = []
    for line in lines[2:]:
        if not line.strip():
            continue
        cols = []
        for s, e in col_ranges:
            cols.append(line[s:e].strip() if s < len(line) else "")

        # Expect at least 7 columns: name, type, encoding, emb, sub, uni, object ID
        if len(cols) < 6:
            continue

        fonts.append({
            "name": cols[0],
            "type": cols[1],
            "encoding": cols[2],
            "embedded": cols[3] == "yes",
            "subset": cols[4] == "yes",
            "unicode": cols[5] == "yes",
        })

    return fonts


# ---------------------------------------------------------------------------
# is_digital_pdf
# ---------------------------------------------------------------------------

def is_digital_pdf(pdf_path: str) -> bool:
    """Detect whether a PDF is digital (has embedded fonts AND sufficient text).

    Digital = has embedded fonts AND non_whitespace_chars / page_count >= DIGITAL_MIN_CHARS_PER_PAGE.
    Uses character count (not word count) for CJK language compatibility.
    """
    # Check fonts
    fonts = parse_pdffonts(pdf_path)
    has_embedded_fonts = any(f["embedded"] for f in fonts)
    if not has_embedded_fonts:
        return False

    # Count non-whitespace characters (language-agnostic: works for CJK, Latin, etc.)
    result = _run_cmd(["pdftotext", pdf_path, "-"])
    if result.returncode != 0:
        return False
    char_count = sum(1 for ch in result.stdout if not ch.isspace())

    # Count pages (from pdftohtml xml, just count <page> tags)
    xml_result = _run_cmd(["pdftohtml", "-xml", "-stdout", pdf_path])
    if xml_result.returncode != 0:
        return False
    page_count = xml_result.stdout.count("<page ")
    if page_count == 0:
        return False

    chars_per_page = char_count / page_count
    return chars_per_page >= DIGITAL_MIN_CHARS_PER_PAGE


# ---------------------------------------------------------------------------
# parse_pdftohtml_xml
# ---------------------------------------------------------------------------

def parse_pdftohtml_xml(pdf_path: str) -> dict:
    """Run pdftohtml -xml and parse into structured dict.

    Returns:
      {"pages": [
        {"number": int, "width_px": int, "height_px": int,
         "width_pts": float, "height_pts": float,
         "fontspecs": [{"id": str, "size_px": int, "size_pts": float,
                        "family": str, "raw_family": str, "color": str}],
         "text_elements": [{"top": int, "left": int, "width": int, "height": int,
                            "font_id": str, "text": str, "bold": bool, "italic": bool}],
         "images": [{"top": int, "left": int, "width": int, "height": int, "src": str}]
        }
      ]}
    """
    result = _run_cmd(["pdftohtml", "-xml", "-stdout", pdf_path])
    if result.returncode != 0:
        return {"pages": []}

    return _parse_xml_string(result.stdout)


def _parse_xml_string(xml_str: str) -> dict:
    """Parse pdftohtml XML string into structured data. Pure function."""
    # pdftohtml XML may have HTML entities and inline markup (<b>, <i>)
    # ET can't handle <b>/<i> as they aren't self-closing or declared.
    # Strategy: parse page-by-page using string extraction + targeted ET parsing.

    pages = []
    # Global fontspec lookup — pdftohtml declares fontspecs on the first page
    # where they appear, but text elements on later pages reference them by id.
    global_fontspec_by_id: dict[str, dict] = {}

    # Split by page boundaries
    page_pattern = re.compile(
        r'<page\s+number="(\d+)"\s+position="absolute"\s+'
        r'top="(\d+)"\s+left="(\d+)"\s+height="(\d+)"\s+width="(\d+)">'
    )
    page_starts = list(page_pattern.finditer(xml_str))

    for i, match in enumerate(page_starts):
        page_num = int(match.group(1))
        height_px = int(match.group(4))
        width_px = int(match.group(5))

        # Extract page content between this <page> and the next (or </pdf2xml>)
        start = match.end()
        if i + 1 < len(page_starts):
            end = page_starts[i + 1].start()
        else:
            end = xml_str.find("</pdf2xml>", start)
            if end == -1:
                end = len(xml_str)

        page_content = xml_str[start:end]

        # Parse fontspecs declared on this page
        fontspecs = _parse_fontspecs(page_content)

        # Accumulate into global lookup (new fontspecs may appear on any page)
        for fs in fontspecs:
            global_fontspec_by_id[fs["id"]] = fs

        # Parse text elements using global fontspec lookup
        text_elements = _parse_text_elements(page_content, global_fontspec_by_id)

        # Parse images
        images = _parse_images(page_content)

        pages.append({
            "number": page_num,
            "width_px": width_px,
            "height_px": height_px,
            "width_pts": _px_to_pts(width_px),
            "height_pts": _px_to_pts(height_px),
            "fontspecs": fontspecs,
            "text_elements": text_elements,
            "images": images,
        })

    return {"pages": pages}


def _parse_fontspecs(page_content: str) -> list[dict]:
    """Extract fontspec elements from a page's XML content."""
    fontspec_re = re.compile(
        r'<fontspec\s+id="(\d+)"\s+size="(\d+)"\s+family="([^"]+)"\s+color="([^"]+)"\s*/>'
    )
    fontspecs = []
    for m in fontspec_re.finditer(page_content):
        raw_family = m.group(3)
        family = _clean_font_family(raw_family)
        size_px = int(m.group(2))

        # NOTE: Bold detection from font family names is unreliable because
        # pdftohtml strips style suffixes (e.g., "-Bold") from the family
        # attribute. Bold is reliably detected from <b> tags in text elements.
        fontspecs.append({
            "id": m.group(1),
            "size_px": size_px,
            "size_pts": _px_to_pts(size_px),
            "family": family,
            "raw_family": raw_family,
            "color": m.group(4),
        })
    return fontspecs


def _parse_text_elements(page_content: str, fontspec_by_id: dict) -> list[dict]:
    """Extract text elements from a page's XML content.

    Handles <b>...</b> and <i>...</i> inline markup.
    """
    text_re = re.compile(
        r'<text\s+top="(\d+)"\s+left="(\d+)"\s+width="(\d+)"\s+height="(\d+)"\s+'
        r'font="(\d+)">(.*?)</text>',
        re.DOTALL,
    )

    elements = []
    for m in text_re.finditer(page_content):
        font_id = m.group(5)
        raw_inner = m.group(6)

        # Detect bold/italic from inline markup
        has_b_tag = "<b>" in raw_inner
        has_i_tag = "<i>" in raw_inner

        # Strip inline markup to get plain text, then decode HTML entities
        text = raw_inner
        text = re.sub(r"</?b>", "", text)
        text = re.sub(r"</?i>", "", text)
        text = html.unescape(text)

        elements.append({
            "top": int(m.group(1)),
            "left": int(m.group(2)),
            "width": int(m.group(3)),
            "height": int(m.group(4)),
            "font_id": font_id,
            "text": text,
            "bold": has_b_tag,
            "italic": has_i_tag,
        })

    return elements


def _parse_images(page_content: str) -> list[dict]:
    """Extract image elements from a page's XML content."""
    img_re = re.compile(
        r'<image\s+top="(\d+)"\s+left="(\d+)"\s+width="(\d+)"\s+height="(\d+)"\s+'
        r'src="([^"]+)"\s*/>'
    )
    images = []
    for m in img_re.finditer(page_content):
        images.append({
            "top": int(m.group(1)),
            "left": int(m.group(2)),
            "width": int(m.group(3)),
            "height": int(m.group(4)),
            "src": m.group(5),
        })
    return images


# ---------------------------------------------------------------------------
# Paragraph grouping
# ---------------------------------------------------------------------------

def _median_width(elements: list[dict], min_top: int, max_top: int) -> float:
    """Compute the median width of elements in a top range. Pure function."""
    widths = sorted(e["width"] for e in elements if min_top <= e["top"] <= max_top)
    if not widths:
        return 0.0
    mid = len(widths) // 2
    if len(widths) % 2 == 1:
        return float(widths[mid])
    return (widths[mid - 1] + widths[mid]) / 2.0


def _coverage_ratio(elements: list[dict], min_top: int, max_top: int) -> float:
    """Compute width coverage ratio: total element width / horizontal span.

    Paragraph text tiles contiguously (~1.0 coverage).
    Table rows have gaps between cells (~0.3-0.7 coverage).
    Pure function.
    """
    row_elems = [e for e in elements if min_top <= e["top"] <= max_top]
    if not row_elems:
        return 0.0
    total_width = sum(e["width"] for e in row_elems)
    min_left = min(e["left"] for e in row_elems)
    max_right = max(e["left"] + e["width"] for e in row_elems)
    span = max_right - min_left
    if span <= 0:
        return 0.0
    return total_width / span


def _find_table_row_tops(elements: list[dict]) -> set[int]:
    """Identify top positions that correspond to table rows.

    Two-pass algorithm:
      Pass 1: A top-group is a table row if it has >= TABLE_ROW_MIN_ELEMENTS,
              <= TABLE_ROW_MAX_ELEMENTS, and median element width >= TABLE_ROW_MEDIAN_WIDTH_MIN_PX.
              The element count ceiling and median width filter reject CJK character
              fragments (pdftohtml splits CJK text into individual narrow elements).
      Pass 2: Expand — any top-group within a LOCAL expansion threshold of a
              confirmed table row that has >= 2 elements is also a table row.
              The expansion threshold is computed per-region (contiguous cluster
              of confirmed rows), NOT globally, to prevent distant false-positive
              rows from inflating the gap.

    Returns a set of top values that are table rows.
    """
    if not elements:
        return set()

    # Bucket elements by top position (with tolerance)
    tops = sorted(set(e["top"] for e in elements))
    top_groups: list[list[int]] = []
    for t in tops:
        if top_groups and t - top_groups[-1][-1] <= TABLE_ROW_TOP_TOLERANCE_PX:
            top_groups[-1].append(t)
        else:
            top_groups.append([t])

    # Count elements per group and find their representative top
    group_info = []  # (min_top, max_top, count, tops_in_group)
    for group in top_groups:
        min_t = min(group)
        max_t = max(group)
        count = sum(1 for e in elements if min_t <= e["top"] <= max_t)
        group_info.append((min_t, max_t, count, group))

    # Pass 1: definite table rows — with CJK fragment rejection
    table_tops: set[int] = set()
    confirmed_table_ranges: list[tuple[int, int]] = []
    rejected_tops: set[int] = set()  # tops explicitly rejected by CJK heuristics
    for min_t, max_t, count, group in group_info:
        if count < TABLE_ROW_MIN_ELEMENTS:
            continue
        # Reject: too many elements (CJK character fragmentation)
        if count > TABLE_ROW_MAX_ELEMENTS:
            rejected_tops.update(group)
            continue
        # Reject: median width too narrow (CJK character fragments are ~17px)
        med_w = _median_width(elements, min_t, max_t)
        if med_w < TABLE_ROW_MEDIAN_WIDTH_MIN_PX:
            rejected_tops.update(group)
            continue
        # Reject: high coverage ratio (paragraph text tiles ~1.0; table rows have gaps ~0.3-0.7)
        coverage = _coverage_ratio(elements, min_t, max_t)
        if coverage > TABLE_ROW_COVERAGE_MAX:
            rejected_tops.update(group)
            continue
        table_tops.update(group)
        confirmed_table_ranges.append((min_t, max_t))

    # Pass 2: expand to neighbors with >= 2 elements (catches wrapped table headers)
    # Uses LOCAL gap computation per contiguous region to prevent
    # distant rows from inflating expansion threshold.
    # Groups explicitly rejected in Pass 1 (CJK fragments) are never re-added.
    if confirmed_table_ranges:
        table_elements = [e for e in elements if e["top"] in table_tops]
        line_height = (
            max(e["height"] for e in table_elements) if table_elements else 20
        )

        # Cluster confirmed ranges into contiguous regions.
        # Two confirmed rows are in the same region if their gap <= 3x line_height.
        region_boundary = line_height * 3
        sorted_ranges = sorted(confirmed_table_ranges, key=lambda r: r[0])
        regions: list[list[tuple[int, int]]] = [[sorted_ranges[0]]]
        for rng in sorted_ranges[1:]:
            prev_rng = regions[-1][-1]
            if rng[0] - prev_rng[1] <= region_boundary:
                regions[-1].append(rng)
            else:
                regions.append([rng])

        # For each region, compute a local expansion threshold
        # and expand neighboring groups.
        for region in regions:
            region_tops = sorted(r[0] for r in region)
            if len(region_tops) >= 2:
                local_max_gap = max(
                    region_tops[i + 1] - region_tops[i]
                    for i in range(len(region_tops) - 1)
                )
            else:
                # Single confirmed row: allow broader expansion (2x line_height)
                # to reach nearby 2-element rows that form the rest of the table.
                local_max_gap = line_height * 2
            local_threshold = max(line_height, local_max_gap) * 1.2

            for min_t, max_t, count, group in group_info:
                if min_t in table_tops:
                    continue  # already confirmed
                if min_t in rejected_tops:
                    continue  # explicitly rejected by CJK heuristics — do not re-add
                if count < 2:
                    continue
                # Check if this group is near THIS region (not any distant region)
                for conf_min, conf_max in region:
                    if (abs(min_t - conf_max) <= local_threshold
                            or abs(conf_min - max_t) <= local_threshold):
                        table_tops.update(group)
                        break

    return table_tops


def _is_list_item(text: str) -> bool:
    """Return True if text looks like the start of a list item."""
    return bool(LIST_ITEM_RE.match(text))


def _detect_alignment(lines: list[dict], page_width_px: int) -> str:
    """Detect text alignment from a group of lines.

    Returns "left", "right", "center", or "justify".
    """
    if not lines or page_width_px <= 0:
        return "left"

    if len(lines) == 1:
        return "left"  # Can't determine alignment from a single line

    lefts = [ln["left"] for ln in lines]
    rights = [ln["left"] + ln["width"] for ln in lines]

    left_range = max(lefts) - min(lefts)
    right_range = max(rights) - min(rights)

    left_aligned = left_range <= ALIGNMENT_TOLERANCE_PX
    right_aligned = right_range <= ALIGNMENT_TOLERANCE_PX

    if left_aligned and right_aligned:
        return "justify"
    elif left_aligned:
        return "left"
    elif right_aligned:
        return "right"

    # Check center: each line's center should be similar
    centers = [(ln["left"] + ln["width"] / 2) for ln in lines]
    center_range = max(centers) - min(centers)
    if center_range <= ALIGNMENT_TOLERANCE_PX:
        return "center"

    return "left"  # default fallback


def _build_paragraph(lines: list[dict], page_width_px: int) -> dict:
    """Build a paragraph dict from a list of line elements.

    Pure function — takes lines and page width, returns paragraph data.
    """
    tops = [ln["top"] for ln in lines]
    lefts = [ln["left"] for ln in lines]
    rights = [ln["left"] + ln["width"] for ln in lines]
    bottoms = [ln["top"] + ln["height"] for ln in lines]

    # Dominant font_id: most frequent among lines
    font_counts: dict[str, int] = {}
    for ln in lines:
        fid = ln["font_id"]
        font_counts[fid] = font_counts.get(fid, 0) + 1
    dominant_font_id = max(font_counts, key=font_counts.get)

    # Bold/italic: True if any line has it
    has_bold = any(ln["bold"] for ln in lines)
    has_italic = any(ln["italic"] for ln in lines)

    # Merge text with space between lines
    merged_text = " ".join(ln["text"].rstrip() for ln in lines)

    alignment = _detect_alignment(lines, page_width_px)

    return {
        "type": "paragraph",
        "text": merged_text,
        "lines": lines,
        "top": min(tops),
        "left": min(lefts),
        "width": max(rights) - min(lefts),
        "height": max(bottoms) - min(tops),
        "font_id": dominant_font_id,
        "bold": has_bold,
        "italic": has_italic,
        "alignment": alignment,
    }


def group_text_into_paragraphs(page_data: dict) -> list[dict]:
    """Group a page's text elements into logical paragraphs.

    Takes a page dict (from parse_pdftohtml_xml) and returns a list of
    paragraph dicts. Pure function.

    Algorithm:
      1. Sort elements by (top, left)
      2. Identify table rows (>=3 elements at same top)
      3. Table-row elements become individual single-line paragraphs
      4. Non-table elements are merged greedily:
         - Consecutive lines join if vertically close, same font/style,
           and not a new list item
    """
    text_elements = page_data.get("text_elements", [])
    page_width_px = page_data.get("width_px", 0)

    if not text_elements:
        return []

    # Sort by vertical then horizontal position
    sorted_elems = sorted(text_elements, key=lambda e: (e["top"], e["left"]))

    # Identify table row tops
    table_tops = _find_table_row_tops(sorted_elems)

    # Separate into table-cell elements and flow elements
    table_elems = []
    flow_elems = []
    for e in sorted_elems:
        if e["top"] in table_tops:
            table_elems.append(e)
        else:
            flow_elems.append(e)

    paragraphs = []

    # Table cells: each becomes its own paragraph, tagged for downstream use
    for e in table_elems:
        para = _build_paragraph([e], page_width_px)
        para["is_table_cell"] = True
        paragraphs.append(para)

    # Flow elements: greedy merge
    if flow_elems:
        current_lines = [flow_elems[0]]

        for elem in flow_elems[1:]:
            prev = current_lines[-1]
            prev_bottom = prev["top"] + prev["height"]
            vertical_gap = elem["top"] - prev_bottom
            line_height = max(prev["height"], 1)  # avoid div by zero

            # Conditions to merge with current paragraph
            vertically_close = vertical_gap < line_height * PARA_VERTICAL_GAP_FACTOR
            same_font = elem["font_id"] == prev["font_id"]
            same_style = (elem["bold"] == prev["bold"] and elem["italic"] == prev["italic"])
            left_compatible = abs(elem["left"] - current_lines[0]["left"]) <= PARA_LEFT_TOLERANCE_PX
            # Also check center alignment: lines with varying left but similar center
            prev_center = prev["left"] + prev["width"] / 2
            elem_center = elem["left"] + elem["width"] / 2
            center_compatible = abs(prev_center - elem_center) <= ALIGNMENT_TOLERANCE_PX
            position_compatible = left_compatible or center_compatible
            not_list_start = not _is_list_item(elem["text"])

            if vertically_close and same_font and same_style and position_compatible and not_list_start:
                current_lines.append(elem)
            else:
                paragraphs.append(_build_paragraph(current_lines, page_width_px))
                current_lines = [elem]

        # Don't forget the last group
        paragraphs.append(_build_paragraph(current_lines, page_width_px))

    # Sort paragraphs by position (table cells + flow, reintegrated)
    paragraphs.sort(key=lambda p: (p["top"], p["left"]))

    return paragraphs


def group_all_pages(parsed_data: dict) -> dict:
    """Apply paragraph grouping to all pages in parsed data.

    Takes the output of parse_pdftohtml_xml and returns an enriched copy
    with a 'paragraphs' key added to each page.
    """
    enriched_pages = []
    for page in parsed_data.get("pages", []):
        paragraphs = group_text_into_paragraphs(page)
        enriched_page = {**page, "paragraphs": paragraphs}
        enriched_pages.append(enriched_page)

    return {**parsed_data, "pages": enriched_pages}


# ---------------------------------------------------------------------------
# Paragraph classification
# ---------------------------------------------------------------------------

def _build_fontspec_lookup(fontspecs: list[dict]) -> dict[str, dict]:
    """Build a fontspec lookup dict keyed by id. Pure function."""
    return {fs["id"]: fs for fs in fontspecs}


def _determine_body_font_size(paragraphs: list[dict], fontspec_by_id: dict) -> float:
    """Determine the most common (body) font size on a page.

    Returns the font size in pts that appears most frequently among paragraphs.
    Falls back to 11.0 if no fontspec data available.
    """
    size_counts: dict[float, int] = {}
    for para in paragraphs:
        fs = fontspec_by_id.get(para["font_id"])
        if fs:
            size = fs["size_pts"]
            size_counts[size] = size_counts.get(size, 0) + 1

    if not size_counts:
        return 11.0  # sensible default

    return max(size_counts, key=size_counts.get)


def _classify_single_paragraph(
    para: dict,
    fontspec_by_id: dict,
    page_height_pts: float,
    body_font_size_pts: float,
) -> dict:
    """Classify a single paragraph and return an enriched copy.

    Pure function — does not mutate the input paragraph dict.
    """
    fs = fontspec_by_id.get(para["font_id"])
    font_size_pts = fs["size_pts"] if fs else 0.0
    font_family = fs["family"] if fs else ""
    color = fs["color"] if fs else "#000000"

    # Convert paragraph position from px to pts for threshold comparison
    top_pts = _px_to_pts(para["top"])

    # Default classification
    para_type = "paragraph"
    heading_level = None

    # --- Footer / Header detection (checked first — overrides heading) ---
    if font_size_pts <= HEADER_FOOTER_MAX_FONT_PTS:
        if top_pts > page_height_pts * FOOTER_TOP_RATIO:
            para_type = "footer"
        elif top_pts < page_height_pts * HEADER_BOTTOM_RATIO:
            para_type = "header"

    # --- Heading detection (only if not header/footer) ---
    # Skip whitespace-only paragraphs — CJK fullwidth spaces at large font sizes
    # can produce phantom headings (T3 evaluator finding #4).
    text_stripped = para["text"].strip()
    is_whitespace_only = not text_stripped

    # Skip trivially short text — date fragments ("月", "日") and form placeholders
    # ("___") at heading font sizes should not become headings (T5 evaluator finding #1).
    content_chars = text_stripped.replace("_", "")
    is_too_short = len(content_chars) < HEADING_MIN_CONTENT_CHARS

    # Skip table-cell paragraphs — bold cells should not be promoted to headings
    # (T4 revisit: T2→T3 flow interaction fix).
    is_table_cell = para.get("is_table_cell", False)

    # Level 1 uses strict > (not >=) because the heading font size map renders
    # level 1 AT 18pt. Exactly-18pt text is article-level (level 2), while
    # >18pt is document-title-level (level 1).
    if para_type == "paragraph" and not is_whitespace_only and not is_too_short and not is_table_cell:
        if font_size_pts > HEADING_LEVEL_1_MIN_PTS:
            para_type = "heading"
            heading_level = 1
        elif font_size_pts >= HEADING_LEVEL_2_MIN_PTS:
            para_type = "heading"
            heading_level = 2
        elif font_size_pts >= HEADING_LEVEL_3_MIN_PTS:
            para_type = "heading"
            heading_level = 3
        elif font_size_pts >= HEADING_LEVEL_4_MIN_PTS and para["bold"]:
            para_type = "heading"
            heading_level = 4
        # Also classify as heading if bold + larger than body text.
        # Note: this branch only fires when font_size_pts < HEADING_LEVEL_4_MIN_PTS
        # (since all sizes >= 11 with bold are caught above), so the only reachable
        # level is 4. The size must be between body_font_size and 11pt.
        elif para["bold"] and font_size_pts > body_font_size_pts:
            para_type = "heading"
            heading_level = 4

    # Build enriched paragraph — new dict, no mutation
    enriched = {
        **para,
        "type": para_type,
        "font_size_pts": font_size_pts,
        "font_family": font_family,
        "color": color,
    }
    if heading_level is not None:
        enriched["heading_level"] = heading_level

    return enriched


def classify_paragraphs(page_data: dict, fontspec_by_id: dict) -> list[dict]:
    """Classify all paragraphs on a page.

    Takes a page dict (output of group_all_pages per-page) and a cumulative
    fontspec lookup. Returns a new list of enriched paragraph dicts.

    Pure function — does not mutate page_data.
    """
    paragraphs = page_data.get("paragraphs", [])
    if not paragraphs:
        return []

    page_height_pts = page_data.get("height_pts", _px_to_pts(page_data.get("height_px", 0)))
    body_font_size_pts = _determine_body_font_size(paragraphs, fontspec_by_id)

    return [
        _classify_single_paragraph(para, fontspec_by_id, page_height_pts, body_font_size_pts)
        for para in paragraphs
    ]


def classify_all_pages(grouped_data: dict) -> dict:
    """Apply classification to all pages. Composition: parse -> group -> classify.

    Takes the output of group_all_pages and returns an enriched copy with
    each paragraph classified. Builds a cumulative fontspec lookup across
    pages (mirroring pdftohtml's global fontspec behavior).

    Pure function — does not mutate grouped_data.
    """
    # Build cumulative fontspec lookup across all pages
    # (fontspecs may be declared on any page but referenced from later pages)
    cumulative_fontspecs: dict[str, dict] = {}

    enriched_pages = []
    for page in grouped_data.get("pages", []):
        # Accumulate fontspecs from this page using the shared helper
        page_fontspec_lookup = _build_fontspec_lookup(page.get("fontspecs", []))
        cumulative_fontspecs.update(page_fontspec_lookup)

        classified_paragraphs = classify_paragraphs(page, cumulative_fontspecs)
        enriched_page = {**page, "paragraphs": classified_paragraphs}
        enriched_pages.append(enriched_page)

    return {**grouped_data, "pages": enriched_pages}


# ---------------------------------------------------------------------------
# Table structure detection
# ---------------------------------------------------------------------------

# Table detection constants
TABLE_CONSECUTIVE_GAP_FACTOR = 2.0  # max vertical gap / line_height for consecutive rows
TABLE_COL_CLUSTER_TOLERANCE_PX = 50  # max left-position difference for same column


def _cluster_columns(left_positions: list[int]) -> list[int]:
    """Cluster left positions into column boundaries.

    Returns sorted list of representative left positions (one per column).
    Uses anchor-based clustering: each cluster is anchored to its first (min) value.
    A new position joins a cluster if it's within tolerance of the cluster's ANCHOR
    (not the last-added value), preventing chain-drift.
    """
    if not left_positions:
        return []

    sorted_lefts = sorted(set(left_positions))
    # Each cluster is (anchor, members)
    clusters: list[tuple[int, list[int]]] = [(sorted_lefts[0], [sorted_lefts[0]])]

    for left in sorted_lefts[1:]:
        anchor, members = clusters[-1]
        if left - anchor <= TABLE_COL_CLUSTER_TOLERANCE_PX:
            members.append(left)
        else:
            clusters.append((left, [left]))

    # Representative = anchor (minimum) of each cluster
    return [anchor for anchor, _ in clusters]


def _assign_col_index(left: int, col_boundaries: list[int]) -> int:
    """Assign a column index to an element based on its left position.

    Finds the closest column boundary. Returns 0-based column index.
    """
    best_col = 0
    best_dist = abs(left - col_boundaries[0])
    for i, boundary in enumerate(col_boundaries[1:], 1):
        dist = abs(left - boundary)
        if dist < best_dist:
            best_dist = dist
            best_col = i
    return best_col


def _group_elements_into_rows(
    elements: list[dict], tolerance: int = TABLE_ROW_TOP_TOLERANCE_PX
) -> list[list[dict]]:
    """Group elements by their top position (with tolerance) into rows.

    Returns a list of rows, where each row is a list of elements sorted by left.
    Rows are sorted by their minimum top position.
    """
    if not elements:
        return []

    sorted_elems = sorted(elements, key=lambda e: (e["top"], e["left"]))

    rows: list[list[dict]] = [[sorted_elems[0]]]
    for elem in sorted_elems[1:]:
        # Check if this element is on the same row as the current one
        current_row_top = rows[-1][0]["top"]
        if abs(elem["top"] - current_row_top) <= tolerance:
            rows[-1].append(elem)
        else:
            rows.append([elem])

    # Sort elements within each row by left position
    for row in rows:
        row.sort(key=lambda e: e["left"])

    return rows


def detect_tables(page_data: dict, fontspec_by_id: dict) -> list[dict]:
    """Detect and reconstruct table structures from a page's text elements.

    Algorithm:
      1. Find table row tops (from _find_table_row_tops)
      2. Collect table elements and interleaved non-table elements
      3. Cluster into contiguous table regions (gap < 2x line_height)
      4. For each region, detect columns and build structured table

    Returns a list of table dicts. Pure function.
    """
    text_elements = page_data.get("text_elements", [])
    if not text_elements:
        return []

    table_tops = _find_table_row_tops(text_elements)
    if not table_tops:
        return []

    # Collect all elements at table tops
    table_elements = [e for e in text_elements if e["top"] in table_tops]
    if not table_elements:
        return []

    # Typical line height for gap calculation
    line_height = max(e["height"] for e in table_elements)

    # Group table elements into rows
    table_rows = _group_elements_into_rows(table_elements)

    # Split into contiguous table regions (gap between rows < threshold)
    regions: list[list[list[dict]]] = [[table_rows[0]]]
    for row in table_rows[1:]:
        prev_row = regions[-1][-1]
        prev_bottom = max(e["top"] + e["height"] for e in prev_row)
        curr_top = min(e["top"] for e in row)
        gap = curr_top - prev_bottom

        if gap < line_height * TABLE_CONSECUTIVE_GAP_FACTOR:
            regions[-1].append(row)
        else:
            regions.append([row])

    # Build table structures from each region
    # Filter: a table must have at least 2 rows to be meaningful
    tables = []
    for region_rows in regions:
        if len(region_rows) < 2:
            continue  # single-row "tables" are not tables
        table = _build_table_from_region(region_rows, text_elements, fontspec_by_id,
                                         line_height)
        if table is not None:
            tables.append(table)

    return tables


def _build_table_from_region(
    region_rows: list[list[dict]],
    all_elements: list[dict],
    fontspec_by_id: dict,
    line_height: int,
) -> dict | None:
    """Build a structured table dict from a contiguous region of table rows.

    Also absorbs non-table elements that fall within the table's vertical extent
    (multi-line cell continuation text).

    Pure function.
    """
    if not region_rows:
        return None

    # Flatten to get all table elements in this region
    flat_elements = [e for row in region_rows for e in row]

    # Determine the vertical extent of this table
    table_top = min(e["top"] for e in flat_elements)
    table_bottom = max(e["top"] + e["height"] for e in flat_elements)

    # Absorb non-table elements that fall within the table's vertical extent
    # These are multi-line cell continuations (e.g., "Requirement" at top=283)
    table_tops_in_region = set()
    for row in region_rows:
        for e in row:
            table_tops_in_region.add(e["top"])

    absorbed = []
    for e in all_elements:
        if e["top"] not in table_tops_in_region and table_top <= e["top"] <= table_bottom:
            absorbed.append(e)

    # Combine all elements (original + absorbed)
    all_table_elements = flat_elements + absorbed

    # Re-group into rows with all elements
    all_rows = _group_elements_into_rows(all_table_elements)

    # Detect columns: collect all left positions across all rows
    all_lefts = [e["left"] for e in all_table_elements]
    col_boundaries = _cluster_columns(all_lefts)

    if len(col_boundaries) < 2:
        # Not a meaningful table (single column)
        return None

    col_count = len(col_boundaries)

    # Build structured rows with column assignments
    structured_rows = []
    for row_elements in all_rows:
        cells = []
        for elem in row_elements:
            col_idx = _assign_col_index(elem["left"], col_boundaries)
            fs = fontspec_by_id.get(elem["font_id"])
            font_size_pts = fs["size_pts"] if fs else 0.0
            cells.append({
                "text": elem["text"],
                "bold": elem["bold"],
                "font_size_pts": font_size_pts,
                "col": col_idx,
            })
        structured_rows.append(cells)

    # Merge multi-line cells: consecutive rows where a column has text at same col
    # but represents a continuation. Detected by: same column, immediately adjacent rows.
    merged_rows = _merge_multiline_cells(structured_rows)

    # Post-merge guard: if merge collapsed to fewer than 2 rows, not a table
    if len(merged_rows) < 2:
        return None

    # Compute bbox in px
    all_lefts_full = [e["left"] for e in all_table_elements]
    all_rights = [e["left"] + e["width"] for e in all_table_elements]
    all_tops = [e["top"] for e in all_table_elements]
    all_bottoms = [e["top"] + e["height"] for e in all_table_elements]

    min_left = min(all_lefts_full)
    min_top = min(all_tops)
    max_right = max(all_rights)
    max_bottom = max(all_bottoms)

    bbox_px = {
        "top": min_top,
        "left": min_left,
        "width": max_right - min_left,
        "height": max_bottom - min_top,
    }
    bbox_pts = {
        "top": _px_to_pts(min_top),
        "left": _px_to_pts(min_left),
        "width": _px_to_pts(max_right - min_left),
        "height": _px_to_pts(max_bottom - min_top),
    }

    # Column boundaries: left edges + right edge of last column
    right_edge = max(all_rights)
    col_boundaries_px = col_boundaries + [right_edge]

    return {
        "type": "table",
        "rows": merged_rows,
        "col_count": col_count,
        "row_count": len(merged_rows),
        "bbox_px": bbox_px,
        "bbox_pts": bbox_pts,
        "col_boundaries_px": col_boundaries_px,
    }


def _merge_multiline_cells(rows: list[list[dict]]) -> list[list[dict]]:
    """Merge cells that span multiple rows at the same column.

    When a cell in row N+1 has the same column index as a cell in row N,
    and the row only adds text to existing columns (no new column),
    the text is appended to the cell in row N.

    Returns a new list of merged rows. Pure function.
    """
    if len(rows) <= 1:
        return [list(row) for row in rows]

    merged: list[list[dict]] = []

    for row in rows:
        # Check if this row should merge with the previous one
        if merged and _should_merge_rows(merged[-1], row):
            # Merge: append text to existing cells by column
            prev_row = merged[-1]
            prev_by_col: dict[int, dict] = {}
            for cell in prev_row:
                prev_by_col[cell["col"]] = cell

            for cell in row:
                if cell["col"] in prev_by_col:
                    # Append text with space separator
                    prev_by_col[cell["col"]]["text"] += " " + cell["text"]
                else:
                    # New column in continuation — add it
                    prev_row.append(dict(cell))
        else:
            # New row — deep copy cells
            merged.append([dict(cell) for cell in row])

    return merged


def _should_merge_rows(prev_row: list[dict], curr_row: list[dict]) -> bool:
    """Determine if curr_row should merge into prev_row (multi-line cell).

    Heuristic: merge if curr_row's columns are a STRICT PROPER subset of
    prev_row's columns — i.e., the continuation row fills fewer columns.
    A row with the same column set is a new data row, not a continuation.
    """
    prev_cols = {cell["col"] for cell in prev_row}
    curr_cols = {cell["col"] for cell in curr_row}

    # Strict proper subset: must be a subset AND have fewer columns
    return curr_cols.issubset(prev_cols) and len(curr_cols) < len(prev_cols)


def detect_tables_all_pages(classified_data: dict) -> dict:
    """Apply table detection to all pages.

    Takes the output of classify_all_pages and returns an enriched copy
    with a 'tables' key added to each page. Pure function.
    """
    # Build cumulative fontspec lookup
    cumulative_fontspecs: dict[str, dict] = {}

    enriched_pages = []
    for page in classified_data.get("pages", []):
        page_fontspec_lookup = _build_fontspec_lookup(page.get("fontspecs", []))
        cumulative_fontspecs.update(page_fontspec_lookup)

        tables = detect_tables(page, cumulative_fontspecs)
        enriched_page = {**page, "tables": tables}
        enriched_pages.append(enriched_page)

    return {**classified_data, "pages": enriched_pages}


# ---------------------------------------------------------------------------
# Image enrichment
# ---------------------------------------------------------------------------

def enrich_images(page_data: dict) -> list[dict]:
    """Enrich image elements with bbox_pts (converted from px).

    Takes a page dict and returns a new list of enriched image dicts.
    Pure function — does not mutate page_data.
    """
    images = page_data.get("images", [])
    enriched = []
    for img in images:
        enriched.append({
            **img,
            "bbox_px": {
                "top": img["top"],
                "left": img["left"],
                "width": img["width"],
                "height": img["height"],
            },
            "bbox_pts": {
                "top": _px_to_pts(img["top"]),
                "left": _px_to_pts(img["left"]),
                "width": _px_to_pts(img["width"]),
                "height": _px_to_pts(img["height"]),
            },
        })
    return enriched


def enrich_images_all_pages(data: dict) -> dict:
    """Apply image enrichment to all pages. Pure function."""
    enriched_pages = []
    for page in data.get("pages", []):
        enriched = enrich_images(page)
        enriched_page = {**page, "images": enriched}
        enriched_pages.append(enriched_page)
    return {**data, "pages": enriched_pages}


# ---------------------------------------------------------------------------
# Table-proximity heading suppression
# ---------------------------------------------------------------------------

# Vertical margin (px) around a table's bbox within which bold H4 headings
# are demoted to paragraphs. Prevents bold table row labels (e.g., "Party A",
# "Company Name") from being promoted to headings when pdftohtml splits them
# into fewer than TABLE_ROW_MIN_ELEMENTS per row.
TABLE_HEADING_PROXIMITY_PX = 50


def _suppress_headings_near_tables(page_data: dict) -> list[dict]:
    """Demote H4 headings that fall within the vertical extent of a table.

    Bold row labels in party-info tables often appear with only 1-2 elements
    per row, failing table-row detection. After table bboxes are known, we
    demote H4 headings that are vertically adjacent to a detected table.

    Only H4 is demoted — H1-H3 headings near tables are structurally meaningful
    (e.g., "Article 3: Payment Schedule" above a payment table).

    Returns a new list of paragraphs. Pure function.
    """
    tables = page_data.get("tables", [])
    paragraphs = page_data.get("paragraphs", [])

    if not tables or not paragraphs:
        return paragraphs

    # Compute extended vertical zones for each table
    table_zones = []
    for t in tables:
        bbox = t["bbox_px"]
        zone_top = bbox["top"] - TABLE_HEADING_PROXIMITY_PX
        zone_bottom = bbox["top"] + bbox["height"] + TABLE_HEADING_PROXIMITY_PX
        zone_left = bbox["left"]
        zone_right = bbox["left"] + bbox["width"]
        table_zones.append((zone_top, zone_bottom, zone_left, zone_right))

    result = []
    for para in paragraphs:
        if (para.get("type") == "heading"
                and para.get("heading_level") == 4):
            para_top = para["top"]
            para_left = para["left"]
            para_right = para["left"] + para["width"]

            in_zone = any(
                zt <= para_top <= zb and para_right >= zl and para_left <= zr
                for zt, zb, zl, zr in table_zones
            )
            if in_zone:
                # Demote: remove heading_level, set type back to paragraph
                demoted = {k: v for k, v in para.items() if k != "heading_level"}
                demoted["type"] = "paragraph"
                result.append(demoted)
                continue

        result.append(para)

    return result


def suppress_headings_near_tables_all_pages(data: dict) -> dict:
    """Apply table-proximity heading suppression to all pages. Pure function."""
    enriched_pages = []
    for page in data.get("pages", []):
        adjusted_paragraphs = _suppress_headings_near_tables(page)
        enriched_page = {**page, "paragraphs": adjusted_paragraphs}
        enriched_pages.append(enriched_page)
    return {**data, "pages": enriched_pages}


# ---------------------------------------------------------------------------
# Complete pipeline
# ---------------------------------------------------------------------------

def parse_digital_pdf(pdf_path: str) -> dict:
    """Complete digital PDF parsing pipeline.

    Composes: parse → group → classify → detect_tables → suppress_headings → enrich_images

    Returns the fully enriched structure with:
      - pages[].text_elements (raw)
      - pages[].paragraphs (grouped + classified, headings near tables suppressed)
      - pages[].tables (detected table structures)
      - pages[].images (enriched with bbox_pts)
      - pages[].fontspecs

    This is the primary entry point for downstream consumers (e.g., poppler_to_dsl.py).
    """
    parsed = parse_pdftohtml_xml(pdf_path)
    grouped = group_all_pages(parsed)
    classified = classify_all_pages(grouped)
    with_tables = detect_tables_all_pages(classified)
    suppressed = suppress_headings_near_tables_all_pages(with_tables)
    with_images = enrich_images_all_pages(suppressed)
    return with_images


# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------

def format_summary(data: dict) -> str:
    """Format a human-readable summary of parsed PDF structure.

    Takes the output of parse_digital_pdf and returns a multi-line string
    with page count, heading/paragraph/table/image counts per page.
    Pure function.
    """
    pages = data.get("pages", [])
    lines = [f"PDF Structure Summary ({len(pages)} pages)"]
    lines.append("=" * 50)

    total_headings = 0
    total_paragraphs = 0
    total_tables = 0
    total_images = 0
    total_footers = 0

    for page in pages:
        pnum = page.get("number", "?")
        paras = page.get("paragraphs", [])
        tables = page.get("tables", [])
        images = page.get("images", [])

        headings = [p for p in paras if p.get("type") == "heading"]
        body_paras = [p for p in paras if p.get("type") == "paragraph"]
        footers = [p for p in paras if p.get("type") in ("footer", "header")]

        total_headings += len(headings)
        total_paragraphs += len(body_paras)
        total_tables += len(tables)
        total_images += len(images)
        total_footers += len(footers)

        lines.append(f"  Page {pnum}: "
                     f"{len(headings)} headings, "
                     f"{len(body_paras)} paragraphs, "
                     f"{len(tables)} tables, "
                     f"{len(images)} images, "
                     f"{len(footers)} footer/header")

        for h in headings:
            level = h.get("heading_level", "?")
            lines.append(f"    H{level}: {h['text'][:70]}")

    lines.append("-" * 50)
    lines.append(f"  Totals: "
                 f"{total_headings} headings, "
                 f"{total_paragraphs} paragraphs, "
                 f"{total_tables} tables, "
                 f"{total_images} images, "
                 f"{total_footers} footer/header")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Parse digital PDF structure using poppler tools."
    )
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--pdf", type=str, help="Path to PDF file — outputs full structure as JSON")
    mode.add_argument("--check-digital", type=str, help="Check if a PDF is digital")
    mode.add_argument("--fonts", type=str, help="Parse and display font info")
    parser.add_argument("--output", type=str, help="Output file path (default: stdout, only with --pdf)")
    parser.add_argument("--summary", action="store_true",
                        help="Print human-readable summary instead of JSON (only with --pdf)")
    args = parser.parse_args()

    if args.check_digital:
        pdf_path = args.check_digital
        if not Path(pdf_path).exists():
            print(f"Error: {pdf_path} not found", file=sys.stderr)
            sys.exit(1)
        is_digital = is_digital_pdf(pdf_path)
        print(json.dumps({"path": pdf_path, "is_digital": is_digital}, indent=2))

    elif args.fonts:
        pdf_path = args.fonts
        if not Path(pdf_path).exists():
            print(f"Error: {pdf_path} not found", file=sys.stderr)
            sys.exit(1)
        fonts = parse_pdffonts(pdf_path)
        print(json.dumps(fonts, indent=2))

    elif args.pdf:
        pdf_path = args.pdf
        if not Path(pdf_path).exists():
            print(f"Error: {pdf_path} not found", file=sys.stderr)
            sys.exit(1)
        data = parse_digital_pdf(pdf_path)

        if args.summary:
            print(format_summary(data))
        else:
            output = json.dumps(data, indent=2, ensure_ascii=False)
            if args.output:
                Path(args.output).write_text(output, encoding="utf-8")
                print(f"Written to {args.output}", file=sys.stderr)
            else:
                print(output)

    else:
        # Unreachable due to mutually_exclusive_group(required=True),
        # but kept for defensive programming.
        parser.print_help()
        sys.exit(1)


if __name__ == "__main__":
    main()
