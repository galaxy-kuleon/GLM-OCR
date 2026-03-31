#!/usr/bin/env python3
"""build_page_dsl.py - Convert OCR JSON + style JSON for a single page into XML DSL.

Usage:
    python build_page_dsl.py --workspace PATH --page N \
        --page-width-pts F --page-height-pts F \
        [--margin-top-cm F] [--margin-bottom-cm F] \
        [--margin-left-cm F] [--margin-right-cm F] \
        [--font-latin NAME] [--font-cjk NAME]

Reads $WORKSPACE/ocr-output/input/input.json (page N, 1-based) and
$WORKSPACE/ocr-output/input/style-page-{N}.json, then outputs
$WORKSPACE/dsl/page-{N}.xml.
"""

import argparse
import json
import logging
import re
import sys
from pathlib import Path

from lxml import etree
from lxml import html as lxml_html

# Configure logging
logging.basicConfig(level=logging.WARNING, format="%(levelname)s: %(message)s")
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# LaTeX/Math pattern cleaning
# ---------------------------------------------------------------------------

LATEX_PATTERNS = [
    # Inline math: $ ... $ → extract content
    (r"\$\s*([^$]+?)\s*\$", r"\1"),
    # Superscript: ^{text} → Unicode or plain
    (r"\^\{\s*\\mathrm\s*\{\s*([^}]+)\s*\}\s*\}", r"\1"),  # ^{ \mathrm {th} } → th
    (r"\^\{\s*([^}]+)\s*\}", r"\1"),  # ^{th} → th
    # Subscript: _{text} → plain
    (r"_\{\s*([^}]+)\s*\}", r"\1"),
    # Common LaTeX commands
    (r"\\mathrm\s*\{\s*([^}]+)\s*\}", r"\1"),  # \mathrm{th} → th
    (r"\\text\s*\{\s*([^}]+)\s*\}", r"\1"),  # \text{...} → ...
    (
        r"\\textbf\s*\{\s*([^}]+)\s*\}",
        r"\1",
    ),  # \textbf{...} → ... (bold handled separately)
    (
        r"\\textit\s*\{\s*([^}]+)\s*\}",
        r"\1",
    ),  # \textit{...} → ... (italic handled separately)
    # Fractions: \frac{a}{b} → a/b
    (r"\\frac\s*\{\s*([^}]+)\s*\}\s*\{\s*([^}]+)\s*\}", r"\1/\2"),
    # Remove extra spaces from LaTeX artifacts
    (r"\$\s*", ""),  # Remove $ with spaces
    (r"\s*\$", ""),  # Remove trailing $ with spaces
]

# Math detection patterns (preserve as equations)
MATH_INDICATORS = [
    r"\\frac",
    r"\\sum",
    r"\\int",
    r"\\prod",
    r"\\sqrt",
    r"\\alpha",
    r"\\beta",
    r"\\gamma",
    r"\\delta",
    r"\\pi",
    r"\\times",
    r"\\div",
    r"\\pm",
    r"\\leq",
    r"\\geq",
    r"\\infty",
    r"\\partial",
    r"\\nabla",
    r"\\cdot",
]

# Validation patterns for content checking
VALIDATION_PATTERNS = {
    "latex_inline": r"\$[^$]+\$",
    "latex_command": r"\\[a-zA-Z]+\s*\{",
    "latex_frac": r"\\frac\s*\{",
    "latex_mathrm": r"\\mathrm\s*\{",
    "latex_superscript": r"\^\s*\{",
    "latex_subscript": r"_\s*\{",
}


def clean_latex_markup(text: str) -> str:
    """Remove LaTeX markup artifacts from OCR text."""
    if not text:
        return text

    cleaned = text
    for pattern, replacement in LATEX_PATTERNS:
        cleaned = re.sub(pattern, replacement, cleaned)

    # Clean up multiple spaces
    cleaned = re.sub(r"\s+", " ", cleaned).strip()

    return cleaned


def detect_math_content(text: str) -> tuple:
    """
    Detect if text contains mathematical formulas.
    Returns (is_math, cleaned_text).
    """
    if not text:
        return False, text

    is_math = any(re.search(pattern, text) for pattern in MATH_INDICATORS)

    if is_math:
        # Don't clean math content - preserve for equation conversion
        return True, text

    # Not math - clean LaTeX artifacts
    return False, clean_latex_markup(text)


def validate_content(content: str, region_index: int, page_num: int) -> list:
    """
    Validate content for artifacts and issues.
    Returns list of issues found.
    """
    issues = []

    for issue_type, pattern in VALIDATION_PATTERNS.items():
        matches = re.findall(pattern, content)
        if matches:
            issues.append(
                {
                    "type": issue_type,
                    "matches": matches[:3],  # First 3 matches
                    "region_index": region_index,
                    "page": page_num,
                }
            )

    return issues


def log_validation_issues(issues: list, workspace: str, page_num: int):
    """Log validation issues to file for review."""
    if not issues:
        return

    log_path = Path(workspace) / "dsl" / "content_validation.log"

    with open(log_path, "a", encoding="utf-8") as f:
        for issue in issues:
            f.write(
                f"Page {issue['page']}, Region {issue['region_index']}: "
                f"{issue['type']} - {issue['matches']}\n"
            )

    # Also log warnings
    for issue in issues:
        logger.warning(
            f"Content validation: {issue['type']} found in "
            f"page {issue['page']}, region {issue['region_index']}"
        )


# ---------------------------------------------------------------------------
# Style defaults by native_label
# ---------------------------------------------------------------------------

STYLE_DEFAULTS = {
    "doc_title": {
        "fs": 18,
        "b": True,
        "a": "center",
        "c": [0, 0, 0],
        "ff": "serif",
        "sb": 12,
        "sa": 6,
        "ls": 1.0,
    },
    "paragraph_title": {
        "fs": 14,
        "b": True,
        "a": "left",
        "c": [0, 0, 0],
        "ff": "serif",
        "sb": 12,
        "sa": 6,
        "ls": 1.0,
    },
    "text": {
        "fs": 11,
        "b": False,
        "a": "left",
        "c": [0, 0, 0],
        "ff": "serif",
        "sb": 0,
        "sa": 0,
        "ls": 1.0,
    },
    "figure_title": {
        "fs": 10,
        "b": False,
        "a": "center",
        "c": [0, 0, 0],
        "ff": "serif",
        "sb": 0,
        "sa": 0,
        "ls": 1.0,
    },
    "vision_footnote": {
        "fs": 9,
        "b": False,
        "a": "left",
        "c": [0, 0, 0],
        "ff": "serif",
        "sb": 0,
        "sa": 0,
        "ls": 1.0,
    },
    "table": {
        "fs": 9,
        "b": False,
        "a": "left",
        "c": [0, 0, 0],
        "ff": "serif",
        "th": True,
        "sb": 0,
        "sa": 0,
        "ls": 1.0,
    },
    "display_formula": {
        "fs": 11,
        "b": False,
        "a": "center",
        "c": [0, 0, 0],
        "ff": "serif",
        "sb": 6,
        "sa": 6,
        "ls": 1.0,
    },
    "inline_formula": {
        "fs": 11,
        "b": False,
        "a": "left",
        "c": [0, 0, 0],
        "ff": "serif",
        "sb": 0,
        "sa": 0,
        "ls": 1.0,
    },
    "abstract": {
        "fs": 11,
        "b": False,
        "a": "left",
        "c": [0, 0, 0],
        "ff": "serif",
        "sb": 0,
        "sa": 0,
        "ls": 1.0,
    },
    "reference_content": {
        "fs": 9,
        "b": False,
        "a": "left",
        "c": [0, 0, 0],
        "ff": "serif",
        "sb": 0,
        "sa": 0,
        "ls": 1.0,
    },
    "vertical_text": {
        "fs": 11,
        "b": False,
        "a": "left",
        "c": [0, 0, 0],
        "ff": "serif",
        "sb": 0,
        "sa": 0,
        "ls": 1.0,
    },
}


def get_style_for_region(region, styles_by_index):
    """Get merged style for a region: VLM style overrides defaults."""
    native = region.get("native_label", "text")
    defaults = STYLE_DEFAULTS.get(native, STYLE_DEFAULTS["text"]).copy()

    idx = region.get("index", -1)
    vlm_style = styles_by_index.get(idx, {})

    # Merge VLM style over defaults
    if "fs" in vlm_style:
        defaults["fs"] = vlm_style["fs"]
    if "b" in vlm_style:
        defaults["b"] = vlm_style["b"]
    if "a" in vlm_style:
        defaults["a"] = vlm_style["a"]
    if "c" in vlm_style:
        defaults["c"] = vlm_style["c"]
    if "th" in vlm_style:
        defaults["th"] = vlm_style["th"]
    if "tb" in vlm_style:
        defaults["tb"] = vlm_style["tb"]
    if "bd" in vlm_style:
        defaults["bd"] = vlm_style["bd"]
    if "cell_overrides" in vlm_style:
        defaults["cell_overrides"] = vlm_style["cell_overrides"]
    if "bg_rgb" in vlm_style:
        defaults["bg_rgb"] = vlm_style["bg_rgb"]
    if "border_style" in vlm_style:
        defaults["border_style"] = vlm_style["border_style"]
    if "ff" in vlm_style:
        defaults["ff"] = vlm_style["ff"]
    if "sb" in vlm_style:
        defaults["sb"] = vlm_style["sb"]
    if "sa" in vlm_style:
        defaults["sa"] = vlm_style["sa"]
    if "ls" in vlm_style:
        defaults["ls"] = vlm_style["ls"]

    return defaults


# ---------------------------------------------------------------------------
# Markdown parsing → multiple runs
# ---------------------------------------------------------------------------


def parse_markdown_runs(text, base_style):
    """Parse markdown bold/italic into a list of run dicts.

    Returns list of: {"text": str, "bold": bool, "italic": bool}
    """
    runs = []
    # Pattern: ***bold+italic***, **bold**, *italic*, plain text
    pattern = r"(\*\*\*(.+?)\*\*\*|\*\*(.+?)\*\*|\*(.+?)\*|([^*]+))"

    for match in re.finditer(pattern, text):
        if match.group(2):  # ***bold+italic***
            runs.append({"text": match.group(2), "bold": True, "italic": True})
        elif match.group(3):  # **bold**
            runs.append({"text": match.group(3), "bold": True, "italic": False})
        elif match.group(4):  # *italic*
            runs.append({"text": match.group(4), "bold": False, "italic": True})
        elif match.group(5):  # plain text
            runs.append(
                {
                    "text": match.group(5),
                    "bold": base_style.get("b", False),
                    "italic": False,
                }
            )

    if not runs:
        runs.append({"text": text, "bold": base_style.get("b", False), "italic": False})

    return runs


# ---------------------------------------------------------------------------
# HTML table parsing
# ---------------------------------------------------------------------------


def get_cell_text_with_breaks(cell_element):
    """Extract text from HTML cell, preserving <br> as newlines."""
    parts = []
    if cell_element.text:
        parts.append(cell_element.text)
    for child in cell_element:
        if child.tag == "br":
            parts.append("\n")
        else:
            parts.append(child.text_content())
        if child.tail:
            parts.append(child.tail)
    return "".join(parts).strip()


def parse_html_table(html_content):
    """Parse HTML table into structured rows.

    Returns list of rows, each row is list of dicts:
        {"text": str, "rowspan": int, "colspan": int, "is_header": bool}
    """
    try:
        tree = lxml_html.fromstring(html_content)
    except Exception:
        return []

    rows = []
    for tr in tree.iter("tr"):
        row = []
        for cell in tr:
            if cell.tag in ("td", "th"):
                text = get_cell_text_with_breaks(cell)
                rowspan = int(cell.get("rowspan", 1))
                colspan = int(cell.get("colspan", 1))
                row.append(
                    {
                        "text": text,
                        "rowspan": rowspan,
                        "colspan": colspan,
                        "is_header": cell.tag == "th",
                    }
                )
        if row:
            rows.append(row)
    return rows


# ---------------------------------------------------------------------------
# Floating text detection
# ---------------------------------------------------------------------------


def detect_floating_regions(page_regions):
    """Detect which regions are floating (text frames or side-by-side).

    Returns a set of region indices that should be treated as floating.
    """
    floating = set()
    for i, region in enumerate(page_regions):
        bbox = region.get("bbox_2d")
        if not bbox or region.get("label") != "text":
            continue

        x1, y1, x2, y2 = bbox
        for j, other in enumerate(page_regions):
            if i == j or other.get("label") != "text":
                continue
            other_bbox = other.get("bbox_2d")
            if not other_bbox:
                continue
            ox1, oy1, ox2, oy2 = other_bbox

            # Side-by-side: similar Y, non-overlapping X
            if abs(y1 - oy1) < 50 and (x1 > ox2 or x2 < ox1):
                floating.add(region["index"])
                floating.add(other["index"])

    return floating


def group_side_by_side(floating_regions):
    """Group floating regions into side-by-side clusters.

    Returns list of lists of regions that should be displayed side-by-side.
    """
    if not floating_regions:
        return []

    # Sort by Y then X
    sorted_regions = sorted(
        floating_regions, key=lambda r: (r["bbox_2d"][1], r["bbox_2d"][0])
    )

    groups = []
    current_group = [sorted_regions[0]]

    for region in sorted_regions[1:]:
        prev = current_group[-1]
        prev_bbox = prev["bbox_2d"]
        curr_bbox = region["bbox_2d"]

        # Same row if Y coordinates are close
        if abs(prev_bbox[1] - curr_bbox[1]) < 50:
            current_group.append(region)
        else:
            groups.append(current_group)
            current_group = [region]

    if current_group:
        groups.append(current_group)

    return groups


# ---------------------------------------------------------------------------
# VLM-driven text box detection & grouping
# ---------------------------------------------------------------------------


def detect_textbox_regions_from_vlm(page_regions, styles_by_index):
    """Identify regions marked as text box (tb=true) by VLM style data.

    Only considers regions with label="text". Image/table/formula regions
    are never treated as text boxes, even if VLM incorrectly marks them.

    Returns a set of region indices.
    """
    tb_indices = set()
    for r in page_regions:
        if r.get("label") != "text":
            continue
        idx = r.get("index", -1)
        style = styles_by_index.get(idx, {})
        if style.get("tb", False):
            tb_indices.add(idx)
    return tb_indices


def group_textbox_regions(page_regions, tb_indices):
    """Group tb=true regions that are spatially adjacent into text box groups.

    Grouping criteria (normalized 0-1000 coordinates):
    - X range overlap > 50%
    - Y gap < 80

    Returns list of lists of regions (each list is one text box group).
    """
    if not tb_indices:
        return []

    tb_regions = [
        r for r in page_regions if r.get("index") in tb_indices and r.get("bbox_2d")
    ]
    if not tb_regions:
        return []

    # Sort by Y then X
    tb_regions.sort(key=lambda r: (r["bbox_2d"][1], r["bbox_2d"][0]))

    groups = [[tb_regions[0]]]

    for region in tb_regions[1:]:
        bbox = region["bbox_2d"]
        x1, y1, x2, y2 = bbox

        merged = False
        for group in groups:
            # Check against all regions in the group for adjacency
            for member in group:
                mb = member["bbox_2d"]
                mx1, my1, mx2, my2 = mb

                # X overlap ratio
                overlap_start = max(x1, mx1)
                overlap_end = min(x2, mx2)
                overlap_width = max(0, overlap_end - overlap_start)
                min_width = min(x2 - x1, mx2 - mx1)
                x_overlap_ratio = overlap_width / min_width if min_width > 0 else 0

                # Y gap: distance between bottom of one and top of the other
                y_gap = y1 - my2  # current region is below member

                if x_overlap_ratio > 0.5 and 0 <= y_gap < 80:
                    group.append(region)
                    merged = True
                    break
            if merged:
                break

        if not merged:
            groups.append([region])

    # Sort regions within each group by Y
    for group in groups:
        group.sort(key=lambda r: r["bbox_2d"][1])

    return groups


def build_grouped_textframe(
    parent, group, styles_by_index, page_width_pts, page_height_pts
):
    """Build a single <text-frame> element from a group of tb=true regions.

    Merges bboxes to get the union, sets has-border from any region's bd field,
    and creates one <paragraph> per region.
    """
    # Compute union bbox
    all_x1 = min(r["bbox_2d"][0] for r in group)
    all_y1 = min(r["bbox_2d"][1] for r in group)
    all_x2 = max(r["bbox_2d"][2] for r in group)
    all_y2 = max(r["bbox_2d"][3] for r in group)

    # Convert to TWIPS
    x_twips = int(all_x1 * page_width_pts / 1000 * 20)
    y_twips = int(all_y1 * page_height_pts / 1000 * 20)
    w_twips = int((all_x2 - all_x1) * page_width_pts / 1000 * 20)
    h_twips = int((all_y2 - all_y1) * page_height_pts / 1000 * 20)

    # Determine border from any region's bd field
    has_border = any(
        styles_by_index.get(r.get("index", -1), {}).get("bd", False) for r in group
    )

    frame = etree.SubElement(parent, "text-frame")
    frame.set("x-twips", str(x_twips))
    frame.set("y-twips", str(y_twips))
    frame.set("width-twips", str(w_twips))
    frame.set("height-twips", str(h_twips))
    frame.set("has-border", str(has_border).lower())
    frame.set("border-color", "000000")

    for region in group:
        style = get_style_for_region(region, styles_by_index)
        content = region.get("content", "") or ""

        para = etree.SubElement(frame, "paragraph")
        para.set("alignment", style.get("a", "left"))

        runs = parse_markdown_runs(content, style)
        for r in runs:
            run_elem = etree.SubElement(para, "run")
            run_elem.text = r["text"]
            run_elem.set("font-size-pt", str(style.get("fs", 11)))
            run_elem.set(
                "color-rgb", ",".join(str(v) for v in style.get("c", [0, 0, 0]))
            )
            if r.get("bold"):
                run_elem.set("bold", "true")
            if r.get("italic"):
                run_elem.set("italic", "true")


# ---------------------------------------------------------------------------
# XML DSL builder
# ---------------------------------------------------------------------------


def make_run_element(text, style, bold_override=None, italic_override=None):
    """Create a <run> XML element."""
    run = etree.SubElement(etree.Element("dummy"), "run")
    run.text = text

    run.set("font-size-pt", str(style.get("fs", 11)))
    run.set("color-rgb", ",".join(str(v) for v in style.get("c", [0, 0, 0])))

    bold = bold_override if bold_override is not None else style.get("b", False)
    italic = italic_override if italic_override is not None else False

    if bold:
        run.set("bold", "true")
    if italic:
        run.set("italic", "true")

    return run


def build_text_region(parent, region, style):
    """Build XML elements for a text-type region."""
    content = region.get("content", "") or ""
    native = region.get("native_label", "text")

    # Detect and handle math content vs LaTeX artifacts
    is_math = False
    if native in ("display_formula", "inline_formula"):
        # Preserve formulas as-is
        is_math = True
        cleaned_content = content
    else:
        # Clean LaTeX artifacts from regular text
        is_math, cleaned_content = detect_math_content(content)

    # Use cleaned content
    content = cleaned_content

    # Strip heading markdown prefixes
    if native == "doc_title":
        content = re.sub(r"^#{1,6}\s+", "", content)
        heading = etree.SubElement(parent, "heading")
        heading.set("level", "1")
        heading.set("alignment", style.get("a", "center"))
        bg_rgb = style.get("bg_rgb")
        if bg_rgb:
            heading.set("bg-color", "".join(f"{v:02X}" for v in bg_rgb))
        # Add font family
        if style.get("ff"):
            heading.set("font-family", style["ff"])
        runs = parse_markdown_runs(content, style)
        for r in runs:
            run_elem = etree.SubElement(heading, "run")
            run_elem.text = r["text"]
            run_elem.set("font-size-pt", str(style.get("fs", 18)))
            run_elem.set("bold", "true")
            run_elem.set(
                "color-rgb", ",".join(str(v) for v in style.get("c", [0, 0, 0]))
            )
            if r.get("italic"):
                run_elem.set("italic", "true")
        return

    if native == "paragraph_title":
        content = re.sub(r"^#{1,6}\s+", "", content)
        heading = etree.SubElement(parent, "heading")
        heading.set("level", "2")
        heading.set("alignment", style.get("a", "left"))
        bg_rgb = style.get("bg_rgb")
        if bg_rgb:
            heading.set("bg-color", "".join(f"{v:02X}" for v in bg_rgb))
        # Add font family
        if style.get("ff"):
            heading.set("font-family", style["ff"])
        runs = parse_markdown_runs(content, style)
        for r in runs:
            run_elem = etree.SubElement(heading, "run")
            run_elem.text = r["text"]
            run_elem.set("font-size-pt", str(style.get("fs", 14)))
            run_elem.set("bold", str(r.get("bold", True)).lower())
            run_elem.set(
                "color-rgb", ",".join(str(v) for v in style.get("c", [0, 0, 0]))
            )
            if r.get("italic"):
                run_elem.set("italic", "true")
        return

    para_style = None

    # Figure caption
    if native == "figure_title":
        para_style = "figure-caption"

    # Footnote
    if native == "vision_footnote":
        para_style = "footnote"

    # Formula
    if native in ("display_formula", "inline_formula") or is_math:
        para_style = "formula"

    para = etree.SubElement(parent, "paragraph")
    if para_style:
        para.set("style", para_style)
    para.set("alignment", style.get("a", "left"))
    para.set("space-before-pt", str(style.get("sb", 0)))
    para.set("space-after-pt", str(style.get("sa", 0)))
    para.set("line-spacing", str(style.get("ls", 1.0)))
    bg_rgb = style.get("bg_rgb")
    if bg_rgb:
        para.set("bg-color", "".join(f"{v:02X}" for v in bg_rgb))
    # Add font family
    if style.get("ff"):
        para.set("font-family", style["ff"])

    runs = parse_markdown_runs(content, style)
    for r in runs:
        run_elem = etree.SubElement(para, "run")
        run_elem.text = r["text"]
        run_elem.set("font-size-pt", str(style.get("fs", 11)))
        run_elem.set("color-rgb", ",".join(str(v) for v in style.get("c", [0, 0, 0])))
        if r.get("bold"):
            run_elem.set("bold", "true")
        if r.get("italic"):
            run_elem.set("italic", "true")

        # Special handling for formulas
        if native in ("display_formula", "inline_formula") or is_math:
            run_elem.set("font-name", "Cambria Math")
            run_elem.set("is-math", "true")
            # Preserve original LaTeX for equation conversion
            if region.get("content") and "$" in region.get("content"):
                run_elem.set("latex", region.get("content"))


def _build_cell_runs(
    cell_elem, text, kw_entries, font_size, default_color, default_bold, default_text_bg
):
    """Split cell text by keywords and create <run> children with per-keyword styling."""
    keywords = [e["keyword"] for e in kw_entries if e.get("keyword")]
    if not keywords:
        cell_elem.text = text
        return
    kw_style_map = {e["keyword"]: e for e in kw_entries if e.get("keyword")}
    # Longest-first to avoid partial matches
    keywords_sorted = sorted(kw_style_map.keys(), key=len, reverse=True)
    pattern = "(" + "|".join(re.escape(kw) for kw in keywords_sorted) + ")"
    parts = re.split(pattern, text)
    for part in parts:
        if not part:
            continue
        run_elem = etree.SubElement(cell_elem, "run")
        run_elem.text = part
        run_elem.set("font-size-pt", str(font_size))
        if part in kw_style_map:
            kw = kw_style_map[part]
            if kw.get("c"):
                run_elem.set("color-rgb", ",".join(str(v) for v in kw["c"]))
            elif default_color and default_color != [0, 0, 0]:
                run_elem.set("color-rgb", ",".join(str(v) for v in default_color))
            if kw.get("bold"):
                run_elem.set("bold", "true")
            elif default_bold:
                run_elem.set("bold", "true")
            if kw.get("text_bg"):
                run_elem.set(
                    "text-bg-color", "".join(f"{v:02X}" for v in kw["text_bg"])
                )
        else:
            # Non-keyword part inherits cell defaults
            if default_color and default_color != [0, 0, 0]:
                run_elem.set("color-rgb", ",".join(str(v) for v in default_color))
            if default_bold:
                run_elem.set("bold", "true")


def build_table_region(parent, region, style, page_width_pts):
    """Build XML <table> element from an OCR table region."""
    content = region.get("content", "") or ""
    bbox = region.get("bbox_2d")

    rows_data = parse_html_table(content)
    if not rows_data:
        return

    num_rows = len(rows_data)
    max_cols = max(sum(c["colspan"] for c in r) for r in rows_data) if rows_data else 0

    if max_cols == 0:
        return

    table_elem = etree.SubElement(parent, "table")
    table_elem.set("rows", str(num_rows))
    table_elem.set("cols", str(max_cols))
    border_style = style.get("border_style", "single")
    if border_style not in ("single", "double", "none"):
        border_style = "single"
    table_elem.set("border-style", border_style)

    if bbox:
        table_elem.set("bbox", ",".join(str(v) for v in bbox))
    table_elem.set("page-width-pts", str(page_width_pts))

    # Equal column widths by default
    ratios = [round(1.0 / max_cols, 2)] * max_cols
    # Adjust last to ensure sum = 1.0
    ratios[-1] = round(1.0 - sum(ratios[:-1]), 2)
    col_widths_elem = etree.SubElement(table_elem, "col-widths")
    col_widths_elem.text = ",".join(str(r) for r in ratios)

    has_header = style.get("th", True) if num_rows > 1 else False
    font_size = style.get("fs", 9)
    region_color = style.get("c", [0, 0, 0])

    # Build cell_overrides lookups (coerce keys to int for consistent lookup)
    cell_overrides = style.get("cell_overrides", {})
    co_col_map = {}  # col_index -> {c, bg, type}
    co_row_map = {}  # row_index -> {c, bg, type}
    co_cell_map = {}  # (row, col) -> {c, bg, type}
    for entry in cell_overrides.get("col_colors", []):
        try:
            co_col_map[int(entry.get("col"))] = entry
        except (TypeError, ValueError):
            pass
    for entry in cell_overrides.get("row_colors", []):
        try:
            co_row_map[int(entry.get("row"))] = entry
        except (TypeError, ValueError):
            pass
    for entry in cell_overrides.get("cell_colors", []):
        try:
            co_cell_map[(int(entry.get("row")), int(entry.get("col")))] = entry
        except (TypeError, ValueError):
            pass
    co_kw_map = {}  # (row, col) -> list of keyword entries
    for entry in cell_overrides.get("keyword_styles", []):
        try:
            key = (int(entry.get("row")), int(entry.get("col")))
            co_kw_map.setdefault(key, []).append(entry)
        except (TypeError, ValueError):
            pass

    # Build occupancy grid to compute cell column indices
    occupied = [[False] * max_cols for _ in range(num_rows)]

    for r_idx, row in enumerate(rows_data):
        row_elem = etree.SubElement(table_elem, "row")
        row_elem.set("index", str(r_idx))
        if r_idx == 0 and has_header:
            row_elem.set("is-header", "true")

        col_cursor = 0
        for cell_data in row:
            while col_cursor < max_cols and occupied[r_idx][col_cursor]:
                col_cursor += 1
            if col_cursor >= max_cols:
                break

            cell_elem = etree.SubElement(row_elem, "cell")
            cell_elem.set("row", str(r_idx))
            cell_elem.set("col", str(col_cursor))
            cell_elem.set("font-size-pt", str(font_size))

            if cell_data["colspan"] > 1:
                cell_elem.set("colspan", str(cell_data["colspan"]))
            if cell_data["rowspan"] > 1:
                cell_elem.set("rowspan", str(cell_data["rowspan"]))

            # Header row styling (bold + center only, no hardcoded bg-color)
            if r_idx == 0 and has_header:
                cell_elem.set("bold", "true")
                cell_elem.set("alignment", "center")

            # Apply cell_overrides: cell-specific > column-level > row-level > region default
            # Determine text color for this cell
            text_color = None
            bg_color = None
            text_bg_color = None

            # Check cell-specific override (highest priority)
            cell_entry = co_cell_map.get((r_idx, col_cursor))
            if cell_entry:
                if cell_entry.get("type") == "text" and "c" in cell_entry:
                    text_color = cell_entry["c"]
                if cell_entry.get("type") == "bg" and "bg" in cell_entry:
                    bg_color = cell_entry["bg"]
                # Some entries may have both c and bg regardless of type
                if (
                    text_color is None
                    and "c" in cell_entry
                    and cell_entry.get("type") != "bg"
                ):
                    text_color = cell_entry["c"]
                if bg_color is None and "bg" in cell_entry:
                    bg_color = cell_entry["bg"]
                if "text_bg" in cell_entry:
                    text_bg_color = cell_entry["text_bg"]

            # Check column-level override
            col_entry = co_col_map.get(col_cursor)
            if col_entry:
                if (
                    text_color is None
                    and col_entry.get("type") == "text"
                    and "c" in col_entry
                ):
                    text_color = col_entry["c"]
                if (
                    bg_color is None
                    and col_entry.get("type") == "bg"
                    and "bg" in col_entry
                ):
                    bg_color = col_entry["bg"]
                if (
                    text_color is None
                    and "c" in col_entry
                    and col_entry.get("type") != "bg"
                ):
                    text_color = col_entry["c"]
                if bg_color is None and "bg" in col_entry:
                    bg_color = col_entry["bg"]
                if text_bg_color is None and "text_bg" in col_entry:
                    text_bg_color = col_entry["text_bg"]

            # Check row-level override
            row_entry = co_row_map.get(r_idx)
            if row_entry:
                if (
                    text_color is None
                    and row_entry.get("type") == "text"
                    and "c" in row_entry
                ):
                    text_color = row_entry["c"]
                if (
                    bg_color is None
                    and row_entry.get("type") == "bg"
                    and "bg" in row_entry
                ):
                    bg_color = row_entry["bg"]
                if (
                    text_color is None
                    and "c" in row_entry
                    and row_entry.get("type") != "bg"
                ):
                    text_color = row_entry["c"]
                if bg_color is None and "bg" in row_entry:
                    bg_color = row_entry["bg"]
                if text_bg_color is None and "text_bg" in row_entry:
                    text_bg_color = row_entry["text_bg"]

            # Apply text color (use override or region default)
            effective_color = text_color if text_color else region_color

            # Apply background color (stays on cell level regardless)
            if bg_color:
                hex_bg = "".join(f"{v:02X}" for v in bg_color)
                cell_elem.set("bg-color", hex_bg)

            # Check for keyword-level styling
            kw_entries = co_kw_map.get((r_idx, col_cursor), [])
            if kw_entries and cell_data["text"]:
                _build_cell_runs(
                    cell_elem,
                    cell_data["text"],
                    kw_entries,
                    font_size,
                    effective_color,
                    (r_idx == 0 and has_header),
                    text_bg_color,
                )
            else:
                cell_elem.text = cell_data["text"]
                if effective_color and effective_color != [0, 0, 0]:
                    cell_elem.set(
                        "color-rgb", ",".join(str(v) for v in effective_color)
                    )
                if text_bg_color:
                    hex_text_bg = "".join(f"{v:02X}" for v in text_bg_color)
                    cell_elem.set("text-bg-color", hex_text_bg)

            # Mark occupancy
            for mr in range(r_idx, min(r_idx + cell_data["rowspan"], num_rows)):
                for mc in range(
                    col_cursor, min(col_cursor + cell_data["colspan"], max_cols)
                ):
                    occupied[mr][mc] = True

            col_cursor += cell_data["colspan"]


def build_image_region(parent, page_idx, image_counter, bbox, page_width_pts):
    """Build XML <image> element."""
    img_src = f"ocr-output/input/imgs/cropped_page{page_idx}_idx{image_counter}.jpg"

    img_elem = etree.SubElement(parent, "image")
    img_elem.set("src", img_src)
    if bbox:
        img_elem.set("bbox", ",".join(str(v) for v in bbox))
    img_elem.set("page-width-pts", str(page_width_pts))
    img_elem.set("alignment", "center")


def build_floating_region(parent, region, style, page_width_pts, page_height_pts):
    """Build a <text-frame> element for a floating text region."""
    bbox = region.get("bbox_2d")
    if not bbox:
        return

    x1, y1, x2, y2 = bbox
    # Convert to TWIPS: normalized 0-1000 → pts → twips (1 pt = 20 twips)
    x_twips = int(x1 * page_width_pts / 1000 * 20)
    y_twips = int(y1 * page_height_pts / 1000 * 20)
    w_twips = int((x2 - x1) * page_width_pts / 1000 * 20)
    h_twips = int((y2 - y1) * page_height_pts / 1000 * 20)

    has_border = style.get("bd", False)

    frame = etree.SubElement(parent, "text-frame")
    frame.set("x-twips", str(x_twips))
    frame.set("y-twips", str(y_twips))
    frame.set("width-twips", str(w_twips))
    frame.set("height-twips", str(h_twips))
    frame.set("has-border", str(has_border).lower())
    frame.set("border-color", "000000")

    content = region.get("content", "") or ""
    para = etree.SubElement(frame, "paragraph")
    para.set("alignment", style.get("a", "left"))

    runs = parse_markdown_runs(content, style)
    for r in runs:
        run_elem = etree.SubElement(para, "run")
        run_elem.text = r["text"]
        run_elem.set("font-size-pt", str(style.get("fs", 11)))
        run_elem.set("color-rgb", ",".join(str(v) for v in style.get("c", [0, 0, 0])))
        if r.get("bold"):
            run_elem.set("bold", "true")
        if r.get("italic"):
            run_elem.set("italic", "true")


def build_side_by_side_group(parent, group, styles_by_index):
    """Build a <side-by-side> element from a group of floating regions."""
    sbs = etree.SubElement(parent, "side-by-side")
    sbs.set("cols", str(len(group)))

    for col_idx, region in enumerate(group):
        col_elem = etree.SubElement(sbs, "column")
        col_elem.set("index", str(col_idx))

        style = get_style_for_region(region, styles_by_index)
        content = region.get("content", "") or ""

        # Split content by lines for multiple paragraphs
        lines = content.split("\n") if content else [""]
        for line in lines:
            line = line.strip()
            if not line:
                continue
            para = etree.SubElement(col_elem, "paragraph")
            runs = parse_markdown_runs(line, style)
            for r in runs:
                run_elem = etree.SubElement(para, "run")
                run_elem.text = r["text"]
                run_elem.set("font-size-pt", str(style.get("fs", 11)))
                run_elem.set(
                    "color-rgb", ",".join(str(v) for v in style.get("c", [0, 0, 0]))
                )
                if r.get("bold"):
                    run_elem.set("bold", "true")
                if r.get("italic"):
                    run_elem.set("italic", "true")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def main():
    parser = argparse.ArgumentParser(
        description="Build per-page XML DSL from OCR + style data"
    )
    parser.add_argument("--workspace", required=True, help="Workspace directory path")
    parser.add_argument("--page", required=True, type=int, help="Page number (1-based)")
    parser.add_argument(
        "--page-width-pts", required=True, type=float, help="Page width in points"
    )
    parser.add_argument(
        "--page-height-pts", required=True, type=float, help="Page height in points"
    )
    parser.add_argument("--margin-top-cm", type=float, default=1.27)
    parser.add_argument("--margin-bottom-cm", type=float, default=1.27)
    parser.add_argument("--margin-left-cm", type=float, default=1.27)
    parser.add_argument("--margin-right-cm", type=float, default=1.27)
    parser.add_argument("--font-latin", default="Arial")
    parser.add_argument("--font-cjk", default="SimSun")
    args = parser.parse_args()

    workspace = Path(args.workspace)
    page_num = args.page  # 1-based

    # Load OCR JSON
    ocr_path = workspace / "ocr-output" / "input" / "input.json"
    if not ocr_path.exists():
        print(f"Error: OCR JSON not found at {ocr_path}", file=sys.stderr)
        sys.exit(1)

    with open(ocr_path) as f:
        ocr_data = json.load(f)

    if page_num < 1 or page_num > len(ocr_data):
        print(
            f"Error: page {page_num} out of range (1-{len(ocr_data)})", file=sys.stderr
        )
        sys.exit(1)

    page_regions = ocr_data[page_num - 1]  # 0-indexed

    # Load style JSON (optional)
    style_path = workspace / "ocr-output" / "input" / f"style-page-{page_num}.json"
    styles_by_index = {}
    if style_path.exists():
        try:
            with open(style_path) as f:
                style_data = json.load(f)
            for s in style_data:
                idx = s.get("region_index", s.get("i", -1))
                # Normalize field names from expanded or short format
                normalized = {}
                if "font_size_pt" in s:
                    normalized["fs"] = s["font_size_pt"]
                elif "fs" in s:
                    normalized["fs"] = s["fs"]
                if "bold" in s:
                    normalized["b"] = s["bold"]
                elif "b" in s:
                    normalized["b"] = s["b"]
                if "alignment" in s:
                    normalized["a"] = s["alignment"]
                elif "a" in s:
                    normalized["a"] = s["a"]
                if "color_rgb" in s:
                    normalized["c"] = s["color_rgb"]
                elif "c" in s:
                    normalized["c"] = s["c"]
                if "th" in s:
                    normalized["th"] = s["th"]
                if "tb" in s or "text_box" in s:
                    normalized["tb"] = s.get("tb", s.get("text_box", False))
                if "bd" in s or "border" in s:
                    normalized["bd"] = s.get("bd", s.get("border", False))
                if "cell_overrides" in s:
                    normalized["cell_overrides"] = s["cell_overrides"]
                if "bg_rgb" in s:
                    normalized["bg_rgb"] = s["bg_rgb"]
                if "border_style" in s:
                    normalized["border_style"] = s["border_style"]
                if "font_family" in s:
                    normalized["ff"] = s["font_family"]
                elif "ff" in s:
                    normalized["ff"] = s["ff"]
                if "space_before_pt" in s:
                    normalized["sb"] = s["space_before_pt"]
                elif "sb" in s:
                    normalized["sb"] = s["sb"]
                if "space_after_pt" in s:
                    normalized["sa"] = s["space_after_pt"]
                elif "sa" in s:
                    normalized["sa"] = s["sa"]
                if "line_spacing" in s:
                    normalized["ls"] = s["line_spacing"]
                elif "ls" in s:
                    normalized["ls"] = s["ls"]
                styles_by_index[idx] = normalized
        except Exception as e:
            print(
                f"Warning: failed to load style JSON {style_path}: {e}", file=sys.stderr
            )

    # --- Three-layer floating detection ---

    # Layer 1: VLM text-box detection (primary)
    tb_indices = detect_textbox_regions_from_vlm(page_regions, styles_by_index)
    tb_groups = group_textbox_regions(page_regions, tb_indices)

    # Layer 2: Geometric side-by-side detection (always)
    sbs_floating_indices = detect_floating_regions(page_regions)
    # Exclude regions already claimed by text-box groups
    sbs_floating_indices -= tb_indices
    sbs_floating_regions = [
        r for r in page_regions if r["index"] in sbs_floating_indices
    ]

    # Collect all non-flow indices
    all_floating_indices = tb_indices | sbs_floating_indices
    flow_regions = [r for r in page_regions if r["index"] not in all_floating_indices]

    # Build XML
    page_elem = etree.Element("page")
    page_elem.set("number", str(page_num))
    page_elem.set("width-pts", str(args.page_width_pts))
    page_elem.set("height-pts", str(args.page_height_pts))
    page_elem.set("margin-top-cm", str(args.margin_top_cm))
    page_elem.set("margin-bottom-cm", str(args.margin_bottom_cm))
    page_elem.set("margin-left-cm", str(args.margin_left_cm))
    page_elem.set("margin-right-cm", str(args.margin_right_cm))
    page_elem.set("font-latin", args.font_latin)
    page_elem.set("font-cjk", args.font_cjk)

    # Process flow regions
    image_counter = 0
    page_idx = page_num - 1  # 0-based for image naming
    all_validation_issues = []

    for region in flow_regions:
        label = region.get("label", "text")
        style = get_style_for_region(region, styles_by_index)

        # Validate content for artifacts
        content = region.get("content", "")
        if content:
            issues = validate_content(content, region.get("index", -1), page_num)
            all_validation_issues.extend(issues)

        if label == "text":
            build_text_region(page_elem, region, style)
        elif label == "table":
            build_table_region(page_elem, region, style, args.page_width_pts)
        elif label == "image":
            build_image_region(
                page_elem,
                page_idx,
                image_counter,
                region.get("bbox_2d"),
                args.page_width_pts,
            )
            image_counter += 1
        elif label == "formula":
            build_text_region(page_elem, region, style)

    # Log validation issues
    if all_validation_issues:
        log_validation_issues(all_validation_issues, str(workspace), page_num)

    # Process side-by-side floating regions
    if sbs_floating_regions:
        sbs_groups = group_side_by_side(sbs_floating_regions)
        for group in sbs_groups:
            if len(group) >= 2:
                build_side_by_side_group(page_elem, group, styles_by_index)
            else:
                # Single floating region → text-frame
                region = group[0]
                style = get_style_for_region(region, styles_by_index)
                build_floating_region(
                    page_elem, region, style, args.page_width_pts, args.page_height_pts
                )

    # Process VLM text-box groups
    for group in tb_groups:
        build_grouped_textframe(
            page_elem, group, styles_by_index, args.page_width_pts, args.page_height_pts
        )

    # Count images in floating regions too
    for region in page_regions:
        if region["index"] in all_floating_indices and region.get("label") == "image":
            build_image_region(
                page_elem,
                page_idx,
                image_counter,
                region.get("bbox_2d"),
                args.page_width_pts,
            )
            image_counter += 1

    # Write output
    dsl_dir = workspace / "dsl"
    dsl_dir.mkdir(parents=True, exist_ok=True)
    output_path = dsl_dir / f"page-{page_num}.xml"

    tree = etree.ElementTree(page_elem)
    tree.write(
        str(output_path),
        xml_declaration=True,
        encoding="UTF-8",
        pretty_print=True,
    )

    n_tb = sum(len(g) for g in tb_groups)
    n_sbs = len(sbs_floating_regions)
    print(
        f"Generated {output_path} ({len(flow_regions)} flow + {n_sbs} side-by-side + {n_tb} text-box regions)"
    )


if __name__ == "__main__":
    main()
