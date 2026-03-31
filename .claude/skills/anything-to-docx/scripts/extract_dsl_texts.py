#!/usr/bin/env python3
"""extract_dsl_texts.py - Extract translatable text from XML DSL files into JSON.

Usage:
    uv run --with lxml extract_dsl_texts.py --workspace PATH --pages N --output texts.json

For each $WORKSPACE/dsl/page-{N}.xml (N=1..pages):
  - Traverses heading, paragraph, table, text-frame, side-by-side, page-header, page-footer
  - Extracts run text (skipping is-math="true") and bare cell text
  - Outputs JSON with unique segment IDs like "p1:heading[0]/run[0]"
"""

import argparse
import json
import re
import sys
from pathlib import Path

from lxml import etree


# ---------------------------------------------------------------------------
# Element tag sets for traversal
# ---------------------------------------------------------------------------

# Top-level children of <page> that contain translatable text
CONTAINER_TAGS = frozenset({
    "heading", "paragraph", "table", "text-frame",
    "side-by-side", "page-header", "page-footer",
})

# Tags that hold inline text via <run> children
INLINE_PARENT_TAGS = frozenset({"heading", "paragraph"})

# Tags we recurse into without generating IDs themselves
STRUCTURAL_TAGS = frozenset({
    "side-by-side", "column", "page-header", "page-footer",
    "text-frame",
})


# ---------------------------------------------------------------------------
# Pure functions: tree traversal -> segment extraction
# ---------------------------------------------------------------------------


def _count_preceding_siblings_with_tag(elem):
    """Count how many preceding siblings share the same tag name.

    Returns the 0-based index among same-tag siblings.
    """
    tag = elem.tag
    parent = elem.getparent()
    if parent is None:
        return 0
    index = 0
    for sibling in parent:
        if sibling is elem:
            break
        if sibling.tag == tag:
            index += 1
    return index


def _build_step(elem):
    """Build a single path step like 'heading[0]' or 'page-header'.

    page-header and page-footer are singletons so they have no index.
    """
    tag = elem.tag
    if tag in ("page-header", "page-footer"):
        return tag
    index = _count_preceding_siblings_with_tag(elem)
    return f"{tag}[{index}]"


def _is_math_run(run_elem):
    """Check if a <run> element is a math run that should be skipped."""
    return run_elem.get("is-math") == "true"


def _clean_ocr_artifacts(text):
    """Remove common OCR artifacts from extracted text.

    Strips markdown code fences, stray leading slashes from figure captions,
    and other noise that glm-ocr may inject.
    """
    if not text:
        return text
    # Remove markdown code fence blocks: ```markdown ... ```, ``` ... ```
    cleaned = re.sub(r"```(?:markdown|xml|json|html)?\s*```", "", text)
    cleaned = re.sub(r"```(?:markdown|xml|json|html)?\s*\n?\s*```", "", cleaned)
    # Remove standalone ``` fences
    cleaned = re.sub(r"^```(?:markdown|xml|json|html)?$", "", cleaned, flags=re.MULTILINE)
    # Strip stray leading "/" from figure captions (OCR artifact)
    cleaned = re.sub(r"^[/／]\s+", "", cleaned.strip())
    return cleaned.strip()


def _get_text(elem):
    """Get direct text content of an element (not including children's text)."""
    text = elem.text
    if text is None:
        return ""
    return _clean_ocr_artifacts(text)


def _text_is_empty(text):
    """Check if text is empty or whitespace-only."""
    return not text or text.strip() == ""


def _determine_type(path_parts):
    """Determine segment type from the xpath parts.

    Returns one of: 'heading', 'paragraph', 'table_cell', 'table_cell_run',
    'page_header', 'page_footer', 'text_frame'.
    """
    if not path_parts:
        return "unknown"

    # Check for structural containers first
    has_page_header = any(p == "page-header" for p in path_parts)
    has_page_footer = any(p == "page-footer" for p in path_parts)
    has_text_frame = any(p.startswith("text-frame") for p in path_parts)

    # The leaf determines the type
    leaf = path_parts[-1]
    leaf_tag = leaf.split("[")[0] if "[" in leaf else leaf

    if leaf_tag == "run":
        # Check parent context
        if any(p.startswith("cell") for p in path_parts):
            return "table_cell_run"
        if has_page_header:
            return "page_header"
        if has_page_footer:
            return "page_footer"
        if has_text_frame:
            return "text_frame"
        if any(p.startswith("heading") for p in path_parts):
            return "heading"
        return "paragraph"
    elif leaf_tag == "cell":
        return "table_cell"
    else:
        return "unknown"


def _extract_from_element(elem, path_parts, page_num, segments):
    """Recursively extract translatable segments from an element.

    Args:
        elem: lxml element to process
        path_parts: list of path steps leading to this element (e.g. ['heading[0]'])
        page_num: 1-based page number
        segments: accumulator list (mutated -- system boundary)
    """
    tag = elem.tag

    # Skip non-content elements
    if tag == "col-widths":
        return
    if tag == "image":
        return

    # Structural containers: recurse into children
    if tag in STRUCTURAL_TAGS:
        step = _build_step(elem)
        child_path = path_parts + [step]
        for child in elem:
            _extract_from_element(child, child_path, page_num, segments)
        return

    # Inline parents (heading, paragraph): extract runs
    if tag in INLINE_PARENT_TAGS:
        step = _build_step(elem)
        child_path = path_parts + [step]
        runs = [c for c in elem if c.tag == "run"]
        if runs:
            for run in runs:
                if _is_math_run(run):
                    continue
                text = _get_text(run)
                if _text_is_empty(text):
                    continue
                run_step = _build_step(run)
                seg_path = child_path + [run_step]
                xpath_str = "/".join(seg_path)
                seg_type = _determine_type(seg_path)
                segments.append({
                    "id": f"p{page_num}:{xpath_str}",
                    "page": page_num,
                    "type": seg_type,
                    "text": text,
                })
        return

    # Table: traverse rows -> cells -> optional runs
    if tag == "table":
        step = _build_step(elem)
        child_path = path_parts + [step]
        for row in elem:
            if row.tag != "row":
                continue
            row_step = _build_step(row)
            row_path = child_path + [row_step]
            for cell in row:
                if cell.tag != "cell":
                    continue
                cell_step = _build_step(cell)
                cell_path = row_path + [cell_step]
                cell_runs = [c for c in cell if c.tag == "run"]
                if cell_runs:
                    # Cell has run children: extract each run
                    for run in cell_runs:
                        if _is_math_run(run):
                            continue
                        text = _get_text(run)
                        if _text_is_empty(text):
                            continue
                        run_step = _build_step(run)
                        seg_path = cell_path + [run_step]
                        xpath_str = "/".join(seg_path)
                        seg_type = _determine_type(seg_path)
                        segments.append({
                            "id": f"p{page_num}:{xpath_str}",
                            "page": page_num,
                            "type": seg_type,
                            "text": text,
                        })
                else:
                    # Cell without run children: extract cell text directly
                    text = _get_text(cell)
                    if _text_is_empty(text):
                        continue
                    xpath_str = "/".join(cell_path)
                    seg_type = _determine_type(cell_path)
                    segments.append({
                        "id": f"p{page_num}:{xpath_str}",
                        "page": page_num,
                        "type": seg_type,
                        "text": text,
                    })
        return


def extract_page_segments(xml_path, page_num):
    """Extract all translatable segments from a single page XML file.

    Args:
        xml_path: Path to the page XML file
        page_num: 1-based page number

    Returns:
        List of segment dicts with id, page, type, text
    """
    tree = etree.parse(str(xml_path))
    root = tree.getroot()
    segments = []  # mutable accumulator at system boundary

    for child in root:
        _extract_from_element(child, [], page_num, segments)

    return segments


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def main():
    parser = argparse.ArgumentParser(
        description="Extract translatable text from XML DSL files"
    )
    parser.add_argument("--workspace", required=True, help="Workspace directory path")
    parser.add_argument("--pages", required=True, type=int, help="Total page count")
    parser.add_argument("--output", required=True, help="Output JSON file path")
    args = parser.parse_args()

    workspace = Path(args.workspace)
    total_pages = args.pages

    all_segments = []

    for page_num in range(1, total_pages + 1):
        xml_path = workspace / "dsl" / f"page-{page_num}.xml"
        if not xml_path.exists():
            print(f"Warning: {xml_path} not found, skipping page {page_num}",
                  file=sys.stderr)
            continue

        page_segments = extract_page_segments(xml_path, page_num)
        all_segments.extend(page_segments)
        print(f"Page {page_num}: extracted {len(page_segments)} segments")

    output_data = {
        "total_segments": len(all_segments),
        "segments": all_segments,
    }

    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(
        json.dumps(output_data, ensure_ascii=False, indent=2)
    )

    print(f"Total: {len(all_segments)} segments written to {args.output}")


if __name__ == "__main__":
    main()
