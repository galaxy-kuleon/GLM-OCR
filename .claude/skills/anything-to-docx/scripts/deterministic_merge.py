#!/usr/bin/env python3
"""deterministic_merge.py - Merge VLM XML with OCR JSON without VLM calls.

Deterministic text replacement merge for weak VLM models:
1. Keep VLM XML structure (layout, styling, element ordering)
2. Replace text with OCR text using sequential matching + text similarity
3. Resolve image paths from OCR cropped images
4. Clean OCR artifacts at merge time

Designed as a drop-in alternative to vlm_merge_dsl.py when the VLM
is too weak to perform the merge step reliably.

Usage:
    uv run --with lxml \
        .claude/skills/anything-to-docx/scripts/deterministic_merge.py \
        --workspace /path/to/workspace --pages 10
"""

import argparse
import json
import os
import re
import shutil
import sys
from copy import deepcopy
from pathlib import Path

from lxml import etree


# ---------------------------------------------------------------------------
# OCR artifact cleanup (shared pattern with other pipeline scripts)
# ---------------------------------------------------------------------------


def _clean_ocr_text(text):
    """Remove common OCR artifacts from content strings."""
    if not text or not isinstance(text, str):
        return text or ""
    cleaned = re.sub(r"```(?:markdown|xml|json|html)?\s*```", "", text)
    cleaned = re.sub(r"```(?:markdown|xml|json|html)?\s*\n?\s*```", "", cleaned)
    cleaned = re.sub(r"^```(?:markdown|xml|json|html)?$", "", cleaned, flags=re.MULTILINE)
    cleaned = cleaned.strip()
    if cleaned.startswith("/ ") or cleaned.startswith("\uff0f "):
        cleaned = cleaned[2:]
    # Remove markdown heading markers from OCR text (structure comes from VLM)
    cleaned = re.sub(r"^#{1,6}\s+", "", cleaned)
    return cleaned.strip()


# ---------------------------------------------------------------------------
# Text similarity
# ---------------------------------------------------------------------------


def _normalize_for_match(text):
    """Normalize text for similarity matching: lowercase, collapse whitespace."""
    if not text:
        return ""
    t = text.lower().strip()
    t = re.sub(r"\s+", " ", t)
    return t


def text_similarity(a, b):
    """Compute word-level Jaccard similarity between two strings.

    Returns float in [0, 1]. Fast and effective for matching the same
    text from two different OCR/VLM sources.

    >>> text_similarity("Hello World", "hello world")
    1.0
    >>> text_similarity("", "")
    0.0
    >>> text_similarity("The quick brown fox", "quick brown fox jumps")
    0.6
    """
    na = _normalize_for_match(a)
    nb = _normalize_for_match(b)
    if not na or not nb:
        return 0.0
    words_a = set(na.split())
    words_b = set(nb.split())
    if not words_a or not words_b:
        return 0.0
    intersection = words_a & words_b
    union = words_a | words_b
    return len(intersection) / len(union)


def _char_overlap_ratio(a, b):
    """Character-level overlap ratio for short text matching.

    More robust than word Jaccard for very short strings (1-3 words).
    Returns float in [0, 1].
    """
    na = _normalize_for_match(a)
    nb = _normalize_for_match(b)
    if not na or not nb:
        return 0.0
    chars_a = set(na)
    chars_b = set(nb)
    intersection = chars_a & chars_b
    union = chars_a | chars_b
    return len(intersection) / len(union) if union else 0.0


def combined_similarity(a, b):
    """Combined similarity score using both word Jaccard and char overlap.

    For short texts (< 5 words), char overlap gets more weight.
    For longer texts, word Jaccard is primary.
    """
    word_sim = text_similarity(a, b)
    char_sim = _char_overlap_ratio(a, b)
    word_count = len(_normalize_for_match(a).split())
    if word_count < 5:
        return 0.4 * word_sim + 0.6 * char_sim
    return 0.7 * word_sim + 0.3 * char_sim


# ---------------------------------------------------------------------------
# Bbox utilities
# ---------------------------------------------------------------------------


def compute_bbox_iou(bbox_a, bbox_b):
    """Compute IoU between two bboxes [x1, y1, x2, y2] (normalized 0-1000).

    >>> compute_bbox_iou([0, 0, 500, 500], [250, 250, 750, 750])
    0.14285714285714285
    >>> compute_bbox_iou([0, 0, 100, 100], [200, 200, 300, 300])
    0.0
    """
    ax1, ay1, ax2, ay2 = bbox_a
    bx1, by1, bx2, by2 = bbox_b

    ix1 = max(ax1, bx1)
    iy1 = max(ay1, by1)
    ix2 = min(ax2, bx2)
    iy2 = min(ay2, by2)

    if ix2 <= ix1 or iy2 <= iy1:
        return 0.0

    inter = (ix2 - ix1) * (iy2 - iy1)
    area_a = (ax2 - ax1) * (ay2 - ay1)
    area_b = (bx2 - bx1) * (by2 - by1)
    union = area_a + area_b - inter
    return inter / union if union > 0 else 0.0


def _parse_bbox_attr(bbox_str):
    """Parse a bbox attribute string 'x1,y1,x2,y2' into a list of floats."""
    if not bbox_str:
        return None
    parts = [float(x.strip()) for x in bbox_str.split(",")]
    return parts if len(parts) == 4 else None


# ---------------------------------------------------------------------------
# VLM XML element collection
# ---------------------------------------------------------------------------


def _get_element_text(elem):
    """Get concatenated text from an element's run children or direct text."""
    runs = elem.findall("run")
    if runs:
        return " ".join((r.text or "") for r in runs if r.text)
    return (elem.text or "").strip()


def _label_matches_type(native_label, elem_type):
    """Check if an OCR native_label matches a VLM element type."""
    heading_labels = {"doc_title", "paragraph_title"}
    text_labels = {"text", "vision_footnote", "figure_title"}

    if elem_type == "heading":
        return native_label in heading_labels
    if elem_type == "paragraph":
        return native_label in text_labels or native_label in heading_labels
    return False


def collect_text_elements(page_el):
    """Collect text-bearing elements from VLM XML in document order.

    Returns list of dicts:
        {
            "element": lxml Element (the paragraph/heading element),
            "text": concatenated text content,
            "type": "heading" | "paragraph",
            "bbox": [x1,y1,x2,y2] or None,
        }

    Tables and images are handled separately.
    """
    results = []
    for elem in page_el:
        tag = elem.tag
        if tag in ("heading", "paragraph"):
            text = _get_element_text(elem)
            results.append({
                "element": elem,
                "text": text,
                "type": tag,
                "bbox": None,  # headings/paragraphs don't have bbox in VLM schema
            })
        elif tag in ("page-header", "page-footer", "text-frame"):
            # Recurse into structural containers
            for child in elem:
                if child.tag in ("heading", "paragraph"):
                    text = _get_element_text(child)
                    bbox = _parse_bbox_attr(elem.get("bbox"))
                    results.append({
                        "element": child,
                        "text": text,
                        "type": child.tag,
                        "bbox": bbox,
                    })
        elif tag == "side-by-side":
            for col in elem.findall("column"):
                for child in col:
                    if child.tag in ("heading", "paragraph"):
                        text = _get_element_text(child)
                        results.append({
                            "element": child,
                            "text": text,
                            "type": child.tag,
                            "bbox": None,
                        })
        # Skip: table, image, horizontal-rule, col-widths
    return results


# ---------------------------------------------------------------------------
# Text replacement
# ---------------------------------------------------------------------------


def _replace_element_text(elem, new_text):
    """Replace all text in a heading/paragraph element with new_text.

    If the element has <run> children, replaces the first run's text
    and clears the rest. Preserves styling of the first run.
    """
    runs = elem.findall("run")
    if runs:
        # Put all text in first run, remove text from others
        runs[0].text = new_text
        for r in runs[1:]:
            r.text = ""
    else:
        elem.text = new_text


def match_and_replace_text(page_el, ocr_regions, page_num):
    """Match OCR text regions to VLM text elements and replace.

    Uses greedy sequential matching with a look-ahead window.
    Modifies page_el in place.

    Returns (replaced_count, total_vlm_elements).
    """
    # Separate OCR regions by type
    text_regions = []
    for r in ocr_regions:
        label = r.get("label", "")
        native = r.get("native_label", "")
        content = _clean_ocr_text(r.get("content", ""))
        if not content:
            continue
        if label in ("text",) or native in ("doc_title", "paragraph_title",
                                             "text", "vision_footnote",
                                             "figure_title"):
            text_regions.append({
                "content": content,
                "native_label": native,
                "bbox": r.get("bbox_2d", [0, 0, 0, 0]),
                "index": r.get("index", 0),
            })

    # Sort OCR text regions by vertical position (top to bottom)
    text_regions.sort(key=lambda r: (r["bbox"][1], r["bbox"][0]))

    # Collect VLM text elements
    vlm_elements = collect_text_elements(page_el)

    if not vlm_elements or not text_regions:
        return 0, len(vlm_elements)

    # Greedy sequential matching with look-ahead
    consumed = set()
    ocr_cursor = 0
    replaced = 0
    LOOK_AHEAD = 6  # search window around cursor

    for ve in vlm_elements:
        vlm_text = ve["text"]
        if not vlm_text.strip():
            continue

        best_score = 0.0
        best_idx = -1

        search_start = max(0, ocr_cursor - 2)
        search_end = min(len(text_regions), ocr_cursor + LOOK_AHEAD)

        for i in range(search_start, search_end):
            if i in consumed:
                continue

            ocr_r = text_regions[i]
            sim = combined_similarity(vlm_text, ocr_r["content"])

            # Label matching bonus
            if _label_matches_type(ocr_r["native_label"], ve["type"]):
                sim += 0.1

            # Proximity bonus (prefer closer to cursor)
            distance = abs(i - ocr_cursor)
            proximity_bonus = 0.05 * max(0, 3 - distance)
            sim += proximity_bonus

            if sim > best_score:
                best_score = sim
                best_idx = i

        # Threshold for accepting a match
        if best_idx >= 0 and best_score > 0.25:
            _replace_element_text(ve["element"], text_regions[best_idx]["content"])
            consumed.add(best_idx)
            ocr_cursor = best_idx + 1
            replaced += 1

    return replaced, len(vlm_elements)


# ---------------------------------------------------------------------------
# Image resolution
# ---------------------------------------------------------------------------


def resolve_images(page_el, ocr_regions, page_num):
    """Set image src paths from OCR image regions by bbox matching.

    For each VLM <image> element, find the best matching OCR image
    region by bbox IoU and set the src to the OCR cropped image path.

    Modifies page_el in place. Returns count of resolved images.
    """
    image_regions = [
        r for r in ocr_regions
        if r.get("label") == "image"
    ]

    vlm_images = page_el.findall(".//image")
    resolved = 0

    # Track image index for this page
    image_counter = 0
    for r in sorted(ocr_regions, key=lambda x: x.get("index", 0)):
        if r.get("label") == "image":
            r["_image_idx"] = image_counter
            image_counter += 1

    for img_el in vlm_images:
        img_bbox = _parse_bbox_attr(img_el.get("bbox"))

        if img_bbox and image_regions:
            best_iou = 0.0
            best_region = None

            for ocr_img in image_regions:
                ocr_bbox = ocr_img.get("bbox_2d", [0, 0, 0, 0])
                iou = compute_bbox_iou(img_bbox, ocr_bbox)
                if iou > best_iou:
                    best_iou = iou
                    best_region = ocr_img

            if best_region and best_iou > 0.05:
                page_idx = page_num - 1  # 0-based
                img_idx = best_region.get("_image_idx", best_region.get("index", 0))
                src = f"ocr-output/input/imgs/cropped_page{page_idx}_idx{img_idx}.jpg"
                img_el.set("src", src)
                resolved += 1
                continue

        # Fallback: sequential assignment
        if image_regions:
            region = image_regions.pop(0)
            page_idx = page_num - 1
            img_idx = region.get("_image_idx", region.get("index", 0))
            src = f"ocr-output/input/imgs/cropped_page{page_idx}_idx{img_idx}.jpg"
            img_el.set("src", src)
            resolved += 1

    return resolved


# ---------------------------------------------------------------------------
# Table text replacement
# ---------------------------------------------------------------------------


def _try_replace_table_text(table_el, ocr_regions, page_num):
    """Attempt to improve table cell text using OCR regions.

    For tables, OCR gives flat text (not cell-structured), so we can only
    do limited correction. Strategy:
    - Find OCR regions that overlap with the table bbox
    - If a single OCR region covers the table, we can't split into cells
    - If OCR has individual cell-like regions (small bbox within table),
      try to match by cell bbox position

    Returns count of cells improved.
    """
    table_bbox = _parse_bbox_attr(table_el.get("bbox"))
    if not table_bbox:
        return 0

    # Find OCR text regions that fall within the table bbox
    overlapping = []
    for r in ocr_regions:
        if r.get("label") == "image":
            continue
        ocr_bbox = r.get("bbox_2d", [0, 0, 0, 0])
        iou = compute_bbox_iou(table_bbox, ocr_bbox)
        if iou > 0.01:
            content = _clean_ocr_text(r.get("content", ""))
            if content:
                overlapping.append({
                    "content": content,
                    "bbox": ocr_bbox,
                })

    if not overlapping:
        return 0

    # If there's only one large OCR region covering the whole table,
    # we can't split it into cells — keep VLM text
    if len(overlapping) == 1:
        return 0

    # Multiple OCR regions within table — try to match to cells
    # This works when OCR detected individual cells
    improved = 0
    for row_el in table_el.findall("row"):
        for cell_el in row_el.findall("cell"):
            cell_text = _get_element_text(cell_el)
            if not cell_text.strip():
                continue

            # Try to find matching OCR region by text similarity
            best_sim = 0.0
            best_content = None
            for ocr_r in overlapping:
                sim = combined_similarity(cell_text, ocr_r["content"])
                if sim > best_sim:
                    best_sim = sim
                    best_content = ocr_r["content"]

            if best_content and best_sim > 0.5:
                run_children = cell_el.findall("run")
                if run_children:
                    run_children[0].text = best_content
                    for r in run_children[1:]:
                        r.text = ""
                else:
                    cell_el.text = best_content
                improved += 1

    return improved


# ---------------------------------------------------------------------------
# Data loading
# ---------------------------------------------------------------------------


def load_ocr_data(workspace):
    """Load glm-ocr JSON data. Returns list of page lists (0-indexed)."""
    ocr_path = Path(workspace) / "ocr-output" / "input" / "input.json"
    if not ocr_path.exists():
        print(f"Warning: OCR data not found at {ocr_path}", file=sys.stderr)
        print(f"FIX: Run glmocr first (Step B2). Expected file: {ocr_path}", file=sys.stderr)
        return []
    with open(ocr_path, "r", encoding="utf-8") as f:
        return json.load(f)


def load_vlm_xml(workspace, page_num):
    """Load VLM-generated XML for a page. Returns etree Element or None."""
    xml_path = Path(workspace) / "dsl-vlm" / f"page-{page_num}.xml"
    if not xml_path.exists():
        return None
    with open(xml_path, "r", encoding="utf-8") as f:
        text = f.read().strip()
    try:
        return etree.fromstring(text.encode("utf-8"))
    except etree.XMLSyntaxError:
        # Try recovery parsing
        try:
            parser = etree.XMLParser(recover=True)
            return etree.fromstring(text.encode("utf-8"), parser=parser)
        except Exception as e:
            print(f"Warning: Cannot parse VLM XML for page {page_num}: {e}",
                  file=sys.stderr)
            print(f"FIX: Re-run vlm_generate_dsl.py for this page, or check dsl-vlm/page-{page_num}.xml manually.",
                  file=sys.stderr)
            return None


def _clean_vlm_xml_artifacts(page_el):
    """Clean OCR artifacts from VLM XML text content."""
    for elem in page_el.iter():
        if elem.text:
            original = elem.text
            cleaned = re.sub(r"```(?:markdown|xml|json|html)?\s*```", "", original)
            cleaned = re.sub(r"```(?:markdown|xml|json|html)?", "", cleaned)
            cleaned = cleaned.strip()
            if cleaned != original:
                elem.text = cleaned if cleaned else None


# ---------------------------------------------------------------------------
# Per-page merge
# ---------------------------------------------------------------------------


def merge_page(workspace, page_num, ocr_data):
    """Merge a single page deterministically.

    Args:
        workspace: workspace directory path
        page_num: 1-based page number
        ocr_data: full OCR data (list of page lists, 0-indexed)

    Returns:
        XML string of merged page, or None on failure.
    """
    # Load VLM XML
    page_el = load_vlm_xml(workspace, page_num)
    if page_el is None:
        return None

    # Deep copy to avoid modifying original
    page_el = deepcopy(page_el)

    # Clean VLM artifacts first
    _clean_vlm_xml_artifacts(page_el)

    # Get OCR regions for this page (0-indexed)
    page_idx = page_num - 1
    ocr_regions = ocr_data[page_idx] if page_idx < len(ocr_data) else []

    if not ocr_regions:
        print(f"  Page {page_num}: no OCR data, using VLM XML as-is")
        return etree.tostring(page_el, encoding="unicode", pretty_print=True)

    # Step 1: Replace text in headings/paragraphs
    replaced, total = match_and_replace_text(page_el, ocr_regions, page_num)
    print(f"  Page {page_num}: replaced {replaced}/{total} text elements")

    # Step 2: Improve table cell text
    for table_el in page_el.findall(".//table"):
        improved = _try_replace_table_text(table_el, ocr_regions, page_num)
        if improved:
            print(f"  Page {page_num}: improved {improved} table cells")

    # Step 3: Resolve image paths
    img_count = resolve_images(page_el, ocr_regions, page_num)
    if img_count:
        print(f"  Page {page_num}: resolved {img_count} image paths")

    return etree.tostring(page_el, encoding="unicode", pretty_print=True)


# ---------------------------------------------------------------------------
# Fallback: copy VLM XML as-is
# ---------------------------------------------------------------------------


def fallback_copy_vlm(workspace, page_num):
    """Copy VLM XML to dsl/ directory as-is when merge fails."""
    src = Path(workspace) / "dsl-vlm" / f"page-{page_num}.xml"
    dst = Path(workspace) / "dsl" / f"page-{page_num}.xml"
    if src.exists():
        shutil.copy2(src, dst)
        print(f"  Page {page_num}: fallback copy from VLM XML")
    else:
        print(f"  Page {page_num}: WARNING - no VLM XML to copy",
              file=sys.stderr)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def main():
    parser = argparse.ArgumentParser(
        description="Deterministic merge: VLM XML structure + OCR text (no VLM calls)"
    )
    parser.add_argument("--workspace", required=True,
                        help="Workspace directory path")
    parser.add_argument("--pages", required=True, type=int,
                        help="Total page count")
    args = parser.parse_args()

    workspace = args.workspace
    total_pages = args.pages

    # Create output directory
    output_dir = Path(workspace) / "dsl"
    os.makedirs(output_dir, exist_ok=True)

    # Load OCR data
    ocr_data = load_ocr_data(workspace)
    print(f"Loaded OCR data: {len(ocr_data)} pages")

    # Process each page
    success_count = 0
    for page_num in range(1, total_pages + 1):
        merged_xml = merge_page(workspace, page_num, ocr_data)

        if merged_xml:
            out_path = output_dir / f"page-{page_num}.xml"
            with open(out_path, "w", encoding="utf-8") as f:
                f.write(merged_xml)
            success_count += 1
        else:
            fallback_copy_vlm(workspace, page_num)

    print(f"\nDeterministic merge complete: {success_count}/{total_pages} pages merged")
    print(f"Output in: {output_dir}")
    if success_count < total_pages:
        print(f"WARNING: {total_pages - success_count} pages failed merge. Check dsl-vlm/ XML files for those pages.",
              file=sys.stderr)
        print(f"FIX: Re-run vlm_generate_dsl.py for failed pages, then re-run this merge.",
              file=sys.stderr)


if __name__ == "__main__":
    main()
