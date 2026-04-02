#!/usr/bin/env python3
"""deterministic_merge.py - Merge VLM XML with OCR JSON without VLM calls.

Deterministic text replacement merge for weak VLM models:
1. Keep VLM XML structure (layout, styling, element ordering)
2. Replace text with OCR text using sequential matching + text similarity
3. Resolve image paths from OCR cropped images
4. Clean OCR artifacts at merge time

When --pdf-structure is provided (digital PDFs), adds poppler as a third
data source with priority: poppler text > OCR text > VLM text.
Poppler also provides exact font size, color, bold/italic styling.

Designed as a drop-in alternative to vlm_merge_dsl.py when the VLM
is too weak to perform the merge step reliably.

Usage:
    uv run --with lxml \
        .claude/skills/anything-to-docx/scripts/deterministic_merge.py \
        --workspace /path/to/workspace --pages 10

    # With poppler enhancement (digital PDFs):
    uv run --with lxml \
        .claude/skills/anything-to-docx/scripts/deterministic_merge.py \
        --workspace /path/to/workspace --pages 10 \
        --pdf-structure /path/to/pdf-structure.json
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

from constants import (
    IMAGE_JPEG_QUALITY,
    MERGE_ACCEPT_THRESHOLD,
    MERGE_BASE_SIM_THRESHOLD,
    MERGE_DEFAULT_FONT_SIZE_PT,
    MERGE_DUPLICATE_SIMILARITY,
    MERGE_GAP_MIN_CHARS,
    MERGE_HEADING_FONT_SIZES,
    MERGE_IMAGE_IOU_THRESHOLD,
    MERGE_LABEL_BONUS,
    MERGE_LENGTH_RATIO_MAX,
    MERGE_LENGTH_RATIO_SIM_OVERRIDE,
    MERGE_LONG_TEXT_CHAR_WEIGHT,
    MERGE_LONG_TEXT_WORD_WEIGHT,
    MERGE_LOOK_AROUND,
    MERGE_POPPLER_H1_MIN_PTS,
    MERGE_POPPLER_H2_MIN_PTS,
    MERGE_POPPLER_H3_MIN_PTS,
    MERGE_POPPLER_H4_MIN_PTS,
    MERGE_POPPLER_MATCH_THRESHOLD,
    MERGE_PROXIMITY_BONUS_PER_UNIT,
    MERGE_PROXIMITY_DISTANCE,
    MERGE_QUALITY_GATE_MATCH_RATE,
    MERGE_QUALITY_GATE_MIN_OCR_REGIONS,
    MERGE_SHORT_TEXT_CHAR_WEIGHT,
    MERGE_SHORT_TEXT_WORD_THRESHOLD,
    MERGE_SHORT_TEXT_WORD_WEIGHT,
    MERGE_SUPERSET_CONTAINMENT_THRESHOLD,
    MERGE_SUPERSET_MIN_CONTAINED,
    MERGE_TABLE_CELL_IOU_THRESHOLD,
    MERGE_TABLE_CELL_SIMILARITY,
    PAGE_DEFAULT_FONT_CJK,
    PAGE_DEFAULT_FONT_LATIN,
    PAGE_DEFAULT_HEIGHT_PTS,
    PAGE_DEFAULT_MARGIN_CM,
    PAGE_DEFAULT_WIDTH_PTS,
)


# ---------------------------------------------------------------------------
# OCR artifact cleanup (shared pattern with other pipeline scripts)
# ---------------------------------------------------------------------------


def _clean_ocr_text(text):
    """Remove common OCR artifacts from content strings.

    Returns (cleaned_text, heading_level) where heading_level is 0 for
    non-headings, or 1-6 based on markdown # markers in OCR output.
    """
    if not text or not isinstance(text, str):
        return (text or ""), 0
    cleaned = re.sub(r"```(?:markdown|xml|json|html)?\s*```", "", text)
    cleaned = re.sub(r"```(?:markdown|xml|json|html)?\s*\n?\s*```", "", cleaned)
    cleaned = re.sub(r"^```(?:markdown|xml|json|html)?$", "", cleaned, flags=re.MULTILINE)
    cleaned = cleaned.strip()
    if cleaned.startswith("/ ") or cleaned.startswith("\uff0f "):
        cleaned = cleaned[2:]
    # Extract heading level from markdown markers before removing them
    heading_level = 0
    m = re.match(r"^(#{1,6})\s+", cleaned)
    if m:
        heading_level = len(m.group(1))
        cleaned = cleaned[m.end():]
    return cleaned.strip(), heading_level


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
    if word_count < MERGE_SHORT_TEXT_WORD_THRESHOLD:
        return MERGE_SHORT_TEXT_WORD_WEIGHT * word_sim + MERGE_SHORT_TEXT_CHAR_WEIGHT * char_sim
    return MERGE_LONG_TEXT_WORD_WEIGHT * word_sim + MERGE_LONG_TEXT_CHAR_WEIGHT * char_sim


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
    """Parse a bbox attribute string 'x1,y1,x2,y2' into a list of floats.

    Returns None for missing, empty, malformed, or non-numeric bbox strings.
    """
    if not bbox_str:
        return None
    try:
        parts = [float(x.strip()) for x in bbox_str.split(",")]
    except (ValueError, TypeError):
        return None
    return parts if len(parts) == 4 else None


def _dedup_image_regions(image_regions):
    """Remove superset image regions that contain other smaller regions.

    When OCR detects both individual signature boxes AND a large region
    covering all of them, keep the smaller specific regions and drop the
    superset. Uses containment ratio: if >70% of a smaller region fits
    inside a larger one, the larger is considered a superset.

    Returns filtered list of image regions (preserves order).
    """
    if len(image_regions) <= 1:
        return image_regions

    def _area(bb):
        return max(0, bb[2] - bb[0]) * max(0, bb[3] - bb[1])

    def _containment(inner, outer):
        """Fraction of inner bbox area that falls inside outer bbox."""
        ix1 = max(inner[0], outer[0])
        iy1 = max(inner[1], outer[1])
        ix2 = min(inner[2], outer[2])
        iy2 = min(inner[3], outer[3])
        if ix2 <= ix1 or iy2 <= iy1:
            return 0.0
        inter = (ix2 - ix1) * (iy2 - iy1)
        inner_area = _area(inner)
        return inter / inner_area if inner_area > 0 else 0.0

    # Tag each region with its bbox and area
    tagged = []
    for r in image_regions:
        bb = r.get("bbox_2d", [0, 0, 0, 0])
        tagged.append((r, bb, _area(bb)))

    # A region is a "superset" if ≥2 other regions are ≥70% contained in it
    superset_indices = set()
    for i, (_, bb_i, area_i) in enumerate(tagged):
        contained_count = 0
        for j, (_, bb_j, area_j) in enumerate(tagged):
            if i == j:
                continue
            if area_j < area_i and _containment(bb_j, bb_i) > MERGE_SUPERSET_CONTAINMENT_THRESHOLD:
                contained_count += 1
        if contained_count >= MERGE_SUPERSET_MIN_CONTAINED:
            superset_indices.add(i)

    return [r for i, (r, _, _) in enumerate(tagged) if i not in superset_indices]


# ---------------------------------------------------------------------------
# Poppler data loading and matching (digital PDF enhancement)
# ---------------------------------------------------------------------------


def load_poppler_data(pdf_structure_path):
    """Load parsed PDF structure JSON from file.

    Returns the full dict from parse_digital_pdf(), or None if path is
    missing/invalid. Pure function: path in → data out (or None).
    """
    json_path = Path(pdf_structure_path)
    if not json_path.exists():
        print(f"Warning: --pdf-structure file not found: {pdf_structure_path}",
              file=sys.stderr)
        return None
    try:
        raw = json_path.read_text(encoding="utf-8")
        return json.loads(raw)
    except (json.JSONDecodeError, OSError) as e:
        print(f"Warning: Cannot read --pdf-structure: {e}", file=sys.stderr)
        return None


def build_poppler_page_lookup(poppler_data):
    """Build a per-page lookup of poppler paragraphs from parse_digital_pdf() output.

    Returns dict: page_number (1-indexed) → list of paragraph dicts.
    Only includes paragraphs with non-empty text that are NOT table cells.
    Paragraphs are sorted by (top, left) for stable iteration.
    Pure function: data in → lookup out.
    """
    if not poppler_data or not isinstance(poppler_data, dict):
        return {}
    pages = poppler_data.get("pages", [])
    lookup = {}
    for page in pages:
        page_num = page.get("number", 0)
        paras = page.get("paragraphs", [])
        # Filter: non-empty text, not table cells
        filtered = [
            p for p in paras
            if p.get("text", "").strip()
            and not p.get("is_table_cell", False)
        ]
        # Sort by vertical then horizontal position for stable ordering
        filtered.sort(key=lambda p: (p.get("top", 0), p.get("left", 0)))
        if filtered:
            lookup[page_num] = filtered
    return lookup


def _match_poppler_paragraph(vlm_text, poppler_paras, consumed):
    """Find the best matching poppler paragraph for a VLM text element.

    Uses text similarity (same combined_similarity as OCR matching).
    consumed: set of indices already matched (prevents double-assignment).

    Returns (best_index, best_score, best_para) or (-1, 0.0, None) if no match.
    Pure function (reads consumed, does not modify it).
    """
    if not vlm_text or not poppler_paras:
        return -1, 0.0, None

    best_idx = -1
    best_score = 0.0
    best_para = None

    for i, para in enumerate(poppler_paras):
        if i in consumed:
            continue
        sim = combined_similarity(vlm_text, para.get("text", ""))
        if sim > best_score:
            best_score = sim
            best_idx = i
            best_para = para

    if best_score >= MERGE_POPPLER_MATCH_THRESHOLD:
        return best_idx, best_score, best_para
    return -1, 0.0, None


def heading_level_from_poppler(font_size_pts, is_bold):
    """Determine heading level from poppler font size and bold status.

    Mapping:
        >18pt → H1
        >=14pt → H2
        >=12pt → H3
        >=11pt AND bold → H4
        otherwise → 0 (not a heading)

    Pure function.
    """
    if not font_size_pts or font_size_pts <= 0:
        return 0
    if font_size_pts > MERGE_POPPLER_H1_MIN_PTS:
        return 1
    if font_size_pts >= MERGE_POPPLER_H2_MIN_PTS:
        return 2
    if font_size_pts >= MERGE_POPPLER_H3_MIN_PTS:
        return 3
    if font_size_pts >= MERGE_POPPLER_H4_MIN_PTS and is_bold:
        return 4
    return 0


def _apply_poppler_styling(elem, poppler_para):
    """Apply font styling from a poppler paragraph to a VLM XML element.

    Sets on first <run> child (or all runs):
    - font-size-pt from poppler font_size_pts
    - bold from poppler bold
    - italic from poppler italic
    - color-rgb from poppler color (hex → R,G,B; skip black/white)

    Modifies elem in place. Returns count of attributes set.
    """
    runs = elem.findall("run")
    if not runs:
        return 0

    attrs_set = 0
    font_size = poppler_para.get("font_size_pts")
    is_bold = poppler_para.get("bold", False)
    is_italic = poppler_para.get("italic", False)
    color = poppler_para.get("color", "")

    for run in runs:
        # Font size
        if font_size and font_size > 0:
            run.set("font-size-pt", str(round(font_size, 1)))
            attrs_set += 1

        # Bold
        if is_bold:
            run.set("bold", "true")
            attrs_set += 1
        else:
            # Only clear bold if poppler explicitly says not bold
            if run.get("bold") == "true":
                run.attrib.pop("bold", None)
                attrs_set += 1

        # Italic
        if is_italic:
            run.set("italic", "true")
            attrs_set += 1
        else:
            if run.get("italic") == "true":
                run.attrib.pop("italic", None)
                attrs_set += 1

        # Color (hex → R,G,B; skip black #000000 and white #ffffff)
        if color and color.startswith("#") and len(color) == 7:
            color_lower = color.lower()
            if color_lower not in ("#000000", "#ffffff"):
                try:
                    r = int(color[1:3], 16)
                    g = int(color[3:5], 16)
                    b = int(color[5:7], 16)
                    run.set("color-rgb", f"{r},{g},{b}")
                    attrs_set += 1
                except ValueError:
                    pass
            else:
                # Remove color-rgb if poppler says it's black (default)
                run.attrib.pop("color-rgb", None)

    return attrs_set


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
            # Skip header/footer elements — their text comes from VLM, not OCR
            if elem.get("style") in ("footer", "header"):
                continue
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


_HEADING_FONT_SIZES = MERGE_HEADING_FONT_SIZES


def _fix_heading_level(elem, ocr_level):
    """Correct heading level and font-size using OCR's markdown markers."""
    elem.set("level", str(ocr_level))
    target_size = _HEADING_FONT_SIZES.get(ocr_level, MERGE_DEFAULT_FONT_SIZE_PT)
    for run in elem.findall("run"):
        run.set("font-size-pt", target_size)


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


def match_and_replace_text(page_el, ocr_regions, page_num, poppler_paras=None):
    """Match OCR text regions to VLM text elements and replace.

    Uses greedy sequential matching with a look-ahead window.
    When poppler_paras is provided (digital PDF), tries poppler match first:
    priority is poppler text > OCR text > VLM text.
    Modifies page_el in place.

    Returns (replaced_count, total_vlm_elements, unmatched_ocr_regions).
    unmatched_ocr_regions: list of OCR text regions that had no VLM match.
    """
    # Separate OCR regions by type
    text_regions = []
    for r in ocr_regions:
        label = r.get("label", "")
        native = r.get("native_label", "")
        content, heading_level = _clean_ocr_text(r.get("content", ""))
        if not content:
            continue
        if label in ("text",) or native in ("doc_title", "paragraph_title",
                                             "text", "vision_footnote",
                                             "figure_title"):
            # Use native_label as fallback for heading level when ## markers are absent
            if heading_level == 0 and native == "doc_title":
                heading_level = 1
            elif heading_level == 0 and native == "paragraph_title":
                heading_level = 2
            text_regions.append({
                "content": content,
                "native_label": native,
                "bbox": r.get("bbox_2d", [0, 0, 0, 0]),
                "index": r.get("index", 0),
                "heading_level": heading_level,
            })

    # Sort OCR text regions by vertical position (top to bottom)
    text_regions.sort(key=lambda r: (r["bbox"][1], r["bbox"][0]))

    # Collect VLM text elements
    vlm_elements = collect_text_elements(page_el)

    if not vlm_elements or not text_regions:
        return 0, len(vlm_elements), []

    # Poppler matching state (when digital PDF data available)
    poppler_consumed = set()
    poppler_matched_count = 0

    # Greedy sequential matching with symmetric look-around
    consumed = set()
    ocr_cursor = 0
    replaced = 0

    for ve in vlm_elements:
        vlm_text = ve["text"]
        if not vlm_text.strip():
            continue

        # ── Poppler-first path (digital PDF enhancement) ──
        # Try poppler match before OCR. Poppler text is exact (zero errors).
        if poppler_paras:
            p_idx, p_score, p_para = _match_poppler_paragraph(
                vlm_text, poppler_paras, poppler_consumed)
            if p_idx >= 0 and p_para:
                # Use poppler's exact text
                _replace_element_text(ve["element"], p_para["text"])
                # Apply font styling from poppler
                _apply_poppler_styling(ve["element"], p_para)
                # Override heading level from poppler font size
                p_font_size = p_para.get("font_size_pts", 0)
                p_bold = p_para.get("bold", False)
                p_hlevel = heading_level_from_poppler(p_font_size, p_bold)
                if p_hlevel > 0 and ve["type"] == "heading":
                    ve["element"].set("level", str(p_hlevel))
                poppler_consumed.add(p_idx)
                poppler_matched_count += 1
                replaced += 1
                # Still consume the matching OCR region to keep cursor aligned
                # (prevents OCR from being gap-filled for already-replaced elements)
                for i, ocr_r in enumerate(text_regions):
                    if i in consumed:
                        continue
                    if combined_similarity(p_para["text"], ocr_r["content"]) > MERGE_ACCEPT_THRESHOLD:
                        consumed.add(i)
                        if i >= ocr_cursor:
                            ocr_cursor = i + 1
                        break
                continue

        # ── OCR path (existing behavior) ──
        best_score = 0.0
        best_idx = -1

        search_start = max(0, ocr_cursor - MERGE_LOOK_AROUND)
        search_end = min(len(text_regions), ocr_cursor + MERGE_LOOK_AROUND)

        for i in range(search_start, search_end):
            if i in consumed:
                continue

            ocr_r = text_regions[i]
            base_sim = combined_similarity(vlm_text, ocr_r["content"])

            # Minimum base similarity — prevent bonuses from promoting bad matches
            if base_sim < MERGE_BASE_SIM_THRESHOLD:
                continue

            # Reject matches where text lengths differ wildly.
            len_vlm = max(len(vlm_text), 1)
            len_ocr = max(len(ocr_r["content"]), 1)
            length_ratio = max(len_vlm, len_ocr) / min(len_vlm, len_ocr)
            if length_ratio > MERGE_LENGTH_RATIO_MAX and base_sim < MERGE_LENGTH_RATIO_SIM_OVERRIDE:
                continue

            # OCR heading regions must match VLM heading elements (not paragraphs)
            ocr_hlevel = ocr_r.get("heading_level", 0)
            if ocr_hlevel > 0 and ve["type"] != "heading":
                continue

            sim = base_sim

            # Label matching bonus
            if _label_matches_type(ocr_r["native_label"], ve["type"]):
                sim += MERGE_LABEL_BONUS

            # Proximity bonus (prefer closer to cursor)
            distance = abs(i - ocr_cursor)
            proximity_bonus = MERGE_PROXIMITY_BONUS_PER_UNIT * max(0, MERGE_PROXIMITY_DISTANCE - distance)
            sim += proximity_bonus

            if sim > best_score:
                best_score = sim
                best_idx = i

        # Threshold for accepting a match
        if best_idx >= 0 and best_score > MERGE_ACCEPT_THRESHOLD:
            _replace_element_text(ve["element"], text_regions[best_idx]["content"])
            # Fix heading level from OCR markers (OCR knows ## = h2, etc.)
            ocr_hlevel = text_regions[best_idx].get("heading_level", 0)
            if ocr_hlevel > 0 and ve["type"] == "heading":
                _fix_heading_level(ve["element"], ocr_hlevel)
            consumed.add(best_idx)
            ocr_cursor = best_idx + 1
            replaced += 1

    if poppler_matched_count:
        print(f"  Page {page_num}: {poppler_matched_count} elements matched via poppler (exact text)")

    # Collect unmatched OCR text regions
    unmatched = [text_regions[i] for i in range(len(text_regions)) if i not in consumed]
    return replaced, len(vlm_elements), unmatched


# ---------------------------------------------------------------------------
# Gap filling: append OCR regions that VLM missed
# ---------------------------------------------------------------------------


def _append_gap_elements(page_el, unmatched_regions):
    """Append unmatched OCR text regions as new elements at end of page.

    Only appends regions with substantial content (>5 chars) to avoid noise.
    Skips regions whose text already exists in the page (dedup against
    poppler-matched or VLM-retained elements to prevent duplicates).
    Creates heading or paragraph elements based on OCR heading_level.
    """
    # Build set of existing text on the page for dedup
    existing_texts = set()
    for elem in page_el.iter():
        if elem.tag in ("heading", "paragraph"):
            runs = elem.findall("run")
            if runs:
                text = " ".join((r.text or "") for r in runs).strip()
            else:
                text = (elem.text or "").strip()
            if text:
                existing_texts.add(text)

    added = 0
    for region in unmatched_regions:
        content = region["content"]
        if len(content) <= MERGE_GAP_MIN_CHARS:
            continue
        # Skip if this text (or very similar) already exists on the page
        if content in existing_texts:
            continue
        # Also check similarity — poppler text may differ slightly from OCR text
        skip = False
        for existing in existing_texts:
            if combined_similarity(content, existing) > MERGE_DUPLICATE_SIMILARITY:
                skip = True
                break
        if skip:
            continue

        hlevel = region.get("heading_level", 0)
        if hlevel > 0:
            elem = etree.SubElement(page_el, "heading",
                                    level=str(hlevel), alignment="left")
            size = _HEADING_FONT_SIZES.get(hlevel, MERGE_DEFAULT_FONT_SIZE_PT)
            run = etree.SubElement(elem, "run",
                                   attrib={"font-size-pt": size, "bold": "true"})
        else:
            elem = etree.SubElement(page_el, "paragraph", alignment="left")
            run = etree.SubElement(elem, "run",
                                   attrib={"font-size-pt": MERGE_DEFAULT_FONT_SIZE_PT})
        run.text = content
        existing_texts.add(content)  # prevent further dupes within gap regions
        added += 1
    return added


# ---------------------------------------------------------------------------
# PDF-extracted image support
# ---------------------------------------------------------------------------


def _build_pdf_images_map(workspace, total_pages):
    """Build a map of PDF-extracted original images for all pages.

    Scans pdf-images/ directory once. Filters out:
    1. smask (alpha mask) files (grayscale mode 'L')
    2. Repeating images that appear on 3+ pages (headers, footers, watermarks)
       Detected by file-size fingerprint — same embedded image = same bytes.

    Returns dict: page_num -> [(path, width, height)]
    """
    pdf_dir = Path(workspace) / "pdf-images"
    if not pdf_dir.exists():
        return {}

    import glob as _glob
    import hashlib
    from PIL import Image as _PILImage

    # Collect all non-smask images grouped by page
    all_images = {}  # page_num -> [(path, width, height, file_hash)]
    for page_num in range(1, total_pages + 1):
        pattern = str(pdf_dir / f"img-{page_num:03d}-*.png")
        candidates = sorted(_glob.glob(pattern))
        page_imgs = []
        for path in candidates:
            try:
                img = _PILImage.open(path)
                if img.mode != "L":  # skip smask
                    # Hash by file size + dimensions (fast, good enough for dedup)
                    fsize = os.path.getsize(path)
                    fhash = f"{fsize}_{img.width}x{img.height}"
                    page_imgs.append((path, img.width, img.height, fhash))
                img.close()
            except Exception:
                continue
        all_images[page_num] = page_imgs

    # Count how many pages each image fingerprint appears on
    hash_page_count = {}
    for page_num, imgs in all_images.items():
        for _, _, _, fhash in imgs:
            if fhash not in hash_page_count:
                hash_page_count[fhash] = set()
            hash_page_count[fhash].add(page_num)

    # Filter: exclude images appearing on 3+ pages (repeating headers/footers)
    repeat_threshold = min(3, max(2, total_pages // 2))
    repeating_hashes = {h for h, pages in hash_page_count.items()
                        if len(pages) >= repeat_threshold}

    result = {}
    for page_num, imgs in all_images.items():
        filtered = [(p, w, h) for p, w, h, fhash in imgs
                     if fhash not in repeating_hashes]
        if filtered:
            result[page_num] = filtered

    if repeating_hashes:
        total_filtered = sum(1 for imgs in all_images.values()
                             for _, _, _, fh in imgs if fh in repeating_hashes)
        print(f"  PDF images: filtered {total_filtered} repeating header/footer images across pages")

    return result


def _match_pdf_image(pdf_images, ocr_bbox, page_width_pts=612, page_height_pts=792):
    """Find the best PDF-extracted image matching an OCR bbox by aspect ratio.

    OCR bbox is [x1, y1, x2, y2] normalized 0-1000.
    PDF image has (path, pixel_width, pixel_height).

    Returns (path, index) of best match, or (None, -1) if no good match.
    """
    if not pdf_images or not ocr_bbox:
        return None, -1

    # OCR bbox aspect ratio (in page-relative units)
    ocr_w = max(ocr_bbox[2] - ocr_bbox[0], 1)
    ocr_h = max(ocr_bbox[3] - ocr_bbox[1], 1)
    # Adjust for page aspect ratio (bbox is normalized, page may not be square)
    ocr_aspect = (ocr_w / 1000 * page_width_pts) / (ocr_h / 1000 * page_height_pts)

    best_idx = -1
    best_score = float("inf")
    for i, (path, pw, ph) in enumerate(pdf_images):
        if ph == 0:
            continue
        pdf_aspect = pw / ph
        # Score = how different the aspect ratios are (lower = better)
        score = abs(pdf_aspect - ocr_aspect) / max(pdf_aspect, ocr_aspect, 0.01)
        if score < best_score:
            best_score = score
            best_idx = i

    # Accept if aspect ratio is within 50% — generous threshold for different crops
    if best_idx >= 0 and best_score < 0.5:
        return pdf_images[best_idx][0], best_idx
    return None, -1


# ---------------------------------------------------------------------------
# Image resolution
# ---------------------------------------------------------------------------


def resolve_images(page_el, ocr_regions, page_num, workspace=None,
                    pdf_images_map=None):
    """Set image src paths from OCR image regions by bbox matching.

    For each VLM <image> element, find the best matching OCR image
    region by bbox IoU and set the src to the OCR cropped image path.
    When PDF-extracted originals are available, prefer them over OCR crops.

    Args:
        pdf_images_map: dict from _build_pdf_images_map(), or None.

    Modifies page_el in place. Returns count of resolved images.
    """
    raw_image_regions = [
        r for r in ocr_regions
        if r.get("label") == "image"
    ]
    image_regions = _dedup_image_regions(raw_image_regions)
    if len(image_regions) < len(raw_image_regions):
        print(f"  Page {page_num}: deduped {len(raw_image_regions) - len(image_regions)} superset image regions")

    vlm_images = page_el.findall(".//image")
    resolved = 0

    # PDF-extracted originals for this page (repeating images already filtered)
    pdf_images = (pdf_images_map or {}).get(page_num, [])
    pdf_used = set()  # Track consumed pdf-images by index

    # Get page dimensions from page_el for aspect ratio calculation
    page_width_pts = float(page_el.get("width-pts", "612"))
    page_height_pts = float(page_el.get("height-pts", "792"))

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

            if best_region and best_iou > MERGE_IMAGE_IOU_THRESHOLD:
                page_idx = page_num - 1  # 0-based
                img_idx = best_region.get("_image_idx", best_region.get("index", 0))
                ocr_src = f"ocr-output/input/imgs/cropped_page{page_idx}_idx{img_idx}.jpg"

                # Replace VLM bbox with OCR bbox — OCR is more accurate for sizing
                ocr_bbox = best_region.get("bbox_2d")
                if ocr_bbox and len(ocr_bbox) == 4:
                    img_el.set("bbox", f"{ocr_bbox[0]},{ocr_bbox[1]},{ocr_bbox[2]},{ocr_bbox[3]}")

                # Try to upgrade to PDF-extracted original
                src = ocr_src
                if pdf_images and ocr_bbox:
                    available = [(p, w, h) for i, (p, w, h) in enumerate(pdf_images)
                                 if i not in pdf_used]
                    pdf_path, pdf_idx = _match_pdf_image(
                        available, ocr_bbox, page_width_pts, page_height_pts)
                    if pdf_path:
                        src = os.path.relpath(pdf_path, workspace)
                        # Find original index to mark as used
                        for i, (p, _, _) in enumerate(pdf_images):
                            if p == pdf_path:
                                pdf_used.add(i)
                                break

                img_el.set("src", src)
                resolved += 1
                continue

        # Fallback: sequential assignment
        if image_regions:
            region = image_regions.pop(0)
            page_idx = page_num - 1
            img_idx = region.get("_image_idx", region.get("index", 0))
            ocr_src = f"ocr-output/input/imgs/cropped_page{page_idx}_idx{img_idx}.jpg"

            # Replace VLM bbox with OCR bbox for correct sizing
            ocr_bbox = region.get("bbox_2d")
            if ocr_bbox and len(ocr_bbox) == 4:
                img_el.set("bbox", f"{ocr_bbox[0]},{ocr_bbox[1]},{ocr_bbox[2]},{ocr_bbox[3]}")

            # Try to upgrade to PDF-extracted original
            src = ocr_src
            if pdf_images and ocr_bbox:
                available = [(p, w, h) for i, (p, w, h) in enumerate(pdf_images)
                             if i not in pdf_used]
                pdf_path, pdf_idx = _match_pdf_image(
                    available, ocr_bbox, page_width_pts, page_height_pts)
                if pdf_path:
                    src = os.path.relpath(pdf_path, workspace)
                    for i, (p, _, _) in enumerate(pdf_images):
                        if p == pdf_path:
                            pdf_used.add(i)
                            break

            img_el.set("src", src)
            resolved += 1

    if pdf_used:
        print(f"  Page {page_num}: upgraded {len(pdf_used)} images to PDF originals")

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
        if iou > MERGE_TABLE_CELL_IOU_THRESHOLD:
            content, _ = _clean_ocr_text(r.get("content", ""))
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

            if best_content and best_sim > MERGE_TABLE_CELL_SIMILARITY:
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
# OCR-only page generation (fallback when VLM XML is missing)
# ---------------------------------------------------------------------------


def _generate_page_from_ocr(ocr_regions, page_num, workspace=None,
                            page_width=PAGE_DEFAULT_WIDTH_PTS,
                            page_height=PAGE_DEFAULT_HEIGHT_PTS):
    """Generate a <page> element purely from OCR data when VLM XML is missing.

    Creates headings, paragraphs, and image placeholders based on OCR regions.
    Uses the same heading-level logic as the merge path (markdown markers +
    native_label fallback).
    """
    page_el = etree.Element("page")
    page_el.set("number", str(page_num))
    page_el.set("width-pts", str(page_width))
    page_el.set("height-pts", str(page_height))
    page_el.set("margin-top-cm", PAGE_DEFAULT_MARGIN_CM)
    page_el.set("margin-bottom-cm", PAGE_DEFAULT_MARGIN_CM)
    page_el.set("margin-left-cm", PAGE_DEFAULT_MARGIN_CM)
    page_el.set("margin-right-cm", PAGE_DEFAULT_MARGIN_CM)
    page_el.set("font-latin", PAGE_DEFAULT_FONT_LATIN)
    page_el.set("font-cjk", PAGE_DEFAULT_FONT_CJK)

    page_idx = page_num - 1  # 0-based for image path construction

    # Dedup overlapping image regions (e.g., OCR detects both individual
    # signature boxes AND a superset region covering all of them)
    image_regions = [r for r in ocr_regions if r and r.get("label") == "image"]
    deduped_images = _dedup_image_regions(image_regions)
    if len(deduped_images) < len(image_regions):
        # Replace ocr_regions with deduped version
        dropped = set(id(r) for r in image_regions) - set(id(r) for r in deduped_images)
        ocr_regions = [r for r in ocr_regions if r is None or r.get("label") != "image" or id(r) not in dropped]

    # Pre-discover actual image files for this page (glob is more reliable
    # than computing from JSON index, since glm-ocr uses its own numbering)
    import glob as _glob
    img_pattern = str(Path(workspace) / f"ocr-output/input/imgs/cropped_page{page_idx}_idx*.jpg") if workspace else ""
    available_imgs = sorted(_glob.glob(img_pattern)) if img_pattern else []
    img_queue = list(available_imgs)  # consume sequentially

    for region in ocr_regions:
        if region is None:
            continue

        label = region.get("label", "text")
        native = region.get("native_label", "")
        content = region.get("content") or ""
        bbox = region.get("bbox_2d")

        # --- Image region ---
        if label == "image":
            img_el = etree.SubElement(page_el, "image")
            if img_queue:
                # Use next available file, make path relative to workspace
                abs_path = img_queue.pop(0)
                src = os.path.relpath(abs_path, workspace) if workspace else abs_path
            else:
                src = "PLACEHOLDER"
            img_el.set("src", src)
            if bbox:
                img_el.set("bbox", f"{bbox[0]},{bbox[1]},{bbox[2]},{bbox[3]}")
            img_el.set("page-width-pts", str(page_width))
            continue

        # --- Text region ---
        cleaned, heading_level = _clean_ocr_text(content)
        if not cleaned:
            continue

        # native_label fallback for heading level
        if heading_level == 0 and native == "doc_title":
            heading_level = 1
        elif heading_level == 0 and native == "paragraph_title":
            heading_level = 2

        if heading_level > 0:
            elem = etree.SubElement(page_el, "heading",
                                    level=str(heading_level), alignment="left")
            size = _HEADING_FONT_SIZES.get(heading_level, MERGE_DEFAULT_FONT_SIZE_PT)
            run = etree.SubElement(elem, "run",
                                   attrib={"font-size-pt": size, "bold": "true"})
        else:
            elem = etree.SubElement(page_el, "paragraph", alignment="left")
            run = etree.SubElement(elem, "run",
                                   attrib={"font-size-pt": MERGE_DEFAULT_FONT_SIZE_PT})
        run.text = cleaned

    return page_el


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


def _reorder_by_poppler_position(page_el, poppler_paras):
    """Reorder page elements using poppler's exact vertical positions.

    For digital PDFs, poppler knows the exact (top, left) of every text element.
    VLM often misordering elements (e.g., putting list item 4 before the title).
    This function corrects the ordering deterministically.

    Algorithm:
    1. For each child element of <page>, try to match its text to a poppler paragraph
    2. Assign the poppler paragraph's top_pts as the sort key
    3. Elements without a poppler match keep a sort key based on their current position
    4. Sort all children by (sort_key, original_index) for stable ordering

    Only reorders direct children of <page> (headings, paragraphs, tables, images).
    Does NOT reorder cells within tables.

    Returns count of elements whose position changed.
    """
    children = list(page_el)
    if len(children) <= 1:
        return 0

    # Build poppler text → position lookup
    poppler_by_text = {}
    for para in poppler_paras:
        text = para.get("text", "").strip()
        if text:
            # Use first 100 chars as key (sufficient for matching)
            key = text[:100]
            top_px = para.get("top", 0)
            poppler_by_text[key] = top_px

    # Assign sort keys to each child
    sort_entries = []  # (sort_key, original_index, element)
    for i, child in enumerate(children):
        sort_key = None

        # Try to match element text to poppler
        tag = child.tag
        if tag in ("heading", "paragraph"):
            # Get text from run elements
            runs = child.findall("run")
            if runs:
                elem_text = " ".join(r.text or "" for r in runs).strip()
            else:
                elem_text = (child.text or "").strip()

            if elem_text:
                # Try exact prefix match (first 100 chars)
                key = elem_text[:100]
                if key in poppler_by_text:
                    sort_key = poppler_by_text[key]
                else:
                    # Try similarity match for imperfect VLM text
                    best_match_top = None
                    best_sim = 0
                    for p_key, p_top in poppler_by_text.items():
                        sim = combined_similarity(elem_text, p_key)
                        if sim > best_sim and sim > 0.4:
                            best_sim = sim
                            best_match_top = p_top
                    if best_match_top is not None:
                        sort_key = best_match_top

        elif tag == "table":
            # Use table bbox top as sort key
            bbox_str = child.get("bbox", "")
            if bbox_str:
                parts = bbox_str.split(",")
                if len(parts) >= 2:
                    try:
                        # bbox is normalized 0-1000; map back to px approx
                        sort_key = float(parts[1]) * 1.2  # rough approx
                    except ValueError:
                        pass

        elif tag == "image":
            bbox_str = child.get("bbox", "")
            if bbox_str:
                parts = bbox_str.split(",")
                if len(parts) >= 2:
                    try:
                        sort_key = float(parts[1]) * 1.2
                    except ValueError:
                        pass

        # Fallback: use a large number to keep at current relative position
        if sort_key is None:
            sort_key = 10000 + i

        sort_entries.append((sort_key, i, child))

    # Sort by (sort_key, original_index) for stability
    sort_entries.sort(key=lambda x: (x[0], x[1]))

    # Check how many elements changed position
    new_order = [e[2] for e in sort_entries]
    changes = sum(1 for a, b in zip(children, new_order) if a is not b)

    if changes == 0:
        return 0

    # Rebuild page element with new order
    for child in children:
        page_el.remove(child)
    for child in new_order:
        page_el.append(child)

    return changes


def _reorder_by_ocr_position(page_el, ocr_regions):
    """Reorder page elements using OCR bbox_2d vertical positions.

    For non-digital PDFs (no poppler data), OCR regions have precise bbox_2d
    coordinates from layout detection. Use these to fix VLM element ordering.

    Algorithm:
    1. For each child element of <page>, match its text to an OCR region
    2. Use the OCR region's bbox_2d[1] (y1, top) as sort key
    3. Elements without an OCR match keep their current relative position
    4. Sort all children by (sort_key, original_index)

    Returns count of elements whose position changed.
    """
    children = list(page_el)
    if len(children) <= 1 or not ocr_regions:
        return 0

    # Build OCR text → y1 position lookup
    ocr_by_text = {}
    for region in ocr_regions:
        content = (region.get("content", "") or "").strip()
        bbox = region.get("bbox_2d", [])
        if content and len(bbox) >= 2:
            # Use first 100 chars as key
            key = content[:100]
            ocr_by_text[key] = bbox[1]  # y1 (top position, 0-1000 normalized)

    # Assign sort keys
    sort_entries = []
    for i, child in enumerate(children):
        sort_key = None
        tag = child.tag

        if tag in ("heading", "paragraph"):
            runs = child.findall("run")
            if runs:
                elem_text = " ".join((r.text or "") for r in runs).strip()
            else:
                elem_text = (child.text or "").strip()

            if elem_text:
                key = elem_text[:100]
                if key in ocr_by_text:
                    sort_key = ocr_by_text[key]
                else:
                    # Similarity match for imperfect VLM text
                    best_y = None
                    best_sim = 0
                    for o_key, o_y in ocr_by_text.items():
                        sim = combined_similarity(elem_text, o_key)
                        if sim > best_sim and sim > 0.4:
                            best_sim = sim
                            best_y = o_y
                    if best_y is not None:
                        sort_key = best_y

        elif tag == "table":
            bbox_str = child.get("bbox", "")
            if bbox_str:
                parts = bbox_str.split(",")
                if len(parts) >= 2:
                    try:
                        sort_key = float(parts[1])  # already 0-1000 normalized
                    except ValueError:
                        pass

        elif tag == "image":
            bbox_str = child.get("bbox", "")
            if bbox_str:
                parts = bbox_str.split(",")
                if len(parts) >= 2:
                    try:
                        sort_key = float(parts[1])
                    except ValueError:
                        pass

        if sort_key is None:
            sort_key = 10000 + i

        sort_entries.append((sort_key, i, child))

    sort_entries.sort(key=lambda x: (x[0], x[1]))

    new_order = [e[2] for e in sort_entries]
    changes = sum(1 for a, b in zip(children, new_order) if a is not b)

    if changes == 0:
        return 0

    for child in children:
        page_el.remove(child)
    for child in new_order:
        page_el.append(child)

    return changes


def _cleanup_placeholder_images(page_el):
    """Remove all <image> elements with unresolved PLACEHOLDER/empty src.

    Returns the count of removed elements. Pure function on the page element
    tree — safe to call on any code path before serialization.
    """
    removed = 0
    for img_el in page_el.findall(".//image"):
        if img_el.get("src") in ("PLACEHOLDER", "", None):
            img_el.getparent().remove(img_el)
            removed += 1
    return removed


def _self_crop_placeholder_images(page_el, workspace, page_num):
    """Crop PLACEHOLDER images from the page image using their bbox.

    When VLM detects an image region (e.g., a logo, watermark, or photo)
    but OCR didn't crop it, we can still recover the image by cropping
    directly from the rendered page image.

    Only touches <image> elements where src is still PLACEHOLDER and a
    valid bbox attribute exists. Leaves src unchanged if cropping fails
    (missing page image, bad bbox, zero-area crop) so that
    _cleanup_placeholder_images can remove them later.

    Args:
        page_el: lxml Element — the <page> being processed
        workspace: workspace directory path (str or Path)
        page_num: 1-based page number

    Returns:
        Count of images successfully self-cropped.
    """
    from PIL import Image  # lazy import — Pillow not always in --with

    page_img_path = Path(workspace) / "input-images" / f"page-{page_num}.png"
    if not page_img_path.exists():
        return 0

    # Collect PLACEHOLDER images that have a bbox
    placeholder_imgs = []
    for img_el in page_el.findall(".//image"):
        if img_el.get("src") not in ("PLACEHOLDER", "", None):
            continue
        bbox = _parse_bbox_attr(img_el.get("bbox"))
        if bbox is None:
            continue
        placeholder_imgs.append((img_el, bbox))

    if not placeholder_imgs:
        return 0

    # Open page image once for all crops on this page
    try:
        page_img = Image.open(page_img_path)
    except Exception as e:
        print(f"  Page {page_num}: cannot open page image for self-crop: {e}",
              file=sys.stderr)
        return 0

    img_width, img_height = page_img.size

    # Create output directory on first successful crop
    crop_dir = Path(workspace) / "self-cropped"
    crop_dir_created = False

    cropped_count = 0
    for img_el, bbox in placeholder_imgs:
        # bbox is normalized 0-1000 -> convert to pixel coordinates
        x1, y1, x2, y2 = bbox
        pixel_x1 = int(x1 / 1000 * img_width)
        pixel_y1 = int(y1 / 1000 * img_height)
        pixel_x2 = int(x2 / 1000 * img_width)
        pixel_y2 = int(y2 / 1000 * img_height)

        # Validate: non-zero area
        if pixel_x2 <= pixel_x1 or pixel_y2 <= pixel_y1:
            continue

        try:
            cropped = page_img.crop((pixel_x1, pixel_y1, pixel_x2, pixel_y2))
        except Exception:
            continue

        if not crop_dir_created:
            crop_dir.mkdir(parents=True, exist_ok=True)
            crop_dir_created = True

        # Filename encodes page + bbox for uniqueness and debuggability
        bbox_tag = f"{int(x1)}_{int(y1)}_{int(x2)}_{int(y2)}"
        filename = f"page{page_num}_bbox{bbox_tag}.jpg"
        crop_path = crop_dir / filename

        try:
            cropped.save(str(crop_path), "JPEG", quality=IMAGE_JPEG_QUALITY)
        except Exception:
            continue

        # src is relative to workspace (same convention as OCR images)
        img_el.set("src", f"self-cropped/{filename}")
        cropped_count += 1

    page_img.close()
    return cropped_count


def _dedup_consecutive_elements(page_el):
    """Remove consecutive duplicate heading/paragraph elements.

    Weak VLMs sometimes repeat the same content. If two nearby text elements
    have >80% text similarity, remove the less-styled one (prefer heading
    over paragraph, larger font over smaller).
    """
    removed = 0
    children = list(page_el)
    prev_elem = None
    prev_text = None
    for child in children:
        if child.tag not in ("heading", "paragraph"):
            prev_elem = None
            prev_text = None
            continue
        text = _get_element_text(child)
        if prev_text and text:
            sim = combined_similarity(prev_text, text)
            if sim > MERGE_DUPLICATE_SIMILARITY:
                # Remove the less-styled duplicate: prefer heading over paragraph
                if prev_elem.tag == "heading" and child.tag == "paragraph":
                    page_el.remove(child)
                elif child.tag == "heading" and prev_elem.tag == "paragraph":
                    page_el.remove(prev_elem)
                    prev_elem = child
                    prev_text = text
                else:
                    page_el.remove(child)
                removed += 1
                continue
        prev_elem = child
        prev_text = text
    return removed


# ---------------------------------------------------------------------------
# Per-page merge
# ---------------------------------------------------------------------------


def merge_page(workspace, page_num, ocr_data, pdf_images_map=None,
               poppler_page_lookup=None):
    """Merge a single page deterministically.

    Args:
        workspace: workspace directory path
        page_num: 1-based page number
        ocr_data: full OCR data (list of page lists, 0-indexed)
        pdf_images_map: dict from _build_pdf_images_map(), or None
        poppler_page_lookup: dict from build_poppler_page_lookup(), or None.
            When provided (digital PDF), poppler text/styling enhances merge.

    Returns:
        XML string of merged page, or None on failure.
    """
    # Load VLM XML
    page_el = load_vlm_xml(workspace, page_num)
    if page_el is None:
        # Fallback: generate page purely from OCR data
        page_idx = page_num - 1
        ocr_regions = ocr_data[page_idx] if page_idx < len(ocr_data) else []
        if ocr_regions:
            print(f"  Page {page_num}: no VLM XML — generating from OCR ({len(ocr_regions)} regions)")
            page_el = _generate_page_from_ocr(ocr_regions, page_num, workspace=workspace)
            # Clean PLACEHOLDERs from OCR-generated page (more image regions than cropped files)
            ph_removed = _cleanup_placeholder_images(page_el)
            if ph_removed:
                print(f"  Page {page_num}: removed {ph_removed} unresolved placeholder images")
            return etree.tostring(page_el, encoding="unicode", pretty_print=True)
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
        # Try self-cropping VLM images that have bbox (no OCR crops available)
        self_cropped = _self_crop_placeholder_images(page_el, workspace, page_num)
        if self_cropped:
            print(f"  Page {page_num}: self-cropped {self_cropped} images from page image")
        # Still clean remaining PLACEHOLDERs — VLM may have emitted unresolvable refs
        ph_removed = _cleanup_placeholder_images(page_el)
        if ph_removed:
            print(f"  Page {page_num}: removed {ph_removed} unresolved placeholder images")
        return etree.tostring(page_el, encoding="unicode", pretty_print=True)

    # Get poppler paragraphs for this page (if available)
    poppler_paras = (poppler_page_lookup or {}).get(page_num)

    # Step 1: Replace text in headings/paragraphs
    replaced, total, unmatched = match_and_replace_text(
        page_el, ocr_regions, page_num, poppler_paras=poppler_paras)
    print(f"  Page {page_num}: replaced {replaced}/{total} text elements")

    # Quality gate: if VLM match rate is very low, OCR-only page is likely better
    # (VLM structure is too broken to be useful as scaffold)
    ocr_text_regions = [r for r in ocr_regions if r and r.get("label") != "image"]
    if total > 0 and replaced / total < MERGE_QUALITY_GATE_MATCH_RATE and len(ocr_text_regions) >= MERGE_QUALITY_GATE_MIN_OCR_REGIONS:
        print(f"  Page {page_num}: match rate {replaced/total:.0%} < 30% — switching to OCR-only")
        page_el = _generate_page_from_ocr(ocr_regions, page_num, workspace=workspace)
        # Clean PLACEHOLDERs from OCR-generated page (more image regions than cropped files)
        ph_removed = _cleanup_placeholder_images(page_el)
        if ph_removed:
            print(f"  Page {page_num}: removed {ph_removed} unresolved placeholder images")
        return etree.tostring(page_el, encoding="unicode", pretty_print=True)

    # Step 2: Improve table cell text
    for table_el in page_el.findall(".//table"):
        improved = _try_replace_table_text(table_el, ocr_regions, page_num)
        if improved:
            print(f"  Page {page_num}: improved {improved} table cells")

    # Step 3: Resolve image paths
    img_count = resolve_images(page_el, ocr_regions, page_num,
                               workspace=workspace,
                               pdf_images_map=pdf_images_map)
    if img_count:
        print(f"  Page {page_num}: resolved {img_count} image paths")

    # Step 3b: Self-crop remaining PLACEHOLDER images from page image
    self_cropped = _self_crop_placeholder_images(page_el, workspace, page_num)
    if self_cropped:
        print(f"  Page {page_num}: self-cropped {self_cropped} images from page image")

    # Step 4: Dedup — remove consecutive duplicate text after OCR replacement
    deduped = _dedup_consecutive_elements(page_el)
    if deduped:
        print(f"  Page {page_num}: removed {deduped} duplicate elements")

    # Step 5: Gap filling — append OCR regions VLM missed
    if unmatched:
        gap_count = _append_gap_elements(page_el, unmatched)
        if gap_count:
            print(f"  Page {page_num}: appended {gap_count} OCR gap elements")

    # Step 6: Remove unresolved PLACEHOLDER images (avoids "[Image missing]" in DOCX)
    placeholder_removed = _cleanup_placeholder_images(page_el)
    if placeholder_removed:
        print(f"  Page {page_num}: removed {placeholder_removed} unresolved placeholder images")

    # Step 7: Reorder elements using precise position data
    # VLM often misordering elements — use poppler or OCR positions to fix
    if poppler_paras:
        reordered = _reorder_by_poppler_position(page_el, poppler_paras)
        if reordered:
            print(f"  Page {page_num}: reordered {reordered} elements by poppler position")
    elif ocr_regions:
        reordered = _reorder_by_ocr_position(page_el, ocr_regions)
        if reordered:
            print(f"  Page {page_num}: reordered {reordered} elements by OCR position")

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
    parser.add_argument("--pdf-structure", type=str, default=None,
                        help="Path to JSON output of parse_pdf_structure.py (digital PDF). "
                             "When provided, uses exact poppler text/styling to enhance merge.")
    args = parser.parse_args()

    workspace = args.workspace
    total_pages = args.pages

    # Create output directory
    output_dir = Path(workspace) / "dsl"
    os.makedirs(output_dir, exist_ok=True)

    # Load OCR data
    ocr_data = load_ocr_data(workspace)
    print(f"Loaded OCR data: {len(ocr_data)} pages")

    # Build PDF-extracted images map (if available, for PDF inputs)
    pdf_images_map = _build_pdf_images_map(workspace, total_pages)
    if pdf_images_map:
        total_pdf = sum(len(v) for v in pdf_images_map.values())
        print(f"PDF images: {total_pdf} content images across {len(pdf_images_map)} pages")

    # Load poppler data when --pdf-structure is provided (digital PDFs)
    poppler_page_lookup = None
    if args.pdf_structure:
        poppler_data = load_poppler_data(args.pdf_structure)
        if poppler_data:
            poppler_page_lookup = build_poppler_page_lookup(poppler_data)
            total_paras = sum(len(v) for v in poppler_page_lookup.values())
            print(f"Poppler data: {total_paras} paragraphs across "
                  f"{len(poppler_page_lookup)} pages (digital PDF enhancement)")

    # Process each page
    success_count = 0
    for page_num in range(1, total_pages + 1):
        merged_xml = merge_page(workspace, page_num, ocr_data,
                                pdf_images_map=pdf_images_map,
                                poppler_page_lookup=poppler_page_lookup)

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
