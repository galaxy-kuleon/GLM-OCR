#!/usr/bin/env python3
"""Apply translations back to a DOCX file.

Reads the original DOCX and a translations JSON file, replaces text in the
corresponding XML elements, and writes a new DOCX with translated content.
All formatting, structure, and embedded objects are preserved (zero-loss
round-trip).

Usage:
    uv run --with lxml \
        .claude/skills/another-pure-pure-docx-translate-to-docx/scripts/apply_docx_translations.py \
        --input INPUT.docx \
        --translations workspace/translations.json \
        --output workspace/translated-output.docx
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import zipfile
from io import BytesIO
from pathlib import Path

from lxml import etree

# ── XML namespace map ────────────────────────────────────────────────────────

NSMAP = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "v": "urn:schemas-microsoft-com:vml",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
}

XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def _ns(prefix: str, local: str) -> str:
    return f"{{{NSMAP[prefix]}}}{local}"


W_T = _ns("w", "t")
W_R = _ns("w", "r")
W_P = _ns("w", "p")
W_RPR = _ns("w", "rPr")
W_TBL = _ns("w", "tbl")
W_TR = _ns("w", "tr")
W_TC = _ns("w", "tc")
W_BODY = _ns("w", "body")
W_TXBX_CONTENT = _ns("w", "txbxContent")
WPS_TXBX = _ns("wps", "txbx")
V_TEXTBOX = _ns("v", "textbox")
WP_DOCPR = _ns("wp", "docPr")
A_T = _ns("a", "t")


# ── ID parsing helpers ───────────────────────────────────────────────────────

# Pattern: name[idx]
_IDX_RE = re.compile(r"^([a-zA-Z_]+)\[(\d+)\]$")


def _parse_id_parts(seg_id: str) -> list[tuple[str, int]]:
    """Parse 'tbl[0]:tr[1]:tc[2]:p[0]' → [('tbl',0), ('tr',1), ('tc',2), ('p',0)]."""
    parts = []
    for token in seg_id.split(":"):
        m = _IDX_RE.match(token)
        if m:
            parts.append((m.group(1), int(m.group(2))))
        else:
            # Non-indexed part like 'descr' or 'title'
            parts.append((token, -1))
    return parts


# ── Tag maps for navigation ─────────────────────────────────────────────────

_TAG_MAP = {
    "p": W_P,
    "tbl": W_TBL,
    "tr": W_TR,
    "tc": W_TC,
}


def _nth_child(parent: etree._Element, tag: str, n: int) -> etree._Element | None:
    """Find the n-th child element with the given tag."""
    xml_tag = _TAG_MAP.get(tag)
    if xml_tag is None:
        return None
    count = 0
    for child in parent:
        if child.tag == xml_tag:
            if count == n:
                return child
            count += 1
    return None


# ── Paragraph text replacement ───────────────────────────────────────────────


def _replace_paragraph_text(p_elem: etree._Element, new_text: str) -> bool:
    """Replace all text in a w:p element with new_text.

    Strategy: put all text into the first run's w:t, clear subsequent runs' w:t.
    Preserves all w:rPr formatting.
    """
    runs_with_t: list[tuple[etree._Element, etree._Element]] = []
    for r in p_elem.iterchildren(W_R):
        for t in r.iterchildren(W_T):
            runs_with_t.append((r, t))

    if not runs_with_t:
        return False

    # First run gets all the text
    _, first_t = runs_with_t[0]
    first_t.text = new_text
    first_t.set(XML_SPACE, "preserve")

    # Clear subsequent runs
    for _, t in runs_with_t[1:]:
        t.text = ""
        t.set(XML_SPACE, "preserve")

    return True


# ── Locate and apply translations ────────────────────────────────────────────


def _find_container(tree: etree._Element) -> etree._Element:
    """Find the main container (w:body or root) for navigation."""
    body = tree.find(f".//{_ns('w', 'body')}")
    return body if body is not None else tree


def _apply_wordml_translation(
    tree: etree._Element,
    seg_id: str,
    translated_text: str,
    part_name: str,
) -> bool:
    """Apply a single translation to a WordprocessingML part."""
    parts = _parse_id_parts(seg_id)
    if not parts:
        return False

    # Strip the part-name prefix if present (e.g., "header1:" or "footer2:")
    short = part_name.rsplit("/", 1)[-1].replace(".xml", "")
    first_name, _ = parts[0]
    if first_name == short:
        parts = parts[1:]
    if not parts:
        return False

    # Handle textbox segments
    if parts[0][0] == "txbx":
        return _apply_textbox_translation(tree, parts, translated_text)

    # Handle docPr alt text
    if parts[0][0] == "docPr":
        return _apply_docpr_translation(tree, parts)

    # Handle chart text (should not reach here for chart parts, but just in case)
    if parts[0][0] == "chart":
        return _apply_chart_translation(tree, parts, translated_text)

    # Navigate to the target paragraph
    container = _find_container(tree)
    current = container

    for name, idx in parts[:-1]:
        child = _nth_child(current, name, idx)
        if child is None:
            return False
        current = child

    # Last part should be 'p'
    last_name, last_idx = parts[-1]
    if last_name == "p":
        p_elem = _nth_child(current, "p", last_idx)
        if p_elem is not None:
            return _replace_paragraph_text(p_elem, translated_text)

    return False


def _apply_textbox_translation(
    tree: etree._Element,
    parts: list[tuple[str, int]],
    translated_text: str,
) -> bool:
    """Apply translation to a textbox element."""
    # parts[0] is ('txbx', idx)
    txbx_idx = parts[0][1]
    remaining = parts[1:]

    # Collect all txbxContent elements from both wps:txbx and v:textbox
    contents: list[etree._Element] = []
    for txbx in tree.iter(WPS_TXBX):
        for content in txbx.iterchildren(W_TXBX_CONTENT):
            contents.append(content)
    for vtb in tree.iter(V_TEXTBOX):
        for content in vtb.iterchildren(W_TXBX_CONTENT):
            contents.append(content)

    if txbx_idx >= len(contents):
        return False

    content = contents[txbx_idx]

    # Navigate remaining path within the textbox content
    current = content
    for name, idx in remaining[:-1]:
        child = _nth_child(current, name, idx)
        if child is None:
            return False
        current = child

    last_name, last_idx = remaining[-1]
    if last_name == "p":
        p_elem = _nth_child(current, "p", last_idx)
        if p_elem is not None:
            return _replace_paragraph_text(p_elem, translated_text)

    return False


def _apply_docpr_translation(
    tree: etree._Element,
    parts: list[tuple[str, int]],
) -> bool:
    """Apply translation to a wp:docPr alt text attribute.

    Note: translated_text is unused; we read it from the translations dict
    in the caller. This function is kept for structural consistency but the
    actual application happens in _apply_single_translation.
    """
    # Handled directly in _apply_single_translation
    return True


def _apply_chart_translation(
    tree: etree._Element,
    parts: list[tuple[str, int]],
    translated_text: str,
) -> bool:
    """Apply translation to chart a:t elements."""
    # parts: [('chart', -1), ('t', idx)]
    if len(parts) < 2:
        return False
    _, target_idx = parts[1]

    idx = 0
    for a_t in tree.iter(A_T):
        text = a_t.text or ""
        if text and not text.isspace():
            if idx == target_idx:
                a_t.text = translated_text
                return True
            idx += 1
    return False


def _apply_single_translation(
    trees: dict[str, etree._Element],
    segment: dict,
) -> bool:
    """Apply a single translation segment."""
    seg_id = segment["id"]
    part = segment["part"]
    translated = segment["translated_text"]
    seg_type = segment.get("type", "")

    if part not in trees:
        print(f"  WARNING: part '{part}' not found in DOCX, skipping {seg_id}", file=sys.stderr)
        return False

    tree = trees[part]

    # Handle docPr alt text directly
    parts = _parse_id_parts(seg_id)
    if parts and parts[0][0] == "docPr":
        docpr_idx = parts[0][1]
        attr_name = parts[1][0] if len(parts) > 1 else "descr"
        idx = 0
        for docpr in tree.iter(WP_DOCPR):
            if idx == docpr_idx:
                docpr.set(attr_name, translated)
                return True
            idx += 1
        return False

    # Handle chart text
    if seg_type == "chart":
        return _apply_chart_translation(tree, parts, translated)

    # Handle WordprocessingML parts
    return _apply_wordml_translation(tree, seg_id, translated, part)


# ── ZIP round-trip ───────────────────────────────────────────────────────────


def apply_translations(
    input_path: str,
    translations_path: str,
    output_path: str,
) -> None:
    with open(translations_path, "r", encoding="utf-8") as f:
        translations_data = json.load(f)

    # Support both {"translations": [...]} and {"segments": [...]} formats
    if "translations" in translations_data:
        translation_segments = translations_data["translations"]
    elif "segments" in translations_data:
        translation_segments = translations_data["segments"]
    else:
        print("ERROR: translations JSON must have 'translations' or 'segments' key.", file=sys.stderr)
        sys.exit(1)

    if not translation_segments:
        print("WARNING: No translations to apply.", file=sys.stderr)
        return

    # Collect all parts that need modification
    parts_needed: set[str] = set()
    for seg in translation_segments:
        parts_needed.add(seg["part"])

    # Read and parse needed XML parts from the DOCX
    trees: dict[str, etree._Element] = {}
    modified_parts: set[str] = set()

    with zipfile.ZipFile(input_path, "r") as zf:
        for part_name in parts_needed:
            if part_name in zf.namelist():
                trees[part_name] = etree.fromstring(zf.read(part_name))

    # Apply translations
    applied = 0
    failed = 0
    for seg in translation_segments:
        translated_text = seg.get("translated_text", "")
        if not translated_text and not seg.get("full_text", ""):
            continue
        if _apply_single_translation(trees, seg):
            modified_parts.add(seg["part"])
            applied += 1
        else:
            print(f"  WARNING: Could not apply translation for {seg['id']} in {seg['part']}", file=sys.stderr)
            failed += 1

    # Rebuild the DOCX: copy all files, replacing modified XML parts
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)

    with zipfile.ZipFile(input_path, "r") as zf_in:
        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf_out:
            for item in zf_in.infolist():
                if item.filename in modified_parts:
                    # Write modified XML
                    xml_bytes = etree.tostring(
                        trees[item.filename],
                        xml_declaration=True,
                        encoding="UTF-8",
                        standalone=True,
                    )
                    zf_out.writestr(item, xml_bytes)
                else:
                    # Pass through unchanged
                    zf_out.writestr(item, zf_in.read(item.filename))

    print(f"Applied {applied} translations ({failed} failed) → {output_path}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Apply translations to a DOCX file."
    )
    parser.add_argument("--input", required=True, help="Path to original DOCX file")
    parser.add_argument("--translations", required=True, help="Path to translations JSON file")
    parser.add_argument("--output", required=True, help="Path to output translated DOCX file")
    args = parser.parse_args()

    apply_translations(args.input, args.translations, args.output)


if __name__ == "__main__":
    main()
