#!/usr/bin/env python3
"""Extract translatable text segments from a DOCX file.

Opens the DOCX as a ZIP archive, parses all relevant XML parts, and extracts
translatable text at the paragraph level. Outputs a JSON file with segment
metadata for downstream translation.

Usage:
    uv run --with lxml \
        .claude/skills/another-pure-pure-docx-translate-to-docx/scripts/extract_docx_texts.py \
        --input INPUT.docx \
        --output workspace/texts.json
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import zipfile
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
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
}


def _ns(prefix: str, local: str) -> str:
    """Return a Clark-notation tag name."""
    return f"{{{NSMAP[prefix]}}}{local}"


W_T = _ns("w", "t")
W_R = _ns("w", "r")
W_P = _ns("w", "p")
W_TBL = _ns("w", "tbl")
W_TR = _ns("w", "tr")
W_TC = _ns("w", "tc")
W_BODY = _ns("w", "body")
W_TXBX_CONTENT = _ns("w", "txbxContent")
WPS_TXBX = _ns("wps", "txbx")
V_TEXTBOX = _ns("v", "textbox")
WP_DOCPR = _ns("wp", "docPr")
A_T = _ns("a", "t")
A_R = _ns("a", "r")
A_P = _ns("a", "p")
C_TX = _ns("c", "tx")
C_RICH = _ns("c", "rich")


# ── Helpers ──────────────────────────────────────────────────────────────────


def _paragraph_runs(p_elem: etree._Element) -> list[dict]:
    """Extract runs (w:r/w:t) from a w:p element.

    Returns a list of {"index": int, "text": str} for runs that contain text.
    """
    runs: list[dict] = []
    idx = 0
    for r in p_elem.iterchildren(W_R):
        for t in r.iterchildren(W_T):
            text = t.text or ""
            if text:
                runs.append({"index": idx, "text": text})
        idx += 1
    return runs


def _paragraph_full_text(runs: list[dict]) -> str:
    return "".join(r["text"] for r in runs)


def _is_blank(text: str) -> bool:
    return not text or text.isspace()


# ── Extractors ───────────────────────────────────────────────────────────────

def _extract_body_paragraphs(
    body: etree._Element,
    part_name: str,
    segments: list[dict],
    seg_type: str = "paragraph",
    id_prefix: str = "",
) -> None:
    """Extract paragraphs directly under a container (body, txbxContent, etc.)."""
    p_idx = 0
    for child in body:
        if child.tag == W_P:
            runs = _paragraph_runs(child)
            full = _paragraph_full_text(runs)
            if not _is_blank(full):
                seg_id = f"{id_prefix}p[{p_idx}]"
                segments.append({
                    "id": seg_id,
                    "part": part_name,
                    "type": seg_type,
                    "full_text": full,
                    "runs": runs,
                })
            p_idx += 1


def _extract_tables(
    container: etree._Element,
    part_name: str,
    segments: list[dict],
    id_prefix: str = "",
) -> None:
    """Extract text from w:tbl elements."""
    for tbl_idx, tbl in enumerate(container.iterchildren(W_TBL)):
        for tr_idx, tr in enumerate(tbl.iterchildren(W_TR)):
            for tc_idx, tc in enumerate(tr.iterchildren(W_TC)):
                p_idx = 0
                for p in tc.iterchildren(W_P):
                    runs = _paragraph_runs(p)
                    full = _paragraph_full_text(runs)
                    if not _is_blank(full):
                        seg_id = f"{id_prefix}tbl[{tbl_idx}]:tr[{tr_idx}]:tc[{tc_idx}]:p[{p_idx}]"
                        segments.append({
                            "id": seg_id,
                            "part": part_name,
                            "type": "table_cell",
                            "full_text": full,
                            "runs": runs,
                        })
                    p_idx += 1


def _extract_textboxes(
    root: etree._Element,
    part_name: str,
    segments: list[dict],
) -> None:
    """Extract text from textbox elements (wps:txbx and v:textbox)."""
    txbx_idx = 0

    # wps:txbx/w:txbxContent
    for txbx in root.iter(WPS_TXBX):
        for content in txbx.iterchildren(W_TXBX_CONTENT):
            _extract_body_paragraphs(
                content, part_name, segments,
                seg_type="textbox",
                id_prefix=f"txbx[{txbx_idx}]:",
            )
            _extract_tables(content, part_name, segments, id_prefix=f"txbx[{txbx_idx}]:")
            txbx_idx += 1

    # v:textbox/w:txbxContent
    for vtb in root.iter(V_TEXTBOX):
        for content in vtb.iterchildren(W_TXBX_CONTENT):
            _extract_body_paragraphs(
                content, part_name, segments,
                seg_type="textbox",
                id_prefix=f"txbx[{txbx_idx}]:",
            )
            _extract_tables(content, part_name, segments, id_prefix=f"txbx[{txbx_idx}]:")
            txbx_idx += 1


def _extract_docpr_alt(
    root: etree._Element,
    part_name: str,
    segments: list[dict],
) -> None:
    """Extract image alt text from wp:docPr elements."""
    idx = 0
    for docpr in root.iter(WP_DOCPR):
        descr = docpr.get("descr", "")
        title = docpr.get("title", "")
        if descr and not _is_blank(descr):
            segments.append({
                "id": f"docPr[{idx}]:descr",
                "part": part_name,
                "type": "alt_text",
                "full_text": descr,
                "runs": [{"index": 0, "text": descr}],
            })
        if title and not _is_blank(title):
            segments.append({
                "id": f"docPr[{idx}]:title",
                "part": part_name,
                "type": "alt_text",
                "full_text": title,
                "runs": [{"index": 0, "text": title}],
            })
        idx += 1


def _extract_chart_texts(
    root: etree._Element,
    part_name: str,
    segments: list[dict],
) -> None:
    """Extract text from chart XML (c:tx/c:rich/a:p/a:r/a:t and similar)."""
    idx = 0
    for a_t in root.iter(A_T):
        text = a_t.text or ""
        if not _is_blank(text):
            segments.append({
                "id": f"chart:t[{idx}]",
                "part": part_name,
                "type": "chart",
                "full_text": text,
                "runs": [{"index": 0, "text": text}],
            })
            idx += 1


def _extract_wordml_part(
    tree: etree._Element,
    part_name: str,
    segments: list[dict],
    seg_type: str = "paragraph",
) -> None:
    """Generic extractor for WordprocessingML parts (document, headers, footers, footnotes, endnotes)."""
    short = part_name.rsplit("/", 1)[-1].replace(".xml", "")
    id_prefix = f"{short}:" if seg_type != "paragraph" else ""

    body = tree.find(f".//{_ns('w', 'body')}")
    container = body if body is not None else tree

    # Direct paragraphs (skip those inside tables or textboxes — they get handled separately)
    p_idx = 0
    for child in container:
        if child.tag == W_P:
            runs = _paragraph_runs(child)
            full = _paragraph_full_text(runs)
            if not _is_blank(full):
                seg_id = f"{id_prefix}p[{p_idx}]"
                segments.append({
                    "id": seg_id,
                    "part": part_name,
                    "type": seg_type,
                    "full_text": full,
                    "runs": runs,
                })
            p_idx += 1

    # Tables
    _extract_tables(container, part_name, segments, id_prefix=id_prefix)

    # Textboxes (anywhere in the part)
    _extract_textboxes(tree, part_name, segments)

    # Image alt text
    _extract_docpr_alt(tree, part_name, segments)


# ── Main ─────────────────────────────────────────────────────────────────────


def extract(input_path: str, output_path: str) -> None:
    if not zipfile.is_zipfile(input_path):
        print(f"ERROR: {input_path} is not a valid ZIP/DOCX file.", file=sys.stderr)
        sys.exit(1)

    segments: list[dict] = []

    with zipfile.ZipFile(input_path, "r") as zf:
        names = zf.namelist()

        # ── document.xml ──
        if "word/document.xml" in names:
            tree = etree.fromstring(zf.read("word/document.xml"))
            _extract_wordml_part(tree, "word/document.xml", segments, seg_type="paragraph")

        # ── headers ──
        for name in sorted(names):
            if re.match(r"word/header\d*\.xml$", name):
                tree = etree.fromstring(zf.read(name))
                short = name.rsplit("/", 1)[-1].replace(".xml", "")
                _extract_wordml_part(tree, name, segments, seg_type="header")

        # ── footers ──
        for name in sorted(names):
            if re.match(r"word/footer\d*\.xml$", name):
                tree = etree.fromstring(zf.read(name))
                _extract_wordml_part(tree, name, segments, seg_type="footer")

        # ── footnotes ──
        if "word/footnotes.xml" in names:
            tree = etree.fromstring(zf.read("word/footnotes.xml"))
            _extract_wordml_part(tree, "word/footnotes.xml", segments, seg_type="footnote")

        # ── endnotes ──
        if "word/endnotes.xml" in names:
            tree = etree.fromstring(zf.read("word/endnotes.xml"))
            _extract_wordml_part(tree, "word/endnotes.xml", segments, seg_type="endnote")

        # ── charts ──
        for name in sorted(names):
            if re.match(r"word/charts/chart\d*\.xml$", name):
                tree = etree.fromstring(zf.read(name))
                _extract_chart_texts(tree, name, segments)

    result = {
        "source_file": os.path.basename(input_path),
        "total_segments": len(segments),
        "segments": segments,
    }

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(f"Extracted {len(segments)} translatable segments → {output_path}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Extract translatable text from a DOCX file."
    )
    parser.add_argument("--input", required=True, help="Path to input DOCX file")
    parser.add_argument("--output", required=True, help="Path to output JSON file")
    args = parser.parse_args()

    extract(args.input, args.output)


if __name__ == "__main__":
    main()
