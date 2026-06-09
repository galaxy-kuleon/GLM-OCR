#!/usr/bin/env python3
"""Regression tests for the born-digital PyMuPDF → DocIR fast path."""

from pathlib import Path
import sys

import pymupdf

# Add ir-schema/ to import path.
sys.path.insert(0, str(Path(__file__).parent.parent))

from builder.digital_ir_builder import build_docir, span_is_visible


def test_span_is_visible_filters_transparent_text():
    assert span_is_visible({"text": "visible", "alpha": 255})
    assert span_is_visible({"text": "legacy-no-alpha"})
    assert not span_is_visible({"text": "hidden", "alpha": 0})
    assert not span_is_visible({"text": "near-transparent", "alpha": 1})


def test_digital_builder_suppresses_transparent_decoy_layer(tmp_path):
    """Hidden OCR/decoy layers (alpha=0) must not become visible DOCX text."""
    fixture = tmp_path / "transparent_decoy.pdf"
    doc = pymupdf.open()
    page = doc.new_page(width=612, height=792)
    page.insert_text(
        (72, 72),
        "Visible truth text.",
        fontsize=14,
        fontname="helv",
        color=(0, 0, 0),
        fill_opacity=1,
    )
    page.insert_text(
        (72, 700),
        "DECOY OCR LAYER TOTAL 00",
        fontsize=9,
        fontname="helv",
        color=(0, 0, 0),
        fill_opacity=0,
    )
    doc.save(fixture)
    doc.close()

    output_xml = tmp_path / "transparent_decoy.docir.xml"
    build_docir(fixture, output_xml)

    xml = output_xml.read_text(encoding="utf-8")
    assert "Visible truth text." in xml
    assert "DECOY OCR LAYER TOTAL 00" not in xml
