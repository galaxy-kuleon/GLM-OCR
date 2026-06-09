#!/usr/bin/env python3
"""Tests for PNG-first raster fallback DocIR generation."""

from pathlib import Path
import sys
import zipfile

import pymupdf
import lxml.etree as etree
import pytest

# Add ir-schema/ to import path.
sys.path.insert(0, str(Path(__file__).parent.parent))

from builder.raster_ir_builder import MAX_PAGES, build_docir
from docx_generator.docx_generator import generate_docx
from validator.semantic_validator import run_all_validations

DOCIR_NS = "urn:docir:v0.1"
NS = {"d": DOCIR_NS}


def make_image_only_pdf(path: Path) -> None:
    doc = pymupdf.open()
    page = doc.new_page(width=200, height=100)
    shape = page.new_shape()
    shape.draw_rect(pymupdf.Rect(10, 10, 190, 90))
    shape.finish(color=(0, 0, 0), fill=(0.9, 0.9, 0.9), width=1)
    shape.commit()
    page.insert_text((30, 55), "Raster Truth", fontsize=14)
    doc.save(path)
    doc.close()


def test_raster_builder_pdf_creates_full_page_image_region(tmp_path):
    pdf = tmp_path / "image_only.pdf"
    make_image_only_pdf(pdf)
    output_xml = tmp_path / "image_only.docir.xml"

    build_docir(pdf, output_xml, dpi=144)

    root = etree.parse(str(output_xml)).getroot()
    assert root.get("generator") == "raster-ir-builder-v0.1.0"
    pages = root.findall(".//d:page", NS)
    images = root.findall(".//d:region[@type='image']", NS)
    assets = root.findall(".//d:asset", NS)
    assert len(pages) == 1
    assert len(images) == 1
    assert len(assets) == 1

    page_size = pages[0].find("d:page_size", NS)
    bbox = images[0].find("d:bbox", NS)
    assert page_size is not None and bbox is not None
    assert page_size.get("width_pt") == "200.00"
    assert page_size.get("height_pt") == "100.00"
    assert bbox.get("x") == "0.00"
    assert bbox.get("y") == "0.00"
    assert bbox.get("width") == "200.00"
    assert bbox.get("height") == "100.00"

    visual_features = images[0].find(".//d:visual_features", NS)
    assert visual_features is not None
    assert visual_features.find("d:raster_fallback", NS) is None
    assert "raster_fallback=true" in (visual_features.findtext("d:description", namespaces=NS) or "")

    file_path = assets[0].findtext("d:file_path", namespaces=NS)
    assert file_path
    assert (tmp_path / file_path).exists()

    schema_path = Path(__file__).parent.parent / "docir-v0.1.0.xsd"
    result = run_all_validations(output_xml, schema_path)
    assert result.passed, result.summary()
    assert not [issue for issue in result.issues if issue.rule == "bbox-overflow"]


def test_raster_docir_generates_positioned_anchored_docx(tmp_path):
    pdf = tmp_path / "image_only.pdf"
    make_image_only_pdf(pdf)
    output_xml = tmp_path / "image_only.docir.xml"
    output_docx = tmp_path / "image_only.docx"

    build_docir(pdf, output_xml, dpi=144)
    generate_docx(output_xml, output_docx, positioned=True, base_dir=tmp_path)

    assert output_docx.exists()
    with zipfile.ZipFile(output_docx) as docx_zip:
        document_xml = docx_zip.read("word/document.xml").decode("utf-8")
    assert "<wp:anchor" in document_xml
    assert "<wp:inline" not in document_xml
    assert "Raster Truth" not in document_xml  # raster baseline is image-backed, not fake editable OCR


def test_raster_builder_caps_image_directory_page_count(tmp_path):
    image_dir = tmp_path / "pages"
    image_dir.mkdir()
    for idx in range(MAX_PAGES + 1):
        (image_dir / f"page_{idx:04d}.png").write_bytes(b"")

    with pytest.raises(ValueError, match="raster fallback limit"):
        build_docir(image_dir, tmp_path / "too_many.docir.xml", dpi=144)
