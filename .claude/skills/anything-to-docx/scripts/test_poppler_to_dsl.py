#!/usr/bin/env python3
"""Unit tests for poppler_to_dsl.py — all pure functions.

Tests verify behavioral correctness of each pure function, not just
surface-level type checking. Each test asserts semantic invariants.
"""

import os
import sys
import xml.etree.ElementTree as ET

import pytest

sys.path.insert(0, str(__import__("pathlib").Path(__file__).parent))

from poppler_to_dsl import (
    bbox_pts_to_normalized,
    build_pdf_images_map,
    col_boundaries_to_ratios,
    detect_page_fonts,
    format_font_size,
    generate_all_dsl,
    generate_page_dsl,
    hex_color_to_rgb,
    image_to_xml,
    match_images_to_pdf_images,
    paragraph_to_xml,
    resolve_font_name,
    table_to_xml,
    _detect_font_from_families,
    FONT_FAMILY_TO_LATIN,
    FONT_FAMILY_TO_CJK,
    DEFAULT_FONT_LATIN,
    DEFAULT_FONT_CJK,
)


# ---------------------------------------------------------------------------
# hex_color_to_rgb
# ---------------------------------------------------------------------------


class TestHexColorToRgb:
    def test_valid_color(self):
        assert hex_color_to_rgb("#FF0000") == "255,0,0"

    def test_black_returns_none(self):
        """Black is the default — no need to emit."""
        assert hex_color_to_rgb("#000000") is None

    def test_white_returns_none(self):
        """White is invisible text — skip."""
        assert hex_color_to_rgb("#ffffff") is None
        assert hex_color_to_rgb("#FFFFFF") is None

    def test_empty_returns_none(self):
        assert hex_color_to_rgb("") is None
        assert hex_color_to_rgb(None) is None

    def test_invalid_returns_none(self):
        assert hex_color_to_rgb("#ZZZ") is None
        assert hex_color_to_rgb("red") is None

    def test_mixed_case(self):
        assert hex_color_to_rgb("#1a52b6") == "26,82,182"


# ---------------------------------------------------------------------------
# format_font_size
# ---------------------------------------------------------------------------


class TestFormatFontSize:
    def test_integer_value(self):
        assert format_font_size(11.0) == "11"

    def test_fractional_rounds(self):
        assert format_font_size(9.3) == "9"
        assert format_font_size(9.7) == "10"

    def test_large_size(self):
        assert format_font_size(22.0) == "22"


# ---------------------------------------------------------------------------
# bbox_pts_to_normalized
# ---------------------------------------------------------------------------


class TestBboxPtsToNormalized:
    def test_full_page(self):
        """A bbox covering the full page should produce 0,0,1000,1000."""
        bbox = {"left": 0, "top": 0, "width": 612, "height": 792}
        result = bbox_pts_to_normalized(bbox, 612.0, 792.0)
        assert result == "0,0,1000,1000"

    def test_half_page(self):
        bbox = {"left": 306, "top": 396, "width": 306, "height": 396}
        result = bbox_pts_to_normalized(bbox, 612.0, 792.0)
        assert result == "500,500,1000,1000"

    def test_zero_dimensions_guard(self):
        """Should not raise ZeroDivisionError."""
        bbox = {"left": 10, "top": 20, "width": 100, "height": 50}
        assert bbox_pts_to_normalized(bbox, 0.0, 0.0) == "0,0,0,0"
        assert bbox_pts_to_normalized(bbox, -1.0, 792.0) == "0,0,0,0"
        assert bbox_pts_to_normalized(bbox, 612.0, -1.0) == "0,0,0,0"

    def test_empty_bbox(self):
        result = bbox_pts_to_normalized({}, 612.0, 792.0)
        assert result == "0,0,0,0"


# ---------------------------------------------------------------------------
# _detect_font_from_families
# ---------------------------------------------------------------------------


class TestDetectFontFromFamilies:
    def test_known_font(self):
        result = _detect_font_from_families(["Arial"], FONT_FAMILY_TO_LATIN, "Fallback")
        assert result == "Arial"

    def test_case_insensitive(self):
        result = _detect_font_from_families(["ARIAL"], FONT_FAMILY_TO_LATIN, "Fallback")
        assert result == "Arial"

    def test_unknown_font_returns_default(self):
        result = _detect_font_from_families(["UnknownFont"], FONT_FAMILY_TO_LATIN, "Fallback")
        assert result == "Fallback"

    def test_none_family_skipped(self):
        """None values in families should not crash."""
        result = _detect_font_from_families([None, "Arial"], FONT_FAMILY_TO_LATIN, "Fallback")
        assert result == "Arial"

    def test_all_none_returns_default(self):
        result = _detect_font_from_families([None, None], FONT_FAMILY_TO_LATIN, "Fallback")
        assert result == "Fallback"

    def test_empty_list(self):
        result = _detect_font_from_families([], FONT_FAMILY_TO_LATIN, "Fallback")
        assert result == "Fallback"

    def test_cjk_mapping(self):
        result = _detect_font_from_families(["NotoSansCJKSC"], FONT_FAMILY_TO_CJK, "Fallback")
        assert result == "SimHei"


# ---------------------------------------------------------------------------
# detect_page_fonts
# ---------------------------------------------------------------------------


class TestDetectPageFonts:
    def test_empty_fontspecs(self):
        latin, cjk = detect_page_fonts([])
        assert latin == DEFAULT_FONT_LATIN
        assert cjk == DEFAULT_FONT_CJK

    def test_with_known_fonts(self):
        fontspecs = [{"family": "LiberationSans"}, {"family": "NotoSansCJKSC"}]
        latin, cjk = detect_page_fonts(fontspecs)
        assert latin == "Arial"
        assert cjk == "SimHei"


# ---------------------------------------------------------------------------
# resolve_font_name
# ---------------------------------------------------------------------------


class TestResolveFontName:
    def test_matching_page_default_returns_none(self):
        """No override needed when font matches page default."""
        assert resolve_font_name("Arial", "Arial", "SimSun") is None

    def test_different_font_returns_override(self):
        result = resolve_font_name("Courier", "Arial", "SimSun")
        assert result == "Courier New"

    def test_cjk_override_returned(self):
        """CJK font that differs from page CJK default should return override."""
        result = resolve_font_name("SimHei", "Arial", "SimSun")
        assert result == "SimHei"

    def test_cjk_matching_page_default_returns_none(self):
        """CJK font matching page CJK default should return None."""
        assert resolve_font_name("SimSun", "Arial", "SimSun") is None

    def test_empty_returns_none(self):
        assert resolve_font_name("", "Arial", "SimSun") is None
        assert resolve_font_name(None, "Arial", "SimSun") is None


# ---------------------------------------------------------------------------
# col_boundaries_to_ratios
# ---------------------------------------------------------------------------


class TestColBoundariesToRatios:
    def test_basic(self):
        # [100, 200, 400] → 2 columns, widths [100, 200], total 300
        ratios = col_boundaries_to_ratios([100, 200, 400])
        assert len(ratios) == 2
        assert abs(ratios[0] - 0.3333) < 0.01
        assert abs(ratios[1] - 0.6667) < 0.01

    def test_sum_is_exactly_one(self):
        """After normalization, sum must be *exactly* 1.0 — not just within epsilon."""
        ratios = col_boundaries_to_ratios([119, 177, 342, 477, 537, 694, 777])
        assert sum(ratios) == 1.0

    def test_count_matches_columns(self):
        boundaries = [119, 177, 342, 477, 537, 694, 777]
        ratios = col_boundaries_to_ratios(boundaries)
        assert len(ratios) == len(boundaries) - 1

    def test_empty(self):
        assert col_boundaries_to_ratios([]) == []

    def test_single_boundary(self):
        assert col_boundaries_to_ratios([100]) == []

    def test_uniform_columns(self):
        ratios = col_boundaries_to_ratios([0, 100, 200, 300])
        assert len(ratios) == 3
        for r in ratios:
            assert abs(r - 1 / 3) < 0.01

    def test_degenerate_zero_width(self):
        """All boundaries at same position — should not crash, sum must be 1.0."""
        ratios = col_boundaries_to_ratios([100, 100, 100])
        assert len(ratios) == 2
        # Equal distribution fallback
        assert abs(ratios[0] - 0.5) < 0.01
        # Normalization must apply to degenerate path too
        assert sum(ratios) == 1.0

    def test_degenerate_zero_width_many_cols(self):
        """Degenerate with n=3,6,7,9 columns: sum must be exactly 1.0 after normalization."""
        for n in (3, 6, 7, 9, 11):
            boundaries = [0] * (n + 1)
            ratios = col_boundaries_to_ratios(boundaries)
            assert len(ratios) == n, f"n={n}: expected {n} ratios, got {len(ratios)}"
            assert sum(ratios) == 1.0, f"n={n}: sum is {sum(ratios)}, not 1.0"


# ---------------------------------------------------------------------------
# table_to_xml
# ---------------------------------------------------------------------------


class TestTableToXml:
    @pytest.fixture
    def sample_table(self):
        return {
            "type": "table",
            "rows": [
                [
                    {"text": "Phase", "bold": True, "font_size_pts": 9.3, "col": 0},
                    {"text": "Condition", "bold": True, "font_size_pts": 9.3, "col": 1},
                ],
                [
                    {"text": "1", "bold": False, "font_size_pts": 9.0, "col": 0},
                    {"text": "Signing", "bold": False, "font_size_pts": 9.0, "col": 1},
                ],
            ],
            "col_count": 2,
            "row_count": 2,
            "bbox_pts": {"top": 100, "left": 50, "width": 400, "height": 100},
            "col_boundaries_px": [100, 300, 600],
        }

    def test_root_attributes(self, sample_table):
        elem = table_to_xml(sample_table, 612.0, 792.0)
        assert elem.tag == "table"
        assert elem.get("rows") == "2"
        assert elem.get("cols") == "2"
        assert elem.get("page-width-pts") == "612"
        assert elem.get("border-style") == "full"

    def test_bbox_is_normalized(self, sample_table):
        elem = table_to_xml(sample_table, 612.0, 792.0)
        bbox = elem.get("bbox")
        assert bbox is not None
        parts = [int(x) for x in bbox.split(",")]
        assert len(parts) == 4
        # All values should be in 0-1000 range
        assert all(0 <= p <= 1000 for p in parts)

    def test_col_widths_present(self, sample_table):
        elem = table_to_xml(sample_table, 612.0, 792.0)
        cw = elem.find("col-widths")
        assert cw is not None
        ratios = [float(x) for x in cw.text.split(",")]
        assert len(ratios) == 2
        assert abs(sum(ratios) - 1.0) < 0.01

    def test_header_row_detection(self, sample_table):
        elem = table_to_xml(sample_table, 612.0, 792.0)
        rows = elem.findall("row")
        assert rows[0].get("is-header") == "true"
        assert rows[1].get("is-header") is None

    def test_cell_attributes(self, sample_table):
        elem = table_to_xml(sample_table, 612.0, 792.0)
        rows = elem.findall("row")
        cells_r0 = rows[0].findall("cell")

        cell = cells_r0[0]
        assert cell.get("row") == "0"
        assert cell.get("col") == "0"
        assert cell.get("font-size-pt") == "9"
        assert cell.get("bold") == "true"
        assert cell.text == "Phase"

    def test_cell_text_is_stripped(self, sample_table):
        """Cell text should have leading/trailing whitespace stripped, consistent with paragraphs."""
        table = {
            "rows": [
                [{"text": "  padded text  ", "bold": False, "font_size_pts": 9, "col": 0}],
            ],
            "col_count": 1, "row_count": 1,
            "bbox_pts": {"top": 10, "left": 10, "width": 200, "height": 30},
            "col_boundaries_px": [10, 210],
        }
        elem = table_to_xml(table, 612.0, 792.0)
        cell = elem.findall("row")[0].findall("cell")[0]
        assert cell.text == "padded text"

    def test_non_bold_cell_has_no_bold_attribute(self, sample_table):
        elem = table_to_xml(sample_table, 612.0, 792.0)
        rows = elem.findall("row")
        cells_r1 = rows[1].findall("cell")
        assert cells_r1[0].get("bold") is None

    def test_empty_table(self):
        empty = {"rows": [], "col_count": 0, "row_count": 0,
                 "bbox_pts": {}, "col_boundaries_px": []}
        elem = table_to_xml(empty, 612.0, 792.0)
        assert elem.tag == "table"
        assert elem.get("rows") == "0"

    def test_sparse_row_with_missing_cols(self):
        """Rows with fewer cells than col_count should still work."""
        table = {
            "rows": [
                [{"text": "A", "bold": False, "font_size_pts": 9, "col": 0}],
                [{"text": "B", "bold": False, "font_size_pts": 9, "col": 2}],
            ],
            "col_count": 3,
            "row_count": 2,
            "bbox_pts": {"top": 10, "left": 10, "width": 300, "height": 50},
            "col_boundaries_px": [10, 110, 210, 310],
        }
        elem = table_to_xml(table, 612.0, 792.0)
        rows = elem.findall("row")
        assert len(rows) == 2
        # Row 0 has 1 cell, Row 1 has 1 cell
        assert len(rows[0].findall("cell")) == 1
        assert len(rows[1].findall("cell")) == 1

    def test_is_header_sparse_row_not_vacuously_true(self):
        """1 bold cell out of 5 columns must NOT be marked as header."""
        table = {
            "rows": [
                [{"text": "A", "bold": True, "font_size_pts": 9, "col": 0}],
            ],
            "col_count": 5,
            "row_count": 1,
            "bbox_pts": {"top": 10, "left": 10, "width": 500, "height": 50},
            "col_boundaries_px": [10, 110, 210, 310, 410, 510],
        }
        elem = table_to_xml(table, 612.0, 792.0)
        row0 = elem.findall("row")[0]
        assert row0.get("is-header") is None

    def test_is_header_empty_row_not_vacuously_true(self):
        """Empty first row must NOT be marked as header."""
        table = {
            "rows": [[]],
            "col_count": 3,
            "row_count": 1,
            "bbox_pts": {},
            "col_boundaries_px": [],
        }
        elem = table_to_xml(table, 612.0, 792.0)
        row0 = elem.findall("row")[0]
        assert row0.get("is-header") is None

    def test_is_header_majority_bold(self):
        """3 bold cells out of 5 columns (majority) → header."""
        table = {
            "rows": [
                [{"text": "A", "bold": True, "font_size_pts": 9, "col": 0},
                 {"text": "B", "bold": True, "font_size_pts": 9, "col": 1},
                 {"text": "C", "bold": True, "font_size_pts": 9, "col": 2},
                 {"text": "D", "bold": False, "font_size_pts": 9, "col": 3},
                 {"text": "E", "bold": False, "font_size_pts": 9, "col": 4}],
            ],
            "col_count": 5,
            "row_count": 1,
            "bbox_pts": {"top": 10, "left": 10, "width": 500, "height": 50},
            "col_boundaries_px": [10, 110, 210, 310, 410, 510],
        }
        elem = table_to_xml(table, 612.0, 792.0)
        row0 = elem.findall("row")[0]
        assert row0.get("is-header") == "true"

    def test_is_header_even_columns_50_percent_is_header(self):
        """For even col_count, exactly 50% bold (ceiling threshold) → header.

        With col_count=4: min_bold_needed = (4+1)//2 = 2.
        So 2 bold out of 4 columns is the boundary case — IS header.
        """
        table = {
            "rows": [
                [{"text": "A", "bold": True, "font_size_pts": 9, "col": 0},
                 {"text": "B", "bold": True, "font_size_pts": 9, "col": 1},
                 {"text": "C", "bold": False, "font_size_pts": 9, "col": 2},
                 {"text": "D", "bold": False, "font_size_pts": 9, "col": 3}],
            ],
            "col_count": 4,
            "row_count": 1,
            "bbox_pts": {"top": 10, "left": 10, "width": 400, "height": 50},
            "col_boundaries_px": [10, 110, 210, 310, 410],
        }
        elem = table_to_xml(table, 612.0, 792.0)
        row0 = elem.findall("row")[0]
        assert row0.get("is-header") == "true"

    def test_is_header_even_columns_below_threshold_not_header(self):
        """For even col_count=4, 1 bold out of 4 (below threshold 2) → NOT header."""
        table = {
            "rows": [
                [{"text": "A", "bold": True, "font_size_pts": 9, "col": 0},
                 {"text": "B", "bold": False, "font_size_pts": 9, "col": 1},
                 {"text": "C", "bold": False, "font_size_pts": 9, "col": 2},
                 {"text": "D", "bold": False, "font_size_pts": 9, "col": 3}],
            ],
            "col_count": 4,
            "row_count": 1,
            "bbox_pts": {"top": 10, "left": 10, "width": 400, "height": 50},
            "col_boundaries_px": [10, 110, 210, 310, 410],
        }
        elem = table_to_xml(table, 612.0, 792.0)
        row0 = elem.findall("row")[0]
        assert row0.get("is-header") is None

    def test_is_header_minority_bold_not_header(self):
        """2 bold out of 5 columns (minority) → NOT header."""
        table = {
            "rows": [
                [{"text": "A", "bold": True, "font_size_pts": 9, "col": 0},
                 {"text": "B", "bold": True, "font_size_pts": 9, "col": 1},
                 {"text": "C", "bold": False, "font_size_pts": 9, "col": 2},
                 {"text": "D", "bold": False, "font_size_pts": 9, "col": 3},
                 {"text": "E", "bold": False, "font_size_pts": 9, "col": 4}],
            ],
            "col_count": 5,
            "row_count": 1,
            "bbox_pts": {"top": 10, "left": 10, "width": 500, "height": 50},
            "col_boundaries_px": [10, 110, 210, 310, 410, 510],
        }
        elem = table_to_xml(table, 612.0, 792.0)
        row0 = elem.findall("row")[0]
        assert row0.get("is-header") is None


# ---------------------------------------------------------------------------
# paragraph_to_xml
# ---------------------------------------------------------------------------


class TestParagraphToXml:
    def test_basic_paragraph(self):
        para = {"type": "paragraph", "text": "Hello world", "font_size_pts": 11}
        elem = paragraph_to_xml(para)
        assert elem.tag == "paragraph"
        run = elem.find("run")
        assert run.text == "Hello world"

    def test_table_cell_skipped(self):
        para = {"type": "paragraph", "text": "Cell text", "is_table_cell": True}
        assert paragraph_to_xml(para) is None

    def test_footer_not_skipped_even_if_table_cell(self):
        """Footer content at table-row positions should still render."""
        para = {"type": "footer", "text": "Page 1", "is_table_cell": True, "font_size_pts": 8}
        elem = paragraph_to_xml(para)
        assert elem is not None
        assert elem.tag == "paragraph"

    def test_empty_text_skipped(self):
        para = {"type": "paragraph", "text": "  "}
        assert paragraph_to_xml(para) is None

    def test_heading(self):
        para = {"type": "heading", "text": "Title", "heading_level": 1,
                "alignment": "center", "font_size_pts": 18, "bold": True}
        elem = paragraph_to_xml(para)
        assert elem.tag == "heading"
        assert elem.get("level") == "1"

    def test_bold_italic_attributes(self):
        para = {"type": "paragraph", "text": "Bold italic",
                "font_size_pts": 11, "bold": True, "italic": True}
        elem = paragraph_to_xml(para)
        run = elem.find("run")
        assert run.get("bold") == "true"
        assert run.get("italic") == "true"

    def test_color_attribute(self):
        para = {"type": "paragraph", "text": "Red text",
                "font_size_pts": 11, "color": "#FF0000"}
        elem = paragraph_to_xml(para)
        run = elem.find("run")
        assert run.get("color-rgb") == "255,0,0"


# ---------------------------------------------------------------------------
# match_images_to_pdf_images
# ---------------------------------------------------------------------------


class TestMatchImagesToPdfImages:
    def test_empty_pdf_images(self):
        assert match_images_to_pdf_images([{"width": 100, "height": 50}], []) == []

    def test_aspect_ratio_matching(self):
        pdftohtml = [
            {"width": 200, "height": 100, "top": 50, "bbox_pts": {"left": 10}},
        ]
        pdf_images = [("pdf-images/img-001-001.png", 400, 200)]
        result = match_images_to_pdf_images(pdftohtml, pdf_images)
        assert len(result) == 1
        assert result[0]["src"] == "pdf-images/img-001-001.png"

    def test_no_match_dropped(self):
        """Unmatched pdf-images should be dropped, not included."""
        pdftohtml = [
            {"width": 100, "height": 100, "top": 50, "bbox_pts": {}},  # 1:1
        ]
        pdf_images = [("img.png", 1000, 100)]  # 10:1 — too different
        result = match_images_to_pdf_images(pdftohtml, pdf_images)
        assert len(result) == 0

    def test_multiple_images_dedup(self):
        """Each pdftohtml image should only be used once (no double-matching)."""
        pdftohtml = [
            {"width": 200, "height": 100, "top": 50, "bbox_pts": {"left": 10}},
            {"width": 100, "height": 100, "top": 200, "bbox_pts": {"left": 20}},
        ]
        # Two pdf-images: first matches pdftohtml[0] (2:1), second matches pdftohtml[1] (1:1)
        pdf_images = [
            ("img-a.png", 400, 200),   # 2:1
            ("img-b.png", 300, 300),   # 1:1
        ]
        result = match_images_to_pdf_images(pdftohtml, pdf_images)
        assert len(result) == 2
        srcs = {r["src"] for r in result}
        assert srcs == {"img-a.png", "img-b.png"}


# ---------------------------------------------------------------------------
# image_to_xml
# ---------------------------------------------------------------------------


class TestImageToXml:
    def test_basic_image(self):
        """Image element should have src, bbox, page-width-pts attributes."""
        img_data = {
            "src": "pdf-images/img-001-001.png",
            "bbox_pts": {"left": 50, "top": 100, "width": 400, "height": 300},
            "top": 150,
        }
        elem = image_to_xml(img_data, 612.0, 792.0)
        assert elem.tag == "image"
        assert elem.get("src") == "pdf-images/img-001-001.png"
        assert elem.get("page-width-pts") == "612"

    def test_bbox_normalized(self):
        """bbox must be normalized to 0-1000 range."""
        img_data = {
            "src": "pdf-images/img-001-001.png",
            "bbox_pts": {"left": 306, "top": 396, "width": 306, "height": 396},
            "top": 0,
        }
        elem = image_to_xml(img_data, 612.0, 792.0)
        bbox = elem.get("bbox")
        assert bbox == "500,500,1000,1000"

    def test_empty_bbox_pts_no_bbox_attr(self):
        """If bbox_pts is empty dict, no bbox attribute should be set."""
        img_data = {
            "src": "pdf-images/img-001-001.png",
            "bbox_pts": {},
            "top": 0,
        }
        elem = image_to_xml(img_data, 612.0, 792.0)
        assert elem.get("bbox") is None

    def test_missing_bbox_pts_no_crash(self):
        """If bbox_pts key is missing entirely, should not crash."""
        img_data = {"src": "pdf-images/img-001-001.png", "top": 0}
        elem = image_to_xml(img_data, 612.0, 792.0)
        assert elem.tag == "image"
        assert elem.get("src") == "pdf-images/img-001-001.png"
        assert elem.get("bbox") is None

    def test_missing_src_no_crash(self):
        """Missing src key should not raise KeyError — uses empty string fallback."""
        img_data = {
            "bbox_pts": {"left": 0, "top": 0, "width": 100, "height": 50},
            "top": 0,
        }
        elem = image_to_xml(img_data, 612.0, 792.0)
        assert elem.tag == "image"
        assert elem.get("src") == ""

    def test_full_page_image_bbox(self):
        """Full-page image bbox should produce 0,0,1000,1000."""
        img_data = {
            "src": "img.png",
            "bbox_pts": {"left": 0, "top": 0, "width": 612, "height": 792},
            "top": 0,
        }
        elem = image_to_xml(img_data, 612.0, 792.0)
        assert elem.get("bbox") == "0,0,1000,1000"


# ---------------------------------------------------------------------------
# generate_page_dsl — integration test
# ---------------------------------------------------------------------------


class TestGeneratePageDsl:
    def _make_page(self, paragraphs=None, tables=None, images=None):
        return {
            "number": 1,
            "width_pts": 612.0,
            "height_pts": 792.0,
            "fontspecs": [],
            "paragraphs": paragraphs or [],
            "tables": tables or [],
            "images": images or [],
        }

    def test_empty_page(self):
        xml_str = generate_page_dsl(self._make_page())
        root = ET.fromstring(xml_str)
        assert root.tag == "page"
        assert list(root) == []

    def test_paragraph_ordering_by_top(self):
        paras = [
            {"type": "paragraph", "text": "Second", "top": 200, "font_size_pts": 11},
            {"type": "paragraph", "text": "First", "top": 100, "font_size_pts": 11},
        ]
        xml_str = generate_page_dsl(self._make_page(paragraphs=paras))
        root = ET.fromstring(xml_str)
        children = list(root)
        assert children[0].find("run").text == "First"
        assert children[1].find("run").text == "Second"

    def test_table_interleaved_with_paragraphs(self):
        paras = [
            {"type": "paragraph", "text": "Before table", "top": 50, "font_size_pts": 11},
            {"type": "paragraph", "text": "After table", "top": 500, "font_size_pts": 11},
        ]
        table = {
            "rows": [[{"text": "Cell", "bold": False, "font_size_pts": 9, "col": 0}]],
            "col_count": 1, "row_count": 1,
            "bbox_pts": {"top": 100, "left": 50, "width": 400, "height": 50},
            "bbox_px": {"top": 200, "left": 75, "width": 600, "height": 75},
            "col_boundaries_px": [75, 675],
        }
        xml_str = generate_page_dsl(self._make_page(paragraphs=paras, tables=[table]))
        root = ET.fromstring(xml_str)
        children = list(root)
        assert len(children) == 3
        assert children[0].tag == "paragraph"  # Before table (top=50)
        assert children[1].tag == "table"       # Table (top=200 in px)
        assert children[2].tag == "paragraph"  # After table (top=500)

    def test_table_cell_paragraphs_filtered(self):
        paras = [
            {"type": "paragraph", "text": "Normal", "top": 50, "font_size_pts": 11},
            {"type": "paragraph", "text": "Cell content", "top": 200,
             "font_size_pts": 9, "is_table_cell": True},
        ]
        table = {
            "rows": [[{"text": "Cell content", "bold": False, "font_size_pts": 9, "col": 0}]],
            "col_count": 1, "row_count": 1,
            "bbox_pts": {"top": 100, "left": 50, "width": 400, "height": 50},
            "bbox_px": {"top": 200, "left": 75, "width": 600, "height": 75},
            "col_boundaries_px": [75, 675],
        }
        xml_str = generate_page_dsl(self._make_page(paragraphs=paras, tables=[table]))
        root = ET.fromstring(xml_str)
        children = list(root)
        # Should have 2 elements: paragraph + table (cell paragraph filtered)
        assert len(children) == 2
        tags = [c.tag for c in children]
        assert tags == ["paragraph", "table"]

    def test_page_attributes(self):
        xml_str = generate_page_dsl(self._make_page())
        root = ET.fromstring(xml_str)
        assert root.get("number") == "1"
        assert root.get("width-pts") == "612"
        assert root.get("height-pts") == "792"
        assert root.get("font-latin") == "Arial"
        assert root.get("font-cjk") == "SimSun"

    def test_images_interleaved_with_paragraphs(self):
        """Images from page_images should be matched and interleaved by top position."""
        paras = [
            {"type": "paragraph", "text": "Before image", "top": 50, "font_size_pts": 11},
            {"type": "paragraph", "text": "After image", "top": 500, "font_size_pts": 11},
        ]
        # pdftohtml images provide bbox_pts for the matching step
        pdftohtml_images = [
            {"width": 200, "height": 100, "top": 200,
             "bbox_pts": {"left": 50, "top": 150, "width": 400, "height": 200}},
        ]
        page_data = self._make_page(paragraphs=paras, images=pdftohtml_images)
        # pdf-images (from pdfimages) with matching aspect ratio 2:1
        page_images = [("pdf-images/img-001-001.png", 400, 200)]

        xml_str = generate_page_dsl(page_data, page_images=page_images)
        root = ET.fromstring(xml_str)
        children = list(root)
        assert len(children) == 3
        assert children[0].tag == "paragraph"  # Before image (top=50)
        assert children[1].tag == "image"      # Image (top=200)
        assert children[2].tag == "paragraph"  # After image (top=500)
        # Verify image src
        assert children[1].get("src") == "pdf-images/img-001-001.png"
        # Verify image has bbox
        assert children[1].get("bbox") is not None

    def test_no_page_images_no_image_elements(self):
        """Without page_images, no image elements should appear even if pdftohtml has images."""
        pdftohtml_images = [
            {"width": 200, "height": 100, "top": 200,
             "bbox_pts": {"left": 50, "top": 150, "width": 400, "height": 200}},
        ]
        page_data = self._make_page(images=pdftohtml_images)
        xml_str = generate_page_dsl(page_data, page_images=None)
        root = ET.fromstring(xml_str)
        assert len(list(root)) == 0

    def test_font_name_override_in_paragraph(self):
        """Paragraphs with non-default font should have font-name on the run."""
        paras = [
            {"type": "paragraph", "text": "Courier text", "top": 100,
             "font_size_pts": 11, "font_family": "Courier"},
        ]
        # Page defaults are Arial/SimSun. "Courier" maps to "Courier New" which
        # differs from both defaults, so it should appear as a font-name override.
        xml_str = generate_page_dsl(self._make_page(paragraphs=paras))
        root = ET.fromstring(xml_str)
        run = root.find(".//run")
        assert run.get("font-name") == "Courier New"

    def test_table_with_border_style_in_dsl(self):
        """Tables in the DSL output should have border-style='full'."""
        table = {
            "rows": [[{"text": "A", "bold": False, "font_size_pts": 9, "col": 0}]],
            "col_count": 1, "row_count": 1,
            "bbox_pts": {"top": 100, "left": 50, "width": 400, "height": 50},
            "bbox_px": {"top": 200, "left": 75, "width": 600, "height": 75},
            "col_boundaries_px": [75, 675],
        }
        xml_str = generate_page_dsl(self._make_page(tables=[table]))
        root = ET.fromstring(xml_str)
        table_elem = root.find("table")
        assert table_elem.get("border-style") == "full"

    def test_table_bbox_px_fallback_to_bbox_pts(self):
        """Table without bbox_px should fall back to bbox_pts for ordering."""
        paras = [
            {"type": "paragraph", "text": "After table", "top": 500, "font_size_pts": 11},
        ]
        table = {
            "rows": [[{"text": "A", "bold": False, "font_size_pts": 9, "col": 0}]],
            "col_count": 1, "row_count": 1,
            "bbox_pts": {"top": 100, "left": 50, "width": 400, "height": 50},
            # No bbox_px — testing fallback
            "col_boundaries_px": [50, 450],
        }
        xml_str = generate_page_dsl(self._make_page(paragraphs=paras, tables=[table]))
        root = ET.fromstring(xml_str)
        children = list(root)
        assert len(children) == 2
        # Table (bbox_pts top=100) should be before paragraph (top=500)
        assert children[0].tag == "table"
        assert children[1].tag == "paragraph"

    def test_table_bbox_px_missing_top_falls_back_to_bbox_pts(self):
        """bbox_px present but missing 'top' key should fall back to bbox_pts."""
        paras = [
            {"type": "paragraph", "text": "Before table", "top": 50, "font_size_pts": 11},
        ]
        table = {
            "rows": [[{"text": "A", "bold": False, "font_size_pts": 9, "col": 0}]],
            "col_count": 1, "row_count": 1,
            "bbox_pts": {"top": 300, "left": 50, "width": 400, "height": 50},
            # bbox_px present but WITHOUT "top" — must fall back to bbox_pts.top=300
            "bbox_px": {"left": 75, "width": 600, "height": 75},
            "col_boundaries_px": [75, 675],
        }
        xml_str = generate_page_dsl(self._make_page(paragraphs=paras, tables=[table]))
        root = ET.fromstring(xml_str)
        children = list(root)
        assert len(children) == 2
        # Paragraph (top=50) before table (fallback top=300)
        assert children[0].tag == "paragraph"
        assert children[1].tag == "table"


# ---------------------------------------------------------------------------
# build_pdf_images_map — filesystem integration test
# ---------------------------------------------------------------------------


class TestBuildPdfImagesMap:
    def test_empty_workspace(self, tmp_path):
        """No pdf-images dir → empty map."""
        result = build_pdf_images_map(str(tmp_path), 3)
        assert result == {}

    def test_empty_pdf_images_dir(self, tmp_path):
        """pdf-images dir exists but is empty → empty map."""
        (tmp_path / "pdf-images").mkdir()
        result = build_pdf_images_map(str(tmp_path), 3)
        assert result == {}

    def test_with_real_images(self, tmp_path):
        """Real PNG images should be collected and indexed by page."""
        from PIL import Image as _PILImage

        pdf_dir = tmp_path / "pdf-images"
        pdf_dir.mkdir()

        # Create a valid 100x50 RGB image for page 1
        img = _PILImage.new("RGB", (100, 50), color="red")
        img.save(str(pdf_dir / "img-001-001.png"))

        # Create a valid 200x100 RGB image for page 2
        img2 = _PILImage.new("RGB", (200, 100), color="blue")
        img2.save(str(pdf_dir / "img-002-001.png"))

        result = build_pdf_images_map(str(tmp_path), 2)
        assert 1 in result
        assert 2 in result
        assert len(result[1]) == 1
        assert len(result[2]) == 1
        # Check structure: (relative_path, width, height)
        path, w, h = result[1][0]
        assert path == "pdf-images/img-001-001.png"
        assert w == 100
        assert h == 50

    def test_smask_filtered(self, tmp_path):
        """Grayscale (mode 'L') images should be filtered as smasks."""
        from PIL import Image as _PILImage

        pdf_dir = tmp_path / "pdf-images"
        pdf_dir.mkdir()

        # Create a grayscale smask
        smask = _PILImage.new("L", (100, 50))
        smask.save(str(pdf_dir / "img-001-001.png"))

        result = build_pdf_images_map(str(tmp_path), 1)
        assert result == {}

    def test_repeating_images_filtered(self, tmp_path):
        """Images appearing on 3+ pages should be filtered as repeating."""
        from PIL import Image as _PILImage

        pdf_dir = tmp_path / "pdf-images"
        pdf_dir.mkdir()

        # Create identical images (same size, dimensions → same fingerprint)
        for page_num in range(1, 5):
            img = _PILImage.new("RGB", (50, 20), color="green")
            img.save(str(pdf_dir / f"img-{page_num:03d}-001.png"))

        result = build_pdf_images_map(str(tmp_path), 4)
        # All pages have the same fingerprint → repeating → filtered out
        assert result == {}


# ---------------------------------------------------------------------------
# generate_all_dsl — filesystem integration test
# ---------------------------------------------------------------------------


class TestGenerateAllDsl:
    def test_writes_files(self, tmp_path):
        """generate_all_dsl should write one XML file per page."""
        parsed_data = {
            "pages": [
                {
                    "number": 1,
                    "width_pts": 612.0,
                    "height_pts": 792.0,
                    "fontspecs": [],
                    "paragraphs": [
                        {"type": "paragraph", "text": "Page 1 text",
                         "top": 100, "font_size_pts": 11},
                    ],
                    "tables": [],
                    "images": [],
                },
                {
                    "number": 2,
                    "width_pts": 612.0,
                    "height_pts": 792.0,
                    "fontspecs": [],
                    "paragraphs": [
                        {"type": "heading", "text": "Page 2 Heading",
                         "top": 50, "font_size_pts": 18, "heading_level": 1,
                         "alignment": "center", "bold": True},
                    ],
                    "tables": [],
                    "images": [],
                },
            ],
        }
        written = generate_all_dsl(parsed_data, str(tmp_path))
        assert len(written) == 2

        # Verify files exist and contain valid XML
        for path in written:
            assert os.path.exists(path)
            content = open(path, encoding="utf-8").read()
            root = ET.fromstring(content)
            assert root.tag == "page"

        # Verify page-1.xml content
        root1 = ET.fromstring(open(written[0], encoding="utf-8").read())
        assert root1.get("number") == "1"
        children1 = list(root1)
        assert len(children1) == 1
        assert children1[0].tag == "paragraph"

        # Verify page-2.xml content
        root2 = ET.fromstring(open(written[1], encoding="utf-8").read())
        assert root2.get("number") == "2"
        children2 = list(root2)
        assert len(children2) == 1
        assert children2[0].tag == "heading"

    def test_creates_dsl_directory(self, tmp_path):
        """Should create dsl/ subdirectory if not exists."""
        parsed_data = {
            "pages": [{
                "number": 1,
                "width_pts": 612.0,
                "height_pts": 792.0,
                "fontspecs": [],
                "paragraphs": [],
                "tables": [],
                "images": [],
            }],
        }
        generate_all_dsl(parsed_data, str(tmp_path))
        assert (tmp_path / "dsl").is_dir()
        assert (tmp_path / "dsl" / "page-1.xml").exists()


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
