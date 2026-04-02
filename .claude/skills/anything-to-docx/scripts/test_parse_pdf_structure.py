#!/usr/bin/env python3
"""Tests for parse_pdf_structure.py.

Covers pure functions with synthetic XML inputs + integration tests against real PDFs.

Run: uv run python3 test_parse_pdf_structure.py
  or: uv run python3 -m pytest test_parse_pdf_structure.py -v
"""

import json
import subprocess
import sys
from pathlib import Path

import pytest

# ---------------------------------------------------------------------------
# Import the module under test
# ---------------------------------------------------------------------------
sys.path.insert(0, str(Path(__file__).resolve().parent))

from parse_pdf_structure import (
    _clean_font_family,
    _parse_fontspecs,
    _parse_images,
    _parse_text_elements,
    _parse_xml_string,
    _px_to_pts,
    _find_table_row_tops,
    _is_list_item,
    _detect_alignment,
    _build_paragraph,
    group_text_into_paragraphs,
    group_all_pages,
    _build_fontspec_lookup,
    _determine_body_font_size,
    _classify_single_paragraph,
    classify_paragraphs,
    classify_all_pages,
    _cluster_columns,
    _assign_col_index,
    _group_elements_into_rows,
    _should_merge_rows,
    _merge_multiline_cells,
    detect_tables,
    detect_tables_all_pages,
    enrich_images,
    enrich_images_all_pages,
    _suppress_headings_near_tables,
    suppress_headings_near_tables_all_pages,
    format_summary,
    parse_digital_pdf,
    DIGITAL_MIN_CHARS_PER_PAGE,
    PDFTOHTML_SCALE,
    PARA_VERTICAL_GAP_FACTOR,
    PARA_LEFT_TOLERANCE_PX,
    TABLE_ROW_MIN_ELEMENTS,
    TABLE_ROW_MAX_ELEMENTS,
    TABLE_ROW_MEDIAN_WIDTH_MIN_PX,
    TABLE_ROW_COVERAGE_MAX,
    _median_width,
    _coverage_ratio,
    ALIGNMENT_TOLERANCE_PX,
    HEADING_LEVEL_1_MIN_PTS,
    HEADING_LEVEL_2_MIN_PTS,
    HEADING_LEVEL_3_MIN_PTS,
    HEADING_LEVEL_4_MIN_PTS,
    HEADING_MIN_CONTENT_CHARS,
    TABLE_HEADING_PROXIMITY_PX,
    FOOTER_TOP_RATIO,
    HEADER_BOTTOM_RATIO,
    HEADER_FOOTER_MAX_FONT_PTS,
    TABLE_COL_CLUSTER_TOLERANCE_PX,
)


# ---------------------------------------------------------------------------
# _clean_font_family
# ---------------------------------------------------------------------------

class TestCleanFontFamily:
    def test_strip_subset_prefix(self):
        assert _clean_font_family("BAAAAA+LiberationSans") == "LiberationSans"

    def test_strip_bold_suffix(self):
        assert _clean_font_family("BAAAAA+LiberationSans-Bold") == "LiberationSans"

    def test_strip_italic_suffix(self):
        assert _clean_font_family("LiberationSans-Italic") == "LiberationSans"

    def test_strip_bolditalic_suffix(self):
        assert _clean_font_family("CDEFGH+TimesNewRoman-BoldItalic") == "TimesNewRoman"

    def test_strip_oblique_suffix(self):
        assert _clean_font_family("Helvetica-Oblique") == "Helvetica"

    def test_strip_boldoblique_suffix(self):
        assert _clean_font_family("Helvetica-BoldOblique") == "Helvetica"

    def test_strip_regular_suffix(self):
        assert _clean_font_family("Arial-Regular") == "Arial"

    def test_no_prefix_no_suffix(self):
        assert _clean_font_family("Courier") == "Courier"

    def test_preserves_hyphens_in_family_name(self):
        # Ensure a hyphen in the middle (not a known suffix) is preserved
        assert _clean_font_family("Noto-Sans") == "Noto-Sans"


# ---------------------------------------------------------------------------
# _px_to_pts
# ---------------------------------------------------------------------------

class TestPxToPts:
    def test_basic_conversion(self):
        assert _px_to_pts(15) == round(15 / PDFTOHTML_SCALE, 1)

    def test_zero(self):
        assert _px_to_pts(0) == 0.0

    def test_float_input(self):
        assert _px_to_pts(22.5) == round(22.5 / PDFTOHTML_SCALE, 1)


# ---------------------------------------------------------------------------
# _parse_fontspecs
# ---------------------------------------------------------------------------

class TestParseFontspecs:
    def test_basic_fontspec(self):
        xml = '<fontspec id="0" size="24" family="Arial" color="#000000"/>'
        result = _parse_fontspecs(xml)
        assert len(result) == 1
        fs = result[0]
        assert fs["id"] == "0"
        assert fs["size_px"] == 24
        assert fs["size_pts"] == _px_to_pts(24)
        assert fs["family"] == "Arial"
        assert fs["raw_family"] == "Arial"
        assert fs["color"] == "#000000"

    def test_no_is_bold_field(self):
        """is_bold was removed as dead code — fontspecs should NOT contain it."""
        xml = '<fontspec id="1" size="18" family="BAAAAA+Times-Bold" color="#333333"/>'
        result = _parse_fontspecs(xml)
        assert len(result) == 1
        assert "is_bold" not in result[0]

    def test_subset_prefix_stripped_from_family(self):
        xml = '<fontspec id="2" size="12" family="ABCDEF+Helvetica-Italic" color="#000000"/>'
        result = _parse_fontspecs(xml)
        assert result[0]["family"] == "Helvetica"
        assert result[0]["raw_family"] == "ABCDEF+Helvetica-Italic"

    def test_multiple_fontspecs(self):
        xml = (
            '<fontspec id="0" size="24" family="Arial" color="#000000"/>'
            '<fontspec id="1" size="12" family="Times" color="#333333"/>'
        )
        result = _parse_fontspecs(xml)
        assert len(result) == 2
        assert result[0]["id"] == "0"
        assert result[1]["id"] == "1"

    def test_empty_content(self):
        assert _parse_fontspecs("") == []


# ---------------------------------------------------------------------------
# _parse_text_elements
# ---------------------------------------------------------------------------

class TestParseTextElements:
    FONTSPEC_MAP = {"0": {}, "1": {}}

    def test_plain_text(self):
        xml = '<text top="100" left="50" width="200" height="20" font="0">Hello world</text>'
        result = _parse_text_elements(xml, self.FONTSPEC_MAP)
        assert len(result) == 1
        assert result[0]["text"] == "Hello world"
        assert result[0]["bold"] is False
        assert result[0]["italic"] is False

    def test_bold_tag(self):
        xml = '<text top="10" left="20" width="100" height="12" font="0"><b>Bold</b></text>'
        result = _parse_text_elements(xml, self.FONTSPEC_MAP)
        assert result[0]["bold"] is True
        assert result[0]["text"] == "Bold"

    def test_italic_tag(self):
        xml = '<text top="10" left="20" width="100" height="12" font="0"><i>Italic</i></text>'
        result = _parse_text_elements(xml, self.FONTSPEC_MAP)
        assert result[0]["italic"] is True
        assert result[0]["text"] == "Italic"

    def test_bold_and_italic(self):
        xml = '<text top="10" left="20" width="100" height="12" font="0"><b><i>Both</i></b></text>'
        result = _parse_text_elements(xml, self.FONTSPEC_MAP)
        assert result[0]["bold"] is True
        assert result[0]["italic"] is True
        assert result[0]["text"] == "Both"

    def test_html_entity_amp(self):
        xml = '<text top="10" left="20" width="100" height="12" font="0">AT&amp;T</text>'
        result = _parse_text_elements(xml, self.FONTSPEC_MAP)
        assert result[0]["text"] == "AT&T"

    def test_html_entity_lt_gt(self):
        xml = '<text top="10" left="20" width="100" height="12" font="0">&lt;div&gt;</text>'
        result = _parse_text_elements(xml, self.FONTSPEC_MAP)
        assert result[0]["text"] == "<div>"

    def test_html_entity_quot(self):
        xml = '<text top="10" left="20" width="100" height="12" font="0">&quot;quoted&quot;</text>'
        result = _parse_text_elements(xml, self.FONTSPEC_MAP)
        assert result[0]["text"] == '"quoted"'

    def test_html_entity_apos(self):
        xml = '<text top="10" left="20" width="100" height="12" font="0">it&#39;s</text>'
        result = _parse_text_elements(xml, self.FONTSPEC_MAP)
        assert result[0]["text"] == "it's"

    def test_html_entity_numeric(self):
        xml = '<text top="10" left="20" width="100" height="12" font="0">&#169; 2024</text>'
        result = _parse_text_elements(xml, self.FONTSPEC_MAP)
        assert result[0]["text"] == "\u00a9 2024"  # copyright symbol

    def test_html_entities_with_bold_tags(self):
        """The blocking bug scenario: entities inside <b> tags."""
        xml = '<text top="10" left="20" width="200" height="12" font="0"><b>AT&amp;T</b> &lt;100&gt;</text>'
        result = _parse_text_elements(xml, self.FONTSPEC_MAP)
        assert result[0]["text"] == "AT&T <100>"
        assert result[0]["bold"] is True

    def test_coordinates_parsed(self):
        xml = '<text top="42" left="99" width="300" height="15" font="1">Test</text>'
        result = _parse_text_elements(xml, self.FONTSPEC_MAP)
        elem = result[0]
        assert elem["top"] == 42
        assert elem["left"] == 99
        assert elem["width"] == 300
        assert elem["height"] == 15
        assert elem["font_id"] == "1"

    def test_empty_content(self):
        assert _parse_text_elements("", self.FONTSPEC_MAP) == []

    def test_cjk_text(self):
        xml = '<text top="10" left="20" width="100" height="12" font="0">\u5408\u540c\u7f16\u53f7</text>'
        result = _parse_text_elements(xml, self.FONTSPEC_MAP)
        assert result[0]["text"] == "\u5408\u540c\u7f16\u53f7"


# ---------------------------------------------------------------------------
# _parse_images
# ---------------------------------------------------------------------------

class TestParseImages:
    def test_basic_image(self):
        xml = '<image top="10" left="20" width="300" height="400" src="page1-img1.png"/>'
        result = _parse_images(xml)
        assert len(result) == 1
        assert result[0] == {
            "top": 10, "left": 20, "width": 300, "height": 400,
            "src": "page1-img1.png",
        }

    def test_multiple_images(self):
        xml = (
            '<image top="10" left="20" width="300" height="400" src="a.png"/>'
            '<image top="50" left="60" width="100" height="200" src="b.jpg"/>'
        )
        result = _parse_images(xml)
        assert len(result) == 2

    def test_empty_content(self):
        assert _parse_images("") == []


# ---------------------------------------------------------------------------
# _parse_xml_string (integration of sub-parsers)
# ---------------------------------------------------------------------------

class TestParseXmlString:
    MINIMAL_XML = (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<!DOCTYPE pdf2xml SYSTEM "pdf2xml.dtd">\n'
        '<pdf2xml>\n'
        '<page number="1" position="absolute" top="0" left="0" height="1188" width="918">\n'
        '<fontspec id="0" size="24" family="Arial" color="#000000"/>\n'
        '<text top="100" left="50" width="200" height="20" font="0">Hello &amp; World</text>\n'
        '<image top="300" left="50" width="400" height="200" src="img.png"/>\n'
        '</page>\n'
        '</pdf2xml>'
    )

    def test_single_page(self):
        result = _parse_xml_string(self.MINIMAL_XML)
        assert len(result["pages"]) == 1
        page = result["pages"][0]
        assert page["number"] == 1
        assert page["width_px"] == 918
        assert page["height_px"] == 1188
        assert len(page["fontspecs"]) == 1
        assert len(page["text_elements"]) == 1
        assert page["text_elements"][0]["text"] == "Hello & World"  # entity decoded
        assert len(page["images"]) == 1

    def test_multi_page(self):
        xml = (
            '<pdf2xml>\n'
            '<page number="1" position="absolute" top="0" left="0" height="1000" width="800">\n'
            '<fontspec id="0" size="12" family="Times" color="#000000"/>\n'
            '<text top="10" left="10" width="100" height="12" font="0">Page 1</text>\n'
            '</page>\n'
            '<page number="2" position="absolute" top="0" left="0" height="1000" width="800">\n'
            '<text top="10" left="10" width="100" height="12" font="0">Page 2</text>\n'
            '</page>\n'
            '</pdf2xml>'
        )
        result = _parse_xml_string(xml)
        assert len(result["pages"]) == 2
        assert result["pages"][0]["text_elements"][0]["text"] == "Page 1"
        assert result["pages"][1]["text_elements"][0]["text"] == "Page 2"

    def test_fontspec_shared_across_pages(self):
        """Fontspec declared on page 1 should be available to page 2."""
        xml = (
            '<pdf2xml>\n'
            '<page number="1" position="absolute" top="0" left="0" height="1000" width="800">\n'
            '<fontspec id="0" size="12" family="Times" color="#000000"/>\n'
            '<text top="10" left="10" width="100" height="12" font="0"><b>Bold</b></text>\n'
            '</page>\n'
            '<page number="2" position="absolute" top="0" left="0" height="1000" width="800">\n'
            '<text top="10" left="10" width="100" height="12" font="0">Normal</text>\n'
            '</page>\n'
            '</pdf2xml>'
        )
        result = _parse_xml_string(xml)
        # Page 1 text is bold (from <b> tag)
        assert result["pages"][0]["text_elements"][0]["bold"] is True
        # Page 2 text references same font but no <b> tag — should not be bold
        assert result["pages"][1]["text_elements"][0]["bold"] is False

    def test_empty_xml(self):
        result = _parse_xml_string("")
        assert result == {"pages": []}


# ---------------------------------------------------------------------------
# Constants sanity
# ---------------------------------------------------------------------------

class TestConstants:
    def test_scale_factor(self):
        assert PDFTOHTML_SCALE == 1.5

    def test_chars_threshold_positive(self):
        assert DIGITAL_MIN_CHARS_PER_PAGE > 0

    def test_para_grouping_constants(self):
        assert PARA_VERTICAL_GAP_FACTOR > 1.0
        assert PARA_LEFT_TOLERANCE_PX > 0
        assert TABLE_ROW_MIN_ELEMENTS >= 3
        assert ALIGNMENT_TOLERANCE_PX > 0


# ---------------------------------------------------------------------------
# _find_table_row_tops
# ---------------------------------------------------------------------------

class TestFindTableRowTops:
    def test_simple_table_row(self):
        """3 elements at same top = table row."""
        elems = [
            {"top": 100, "left": 50, "width": 80, "height": 18},
            {"top": 100, "left": 200, "width": 80, "height": 18},
            {"top": 100, "left": 350, "width": 80, "height": 18},
        ]
        assert 100 in _find_table_row_tops(elems)

    def test_two_elements_not_table(self):
        """Only 2 elements at same top — not enough for a table row."""
        elems = [
            {"top": 100, "left": 50, "width": 80, "height": 18},
            {"top": 100, "left": 200, "width": 80, "height": 18},
        ]
        assert _find_table_row_tops(elems) == set()

    def test_tolerance_grouping(self):
        """Elements within top tolerance should be grouped."""
        elems = [
            {"top": 100, "left": 50, "width": 80, "height": 18},
            {"top": 103, "left": 200, "width": 80, "height": 18},
            {"top": 100, "left": 350, "width": 80, "height": 18},
        ]
        tops = _find_table_row_tops(elems)
        assert 100 in tops
        assert 103 in tops

    def test_expansion_catches_wrapped_header(self):
        """Pass 2: rows with 2 elements near confirmed table rows get expanded."""
        # top=232 has 2 elements, top=242 has 3, top=251 has 2
        elems = [
            {"top": 232, "left": 119, "width": 39, "height": 18},
            {"top": 232, "left": 216, "width": 73, "height": 18},
            {"top": 242, "left": 385, "width": 88, "height": 18},
            {"top": 242, "left": 537, "width": 112, "height": 18},
            {"top": 242, "left": 699, "width": 71, "height": 18},
            {"top": 251, "left": 134, "width": 9, "height": 18},
            {"top": 251, "left": 212, "width": 77, "height": 18},
        ]
        tops = _find_table_row_tops(elems)
        assert 232 in tops, "Should be expanded from nearby table row"
        assert 242 in tops, "Definite table row (3 elements)"
        assert 251 in tops, "Should be expanded from nearby table row"

    def test_empty_input(self):
        assert _find_table_row_tops([]) == set()

    def test_body_text_not_table(self):
        """Regular body lines (1 element per top) are not table rows."""
        elems = [
            {"top": 354, "left": 108, "width": 673, "height": 18},
            {"top": 373, "left": 108, "width": 703, "height": 18},
            {"top": 392, "left": 108, "width": 541, "height": 18},
        ]
        assert _find_table_row_tops(elems) == set()


# ---------------------------------------------------------------------------
# _is_list_item
# ---------------------------------------------------------------------------

class TestIsListItem:
    def test_numbered_period(self):
        assert _is_list_item("1. First item") is True

    def test_numbered_paren(self):
        assert _is_list_item("3) Third item") is True

    def test_multi_digit(self):
        assert _is_list_item("10. Tenth item") is True

    def test_dash(self):
        assert _is_list_item("- dash item") is True

    def test_bullet(self):
        assert _is_list_item("\u2022 bullet item") is True

    def test_em_dash(self):
        assert _is_list_item("\u2014 em dash item") is True

    def test_regular_text(self):
        assert _is_list_item("Hello world") is False

    def test_number_in_middle(self):
        assert _is_list_item("The total is 12,000.") is False

    def test_article_heading(self):
        assert _is_list_item("Article 1: Something") is False


# ---------------------------------------------------------------------------
# _detect_alignment
# ---------------------------------------------------------------------------

class TestDetectAlignment:
    def test_left_aligned(self):
        lines = [
            {"left": 108, "width": 673},
            {"left": 108, "width": 703},
            {"left": 108, "width": 541},
        ]
        assert _detect_alignment(lines, 918) == "left"

    def test_justify(self):
        """Same left AND same right = justify."""
        lines = [
            {"left": 108, "width": 700},
            {"left": 108, "width": 695},
            {"left": 108, "width": 690},
        ]
        assert _detect_alignment(lines, 918) == "justify"

    def test_center(self):
        lines = [
            {"left": 126, "width": 675},  # center = 463.5
            {"left": 195, "width": 528},  # center = 459.0
        ]
        assert _detect_alignment(lines, 918) == "center"

    def test_right_aligned(self):
        lines = [
            {"left": 500, "width": 400},  # right = 900
            {"left": 600, "width": 300},  # right = 900
        ]
        assert _detect_alignment(lines, 918) == "right"

    def test_single_line_defaults_left(self):
        assert _detect_alignment([{"left": 108, "width": 500}], 918) == "left"

    def test_empty_defaults_left(self):
        assert _detect_alignment([], 918) == "left"


# ---------------------------------------------------------------------------
# _build_paragraph
# ---------------------------------------------------------------------------

class TestBuildParagraph:
    def test_three_line_paragraph(self):
        lines = [
            {"top": 354, "left": 108, "width": 673, "height": 18,
             "font_id": "3", "text": "Line one. ", "bold": False, "italic": False},
            {"top": 373, "left": 108, "width": 703, "height": 18,
             "font_id": "3", "text": "Line two. ", "bold": False, "italic": False},
            {"top": 392, "left": 108, "width": 541, "height": 18,
             "font_id": "3", "text": "Line three.", "bold": False, "italic": False},
        ]
        p = _build_paragraph(lines, 918)
        assert p["type"] == "paragraph"
        assert p["top"] == 354
        assert p["left"] == 108
        assert p["width"] == 703  # max(108+673, 108+703, 108+541) - 108
        assert p["height"] == 392 + 18 - 354  # 56
        assert p["font_id"] == "3"
        assert p["bold"] is False
        assert p["italic"] is False
        assert p["text"] == "Line one. Line two. Line three."
        assert len(p["lines"]) == 3

    def test_bold_propagation(self):
        """Bold is True if any line is bold."""
        lines = [
            {"top": 10, "left": 10, "width": 100, "height": 12,
             "font_id": "0", "text": "Normal ", "bold": False, "italic": False},
            {"top": 22, "left": 10, "width": 100, "height": 12,
             "font_id": "0", "text": "Bold", "bold": True, "italic": False},
        ]
        p = _build_paragraph(lines, 500)
        assert p["bold"] is True

    def test_dominant_font_id(self):
        """Most frequent font_id wins."""
        lines = [
            {"top": 10, "left": 10, "width": 100, "height": 12,
             "font_id": "5", "text": "A", "bold": False, "italic": False},
            {"top": 22, "left": 10, "width": 100, "height": 12,
             "font_id": "3", "text": "B", "bold": False, "italic": False},
            {"top": 34, "left": 10, "width": 100, "height": 12,
             "font_id": "3", "text": "C", "bold": False, "italic": False},
        ]
        p = _build_paragraph(lines, 500)
        assert p["font_id"] == "3"

    def test_single_line(self):
        lines = [
            {"top": 100, "left": 50, "width": 200, "height": 20,
             "font_id": "1", "text": "Single", "bold": True, "italic": False},
        ]
        p = _build_paragraph(lines, 500)
        assert p["text"] == "Single"
        assert p["alignment"] == "left"


# ---------------------------------------------------------------------------
# group_text_into_paragraphs
# ---------------------------------------------------------------------------

class TestGroupTextIntoParagraphs:
    def test_empty_page(self):
        page = {"text_elements": [], "width_px": 918}
        assert group_text_into_paragraphs(page) == []

    def test_single_element(self):
        page = {
            "text_elements": [
                {"top": 100, "left": 50, "width": 200, "height": 20,
                 "font_id": "0", "text": "Hello", "bold": False, "italic": False},
            ],
            "width_px": 918,
        }
        result = group_text_into_paragraphs(page)
        assert len(result) == 1
        assert result[0]["text"] == "Hello"

    def test_merges_consecutive_lines(self):
        """Three lines with same font/style/position merge into one paragraph."""
        page = {
            "text_elements": [
                {"top": 100, "left": 50, "width": 400, "height": 18,
                 "font_id": "0", "text": "First line.", "bold": False, "italic": False},
                {"top": 119, "left": 50, "width": 400, "height": 18,
                 "font_id": "0", "text": "Second line.", "bold": False, "italic": False},
                {"top": 138, "left": 50, "width": 300, "height": 18,
                 "font_id": "0", "text": "Third line.", "bold": False, "italic": False},
            ],
            "width_px": 918,
        }
        result = group_text_into_paragraphs(page)
        assert len(result) == 1
        assert "First line." in result[0]["text"]
        assert "Third line." in result[0]["text"]

    def test_splits_on_font_change(self):
        """Different font_id = new paragraph."""
        page = {
            "text_elements": [
                {"top": 100, "left": 50, "width": 400, "height": 18,
                 "font_id": "0", "text": "Heading", "bold": True, "italic": False},
                {"top": 119, "left": 50, "width": 400, "height": 18,
                 "font_id": "1", "text": "Body text.", "bold": False, "italic": False},
            ],
            "width_px": 918,
        }
        result = group_text_into_paragraphs(page)
        assert len(result) == 2

    def test_splits_on_large_gap(self):
        """Large vertical gap = new paragraph."""
        page = {
            "text_elements": [
                {"top": 100, "left": 50, "width": 400, "height": 18,
                 "font_id": "0", "text": "Para 1.", "bold": False, "italic": False},
                {"top": 200, "left": 50, "width": 400, "height": 18,
                 "font_id": "0", "text": "Para 2.", "bold": False, "italic": False},
            ],
            "width_px": 918,
        }
        result = group_text_into_paragraphs(page)
        assert len(result) == 2

    def test_splits_on_list_item(self):
        """List items start new paragraphs even if vertically close."""
        page = {
            "text_elements": [
                {"top": 100, "left": 50, "width": 400, "height": 18,
                 "font_id": "0", "text": "1. First item", "bold": False, "italic": False},
                {"top": 119, "left": 50, "width": 400, "height": 18,
                 "font_id": "0", "text": "2. Second item", "bold": False, "italic": False},
                {"top": 138, "left": 50, "width": 400, "height": 18,
                 "font_id": "0", "text": "3. Third item", "bold": False, "italic": False},
            ],
            "width_px": 918,
        }
        result = group_text_into_paragraphs(page)
        assert len(result) == 3

    def test_merges_centered_lines(self):
        """Centered text (different left but same center) should merge."""
        page = {
            "text_elements": [
                {"top": 100, "left": 126, "width": 675, "height": 37,
                 "font_id": "1", "text": "Long title line one", "bold": True, "italic": False},
                {"top": 138, "left": 195, "width": 528, "height": 37,
                 "font_id": "1", "text": "and shorter line two", "bold": True, "italic": False},
            ],
            "width_px": 918,
        }
        result = group_text_into_paragraphs(page)
        assert len(result) == 1
        assert result[0]["alignment"] == "center"

    def test_table_rows_stay_separate(self):
        """Elements detected as table rows should not merge."""
        page = {
            "text_elements": [
                {"top": 100, "left": 50, "width": 80, "height": 18,
                 "font_id": "0", "text": "Col1", "bold": False, "italic": False},
                {"top": 100, "left": 200, "width": 80, "height": 18,
                 "font_id": "0", "text": "Col2", "bold": False, "italic": False},
                {"top": 100, "left": 350, "width": 80, "height": 18,
                 "font_id": "0", "text": "Col3", "bold": False, "italic": False},
            ],
            "width_px": 918,
        }
        result = group_text_into_paragraphs(page)
        assert len(result) == 3

    def test_no_text_loss(self):
        """Every input element must appear in exactly one paragraph's lines."""
        page = {
            "text_elements": [
                {"top": 100, "left": 50, "width": 400, "height": 18,
                 "font_id": "0", "text": "A", "bold": False, "italic": False},
                {"top": 119, "left": 50, "width": 400, "height": 18,
                 "font_id": "0", "text": "B", "bold": False, "italic": False},
                {"top": 200, "left": 50, "width": 400, "height": 18,
                 "font_id": "1", "text": "C", "bold": True, "italic": False},
                {"top": 200, "left": 300, "width": 80, "height": 18,
                 "font_id": "1", "text": "D", "bold": True, "italic": False},
                {"top": 200, "left": 500, "width": 80, "height": 18,
                 "font_id": "1", "text": "E", "bold": True, "italic": False},
            ],
            "width_px": 918,
        }
        result = group_text_into_paragraphs(page)
        total_lines = sum(len(p["lines"]) for p in result)
        assert total_lines == 5

    def test_paragraph_output_shape(self):
        """Verify output dict has all required keys."""
        page = {
            "text_elements": [
                {"top": 100, "left": 50, "width": 400, "height": 18,
                 "font_id": "0", "text": "Test", "bold": False, "italic": False},
            ],
            "width_px": 918,
        }
        result = group_text_into_paragraphs(page)
        p = result[0]
        required_keys = {
            "type", "text", "lines", "top", "left", "width", "height",
            "font_id", "bold", "italic", "alignment",
        }
        assert set(p.keys()) == required_keys

    def test_missing_text_elements_key(self):
        """Graceful handling of missing text_elements."""
        assert group_text_into_paragraphs({"width_px": 918}) == []
        assert group_text_into_paragraphs({}) == []


# ---------------------------------------------------------------------------
# group_all_pages
# ---------------------------------------------------------------------------

class TestGroupAllPages:
    def test_adds_paragraphs_key(self):
        parsed = {
            "pages": [{
                "number": 1,
                "width_px": 918,
                "height_px": 1188,
                "text_elements": [
                    {"top": 100, "left": 50, "width": 400, "height": 18,
                     "font_id": "0", "text": "Hello", "bold": False, "italic": False},
                ],
                "images": [],
                "fontspecs": [],
            }]
        }
        result = group_all_pages(parsed)
        assert "paragraphs" in result["pages"][0]
        assert len(result["pages"][0]["paragraphs"]) == 1

    def test_preserves_original_fields(self):
        parsed = {
            "pages": [{
                "number": 1,
                "width_px": 918,
                "height_px": 1188,
                "text_elements": [],
                "images": [{"top": 0, "left": 0, "width": 100, "height": 100, "src": "a.png"}],
                "fontspecs": [{"id": "0", "size_px": 12}],
            }]
        }
        result = group_all_pages(parsed)
        page = result["pages"][0]
        assert page["number"] == 1
        assert page["width_px"] == 918
        assert len(page["images"]) == 1
        assert len(page["fontspecs"]) == 1

    def test_empty_pages(self):
        assert group_all_pages({"pages": []}) == {"pages": []}

    def test_missing_pages_key(self):
        result = group_all_pages({})
        assert result == {"pages": []}


# ---------------------------------------------------------------------------
# Integration: CLI against real PDFs (if available)
# ---------------------------------------------------------------------------

SCRIPT_DIR = Path(__file__).resolve().parent
SCRIPT_PATH = SCRIPT_DIR / "parse_pdf_structure.py"
# scripts/ -> anything-to-docx/ -> skills/ -> .claude/ -> PROJECT_ROOT
PROJECT_ROOT = SCRIPT_DIR.parent.parent.parent.parent

EN_PDF = PROJECT_ROOT / "contract_en3.pdf"
ZH_PDF = PROJECT_ROOT / "contract_zh3.pdf"


class TestCLIIntegration:
    @pytest.mark.skipif(not EN_PDF.exists(), reason="contract_en3.pdf not found")
    def test_check_digital_english(self):
        result = subprocess.run(
            [sys.executable, str(SCRIPT_PATH), "--check-digital", str(EN_PDF)],
            capture_output=True, text=True, timeout=30,
        )
        assert result.returncode == 0
        data = json.loads(result.stdout)
        assert data["is_digital"] is True

    @pytest.mark.skipif(not ZH_PDF.exists(), reason="contract_zh3.pdf not found")
    def test_check_digital_chinese(self):
        result = subprocess.run(
            [sys.executable, str(SCRIPT_PATH), "--check-digital", str(ZH_PDF)],
            capture_output=True, text=True, timeout=30,
        )
        assert result.returncode == 0
        data = json.loads(result.stdout)
        assert data["is_digital"] is True

    @pytest.mark.skipif(not EN_PDF.exists(), reason="contract_en3.pdf not found")
    def test_full_parse_english(self):
        result = subprocess.run(
            [sys.executable, str(SCRIPT_PATH), "--pdf", str(EN_PDF)],
            capture_output=True, text=True, timeout=30,
        )
        assert result.returncode == 0
        data = json.loads(result.stdout)
        assert "pages" in data
        assert len(data["pages"]) > 0
        # Every text element should have decoded entities (no raw &amp; etc.)
        for page in data["pages"]:
            for elem in page["text_elements"]:
                assert "&amp;" not in elem["text"], f"Raw &amp; in: {elem['text']}"
                assert "&lt;" not in elem["text"], f"Raw &lt; in: {elem['text']}"
                assert "&gt;" not in elem["text"], f"Raw &gt; in: {elem['text']}"

    @pytest.mark.skipif(not ZH_PDF.exists(), reason="contract_zh3.pdf not found")
    def test_full_parse_chinese(self):
        result = subprocess.run(
            [sys.executable, str(SCRIPT_PATH), "--pdf", str(ZH_PDF)],
            capture_output=True, text=True, timeout=30,
        )
        assert result.returncode == 0
        data = json.loads(result.stdout)
        assert "pages" in data
        assert len(data["pages"]) > 0

    @pytest.mark.skipif(not EN_PDF.exists(), reason="contract_en3.pdf not found")
    def test_fonts_flag(self):
        result = subprocess.run(
            [sys.executable, str(SCRIPT_PATH), "--fonts", str(EN_PDF)],
            capture_output=True, text=True, timeout=30,
        )
        assert result.returncode == 0
        fonts = json.loads(result.stdout)
        assert isinstance(fonts, list)
        assert len(fonts) > 0
        # Every font has expected keys
        for f in fonts:
            assert "name" in f
            assert "embedded" in f

    def test_mutually_exclusive_flags(self):
        """--pdf and --check-digital cannot be used together."""
        result = subprocess.run(
            [sys.executable, str(SCRIPT_PATH), "--pdf", "a.pdf", "--check-digital", "b.pdf"],
            capture_output=True, text=True, timeout=10,
        )
        assert result.returncode != 0

    def test_no_flags_errors(self):
        """No mode flag should produce an error."""
        result = subprocess.run(
            [sys.executable, str(SCRIPT_PATH)],
            capture_output=True, text=True, timeout=10,
        )
        assert result.returncode != 0


# ---------------------------------------------------------------------------
# Integration: paragraph grouping against real PDFs
# ---------------------------------------------------------------------------

class TestParagraphGroupingIntegration:
    @pytest.mark.skipif(not EN_PDF.exists(), reason="contract_en3.pdf not found")
    def test_en3_page1_reduces_elements(self):
        """Page 1's 41 elements should collapse into significantly fewer paragraphs."""
        from parse_pdf_structure import parse_pdftohtml_xml
        data = parse_pdftohtml_xml(str(EN_PDF))
        page1 = data["pages"][0]
        paragraphs = group_text_into_paragraphs(page1)
        assert len(page1["text_elements"]) == 41
        assert len(paragraphs) < 41
        # Should be roughly 30-35 (table cells stay separate, body merges)
        assert len(paragraphs) >= 20

    @pytest.mark.skipif(not EN_PDF.exists(), reason="contract_en3.pdf not found")
    def test_en3_title_merged_and_centered(self):
        """Title lines should merge into one centered paragraph."""
        from parse_pdf_structure import parse_pdftohtml_xml
        data = parse_pdftohtml_xml(str(EN_PDF))
        page1 = data["pages"][0]
        paragraphs = group_text_into_paragraphs(page1)
        # Find the title paragraph
        title_paras = [p for p in paragraphs if "Artificial Intelligence" in p["text"]]
        assert len(title_paras) == 1
        title = title_paras[0]
        assert "Service Agreement" in title["text"]
        assert title["alignment"] == "center"
        assert title["bold"] is True
        assert len(title["lines"]) == 2

    @pytest.mark.skipif(not EN_PDF.exists(), reason="contract_en3.pdf not found")
    def test_en3_body_paragraph_merged(self):
        """Article 1 body paragraph (3 lines) should merge."""
        from parse_pdf_structure import parse_pdftohtml_xml
        data = parse_pdftohtml_xml(str(EN_PDF))
        page1 = data["pages"][0]
        paragraphs = group_text_into_paragraphs(page1)
        body = [p for p in paragraphs if "entered into on September" in p["text"]]
        assert len(body) == 1
        assert "good faith" in body[0]["text"]
        assert len(body[0]["lines"]) == 3

    @pytest.mark.skipif(not EN_PDF.exists(), reason="contract_en3.pdf not found")
    def test_en3_list_items_separate(self):
        """Numbered list items 1-4 should be separate paragraphs."""
        from parse_pdf_structure import parse_pdftohtml_xml
        data = parse_pdftohtml_xml(str(EN_PDF))
        page1 = data["pages"][0]
        paragraphs = group_text_into_paragraphs(page1)
        list_items = [p for p in paragraphs if p["text"].startswith(("1.", "2.", "3.", "4."))]
        assert len(list_items) >= 4

    @pytest.mark.skipif(not EN_PDF.exists(), reason="contract_en3.pdf not found")
    def test_en3_no_text_loss(self):
        """All text elements must be accounted for across all pages."""
        from parse_pdf_structure import parse_pdftohtml_xml
        data = parse_pdftohtml_xml(str(EN_PDF))
        enriched = group_all_pages(data)
        for page in enriched["pages"]:
            total = sum(len(p["lines"]) for p in page["paragraphs"])
            assert total == len(page["text_elements"]), \
                f"Page {page['number']}: text loss detected"

    @pytest.mark.skipif(not ZH_PDF.exists(), reason="contract_zh3.pdf not found")
    def test_zh3_no_text_loss(self):
        """All text elements must be accounted for in ZH PDF."""
        from parse_pdf_structure import parse_pdftohtml_xml
        data = parse_pdftohtml_xml(str(ZH_PDF))
        enriched = group_all_pages(data)
        for page in enriched["pages"]:
            total = sum(len(p["lines"]) for p in page["paragraphs"])
            assert total == len(page["text_elements"]), \
                f"Page {page['number']}: text loss detected"

    @pytest.mark.skipif(not EN_PDF.exists(), reason="contract_en3.pdf not found")
    def test_en3_page3_table_cells_separate(self):
        """Page 3 table data rows should remain as individual cells."""
        from parse_pdf_structure import parse_pdftohtml_xml
        data = parse_pdftohtml_xml(str(EN_PDF))
        page3 = data["pages"][2]
        paragraphs = group_text_into_paragraphs(page3)
        # Find cells at table data row top=292 — should be separate
        row_292 = [p for p in paragraphs if p["top"] == 292]
        assert len(row_292) >= 3  # At least phase, condition, amount, date


# ---------------------------------------------------------------------------
# _build_fontspec_lookup
# ---------------------------------------------------------------------------

class TestBuildFontspecLookup:
    def test_basic(self):
        fontspecs = [
            {"id": "0", "size_pts": 12.0},
            {"id": "1", "size_pts": 18.0},
        ]
        lookup = _build_fontspec_lookup(fontspecs)
        assert lookup["0"]["size_pts"] == 12.0
        assert lookup["1"]["size_pts"] == 18.0

    def test_empty(self):
        assert _build_fontspec_lookup([]) == {}

    def test_overwrites_duplicate_ids(self):
        """Last fontspec with same id wins (dict semantics)."""
        fontspecs = [
            {"id": "0", "size_pts": 12.0},
            {"id": "0", "size_pts": 14.0},
        ]
        lookup = _build_fontspec_lookup(fontspecs)
        assert lookup["0"]["size_pts"] == 14.0


# ---------------------------------------------------------------------------
# _determine_body_font_size
# ---------------------------------------------------------------------------

class TestDetermineBodyFontSize:
    def test_most_common_size(self):
        lookup = {
            "0": {"size_pts": 22.0},
            "1": {"size_pts": 11.0},
        }
        paras = [
            {"font_id": "0"},  # 22.0 — heading (1 occurrence)
            {"font_id": "1"},  # 11.0 — body (3 occurrences)
            {"font_id": "1"},
            {"font_id": "1"},
        ]
        assert _determine_body_font_size(paras, lookup) == 11.0

    def test_fallback_when_no_fontspecs(self):
        assert _determine_body_font_size([{"font_id": "99"}], {}) == 11.0

    def test_empty_paragraphs(self):
        assert _determine_body_font_size([], {"0": {"size_pts": 12.0}}) == 11.0


# ---------------------------------------------------------------------------
# _classify_single_paragraph
# ---------------------------------------------------------------------------

class TestClassifySingleParagraph:
    """Unit tests for the pure classification function."""

    FONTSPECS = {
        "title": {"size_pts": 22.0, "family": "Arial", "color": "#000000"},
        "article": {"size_pts": 18.0, "family": "Arial", "color": "#1a5276"},
        "sub_heading": {"size_pts": 14.0, "family": "Arial", "color": "#2e86c1"},
        "body_heading": {"size_pts": 12.0, "family": "Arial", "color": "#000000"},
        "body": {"size_pts": 11.0, "family": "Times", "color": "#333333"},
        "small": {"size_pts": 8.0, "family": "Times", "color": "#999999"},
        "medium": {"size_pts": 11.3, "family": "Times", "color": "#000000"},
        # Boundary-test fonts for mutation coverage
        "just_below_h2": {"size_pts": 13.9, "family": "Arial", "color": "#000000"},
        "just_below_h4": {"size_pts": 10.9, "family": "Times", "color": "#000000"},
    }

    PAGE_HEIGHT_PTS = 792.0
    BODY_SIZE = 11.0

    def _make_para(self, font_id="body", bold=False, top_px=300, **kw):
        """Helper to build a minimal paragraph dict."""
        return {
            "font_id": font_id, "bold": bold, "italic": False,
            "top": top_px, "left": 108, "width": 400, "height": 18,
            "text": kw.get("text", "Some text"), "lines": [],
            "type": "paragraph", "alignment": "left",
        }

    def test_heading_level_1(self):
        """Font size > 18 = heading level 1."""
        para = self._make_para(font_id="title", bold=True)
        result = _classify_single_paragraph(
            para, self.FONTSPECS, self.PAGE_HEIGHT_PTS, self.BODY_SIZE)
        assert result["type"] == "heading"
        assert result["heading_level"] == 1
        assert result["font_size_pts"] == 22.0

    def test_heading_level_2_at_18pts(self):
        """Font size exactly 18.0 = heading level 2 (not 1, since > not >=)."""
        para = self._make_para(font_id="article", bold=True)
        result = _classify_single_paragraph(
            para, self.FONTSPECS, self.PAGE_HEIGHT_PTS, self.BODY_SIZE)
        assert result["type"] == "heading"
        assert result["heading_level"] == 2

    def test_heading_level_2_at_14pts(self):
        para = self._make_para(font_id="sub_heading", bold=True)
        result = _classify_single_paragraph(
            para, self.FONTSPECS, self.PAGE_HEIGHT_PTS, self.BODY_SIZE)
        assert result["type"] == "heading"
        assert result["heading_level"] == 2

    def test_heading_level_3(self):
        para = self._make_para(font_id="body_heading", bold=True)
        result = _classify_single_paragraph(
            para, self.FONTSPECS, self.PAGE_HEIGHT_PTS, self.BODY_SIZE)
        assert result["type"] == "heading"
        assert result["heading_level"] == 3

    def test_heading_level_4_bold_required(self):
        """Level 4 requires bold + size >= 11."""
        bold_para = self._make_para(font_id="body", bold=True)
        result = _classify_single_paragraph(
            bold_para, self.FONTSPECS, self.PAGE_HEIGHT_PTS, self.BODY_SIZE)
        assert result["type"] == "heading"
        assert result["heading_level"] == 4

    def test_no_heading_without_bold_at_11pts(self):
        """11pts without bold = regular paragraph, NOT heading level 4."""
        para = self._make_para(font_id="body", bold=False)
        result = _classify_single_paragraph(
            para, self.FONTSPECS, self.PAGE_HEIGHT_PTS, self.BODY_SIZE)
        assert result["type"] == "paragraph"
        assert "heading_level" not in result

    def test_h2_boundary_below_14pts(self):
        """13.9pts bold = heading level 3, NOT level 2 (catches 14→13 mutation)."""
        para = self._make_para(font_id="just_below_h2", bold=True)
        result = _classify_single_paragraph(
            para, self.FONTSPECS, self.PAGE_HEIGHT_PTS, self.BODY_SIZE)
        assert result["type"] == "heading"
        assert result["heading_level"] == 3, \
            f"13.9pt should be H3, not H{result['heading_level']}"

    def test_h4_boundary_below_11pts(self):
        """10.9pts bold = NOT heading level 4 via fixed threshold (catches 11→10 mutation).
        May still become heading via bold+larger-than-body fallback."""
        para = self._make_para(font_id="just_below_h4", bold=True)
        result = _classify_single_paragraph(
            para, self.FONTSPECS, self.PAGE_HEIGHT_PTS, self.BODY_SIZE)
        # 10.9 < 11 threshold, so the fixed H4 branch should NOT fire.
        # But bold + 10.9 > body(11.0)? No, 10.9 < 11.0. So no heading at all.
        assert result["heading_level"] != 4 if "heading_level" in result else True
        # Actually 10.9 < body_size 11.0, so not heading at all
        assert result["type"] == "paragraph"

    def test_body_paragraph(self):
        para = self._make_para(font_id="body", bold=False)
        result = _classify_single_paragraph(
            para, self.FONTSPECS, self.PAGE_HEIGHT_PTS, self.BODY_SIZE)
        assert result["type"] == "paragraph"
        assert "heading_level" not in result
        assert result["font_family"] == "Times"
        assert result["color"] == "#333333"

    def test_footer_detection(self):
        """Small font at bottom of page = footer."""
        # 792 * 0.92 = 728.64 pts → need top_pts > 728.64 → top_px > 1092.96
        para = self._make_para(font_id="small", top_px=1100)
        result = _classify_single_paragraph(
            para, self.FONTSPECS, self.PAGE_HEIGHT_PTS, self.BODY_SIZE)
        assert result["type"] == "footer"
        assert "heading_level" not in result

    def test_header_detection(self):
        """Small font at top of page = header."""
        # 792 * 0.08 = 63.36 pts → need top_pts < 63.36 → top_px < 95.04
        para = self._make_para(font_id="small", top_px=90)
        result = _classify_single_paragraph(
            para, self.FONTSPECS, self.PAGE_HEIGHT_PTS, self.BODY_SIZE)
        assert result["type"] == "header"

    def test_footer_requires_small_font(self):
        """Normal-sized font at bottom of page = NOT footer."""
        para = self._make_para(font_id="body", top_px=1100)
        result = _classify_single_paragraph(
            para, self.FONTSPECS, self.PAGE_HEIGHT_PTS, self.BODY_SIZE)
        assert result["type"] != "footer"

    def test_header_requires_small_font(self):
        """Normal-sized font at top of page = NOT header."""
        para = self._make_para(font_id="body", top_px=90)
        result = _classify_single_paragraph(
            para, self.FONTSPECS, self.PAGE_HEIGHT_PTS, self.BODY_SIZE)
        assert result["type"] != "header"

    def test_footer_overrides_heading(self):
        """Even if font size >= heading threshold, footer position + small font = footer."""
        # 8pt font at footer position — should be footer, not checked for heading
        para = self._make_para(font_id="small", top_px=1100)
        result = _classify_single_paragraph(
            para, self.FONTSPECS, self.PAGE_HEIGHT_PTS, self.BODY_SIZE)
        assert result["type"] == "footer"

    def test_unknown_font_id_defaults(self):
        """Paragraph with unknown font_id gets default values."""
        para = self._make_para(font_id="nonexistent")
        result = _classify_single_paragraph(
            para, self.FONTSPECS, self.PAGE_HEIGHT_PTS, self.BODY_SIZE)
        assert result["type"] == "paragraph"
        assert result["font_size_pts"] == 0.0
        assert result["font_family"] == ""
        assert result["color"] == "#000000"

    def test_bold_larger_than_body_becomes_heading(self):
        """Bold text larger than body text → heading (even if below fixed threshold)."""
        # medium is 11.3, body is 11.0 — medium is larger
        para = self._make_para(font_id="medium", bold=True)
        result = _classify_single_paragraph(
            para, self.FONTSPECS, self.PAGE_HEIGHT_PTS, self.BODY_SIZE)
        assert result["type"] == "heading"

    def test_preserves_original_fields(self):
        """Classification adds fields but doesn't remove originals."""
        para = self._make_para(font_id="body", text="Hello world")
        result = _classify_single_paragraph(
            para, self.FONTSPECS, self.PAGE_HEIGHT_PTS, self.BODY_SIZE)
        assert result["text"] == "Hello world"
        assert result["left"] == 108
        assert result["alignment"] == "left"
        assert "lines" in result

    def test_whitespace_only_not_heading(self):
        """Whitespace-only text at heading font size should NOT become heading."""
        para = self._make_para(font_id="title", bold=True, text="\u3000 \t")  # CJK space
        result = _classify_single_paragraph(
            para, self.FONTSPECS, self.PAGE_HEIGHT_PTS, self.BODY_SIZE)
        assert result["type"] == "paragraph"
        assert "heading_level" not in result

    def test_table_cell_not_heading(self):
        """Table cell paragraphs should NOT be classified as headings."""
        para = self._make_para(font_id="title", bold=True)
        para["is_table_cell"] = True
        result = _classify_single_paragraph(
            para, self.FONTSPECS, self.PAGE_HEIGHT_PTS, self.BODY_SIZE)
        assert result["type"] == "paragraph"
        assert "heading_level" not in result

    def test_does_not_mutate_input(self):
        """Pure function — input dict must not be modified."""
        para = self._make_para(font_id="title", bold=True)
        original_type = para["type"]
        _classify_single_paragraph(
            para, self.FONTSPECS, self.PAGE_HEIGHT_PTS, self.BODY_SIZE)
        assert para["type"] == original_type  # still "paragraph"


# ---------------------------------------------------------------------------
# classify_paragraphs
# ---------------------------------------------------------------------------

class TestClassifyParagraphs:
    FONTSPECS = {
        "0": {"size_pts": 9.3, "family": "Sans", "color": "#666666"},
        "1": {"size_pts": 22.0, "family": "Sans", "color": "#1a5276"},
        "2": {"size_pts": 18.0, "family": "Sans", "color": "#1a5276"},
        "3": {"size_pts": 11.3, "family": "Sans", "color": "#000000"},
        "7": {"size_pts": 8.0, "family": "Sans", "color": "#999999"},
    }

    def _make_page(self, paragraphs):
        return {
            "height_px": 1188,
            "height_pts": 792.0,
            "width_px": 918,
            "fontspecs": [],
            "paragraphs": paragraphs,
        }

    def test_mixed_classification(self):
        paras = [
            {"font_id": "1", "bold": True, "top": 200, "left": 126, "width": 675,
             "height": 37, "text": "Title", "lines": [], "type": "paragraph",
             "italic": False, "alignment": "center"},
            {"font_id": "3", "bold": False, "top": 400, "left": 108, "width": 700,
             "height": 18, "text": "Body text here.", "lines": [], "type": "paragraph",
             "italic": False, "alignment": "left"},
            {"font_id": "7", "bold": False, "top": 1121, "left": 108, "width": 400,
             "height": 12, "text": "Footer text", "lines": [], "type": "paragraph",
             "italic": False, "alignment": "left"},
        ]
        page = self._make_page(paras)
        result = classify_paragraphs(page, self.FONTSPECS)
        assert len(result) == 3
        assert result[0]["type"] == "heading"
        assert result[0]["heading_level"] == 1
        assert result[1]["type"] == "paragraph"
        assert result[2]["type"] == "footer"

    def test_empty_paragraphs(self):
        page = self._make_page([])
        assert classify_paragraphs(page, self.FONTSPECS) == []

    def test_all_paragraphs_enriched(self):
        """Every classified paragraph must have font_size_pts, font_family, color."""
        paras = [
            {"font_id": "3", "bold": False, "top": 300, "left": 108, "width": 400,
             "height": 18, "text": "A", "lines": [], "type": "paragraph",
             "italic": False, "alignment": "left"},
            {"font_id": "3", "bold": False, "top": 320, "left": 108, "width": 400,
             "height": 18, "text": "B", "lines": [], "type": "paragraph",
             "italic": False, "alignment": "left"},
        ]
        page = self._make_page(paras)
        result = classify_paragraphs(page, self.FONTSPECS)
        for p in result:
            assert "font_size_pts" in p
            assert "font_family" in p
            assert "color" in p
            assert "type" in p


# ---------------------------------------------------------------------------
# classify_all_pages
# ---------------------------------------------------------------------------

class TestClassifyAllPages:
    def test_adds_classification_to_all_pages(self):
        grouped = {
            "pages": [
                {
                    "number": 1,
                    "height_px": 1188, "height_pts": 792.0,
                    "width_px": 918, "width_pts": 612.0,
                    "fontspecs": [
                        {"id": "0", "size_px": 33, "size_pts": 22.0,
                         "family": "Arial", "raw_family": "Arial-Bold", "color": "#000000"},
                        {"id": "1", "size_px": 17, "size_pts": 11.3,
                         "family": "Arial", "raw_family": "Arial", "color": "#000000"},
                    ],
                    "text_elements": [],
                    "images": [],
                    "paragraphs": [
                        {"font_id": "0", "bold": True, "top": 200, "left": 108,
                         "width": 500, "height": 33, "text": "Title",
                         "lines": [], "type": "paragraph", "italic": False,
                         "alignment": "center"},
                        {"font_id": "1", "bold": False, "top": 400, "left": 108,
                         "width": 700, "height": 17, "text": "Body",
                         "lines": [], "type": "paragraph", "italic": False,
                         "alignment": "left"},
                    ],
                },
            ],
        }
        result = classify_all_pages(grouped)
        page1 = result["pages"][0]
        assert page1["paragraphs"][0]["type"] == "heading"
        assert page1["paragraphs"][0]["heading_level"] == 1
        assert page1["paragraphs"][1]["type"] == "paragraph"

    def test_preserves_non_paragraph_fields(self):
        grouped = {
            "pages": [{
                "number": 1,
                "height_px": 1000, "height_pts": 666.7,
                "width_px": 800, "width_pts": 533.3,
                "fontspecs": [],
                "text_elements": [{"fake": True}],
                "images": [{"src": "img.png"}],
                "paragraphs": [],
            }],
        }
        result = classify_all_pages(grouped)
        page = result["pages"][0]
        assert page["number"] == 1
        assert len(page["text_elements"]) == 1
        assert len(page["images"]) == 1

    def test_empty_pages(self):
        assert classify_all_pages({"pages": []}) == {"pages": []}

    def test_cumulative_fontspec_across_pages(self):
        """Fontspec declared on page 1 must be usable on page 2."""
        grouped = {
            "pages": [
                {
                    "number": 1,
                    "height_px": 1188, "height_pts": 792.0,
                    "width_px": 918, "width_pts": 612.0,
                    "fontspecs": [
                        {"id": "5", "size_px": 33, "size_pts": 22.0,
                         "family": "Courier", "raw_family": "Courier", "color": "#ff0000"},
                    ],
                    "text_elements": [],
                    "images": [],
                    "paragraphs": [
                        {"font_id": "5", "bold": True, "top": 200, "left": 108,
                         "width": 500, "height": 33, "text": "Page 1 title",
                         "lines": [], "type": "paragraph", "italic": False,
                         "alignment": "center"},
                    ],
                },
                {
                    "number": 2,
                    "height_px": 1188, "height_pts": 792.0,
                    "width_px": 918, "width_pts": 612.0,
                    "fontspecs": [],  # no fontspecs declared on page 2
                    "text_elements": [],
                    "images": [],
                    "paragraphs": [
                        {"font_id": "5", "bold": True, "top": 200, "left": 108,
                         "width": 500, "height": 33, "text": "Page 2 heading",
                         "lines": [], "type": "paragraph", "italic": False,
                         "alignment": "center"},
                    ],
                },
            ],
        }
        result = classify_all_pages(grouped)
        # Page 2 paragraph references font_id "5" declared on page 1
        p2_para = result["pages"][1]["paragraphs"][0]
        assert p2_para["type"] == "heading"
        assert p2_para["font_family"] == "Courier"
        assert p2_para["color"] == "#ff0000"

    def test_does_not_mutate_input(self):
        """Pure function — input must not be modified."""
        grouped = {
            "pages": [{
                "number": 1,
                "height_px": 1188, "height_pts": 792.0,
                "width_px": 918, "width_pts": 612.0,
                "fontspecs": [
                    {"id": "0", "size_px": 17, "size_pts": 11.3,
                     "family": "Arial", "raw_family": "Arial", "color": "#000000"},
                ],
                "text_elements": [],
                "images": [],
                "paragraphs": [
                    {"font_id": "0", "bold": False, "top": 400, "left": 108,
                     "width": 700, "height": 17, "text": "Body",
                     "lines": [], "type": "paragraph", "italic": False,
                     "alignment": "left"},
                ],
            }],
        }
        original_type = grouped["pages"][0]["paragraphs"][0]["type"]
        classify_all_pages(grouped)
        assert grouped["pages"][0]["paragraphs"][0]["type"] == original_type


# ---------------------------------------------------------------------------
# Classification constants sanity
# ---------------------------------------------------------------------------

class TestClassificationConstants:
    def test_heading_thresholds_ordered(self):
        """Higher heading levels have larger min font sizes."""
        assert HEADING_LEVEL_1_MIN_PTS > HEADING_LEVEL_2_MIN_PTS
        assert HEADING_LEVEL_2_MIN_PTS > HEADING_LEVEL_3_MIN_PTS
        assert HEADING_LEVEL_3_MIN_PTS > HEADING_LEVEL_4_MIN_PTS

    def test_header_footer_ratios(self):
        assert 0 < HEADER_BOTTOM_RATIO < FOOTER_TOP_RATIO < 1
        assert HEADER_FOOTER_MAX_FONT_PTS > 0


# ---------------------------------------------------------------------------
# Integration: classification against real PDFs
# ---------------------------------------------------------------------------

class TestClassificationIntegration:
    @pytest.mark.skipif(not EN_PDF.exists(), reason="contract_en3.pdf not found")
    def test_en3_title_is_heading_level_1(self):
        from parse_pdf_structure import parse_pdftohtml_xml
        data = parse_pdftohtml_xml(str(EN_PDF))
        grouped = group_all_pages(data)
        classified = classify_all_pages(grouped)
        page1 = classified["pages"][0]
        title = [p for p in page1["paragraphs"]
                 if "Artificial Intelligence" in p["text"]]
        assert len(title) == 1
        assert title[0]["type"] == "heading"
        assert title[0]["heading_level"] == 1

    @pytest.mark.skipif(not EN_PDF.exists(), reason="contract_en3.pdf not found")
    def test_en3_articles_are_heading_level_2(self):
        from parse_pdf_structure import parse_pdftohtml_xml
        data = parse_pdftohtml_xml(str(EN_PDF))
        grouped = group_all_pages(data)
        classified = classify_all_pages(grouped)
        page1 = classified["pages"][0]
        articles = [p for p in page1["paragraphs"]
                    if p["text"].startswith("Article")]
        assert len(articles) == 2  # Article 1 and Article 2
        for a in articles:
            assert a["type"] == "heading"
            assert a["heading_level"] == 2

    @pytest.mark.skipif(not EN_PDF.exists(), reason="contract_en3.pdf not found")
    def test_en3_body_is_paragraph(self):
        from parse_pdf_structure import parse_pdftohtml_xml
        data = parse_pdftohtml_xml(str(EN_PDF))
        grouped = group_all_pages(data)
        classified = classify_all_pages(grouped)
        page1 = classified["pages"][0]
        body = [p for p in page1["paragraphs"]
                if "entered into on September" in p["text"]]
        assert len(body) == 1
        assert body[0]["type"] == "paragraph"

    @pytest.mark.skipif(not EN_PDF.exists(), reason="contract_en3.pdf not found")
    def test_en3_footer_detected(self):
        from parse_pdf_structure import parse_pdftohtml_xml
        data = parse_pdftohtml_xml(str(EN_PDF))
        grouped = group_all_pages(data)
        classified = classify_all_pages(grouped)
        page1 = classified["pages"][0]
        footers = [p for p in page1["paragraphs"] if p["type"] == "footer"]
        assert len(footers) >= 1
        footer_texts = " ".join(f["text"] for f in footers)
        assert "Confidential" in footer_texts

    @pytest.mark.skipif(not EN_PDF.exists(), reason="contract_en3.pdf not found")
    def test_en3_all_paragraphs_have_classification_fields(self):
        """Every paragraph across all pages must have classification fields."""
        from parse_pdf_structure import parse_pdftohtml_xml
        data = parse_pdftohtml_xml(str(EN_PDF))
        grouped = group_all_pages(data)
        classified = classify_all_pages(grouped)
        for page in classified["pages"]:
            for p in page["paragraphs"]:
                assert "font_size_pts" in p, f"Missing font_size_pts: {p['text'][:40]}"
                assert "font_family" in p, f"Missing font_family: {p['text'][:40]}"
                assert "color" in p, f"Missing color: {p['text'][:40]}"
                assert p["type"] in ("heading", "paragraph", "footer", "header"), \
                    f"Invalid type {p['type']}: {p['text'][:40]}"
                if p["type"] == "heading":
                    assert "heading_level" in p
                    assert 1 <= p["heading_level"] <= 4

    @pytest.mark.skipif(not EN_PDF.exists(), reason="contract_en3.pdf not found")
    def test_en3_paragraph_count_preserved(self):
        """Classification must not change the number of paragraphs."""
        from parse_pdf_structure import parse_pdftohtml_xml
        data = parse_pdftohtml_xml(str(EN_PDF))
        grouped = group_all_pages(data)
        classified = classify_all_pages(grouped)
        for orig_page, cls_page in zip(grouped["pages"], classified["pages"]):
            assert len(orig_page["paragraphs"]) == len(cls_page["paragraphs"])

    @pytest.mark.skipif(not ZH_PDF.exists(), reason="contract_zh3.pdf not found")
    def test_zh3_all_paragraphs_classified(self):
        """ZH PDF: every paragraph must have classification fields."""
        from parse_pdf_structure import parse_pdftohtml_xml
        data = parse_pdftohtml_xml(str(ZH_PDF))
        grouped = group_all_pages(data)
        classified = classify_all_pages(grouped)
        for page in classified["pages"]:
            for p in page["paragraphs"]:
                assert "font_size_pts" in p
                assert p["type"] in ("heading", "paragraph", "footer", "header")

    @pytest.mark.skipif(not EN_PDF.exists(), reason="contract_en3.pdf not found")
    def test_en3_composition_pipeline(self):
        """Verify the full pipeline: parse -> group -> classify."""
        from parse_pdf_structure import parse_pdftohtml_xml
        data = parse_pdftohtml_xml(str(EN_PDF))
        grouped = group_all_pages(data)
        classified = classify_all_pages(grouped)

        # Pipeline preserves structure
        assert len(classified["pages"]) == len(data["pages"])
        for page in classified["pages"]:
            assert "paragraphs" in page
            assert "text_elements" in page
            assert "images" in page
            assert "fontspecs" in page

    @pytest.mark.skipif(not EN_PDF.exists(), reason="contract_en3.pdf not found")
    def test_en3_footer_not_false_table_row(self):
        """T2 evaluator finding #1: footer elements should NOT be misclassified.

        The footer at bottom of page 1 has 3 elements at similar top positions,
        which could be falsely detected as a table row. After classification,
        these should be marked as footer.
        """
        from parse_pdf_structure import parse_pdftohtml_xml
        data = parse_pdftohtml_xml(str(EN_PDF))
        grouped = group_all_pages(data)
        classified = classify_all_pages(grouped)
        page1 = classified["pages"][0]

        # Find elements near bottom of page (top > 1100 px)
        bottom_elems = [p for p in page1["paragraphs"] if p["top"] > 1100]
        # The small-font ones should be footer
        small_bottom = [p for p in bottom_elems if p["font_size_pts"] <= 10]
        for p in small_bottom:
            assert p["type"] == "footer", \
                f"Expected footer, got {p['type']}: {p['text'][:40]}"


# ---------------------------------------------------------------------------
# _cluster_columns
# ---------------------------------------------------------------------------

class TestClusterColumns:
    def test_five_distinct_columns(self):
        """EN3 page 3 table left positions should cluster into 5 columns."""
        lefts = [119, 134, 177, 212, 216, 342, 348, 385, 537, 556, 694, 699]
        cols = _cluster_columns(lefts)
        assert len(cols) == 5
        # Check anchors are sensible
        assert cols[0] == 119  # Phase
        assert cols[1] == 177  # Payment Condition
        assert cols[2] == 342  # Deliverable
        assert cols[3] == 537  # Amount
        assert cols[4] == 694  # Due Date

    def test_no_chain_drift(self):
        """Anchor-based clustering prevents chaining: 100, 140, 180 should be 2 clusters.
        With tolerance 50: 100→140=40<50 same cluster, but 180-100=80>50 → new cluster."""
        cols = _cluster_columns([100, 140, 180])
        assert len(cols) == 2
        assert cols[0] == 100
        assert cols[1] == 180

    def test_single_position(self):
        assert _cluster_columns([500]) == [500]

    def test_empty(self):
        assert _cluster_columns([]) == []

    def test_duplicates_handled(self):
        """Duplicate positions should not create extra clusters."""
        cols = _cluster_columns([100, 100, 200, 200])
        assert len(cols) == 2

    def test_tolerance_boundary(self):
        """Two positions exactly at tolerance boundary cluster together."""
        cols = _cluster_columns([100, 100 + TABLE_COL_CLUSTER_TOLERANCE_PX])
        assert len(cols) == 1


# ---------------------------------------------------------------------------
# _assign_col_index
# ---------------------------------------------------------------------------

class TestAssignColIndex:
    def test_exact_match(self):
        assert _assign_col_index(119, [119, 177, 342, 537, 694]) == 0
        assert _assign_col_index(342, [119, 177, 342, 537, 694]) == 2

    def test_closest_match(self):
        """134 is closer to 119 (diff=15) than 177 (diff=43)."""
        assert _assign_col_index(134, [119, 177, 342, 537, 694]) == 0

    def test_close_to_second_col(self):
        """216 is closer to 177 (diff=39) than 342 (diff=126)."""
        assert _assign_col_index(216, [119, 177, 342, 537, 694]) == 1


# ---------------------------------------------------------------------------
# _group_elements_into_rows
# ---------------------------------------------------------------------------

class TestGroupElementsIntoRows:
    def test_three_elements_same_top(self):
        elems = [
            {"top": 100, "left": 50, "width": 80, "height": 18},
            {"top": 100, "left": 200, "width": 80, "height": 18},
            {"top": 100, "left": 350, "width": 80, "height": 18},
        ]
        rows = _group_elements_into_rows(elems)
        assert len(rows) == 1
        assert len(rows[0]) == 3
        # Sorted by left
        assert rows[0][0]["left"] == 50
        assert rows[0][2]["left"] == 350

    def test_two_rows(self):
        elems = [
            {"top": 100, "left": 50, "width": 80, "height": 18},
            {"top": 120, "left": 50, "width": 80, "height": 18},
        ]
        rows = _group_elements_into_rows(elems)
        assert len(rows) == 2

    def test_tolerance_grouping(self):
        """Elements within tolerance merge into same row."""
        elems = [
            {"top": 100, "left": 50, "width": 80, "height": 18},
            {"top": 103, "left": 200, "width": 80, "height": 18},
        ]
        rows = _group_elements_into_rows(elems, tolerance=5)
        assert len(rows) == 1

    def test_empty(self):
        assert _group_elements_into_rows([]) == []


# ---------------------------------------------------------------------------
# _should_merge_rows / _merge_multiline_cells
# ---------------------------------------------------------------------------

class TestMultilineCellMerge:
    def test_strict_subset_merges(self):
        """Row with strictly fewer columns merges into previous."""
        prev = [{"col": 0, "text": "A"}, {"col": 1, "text": "B"}, {"col": 2, "text": "C"}]
        curr = [{"col": 1, "text": "continuation"}]
        assert _should_merge_rows(prev, curr) is True

    def test_same_columns_no_merge(self):
        """Row with same column set does NOT merge (it's a new data row)."""
        prev = [{"col": 0, "text": "1"}, {"col": 1, "text": "A"}, {"col": 2, "text": "100"}]
        curr = [{"col": 0, "text": "2"}, {"col": 1, "text": "B"}, {"col": 2, "text": "200"}]
        assert _should_merge_rows(prev, curr) is False

    def test_superset_no_merge(self):
        """Row with additional columns does NOT merge."""
        prev = [{"col": 0, "text": "A"}]
        curr = [{"col": 0, "text": "B"}, {"col": 1, "text": "C"}]
        assert _should_merge_rows(prev, curr) is False

    def test_merge_appends_text(self):
        """Merged text should be appended with space."""
        rows = [
            [{"col": 0, "text": "Hello", "bold": True, "font_size_pts": 9.3},
             {"col": 1, "text": "World", "bold": False, "font_size_pts": 9.3}],
            [{"col": 1, "text": "continued", "bold": False, "font_size_pts": 9.3}],
        ]
        merged = _merge_multiline_cells(rows)
        assert len(merged) == 1
        # Find col 1 cell
        col1_cells = [c for c in merged[0] if c["col"] == 1]
        assert len(col1_cells) == 1
        assert col1_cells[0]["text"] == "World continued"

    def test_no_merge_for_single_row(self):
        rows = [[{"col": 0, "text": "A", "bold": False, "font_size_pts": 9.3}]]
        merged = _merge_multiline_cells(rows)
        assert len(merged) == 1

    def test_empty_rows(self):
        assert _merge_multiline_cells([]) == []


# ---------------------------------------------------------------------------
# detect_tables
# ---------------------------------------------------------------------------

class TestDetectTables:
    """Unit tests with synthetic data."""

    FONTSPECS = {"0": {"size_pts": 9.3, "family": "Sans", "color": "#000"}}

    def _make_elem(self, top, left, text="x", width=80, bold=False):
        return {
            "top": top, "left": left, "width": width, "height": 18,
            "font_id": "0", "text": text, "bold": bold, "italic": False,
        }

    def test_simple_3x3_table(self):
        """3 rows x 3 columns should produce one table."""
        elems = [
            # Row 1
            self._make_elem(100, 50, "A"), self._make_elem(100, 200, "B"),
            self._make_elem(100, 400, "C"),
            # Row 2
            self._make_elem(120, 50, "1"), self._make_elem(120, 200, "2"),
            self._make_elem(120, 400, "3"),
            # Row 3
            self._make_elem(140, 50, "X"), self._make_elem(140, 200, "Y"),
            self._make_elem(140, 400, "Z"),
        ]
        page = {"text_elements": elems}
        tables = detect_tables(page, self.FONTSPECS)
        assert len(tables) == 1
        t = tables[0]
        assert t["col_count"] == 3
        assert t["row_count"] == 3
        assert t["type"] == "table"

    def test_no_table_with_two_elements_per_row(self):
        """Only 2 elements per row — not a table."""
        elems = [
            self._make_elem(100, 50, "A"), self._make_elem(100, 200, "B"),
            self._make_elem(120, 50, "C"), self._make_elem(120, 200, "D"),
        ]
        page = {"text_elements": elems}
        tables = detect_tables(page, self.FONTSPECS)
        assert len(tables) == 0

    def test_empty_page(self):
        page = {"text_elements": []}
        assert detect_tables(page, self.FONTSPECS) == []

    def test_large_gap_splits_tables(self):
        """A large vertical gap between rows creates two separate table regions.
        Single-row regions are filtered out."""
        elems = [
            # Table 1 (rows at 100, 120, 140)
            self._make_elem(100, 50, "A"), self._make_elem(100, 200, "B"),
            self._make_elem(100, 400, "C"),
            self._make_elem(120, 50, "1"), self._make_elem(120, 200, "2"),
            self._make_elem(120, 400, "3"),
            self._make_elem(140, 50, "X"), self._make_elem(140, 200, "Y"),
            self._make_elem(140, 400, "Z"),
            # Gap of 500px
            # Table 2 (rows at 700, 720, 740)
            self._make_elem(700, 50, "D"), self._make_elem(700, 200, "E"),
            self._make_elem(700, 400, "F"),
            self._make_elem(720, 50, "4"), self._make_elem(720, 200, "5"),
            self._make_elem(720, 400, "6"),
            self._make_elem(740, 50, "P"), self._make_elem(740, 200, "Q"),
            self._make_elem(740, 400, "R"),
        ]
        page = {"text_elements": elems}
        tables = detect_tables(page, self.FONTSPECS)
        assert len(tables) == 2

    def test_table_has_bbox(self):
        """Table must have bbox_px and bbox_pts."""
        elems = [
            self._make_elem(100, 50, "A"), self._make_elem(100, 200, "B"),
            self._make_elem(100, 400, "C"),
            self._make_elem(120, 50, "1"), self._make_elem(120, 200, "2"),
            self._make_elem(120, 400, "3"),
        ]
        page = {"text_elements": elems}
        tables = detect_tables(page, self.FONTSPECS)
        assert len(tables) == 1
        t = tables[0]
        assert "bbox_px" in t
        assert "bbox_pts" in t
        assert t["bbox_px"]["top"] == 100
        assert t["bbox_px"]["left"] == 50

    def test_table_has_col_boundaries(self):
        elems = [
            self._make_elem(100, 50, "A"), self._make_elem(100, 200, "B"),
            self._make_elem(100, 400, "C"),
            self._make_elem(120, 50, "1"), self._make_elem(120, 200, "2"),
            self._make_elem(120, 400, "3"),
        ]
        page = {"text_elements": elems}
        tables = detect_tables(page, self.FONTSPECS)
        t = tables[0]
        assert "col_boundaries_px" in t
        assert len(t["col_boundaries_px"]) == t["col_count"] + 1  # left edges + right edge

    def test_single_row_table_filtered(self):
        """A single-row table region should be filtered out."""
        elems = [
            self._make_elem(100, 50, "A"), self._make_elem(100, 200, "B"),
            self._make_elem(100, 400, "C"),
        ]
        page = {"text_elements": elems}
        tables = detect_tables(page, self.FONTSPECS)
        assert len(tables) == 0


# ---------------------------------------------------------------------------
# enrich_images
# ---------------------------------------------------------------------------

class TestEnrichImages:
    def test_adds_bbox_pts(self):
        page = {
            "images": [
                {"top": 53, "left": 108, "width": 203, "height": 61, "src": "img.png"},
            ],
        }
        enriched = enrich_images(page)
        assert len(enriched) == 1
        img = enriched[0]
        assert "bbox_pts" in img
        assert "bbox_px" in img
        assert img["bbox_pts"]["top"] == _px_to_pts(53)
        assert img["bbox_pts"]["left"] == _px_to_pts(108)
        assert img["bbox_pts"]["width"] == _px_to_pts(203)
        assert img["bbox_pts"]["height"] == _px_to_pts(61)
        # Original fields preserved
        assert img["src"] == "img.png"
        assert img["top"] == 53

    def test_empty_images(self):
        assert enrich_images({"images": []}) == []

    def test_no_images_key(self):
        assert enrich_images({}) == []

    def test_does_not_mutate(self):
        original_img = {"top": 53, "left": 108, "width": 203, "height": 61, "src": "a.png"}
        page = {"images": [original_img]}
        enrich_images(page)
        assert "bbox_pts" not in original_img  # original not mutated


# ---------------------------------------------------------------------------
# Integration: detect_tables against real PDFs
# ---------------------------------------------------------------------------

class TestDetectTablesIntegration:
    @pytest.mark.skipif(not EN_PDF.exists(), reason="contract_en3.pdf not found")
    def test_en3_page3_payment_table(self):
        """EN3 page 3 should have exactly one table (payment schedule + Total row)."""
        from parse_pdf_structure import parse_pdftohtml_xml
        data = parse_pdftohtml_xml(str(EN_PDF))
        grouped = group_all_pages(data)
        classified = classify_all_pages(grouped)

        fontspec_by_id = {}
        for page in classified["pages"]:
            fontspec_by_id.update(_build_fontspec_lookup(page.get("fontspecs", [])))

        p3 = classified["pages"][2]
        tables = detect_tables(p3, fontspec_by_id)
        assert len(tables) == 1
        t = tables[0]
        assert t["col_count"] >= 5  # 5-6 cols depending on Total row alignment
        assert t["row_count"] >= 6  # header + 4 data rows + Total row + multi-line
        # Total row must be part of the table
        all_texts = [c["text"] for row in t["rows"] for c in row]
        assert any("Total" in txt for txt in all_texts), "Total row must be in table"
        assert any("12,000,000" in txt for txt in all_texts), "Total amount must be in table"

    @pytest.mark.skipif(not EN_PDF.exists(), reason="contract_en3.pdf not found")
    def test_en3_page3_table_contains_phase_data(self):
        """Table rows should contain phase numbers and amounts."""
        from parse_pdf_structure import parse_pdftohtml_xml
        data = parse_pdftohtml_xml(str(EN_PDF))
        grouped = group_all_pages(data)
        classified = classify_all_pages(grouped)

        fontspec_by_id = {}
        for page in classified["pages"]:
            fontspec_by_id.update(_build_fontspec_lookup(page.get("fontspecs", [])))

        p3 = classified["pages"][2]
        tables = detect_tables(p3, fontspec_by_id)
        t = tables[0]

        # Flatten all cell texts
        all_texts = [c["text"] for row in t["rows"] for c in row]
        combined = " ".join(all_texts)
        assert "2,400,000" in combined
        assert "3,600,000" in combined
        assert "Contract Signing" in combined

    @pytest.mark.skipif(not EN_PDF.exists(), reason="contract_en3.pdf not found")
    def test_en3_other_pages_table_presence(self):
        """Pages 2, 4, 5 should have no tables; page 1 has the party info table."""
        from parse_pdf_structure import parse_pdftohtml_xml
        data = parse_pdftohtml_xml(str(EN_PDF))
        grouped = group_all_pages(data)
        classified = classify_all_pages(grouped)

        fontspec_by_id = {}
        for page in classified["pages"]:
            fontspec_by_id.update(_build_fontspec_lookup(page.get("fontspecs", [])))

        # Page 1 (index 0): party info table is correctly detected
        tables_p1 = detect_tables(classified["pages"][0], fontspec_by_id)
        assert len(tables_p1) >= 1, "Page 1 should have the party info table"

        # Pages 2, 4, 5 should have no tables
        for i in [1, 3, 4]:
            tables = detect_tables(classified["pages"][i], fontspec_by_id)
            assert len(tables) == 0, f"Page {i+1} should have no tables"

    @pytest.mark.skipif(not EN_PDF.exists(), reason="contract_en3.pdf not found")
    def test_en3_table_cells_not_headings(self):
        """Table cell text must NOT be classified as headings after T4 fix."""
        from parse_pdf_structure import parse_pdftohtml_xml
        data = parse_pdftohtml_xml(str(EN_PDF))
        grouped = group_all_pages(data)
        classified = classify_all_pages(grouped)
        p3 = classified["pages"][2]

        # Collect table cell tops from paragraph grouping
        table_cell_paras = [p for p in p3["paragraphs"] if p.get("is_table_cell")]
        heading_cells = [p for p in table_cell_paras if p["type"] == "heading"]
        assert len(heading_cells) == 0, \
            f"Table cells classified as heading: {[p['text'][:20] for p in heading_cells]}"

    @pytest.mark.skipif(not ZH_PDF.exists(), reason="contract_zh3.pdf not found")
    def test_zh3_page1_has_party_table(self):
        """ZH3 page 1 should detect the party info table."""
        from parse_pdf_structure import parse_pdftohtml_xml
        data = parse_pdftohtml_xml(str(ZH_PDF))
        grouped = group_all_pages(data)
        classified = classify_all_pages(grouped)

        fontspec_by_id = {}
        for page in classified["pages"]:
            fontspec_by_id.update(_build_fontspec_lookup(page.get("fontspecs", [])))

        p1 = classified["pages"][0]
        tables = detect_tables(p1, fontspec_by_id)
        assert len(tables) >= 1
        # At least one table should contain party info
        all_texts = " ".join(
            c["text"] for t in tables for row in t["rows"] for c in row
        )
        assert "公司名稱" in all_texts or "甲方" in all_texts


# ---------------------------------------------------------------------------
# Integration: parse_digital_pdf (full pipeline)
# ---------------------------------------------------------------------------

class TestParseDigitalPdfIntegration:
    @pytest.mark.skipif(not EN_PDF.exists(), reason="contract_en3.pdf not found")
    def test_en3_full_pipeline(self):
        """Full pipeline produces complete structure."""
        result = parse_digital_pdf(str(EN_PDF))
        assert len(result["pages"]) == 5
        for page in result["pages"]:
            assert "text_elements" in page
            assert "paragraphs" in page
            assert "tables" in page
            assert "images" in page
            assert "fontspecs" in page

    @pytest.mark.skipif(not EN_PDF.exists(), reason="contract_en3.pdf not found")
    def test_en3_images_enriched(self):
        """All images should have bbox_pts after pipeline."""
        result = parse_digital_pdf(str(EN_PDF))
        for page in result["pages"]:
            for img in page["images"]:
                assert "bbox_pts" in img, f"Missing bbox_pts: {img.get('src')}"
                assert "bbox_px" in img
                assert isinstance(img["bbox_pts"]["top"], float)

    @pytest.mark.skipif(not EN_PDF.exists(), reason="contract_en3.pdf not found")
    def test_en3_page3_has_table(self):
        """Pipeline output: page 3 should have exactly 1 table with Total row."""
        result = parse_digital_pdf(str(EN_PDF))
        p3 = result["pages"][2]
        assert len(p3["tables"]) == 1
        assert p3["tables"][0]["col_count"] >= 5  # 5-6 cols depending on Total row alignment

    @pytest.mark.skipif(not ZH_PDF.exists(), reason="contract_zh3.pdf not found")
    def test_zh3_full_pipeline(self):
        """ZH3 full pipeline works without errors."""
        result = parse_digital_pdf(str(ZH_PDF))
        assert len(result["pages"]) == 4
        for page in result["pages"]:
            assert "tables" in page
            assert "paragraphs" in page
            for img in page["images"]:
                assert "bbox_pts" in img

    @pytest.mark.skipif(not EN_PDF.exists(), reason="contract_en3.pdf not found")
    def test_pipeline_does_not_lose_text(self):
        """Pipeline must not alter the text_elements count."""
        from parse_pdf_structure import parse_pdftohtml_xml
        raw = parse_pdftohtml_xml(str(EN_PDF))
        result = parse_digital_pdf(str(EN_PDF))
        for raw_page, pipe_page in zip(raw["pages"], result["pages"]):
            assert len(raw_page["text_elements"]) == len(pipe_page["text_elements"])

    @pytest.mark.skipif(not EN_PDF.exists(), reason="contract_en3.pdf not found")
    def test_pipeline_paragraphs_classified(self):
        """Every paragraph in pipeline output has classification fields."""
        result = parse_digital_pdf(str(EN_PDF))
        for page in result["pages"]:
            for p in page["paragraphs"]:
                assert "font_size_pts" in p
                assert "type" in p
                assert p["type"] in ("heading", "paragraph", "footer", "header")


# ---------------------------------------------------------------------------
# Regression: Bug #1 — "Total" row not falsely promoted to heading
# ---------------------------------------------------------------------------

class TestBug1TotalRowNotHeading:
    """EN3 page 3: 'Total' and '12,000,000' must be in the table, NOT headings.

    Root cause: Pass 2 expansion in _find_table_row_tops used page-global
    max element height as line_height, which was too small to bridge the
    31px gap from the last confirmed row to the Total row.
    Fix: use max(inter-row gap, element height) * 1.2 as expansion threshold.
    """

    @pytest.mark.skipif(not EN_PDF.exists(), reason="contract_en3.pdf not found")
    def test_total_not_heading(self):
        """'Total' must NOT appear as a heading on EN3 page 3."""
        result = parse_digital_pdf(str(EN_PDF))
        p3 = result["pages"][2]
        heading_texts = [p["text"] for p in p3["paragraphs"] if p["type"] == "heading"]
        assert not any("Total" in t for t in heading_texts), (
            "'Total' should not be a heading — it belongs in the payment table"
        )

    @pytest.mark.skipif(not EN_PDF.exists(), reason="contract_en3.pdf not found")
    def test_12m_not_heading(self):
        """'12,000,000' must NOT appear as a heading on EN3 page 3."""
        result = parse_digital_pdf(str(EN_PDF))
        p3 = result["pages"][2]
        heading_texts = [p["text"] for p in p3["paragraphs"] if p["type"] == "heading"]
        assert not any("12,000,000" in t for t in heading_texts), (
            "'12,000,000' should not be a heading — it belongs in the payment table"
        )

    @pytest.mark.skipif(not EN_PDF.exists(), reason="contract_en3.pdf not found")
    def test_total_in_table(self):
        """'Total' and '12,000,000' must be present in the payment table."""
        result = parse_digital_pdf(str(EN_PDF))
        p3 = result["pages"][2]
        assert len(p3["tables"]) >= 1
        t = p3["tables"][0]
        all_texts = [c["text"] for row in t["rows"] for c in row]
        combined = " ".join(all_texts)
        assert "Total" in combined
        assert "12,000,000" in combined


# ---------------------------------------------------------------------------
# Regression: Bug #2 — ZH3 page 1 no false-positive table from numbered list
# ---------------------------------------------------------------------------

class TestBug2ZH3NoFalseTable:
    """ZH3 page 1: numbered list items with mixed CJK/Latin must NOT form a table.

    Root cause: 6 fragments at same top triggered table detection;
    after _merge_multiline_cells they collapsed to 1 row.
    Fix: post-merge row count check filters out single-row tables.
    """

    @pytest.mark.skipif(not ZH_PDF.exists(), reason="contract_zh3.pdf not found")
    def test_numbered_list_not_in_table(self):
        """Numbered list items ('1. LLM', '2. 智慧行銷') must NOT be in any table."""
        result = parse_digital_pdf(str(ZH_PDF))
        p1 = result["pages"][0]
        all_table_texts = []
        for t in p1["tables"]:
            for row in t["rows"]:
                for c in row:
                    all_table_texts.append(c["text"])
        combined = " ".join(all_table_texts)
        assert "LLM" not in combined, "Numbered list item 'LLM' should not be in a table"
        assert "RAG" not in combined, "Numbered list item 'RAG' should not be in a table"

    @pytest.mark.skipif(not ZH_PDF.exists(), reason="contract_zh3.pdf not found")
    def test_post_merge_single_row_filter(self):
        """Synthetic: a region where all rows merge into one should return None."""
        # 6 elements at same top → Pass 1 detects as table row
        # After merge, collapses to 1 row → post-merge guard filters it out
        from parse_pdf_structure import _merge_multiline_cells
        rows = [
            [{"text": "a", "bold": False, "font_size_pts": 12, "col": 0},
             {"text": "b", "bold": False, "font_size_pts": 12, "col": 1},
             {"text": "c", "bold": False, "font_size_pts": 12, "col": 2}],
            [{"text": "d", "bold": False, "font_size_pts": 12, "col": 0},
             {"text": "e", "bold": False, "font_size_pts": 12, "col": 1}],
        ]
        merged = _merge_multiline_cells(rows)
        # Row 2 cols {0,1} is a strict subset of row 1 cols {0,1,2} → merges
        assert len(merged) == 1, "Continuation row should merge into previous"


# ---------------------------------------------------------------------------
# Regression: Bug #3 — CJK character fragments → false table rows
# ---------------------------------------------------------------------------

class TestBug3CJKFragmentTableFalsePositive:
    """ZH3 page 3: CJK paragraph text split into 7-22 narrow fragments per line
    must NOT be detected as table rows.

    Root cause: pdftohtml splits CJK text with mixed fonts into individual
    character/word fragments (width ~17px each). With 14-22 fragments at
    the same top, Pass 1 (threshold >= 3) falsely confirmed them as table rows.
    Additionally, the global max_row_gap inflated to 278+ px, causing Pass 2
    expansion to absorb everything between the real table and the CJK fragments.

    Fix: Three-pronged:
      1. TABLE_ROW_MAX_ELEMENTS ceiling (8) rejects rows with too many elements
      2. TABLE_ROW_MEDIAN_WIDTH_MIN_PX (30) rejects rows of narrow fragments
      3. TABLE_ROW_COVERAGE_MAX (0.85) rejects contiguous paragraph text
      4. Local max_row_gap per contiguous region prevents distant false-positive
         rows from inflating expansion threshold
    """

    @pytest.mark.skipif(not ZH_PDF.exists(), reason="contract_zh3.pdf not found")
    def test_zh3_article_headings_classified_as_h2(self):
        """All 7 article headings (第一條 through 第七條) must be classified as H2."""
        result = parse_digital_pdf(str(ZH_PDF))
        all_headings = []
        for page in result["pages"]:
            for p in page["paragraphs"]:
                if p["type"] == "heading":
                    all_headings.append(p)

        article_names = "一二三四五六七"
        for i, name in enumerate(article_names):
            keyword = f"第{name}條"
            matches = [h for h in all_headings if keyword in h["text"]]
            assert len(matches) >= 1, f"'{keyword}' not found among headings"
            for m in matches:
                assert m["heading_level"] == 2, (
                    f"'{keyword}' should be H2 but got H{m['heading_level']}: {m['text'][:40]}"
                )

    @pytest.mark.skipif(not ZH_PDF.exists(), reason="contract_zh3.pdf not found")
    def test_zh3_page3_no_false_positive_tables_from_paragraphs(self):
        """ZH3 page 3 paragraph text must NOT form false-positive tables.

        CJK fragment lines (Articles 5-7 body text) at top=633-893
        must NOT be in any table.
        """
        result = parse_digital_pdf(str(ZH_PDF))
        p3 = result["pages"][2]
        # Collect all text in tables on page 3
        all_table_texts = []
        for t in p3["tables"]:
            for row in t["rows"]:
                for c in row:
                    all_table_texts.append(c["text"])
        combined = " ".join(all_table_texts)

        # These are paragraph body text terms that must NOT be in any table
        paragraph_terms = ["保密義務", "違約", "準據法", "智慧財產權"]
        for term in paragraph_terms:
            # Allow the term in a table header if it's a heading that got absorbed,
            # but paragraph BODY text should not be there
            # We check the table body text (all rows after first)
            body_table_texts = []
            for t in p3["tables"]:
                for row in t["rows"][1:]:  # skip header row
                    for c in row:
                        body_table_texts.append(c["text"])
            body_combined = " ".join(body_table_texts)
            assert term not in body_combined, (
                f"Paragraph term '{term}' found in table body — CJK fragment false positive"
            )

    @pytest.mark.skipif(not ZH_PDF.exists(), reason="contract_zh3.pdf not found")
    def test_zh3_page3_payment_table_still_detected(self):
        """ZH3 page 3 payment schedule table must still be correctly detected."""
        result = parse_digital_pdf(str(ZH_PDF))
        p3 = result["pages"][2]
        assert len(p3["tables"]) >= 1, "Payment table must be detected on page 3"
        t = p3["tables"][0]
        all_texts = [c["text"] for row in t["rows"] for c in row]
        combined = " ".join(all_texts)
        assert "期次" in combined or "付款條件" in combined, (
            "Payment table must contain header text like '期次' or '付款條件'"
        )
        assert "2,400,000" in combined, "Payment table must contain amount 2,400,000"

    @pytest.mark.skipif(not ZH_PDF.exists(), reason="contract_zh3.pdf not found")
    def test_zh3_page1_heading_not_absorbed_into_table(self):
        """ZH3 page 1: '第二條 合約標的與範圍' must NOT be absorbed into party table."""
        result = parse_digital_pdf(str(ZH_PDF))
        p1 = result["pages"][0]

        # Check that 第二條 is a heading
        headings = [p for p in p1["paragraphs"]
                    if p["type"] == "heading" and "第二條" in p["text"]]
        assert len(headings) == 1, "'第二條' must appear as exactly one heading"
        assert headings[0]["heading_level"] == 2

        # Check that 第二條 is NOT in any table
        for t in p1["tables"]:
            all_texts = [c["text"] for row in t["rows"] for c in row]
            combined = " ".join(all_texts)
            assert "第二條" not in combined, (
                "'第二條' heading must not be absorbed into party table"
            )


class TestCJKFragmentFilters:
    """Unit tests for the CJK fragment detection heuristics."""

    def test_median_width_rejects_narrow_fragments(self):
        """Narrow CJK fragments (17px each) should fail median width check."""
        # 10 narrow CJK character fragments at same top
        elems = [
            {"top": 100, "left": 108 + i * 17, "width": 17, "height": 22}
            for i in range(10)
        ]
        med = _median_width(elems, 100, 100)
        assert med == 17.0
        assert med < TABLE_ROW_MEDIAN_WIDTH_MIN_PX

    def test_median_width_accepts_table_cells(self):
        """Real table cells (50-80px) should pass median width check."""
        elems = [
            {"top": 100, "left": 100, "width": 50, "height": 18},
            {"top": 100, "left": 300, "width": 80, "height": 18},
            {"top": 100, "left": 500, "width": 60, "height": 18},
        ]
        med = _median_width(elems, 100, 100)
        assert med >= TABLE_ROW_MEDIAN_WIDTH_MIN_PX

    def test_coverage_ratio_rejects_contiguous_text(self):
        """Contiguous paragraph text (coverage ~1.0) should fail coverage check."""
        # Elements tile perfectly: left + width = next left
        elems = [
            {"top": 100, "left": 108, "width": 297, "height": 22},
            {"top": 100, "left": 405, "width": 121, "height": 18},
            {"top": 100, "left": 526, "width": 280, "height": 22},
        ]
        coverage = _coverage_ratio(elems, 100, 100)
        assert coverage > TABLE_ROW_COVERAGE_MAX, (
            f"Contiguous text coverage {coverage} should exceed threshold {TABLE_ROW_COVERAGE_MAX}"
        )

    def test_coverage_ratio_accepts_table_row(self):
        """Table row with gaps between cells should pass coverage check."""
        elems = [
            {"top": 100, "left": 122, "width": 33, "height": 22},
            {"top": 100, "left": 218, "width": 66, "height": 22},
            {"top": 100, "left": 396, "width": 66, "height": 22},
            {"top": 100, "left": 551, "width": 83, "height": 22},
            {"top": 100, "left": 702, "width": 66, "height": 22},
        ]
        coverage = _coverage_ratio(elems, 100, 100)
        assert coverage < TABLE_ROW_COVERAGE_MAX, (
            f"Table row coverage {coverage} should be below threshold {TABLE_ROW_COVERAGE_MAX}"
        )

    def test_element_ceiling_rejects_many_fragments(self):
        """More than TABLE_ROW_MAX_ELEMENTS fragments should not form a table row."""
        # 15 CJK fragments — exceeds ceiling of 8
        elems = [
            {"top": 100, "left": 108 + i * 40, "width": 33, "height": 22}
            for i in range(15)
        ]
        tops = _find_table_row_tops(elems)
        assert 100 not in tops, "15 elements should exceed TABLE_ROW_MAX_ELEMENTS"

    def test_cjk_fragments_not_table_row(self):
        """Synthetic CJK fragment scenario: 20 narrow elements should NOT be a table row."""
        elems = [
            {"top": 100, "left": 108 + i * 17, "width": 17, "height": 22}
            for i in range(20)
        ]
        tops = _find_table_row_tops(elems)
        assert 100 not in tops, "20 narrow CJK fragments must not be a table row"

    def test_constants_consistent(self):
        """Verify CJK filter constants are reasonable."""
        assert TABLE_ROW_MAX_ELEMENTS == 8
        assert TABLE_ROW_MEDIAN_WIDTH_MIN_PX == 30
        assert TABLE_ROW_COVERAGE_MAX == 0.85
        assert TABLE_ROW_MAX_ELEMENTS > TABLE_ROW_MIN_ELEMENTS


# ---------------------------------------------------------------------------
# T5: Heading content filter (HEADING_MIN_CONTENT_CHARS)
# ---------------------------------------------------------------------------

class TestHeadingContentFilter:
    """Unit tests for the heading minimum content length filter.

    Prevents trivially short fragments (date parts, form placeholders)
    from being classified as headings even when at heading font sizes.
    """

    FONTSPECS = {
        "h3": {"size_pts": 12.0, "family": "Arial", "color": "#000000"},
        "h1": {"size_pts": 22.0, "family": "Arial", "color": "#000000"},
        "body": {"size_pts": 11.0, "family": "Times", "color": "#000000"},
    }
    PAGE_HEIGHT_PTS = 792.0
    BODY_SIZE = 11.0

    def _make_para(self, font_id="body", bold=False, text="Some text"):
        return {
            "font_id": font_id, "bold": bold, "italic": False,
            "top": 300, "left": 108, "width": 400, "height": 18,
            "text": text, "lines": [], "type": "paragraph", "alignment": "left",
        }

    def test_single_cjk_char_not_heading(self):
        """Single CJK character '月' at H3 font size = NOT heading."""
        para = self._make_para(font_id="h3", text="月")
        result = _classify_single_paragraph(
            para, self.FONTSPECS, self.PAGE_HEIGHT_PTS, self.BODY_SIZE)
        assert result["type"] == "paragraph"
        assert "heading_level" not in result

    def test_single_char_day_not_heading(self):
        """'日' at H3 font size = NOT heading."""
        para = self._make_para(font_id="h3", text="日")
        result = _classify_single_paragraph(
            para, self.FONTSPECS, self.PAGE_HEIGHT_PTS, self.BODY_SIZE)
        assert result["type"] == "paragraph"

    def test_underscores_only_not_heading(self):
        """'___' at H3 font size = NOT heading."""
        para = self._make_para(font_id="h3", text="___")
        result = _classify_single_paragraph(
            para, self.FONTSPECS, self.PAGE_HEIGHT_PTS, self.BODY_SIZE)
        assert result["type"] == "paragraph"

    def test_space_underscores_not_heading(self):
        """' ___' at H3 font size = NOT heading."""
        para = self._make_para(font_id="h3", text=" ___")
        result = _classify_single_paragraph(
            para, self.FONTSPECS, self.PAGE_HEIGHT_PTS, self.BODY_SIZE)
        assert result["type"] == "paragraph"

    def test_legitimate_short_heading_preserved(self):
        """'簽署欄' (3 chars, no underscores) at H1 size = still heading."""
        para = self._make_para(font_id="h1", bold=True, text="簽署欄")
        result = _classify_single_paragraph(
            para, self.FONTSPECS, self.PAGE_HEIGHT_PTS, self.BODY_SIZE)
        assert result["type"] == "heading"
        assert result["heading_level"] == 1

    def test_two_char_heading_preserved(self):
        """2-char text at heading size is allowed (just meets threshold)."""
        para = self._make_para(font_id="h3", text="OK")
        result = _classify_single_paragraph(
            para, self.FONTSPECS, self.PAGE_HEIGHT_PTS, self.BODY_SIZE)
        assert result["type"] == "heading"
        assert result["heading_level"] == 3

    def test_long_date_heading_preserved(self):
        """'中華民國一一四年' (8 chars) at H3 size = still heading."""
        para = self._make_para(font_id="h3", text="中華民國一一四年")
        result = _classify_single_paragraph(
            para, self.FONTSPECS, self.PAGE_HEIGHT_PTS, self.BODY_SIZE)
        assert result["type"] == "heading"
        assert result["heading_level"] == 3

    def test_constant_value(self):
        assert HEADING_MIN_CONTENT_CHARS == 2


# ---------------------------------------------------------------------------
# T5: Table-proximity heading suppression
# ---------------------------------------------------------------------------

class TestSuppressHeadingsNearTables:
    """Unit tests for _suppress_headings_near_tables.

    Verifies that H4 headings near detected tables are demoted to paragraphs,
    while H1-H3 headings are untouched.
    """

    def _make_table(self, top, left, width, height):
        return {
            "type": "table",
            "bbox_px": {"top": top, "left": left, "width": width, "height": height},
            "rows": [], "col_count": 3, "row_count": 2,
        }

    def _make_heading(self, text, top, left=108, level=4):
        return {
            "type": "heading",
            "heading_level": level,
            "text": text,
            "top": top,
            "left": left,
            "width": 100,
            "height": 18,
            "bold": True,
            "font_size_pts": 11.3,
        }

    def _make_para(self, text, top):
        return {
            "type": "paragraph",
            "text": text,
            "top": top,
            "left": 108,
            "width": 400,
            "height": 18,
            "bold": False,
            "font_size_pts": 11.0,
        }

    def test_h4_near_table_demoted(self):
        """H4 heading within table proximity zone is demoted."""
        page = {
            "tables": [self._make_table(500, 100, 600, 100)],  # top=500, bottom=600
            "paragraphs": [
                self._make_heading("Label A", 480, level=4),  # within 50px above
            ],
        }
        result = _suppress_headings_near_tables(page)
        assert result[0]["type"] == "paragraph"
        assert "heading_level" not in result[0]

    def test_h4_inside_table_demoted(self):
        """H4 heading inside table bbox is demoted."""
        page = {
            "tables": [self._make_table(500, 100, 600, 100)],
            "paragraphs": [
                self._make_heading("Label B", 550, level=4),  # inside bbox
            ],
        }
        result = _suppress_headings_near_tables(page)
        assert result[0]["type"] == "paragraph"

    def test_h4_below_table_demoted(self):
        """H4 heading within proximity zone below table is demoted."""
        page = {
            "tables": [self._make_table(500, 100, 600, 100)],
            "paragraphs": [
                self._make_heading("Label C", 630, level=4),  # within 50px below bottom
            ],
        }
        result = _suppress_headings_near_tables(page)
        assert result[0]["type"] == "paragraph"

    def test_h4_far_from_table_preserved(self):
        """H4 heading far from any table is NOT demoted."""
        page = {
            "tables": [self._make_table(500, 100, 600, 100)],
            "paragraphs": [
                self._make_heading("Real Heading", 200, level=4),  # 300px above
            ],
        }
        result = _suppress_headings_near_tables(page)
        assert result[0]["type"] == "heading"
        assert result[0]["heading_level"] == 4

    def test_h2_near_table_preserved(self):
        """H2 heading near table is NOT demoted (only H4 is demoted)."""
        page = {
            "tables": [self._make_table(500, 100, 600, 100)],
            "paragraphs": [
                self._make_heading("Article 3", 480, level=2),  # near table, but H2
            ],
        }
        result = _suppress_headings_near_tables(page)
        assert result[0]["type"] == "heading"
        assert result[0]["heading_level"] == 2

    def test_h1_near_table_preserved(self):
        """H1 heading near table is NOT demoted."""
        page = {
            "tables": [self._make_table(500, 100, 600, 100)],
            "paragraphs": [
                self._make_heading("Title", 520, level=1),
            ],
        }
        result = _suppress_headings_near_tables(page)
        assert result[0]["type"] == "heading"
        assert result[0]["heading_level"] == 1

    def test_no_tables_passthrough(self):
        """Without tables, all paragraphs pass through unchanged."""
        page = {
            "tables": [],
            "paragraphs": [
                self._make_heading("Heading", 100, level=4),
                self._make_para("Body", 200),
            ],
        }
        result = _suppress_headings_near_tables(page)
        assert len(result) == 2
        assert result[0]["type"] == "heading"

    def test_regular_paragraphs_untouched(self):
        """Non-heading paragraphs near tables are unchanged."""
        page = {
            "tables": [self._make_table(500, 100, 600, 100)],
            "paragraphs": [
                self._make_para("Body text", 520),
            ],
        }
        result = _suppress_headings_near_tables(page)
        assert result[0]["type"] == "paragraph"

    def test_proximity_constant(self):
        assert TABLE_HEADING_PROXIMITY_PX == 50


# ---------------------------------------------------------------------------
# T5: Integration — ZH3 false H3 headings fixed
# ---------------------------------------------------------------------------

class TestBug4ZH3FalseH3Headings:
    """ZH3 page 4: '月', '日', '___' at 12pt were falsely classified as H3.

    Root cause: single-char date fragments and underscore placeholders at
    heading font sizes. Fix: HEADING_MIN_CONTENT_CHARS filter.
    """

    @pytest.mark.skipif(not ZH_PDF.exists(), reason="contract_zh3.pdf not found")
    def test_zh3_page4_no_false_h3_month_day(self):
        """'月' and '日' must NOT be headings on ZH3 page 4."""
        result = parse_digital_pdf(str(ZH_PDF))
        p4 = result["pages"][3]
        headings = [p for p in p4["paragraphs"] if p["type"] == "heading"]
        heading_texts = [h["text"].strip() for h in headings]
        assert "月" not in heading_texts, "'月' should not be a heading"
        assert "日" not in heading_texts, "'日' should not be a heading"

    @pytest.mark.skipif(not ZH_PDF.exists(), reason="contract_zh3.pdf not found")
    def test_zh3_page4_no_false_h3_underscores(self):
        """'___' must NOT be a heading on ZH3 page 4."""
        result = parse_digital_pdf(str(ZH_PDF))
        p4 = result["pages"][3]
        headings = [p for p in p4["paragraphs"] if p["type"] == "heading"]
        for h in headings:
            content = h["text"].strip().replace("_", "")
            assert len(content) >= HEADING_MIN_CONTENT_CHARS, (
                f"False heading from underscore placeholder: {h['text']!r}"
            )

    @pytest.mark.skipif(not ZH_PDF.exists(), reason="contract_zh3.pdf not found")
    def test_zh3_page4_legitimate_headings_preserved(self):
        """'簽署欄' (H2) and '中華民國一一四年' (H3) must remain headings."""
        result = parse_digital_pdf(str(ZH_PDF))
        p4 = result["pages"][3]
        headings = [p for p in p4["paragraphs"] if p["type"] == "heading"]
        heading_texts = " ".join(h["text"] for h in headings)
        assert "簽署欄" in heading_texts
        assert "中華民國" in heading_texts


# ---------------------------------------------------------------------------
# T5: Integration — EN3 bold table labels suppressed
# ---------------------------------------------------------------------------

class TestBug5EN3BoldTableLabels:
    """EN3 page 1: Bold table labels (Party A, Company Name) were H4 headings.

    Root cause: bold row labels appear with only 1-2 elements per row,
    failing table-row detection, then promoted to H4.
    Fix: table-proximity heading suppression demotes H4 near tables.
    """

    @pytest.mark.skipif(not EN_PDF.exists(), reason="contract_en3.pdf not found")
    def test_en3_party_labels_not_headings(self):
        """Bold table labels must NOT be headings after suppression."""
        result = parse_digital_pdf(str(EN_PDF))
        p1 = result["pages"][0]
        headings = [p for p in p1["paragraphs"] if p["type"] == "heading"]
        heading_texts = [h["text"] for h in headings]
        for label in ["Party A", "Party B", "Company Name", "Address"]:
            assert label not in heading_texts, (
                f"'{label}' should not be a heading — it's a table row label"
            )

    @pytest.mark.skipif(not EN_PDF.exists(), reason="contract_en3.pdf not found")
    def test_en3_real_headings_preserved(self):
        """Real headings (H1, H2) must be preserved after suppression."""
        result = parse_digital_pdf(str(EN_PDF))
        p1 = result["pages"][0]
        headings = [p for p in p1["paragraphs"] if p["type"] == "heading"]
        heading_texts = " ".join(h["text"] for h in headings)
        assert "Artificial Intelligence" in heading_texts
        assert "Article 1" in heading_texts
        assert "Article 2" in heading_texts
        assert "Project Workflow" in heading_texts

    @pytest.mark.skipif(not EN_PDF.exists(), reason="contract_en3.pdf not found")
    def test_en3_page1_heading_count(self):
        """EN3 page 1 should have exactly 4 headings after suppression."""
        result = parse_digital_pdf(str(EN_PDF))
        p1 = result["pages"][0]
        headings = [p for p in p1["paragraphs"] if p["type"] == "heading"]
        assert len(headings) == 4, (
            f"Expected 4 headings, got {len(headings)}: "
            f"{[h['text'][:30] for h in headings]}"
        )


# ---------------------------------------------------------------------------
# T5: format_summary
# ---------------------------------------------------------------------------

class TestFormatSummary:
    """Unit tests for the human-readable summary formatter."""

    def test_basic_summary(self):
        data = {
            "pages": [
                {
                    "number": 1,
                    "paragraphs": [
                        {"type": "heading", "heading_level": 1, "text": "Title"},
                        {"type": "paragraph", "text": "Body"},
                        {"type": "footer", "text": "Page 1"},
                    ],
                    "tables": [{"type": "table"}],
                    "images": [{"src": "img.png"}],
                },
            ],
        }
        summary = format_summary(data)
        assert "1 pages" in summary
        assert "1 headings" in summary
        assert "1 paragraphs" in summary
        assert "1 tables" in summary
        assert "1 images" in summary
        assert "H1: Title" in summary

    def test_empty_pdf(self):
        summary = format_summary({"pages": []})
        assert "0 pages" in summary
        assert "0 headings" in summary

    def test_multi_page(self):
        data = {
            "pages": [
                {
                    "number": 1,
                    "paragraphs": [
                        {"type": "heading", "heading_level": 2, "text": "Section 1"},
                    ],
                    "tables": [],
                    "images": [],
                },
                {
                    "number": 2,
                    "paragraphs": [
                        {"type": "paragraph", "text": "Body"},
                    ],
                    "tables": [],
                    "images": [{"src": "a.png"}, {"src": "b.png"}],
                },
            ],
        }
        summary = format_summary(data)
        assert "2 pages" in summary
        assert "Page 1:" in summary
        assert "Page 2:" in summary
        assert "2 images" in summary  # total

    @pytest.mark.skipif(not EN_PDF.exists(), reason="contract_en3.pdf not found")
    def test_en3_summary_structure(self):
        """Summary of EN3 must mention key structural elements."""
        data = parse_digital_pdf(str(EN_PDF))
        summary = format_summary(data)
        assert "5 pages" in summary
        assert "Article" in summary
        assert "tables" in summary


# ---------------------------------------------------------------------------
# T5: CLI --summary flag
# ---------------------------------------------------------------------------

class TestCLISummaryFlag:
    @pytest.mark.skipif(not EN_PDF.exists(), reason="contract_en3.pdf not found")
    def test_summary_flag_produces_text(self):
        result = subprocess.run(
            [sys.executable, str(SCRIPT_PATH), "--pdf", str(EN_PDF), "--summary"],
            capture_output=True, text=True, timeout=30,
        )
        assert result.returncode == 0
        assert "PDF Structure Summary" in result.stdout
        assert "5 pages" in result.stdout
        assert "headings" in result.stdout

    @pytest.mark.skipif(not ZH_PDF.exists(), reason="contract_zh3.pdf not found")
    def test_summary_flag_zh3(self):
        result = subprocess.run(
            [sys.executable, str(SCRIPT_PATH), "--pdf", str(ZH_PDF), "--summary"],
            capture_output=True, text=True, timeout=30,
        )
        assert result.returncode == 0
        assert "4 pages" in result.stdout

    def test_summary_without_pdf_ignored(self):
        """--summary without --pdf should still require a mode flag."""
        result = subprocess.run(
            [sys.executable, str(SCRIPT_PATH), "--summary"],
            capture_output=True, text=True, timeout=10,
        )
        assert result.returncode != 0


# ---------------------------------------------------------------------------
# Runner
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    sys.exit(pytest.main([__file__, "-v"]))
