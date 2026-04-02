#!/usr/bin/env python3
"""Tests for deterministic_merge.py — poppler integration and existing behavior.

Tests the new poppler-enhanced merge path AND verifies backward compatibility
(without --pdf-structure, behavior is identical to before).

Run: uv run --with lxml python -m pytest test_deterministic_merge.py -v
"""

import json
import sys
from pathlib import Path

import pytest
from lxml import etree

# Ensure scripts directory is on path
sys.path.insert(0, str(Path(__file__).parent))

from deterministic_merge import (
    _apply_poppler_styling,
    _get_element_text,
    _match_poppler_paragraph,
    _replace_element_text,
    build_poppler_page_lookup,
    collect_text_elements,
    combined_similarity,
    heading_level_from_poppler,
    load_poppler_data,
    match_and_replace_text,
    text_similarity,
)


# =============================================================================
# Helpers
# =============================================================================


def _make_page_xml(elements_xml):
    """Build a <page> element from inner XML string."""
    xml = (
        f'<page number="1" width-pts="595" height-pts="842" '
        f'margin-top-cm="1.27" margin-bottom-cm="1.27" '
        f'margin-left-cm="1.27" margin-right-cm="1.27" '
        f'font-latin="Arial" font-cjk="SimSun">'
        f'{elements_xml}</page>'
    )
    return etree.fromstring(xml.encode())


def _make_ocr_region(content, native_label="text", bbox=None, index=0, label="text"):
    """Build an OCR region dict."""
    return {
        "label": label,
        "native_label": native_label,
        "content": content,
        "bbox_2d": bbox or [0, 0, 500, 100],
        "index": index,
    }


def _make_poppler_para(text, font_size_pts=11.0, bold=False, italic=False,
                       color="", top=0, left=0, ptype="paragraph",
                       heading_level=None):
    """Build a poppler paragraph dict matching parse_digital_pdf() shape."""
    para = {
        "type": ptype,
        "text": text,
        "top": top,
        "left": left,
        "width": 450,
        "height": 20,
        "font_id": "1",
        "bold": bold,
        "italic": italic,
        "font_size_pts": font_size_pts,
        "font_family": "Arial",
        "color": color,
        "is_table_cell": False,
        "alignment": "left",
    }
    if heading_level is not None:
        para["heading_level"] = heading_level
    return para


# =============================================================================
# build_poppler_page_lookup
# =============================================================================


class TestBuildPopplerPageLookup:
    """Tests for build_poppler_page_lookup."""

    def test_none_input(self):
        assert build_poppler_page_lookup(None) == {}

    def test_empty_dict(self):
        assert build_poppler_page_lookup({}) == {}

    def test_empty_pages(self):
        assert build_poppler_page_lookup({"pages": []}) == {}

    def test_non_dict_input(self):
        assert build_poppler_page_lookup("not a dict") == {}
        assert build_poppler_page_lookup(42) == {}
        assert build_poppler_page_lookup([]) == {}

    def test_filters_table_cells(self):
        data = {
            "pages": [{
                "number": 1,
                "paragraphs": [
                    _make_poppler_para("Normal text", top=100),
                    {**_make_poppler_para("Cell text", top=200), "is_table_cell": True},
                ],
            }]
        }
        lookup = build_poppler_page_lookup(data)
        assert len(lookup[1]) == 1
        assert lookup[1][0]["text"] == "Normal text"

    def test_filters_empty_text(self):
        data = {
            "pages": [{
                "number": 1,
                "paragraphs": [
                    _make_poppler_para("Valid text", top=100),
                    _make_poppler_para("", top=200),
                    _make_poppler_para("   ", top=300),
                ],
            }]
        }
        lookup = build_poppler_page_lookup(data)
        assert len(lookup[1]) == 1

    def test_sorts_by_position(self):
        data = {
            "pages": [{
                "number": 1,
                "paragraphs": [
                    _make_poppler_para("Third", top=300, left=50),
                    _make_poppler_para("First", top=100, left=50),
                    _make_poppler_para("Second", top=200, left=50),
                ],
            }]
        }
        lookup = build_poppler_page_lookup(data)
        texts = [p["text"] for p in lookup[1]]
        assert texts == ["First", "Second", "Third"]

    def test_multiple_pages(self):
        data = {
            "pages": [
                {"number": 1, "paragraphs": [_make_poppler_para("Page one")]},
                {"number": 2, "paragraphs": [_make_poppler_para("Page two")]},
                {"number": 3, "paragraphs": []},  # empty page — excluded
            ]
        }
        lookup = build_poppler_page_lookup(data)
        assert set(lookup.keys()) == {1, 2}

    def test_page_with_only_table_cells(self):
        data = {
            "pages": [{
                "number": 1,
                "paragraphs": [
                    {**_make_poppler_para("Cell 1"), "is_table_cell": True},
                    {**_make_poppler_para("Cell 2"), "is_table_cell": True},
                ],
            }]
        }
        lookup = build_poppler_page_lookup(data)
        assert 1 not in lookup  # all filtered → page excluded


# =============================================================================
# _match_poppler_paragraph
# =============================================================================


class TestMatchPopplerParagraph:
    """Tests for _match_poppler_paragraph."""

    def test_exact_match(self):
        paras = [_make_poppler_para("Contract Agreement")]
        idx, score, para = _match_poppler_paragraph("Contract Agreement", paras, set())
        assert idx == 0
        assert score > 0.9
        assert para["text"] == "Contract Agreement"

    def test_close_match(self):
        paras = [
            _make_poppler_para("Introduction to the legal framework"),
            _make_poppler_para("Terms and Conditions"),
        ]
        idx, score, para = _match_poppler_paragraph(
            "Introduction to the legal framework", paras, set())
        assert idx == 0
        assert score > 0.5

    def test_no_match_below_threshold(self):
        paras = [_make_poppler_para("Completely unrelated different topic here")]
        idx, score, para = _match_poppler_paragraph("XYZ ABC DEF", paras, set())
        assert idx == -1
        assert para is None

    def test_consumed_indices_skipped(self):
        paras = [
            _make_poppler_para("Hello World"),
            _make_poppler_para("Hello World Again"),
        ]
        # Index 0 consumed — should match index 1
        idx, score, para = _match_poppler_paragraph("Hello World", paras, {0})
        assert idx == 1

    def test_empty_vlm_text(self):
        paras = [_make_poppler_para("Some text")]
        idx, _, _ = _match_poppler_paragraph("", paras, set())
        assert idx == -1

    def test_empty_paras(self):
        idx, _, _ = _match_poppler_paragraph("Some text", [], set())
        assert idx == -1

    def test_best_match_wins(self):
        paras = [
            _make_poppler_para("Annual Report 2024"),
            _make_poppler_para("Annual Report 2024 Summary"),
            _make_poppler_para("Quarterly Update"),
        ]
        idx, _, para = _match_poppler_paragraph("Annual Report 2024", paras, set())
        assert idx == 0  # exact match should win
        assert para["text"] == "Annual Report 2024"

    def test_all_consumed(self):
        paras = [_make_poppler_para("Only paragraph")]
        idx, _, para = _match_poppler_paragraph("Only paragraph", paras, {0})
        assert idx == -1
        assert para is None


# =============================================================================
# heading_level_from_poppler
# =============================================================================


class TestHeadingLevelFromPoppler:
    """Tests for heading_level_from_poppler."""

    @pytest.mark.parametrize("size,expected", [
        (24.0, 1), (20.0, 1), (18.1, 1),  # > 18pt → H1
    ])
    def test_h1(self, size, expected):
        assert heading_level_from_poppler(size, False) == expected

    @pytest.mark.parametrize("size,expected", [
        (18.0, 2), (16.0, 2), (14.0, 2),  # >= 14pt → H2
    ])
    def test_h2(self, size, expected):
        assert heading_level_from_poppler(size, False) == expected

    @pytest.mark.parametrize("size,expected", [
        (13.0, 3), (12.0, 3),  # >= 12pt → H3
    ])
    def test_h3(self, size, expected):
        assert heading_level_from_poppler(size, False) == expected

    def test_h4_bold_only(self):
        assert heading_level_from_poppler(11.0, True) == 4
        assert heading_level_from_poppler(11.5, True) == 4

    def test_h4_requires_bold(self):
        # 11pt NOT bold → not a heading
        assert heading_level_from_poppler(11.0, False) == 0

    def test_small_text_not_heading(self):
        assert heading_level_from_poppler(10.0, False) == 0
        assert heading_level_from_poppler(10.0, True) == 0
        assert heading_level_from_poppler(9.0, False) == 0

    def test_edge_cases(self):
        assert heading_level_from_poppler(0, False) == 0
        assert heading_level_from_poppler(None, False) == 0
        assert heading_level_from_poppler(-1, False) == 0

    def test_boundary_at_12(self):
        # 11.9pt NOT bold → not a heading (< 12, not bold enough for H4)
        assert heading_level_from_poppler(11.9, False) == 0
        # 11.9pt bold → H4
        assert heading_level_from_poppler(11.9, True) == 4
        # 12.0pt → H3 (regardless of bold)
        assert heading_level_from_poppler(12.0, False) == 3
        assert heading_level_from_poppler(12.0, True) == 3


# =============================================================================
# _apply_poppler_styling
# =============================================================================


class TestApplyPopplerStyling:
    """Tests for _apply_poppler_styling."""

    def test_font_size(self):
        elem = _make_page_xml(
            '<heading level="1"><run font-size-pt="18">Title</run></heading>'
        ).find("heading")
        para = _make_poppler_para("Title", font_size_pts=20.0)
        _apply_poppler_styling(elem, para)
        assert elem.find("run").get("font-size-pt") == "20.0"

    def test_bold_true(self):
        elem = _make_page_xml(
            '<paragraph><run font-size-pt="11">Text</run></paragraph>'
        ).find("paragraph")
        para = _make_poppler_para("Text", bold=True)
        _apply_poppler_styling(elem, para)
        assert elem.find("run").get("bold") == "true"

    def test_bold_false_removes(self):
        elem = _make_page_xml(
            '<paragraph><run font-size-pt="11" bold="true">Text</run></paragraph>'
        ).find("paragraph")
        para = _make_poppler_para("Text", bold=False)
        _apply_poppler_styling(elem, para)
        assert elem.find("run").get("bold") is None

    def test_italic_true(self):
        elem = _make_page_xml(
            '<paragraph><run font-size-pt="11">Text</run></paragraph>'
        ).find("paragraph")
        para = _make_poppler_para("Text", italic=True)
        _apply_poppler_styling(elem, para)
        assert elem.find("run").get("italic") == "true"

    def test_italic_false_removes(self):
        elem = _make_page_xml(
            '<paragraph><run font-size-pt="11" italic="true">Text</run></paragraph>'
        ).find("paragraph")
        para = _make_poppler_para("Text", italic=False)
        _apply_poppler_styling(elem, para)
        assert elem.find("run").get("italic") is None

    def test_color_hex_to_rgb(self):
        elem = _make_page_xml(
            '<paragraph><run font-size-pt="11">Text</run></paragraph>'
        ).find("paragraph")
        para = _make_poppler_para("Text", color="#1a5276")
        _apply_poppler_styling(elem, para)
        assert elem.find("run").get("color-rgb") == "26,82,118"

    def test_color_black_not_set(self):
        elem = _make_page_xml(
            '<paragraph><run font-size-pt="11">Text</run></paragraph>'
        ).find("paragraph")
        para = _make_poppler_para("Text", color="#000000")
        _apply_poppler_styling(elem, para)
        assert elem.find("run").get("color-rgb") is None

    def test_color_white_not_set(self):
        elem = _make_page_xml(
            '<paragraph><run font-size-pt="11">Text</run></paragraph>'
        ).find("paragraph")
        para = _make_poppler_para("Text", color="#ffffff")
        _apply_poppler_styling(elem, para)
        assert elem.find("run").get("color-rgb") is None

    def test_color_empty_not_set(self):
        elem = _make_page_xml(
            '<paragraph><run font-size-pt="11">Text</run></paragraph>'
        ).find("paragraph")
        para = _make_poppler_para("Text", color="")
        _apply_poppler_styling(elem, para)
        assert elem.find("run").get("color-rgb") is None

    def test_no_runs_returns_zero(self):
        # Element without <run> children
        elem = etree.SubElement(etree.Element("page"), "paragraph")
        elem.text = "Plain text"
        para = _make_poppler_para("Plain text", font_size_pts=14.0)
        assert _apply_poppler_styling(elem, para) == 0

    def test_multiple_runs(self):
        elem = _make_page_xml(
            '<paragraph>'
            '<run font-size-pt="11">First</run>'
            '<run font-size-pt="11">Second</run>'
            '</paragraph>'
        ).find("paragraph")
        para = _make_poppler_para("First Second", font_size_pts=14.0, bold=True,
                                  color="#ff0000")
        _apply_poppler_styling(elem, para)
        for run in elem.findall("run"):
            assert run.get("font-size-pt") == "14.0"
            assert run.get("bold") == "true"
            assert run.get("color-rgb") == "255,0,0"

    def test_color_invalid_hex_ignored(self):
        elem = _make_page_xml(
            '<paragraph><run font-size-pt="11">Text</run></paragraph>'
        ).find("paragraph")
        para = _make_poppler_para("Text", color="#ZZZZZZ")
        _apply_poppler_styling(elem, para)
        assert elem.find("run").get("color-rgb") is None

    def test_color_short_hex_ignored(self):
        elem = _make_page_xml(
            '<paragraph><run font-size-pt="11">Text</run></paragraph>'
        ).find("paragraph")
        para = _make_poppler_para("Text", color="#f00")
        _apply_poppler_styling(elem, para)
        assert elem.find("run").get("color-rgb") is None


# =============================================================================
# match_and_replace_text with poppler
# =============================================================================


class TestMatchAndReplaceWithPoppler:
    """Tests for poppler integration into match_and_replace_text."""

    def test_without_poppler_unchanged(self):
        """Without poppler, behavior is identical to original."""
        page_el = _make_page_xml(
            '<heading level="1" alignment="center">'
            '<run font-size-pt="18" bold="true">Contrat Title</run>'
            '</heading>'
            '<paragraph alignment="left">'
            '<run font-size-pt="11">Body text here.</run>'
            '</paragraph>'
        )
        ocr_regions = [
            _make_ocr_region("Contract Title", "doc_title", [100, 50, 900, 100]),
            _make_ocr_region("Body text here.", "text", [100, 150, 900, 200], index=1),
        ]
        replaced, total, unmatched = match_and_replace_text(
            page_el, ocr_regions, 1, poppler_paras=None)
        assert replaced == 2
        assert _get_element_text(page_el.find("heading")) == "Contract Title"

    def test_poppler_text_takes_priority(self):
        """Poppler text should override both VLM and OCR text."""
        page_el = _make_page_xml(
            '<heading level="1" alignment="center">'
            '<run font-size-pt="18" bold="true">VLM Title</run>'
            '</heading>'
            '<paragraph alignment="left">'
            '<run font-size-pt="11">VLM body text.</run>'
            '</paragraph>'
        )
        ocr_regions = [
            _make_ocr_region("OCR Title", "doc_title", [100, 50, 900, 100]),
            _make_ocr_region("OCR body text.", "text", [100, 150, 900, 200], index=1),
        ]
        poppler_paras = [
            _make_poppler_para("Exact Title from PDF", font_size_pts=20.0,
                               bold=True, top=100),
            _make_poppler_para("Exact body text from PDF.", font_size_pts=11.0,
                               top=200),
        ]
        replaced, total, _ = match_and_replace_text(
            page_el, ocr_regions, 1, poppler_paras=poppler_paras)
        assert replaced == 2
        assert _get_element_text(page_el.find("heading")) == "Exact Title from PDF"
        assert _get_element_text(page_el.find("paragraph")) == "Exact body text from PDF."

    def test_poppler_styling_applied(self):
        """Font size, color, bold should be applied from poppler."""
        page_el = _make_page_xml(
            '<heading level="1" alignment="center">'
            '<run font-size-pt="18" bold="true">Title</run>'
            '</heading>'
        )
        ocr_regions = [
            _make_ocr_region("Title", "doc_title", [100, 50, 900, 100]),
        ]
        poppler_paras = [
            _make_poppler_para("Title", font_size_pts=24.0, bold=True,
                               color="#1a5276", top=100),
        ]
        match_and_replace_text(page_el, ocr_regions, 1, poppler_paras=poppler_paras)
        run = page_el.find("heading/run")
        assert run.get("font-size-pt") == "24.0"
        assert run.get("bold") == "true"
        assert run.get("color-rgb") == "26,82,118"

    def test_poppler_heading_level_override(self):
        """Heading level should be determined from poppler font size."""
        page_el = _make_page_xml(
            '<heading level="3" alignment="left">'
            '<run font-size-pt="12">Section Title</run>'
            '</heading>'
        )
        ocr_regions = [
            _make_ocr_region("Section Title", "paragraph_title", [100, 50, 900, 100]),
        ]
        # Poppler says 14pt → should be H2 (not the VLM's H3)
        poppler_paras = [
            _make_poppler_para("Section Title", font_size_pts=14.0, bold=True, top=100),
        ]
        match_and_replace_text(page_el, ocr_regions, 1, poppler_paras=poppler_paras)
        assert page_el.find("heading").get("level") == "2"

    def test_partial_poppler_ocr_fallback(self):
        """Elements without poppler match should fall back to OCR text."""
        page_el = _make_page_xml(
            '<heading level="1" alignment="center">'
            '<run font-size-pt="18" bold="true">VLM Title</run>'
            '</heading>'
            '<paragraph alignment="left">'
            '<run font-size-pt="11">VLM body text content.</run>'
            '</paragraph>'
        )
        ocr_regions = [
            _make_ocr_region("OCR Title", "doc_title", [100, 50, 900, 100]),
            _make_ocr_region("OCR body text content.", "text", [100, 150, 900, 200], index=1),
        ]
        # Only one poppler paragraph — second element should use OCR
        poppler_paras = [
            _make_poppler_para("Exact Title from PDF", font_size_pts=20.0,
                               bold=True, top=100),
        ]
        replaced, total, _ = match_and_replace_text(
            page_el, ocr_regions, 1, poppler_paras=poppler_paras)
        assert replaced == 2
        assert _get_element_text(page_el.find("heading")) == "Exact Title from PDF"
        # Second element should have OCR text (fallback)
        assert _get_element_text(page_el.find("paragraph")) == "OCR body text content."

    def test_empty_poppler_list_falls_back(self):
        """Empty poppler list = all OCR (existing behavior)."""
        page_el = _make_page_xml(
            '<paragraph alignment="left">'
            '<run font-size-pt="11">Some text here</run>'
            '</paragraph>'
        )
        ocr_regions = [
            _make_ocr_region("Some text here", "text", [0, 0, 500, 100]),
        ]
        replaced, _, _ = match_and_replace_text(
            page_el, ocr_regions, 1, poppler_paras=[])
        assert replaced == 1
        assert _get_element_text(page_el.find("paragraph")) == "Some text here"

    def test_poppler_does_not_change_paragraphs_to_headings(self):
        """If VLM says paragraph, poppler heading detection should NOT upgrade it.

        Heading level override only applies when VLM already says it's a heading.
        """
        page_el = _make_page_xml(
            '<paragraph alignment="left">'
            '<run font-size-pt="11">Big paragraph text</run>'
            '</paragraph>'
        )
        ocr_regions = [
            _make_ocr_region("Big paragraph text", "text", [100, 50, 900, 100]),
        ]
        # Poppler has large font (H1 size) but VLM says paragraph
        poppler_paras = [
            _make_poppler_para("Big paragraph text", font_size_pts=20.0, top=100),
        ]
        match_and_replace_text(page_el, ocr_regions, 1, poppler_paras=poppler_paras)
        # Element should still be a paragraph (not promoted to heading)
        assert page_el.find("paragraph") is not None
        assert page_el.find("heading") is None

    def test_poppler_italic_applied(self):
        """Italic should be applied from poppler."""
        page_el = _make_page_xml(
            '<paragraph alignment="left">'
            '<run font-size-pt="11">Italic text</run>'
            '</paragraph>'
        )
        ocr_regions = [
            _make_ocr_region("Italic text", "text", [0, 0, 500, 100]),
        ]
        poppler_paras = [
            _make_poppler_para("Italic text", italic=True, top=100),
        ]
        match_and_replace_text(page_el, ocr_regions, 1, poppler_paras=poppler_paras)
        assert page_el.find("paragraph/run").get("italic") == "true"


# =============================================================================
# load_poppler_data
# =============================================================================


class TestLoadPopplerData:
    """Tests for load_poppler_data."""

    def test_nonexistent_file(self, tmp_path):
        result = load_poppler_data(str(tmp_path / "nonexistent.json"))
        assert result is None

    def test_valid_json(self, tmp_path):
        data = {"pages": [{"number": 1, "paragraphs": []}]}
        path = tmp_path / "structure.json"
        path.write_text(json.dumps(data), encoding="utf-8")
        result = load_poppler_data(str(path))
        assert result == data

    def test_invalid_json(self, tmp_path):
        path = tmp_path / "bad.json"
        path.write_text("{invalid json", encoding="utf-8")
        result = load_poppler_data(str(path))
        assert result is None

    def test_empty_file(self, tmp_path):
        path = tmp_path / "empty.json"
        path.write_text("", encoding="utf-8")
        result = load_poppler_data(str(path))
        assert result is None


# =============================================================================
# Backward compatibility: existing functions still work
# =============================================================================


class TestBackwardCompatibility:
    """Verify existing functions are not broken by poppler additions."""

    def test_text_similarity(self):
        assert text_similarity("Hello World", "hello world") == 1.0
        assert text_similarity("", "") == 0.0
        assert text_similarity("The quick brown fox", "quick brown fox jumps") == 0.6

    def test_combined_similarity(self):
        assert combined_similarity("a b c", "a b c") > 0.9
        assert combined_similarity("", "") == 0.0

    def test_collect_text_elements(self):
        page_el = _make_page_xml(
            '<heading level="1"><run>Title</run></heading>'
            '<paragraph><run>Body</run></paragraph>'
            '<image src="test.jpg" bbox="0,0,100,100"/>'
        )
        elements = collect_text_elements(page_el)
        assert len(elements) == 2
        assert elements[0]["type"] == "heading"
        assert elements[1]["type"] == "paragraph"

    def test_replace_element_text(self):
        elem = _make_page_xml(
            '<paragraph><run font-size-pt="11">Old</run></paragraph>'
        ).find("paragraph")
        _replace_element_text(elem, "New text")
        assert _get_element_text(elem) == "New text"

    def test_match_and_replace_text_no_poppler(self):
        """Original behavior preserved when poppler_paras is None."""
        page_el = _make_page_xml(
            '<paragraph alignment="left">'
            '<run font-size-pt="11">Original text content</run>'
            '</paragraph>'
        )
        ocr_regions = [
            _make_ocr_region("Original text content", "text", [0, 0, 500, 100]),
        ]
        replaced, total, unmatched = match_and_replace_text(
            page_el, ocr_regions, 1)
        assert replaced == 1
        assert total == 1
        assert len(unmatched) == 0


# =============================================================================
# Property-based style: edge cases and invariants
# =============================================================================


class TestInvariants:
    """Property-based thinking: verify invariants hold across edge cases."""

    def test_heading_level_monotonic(self):
        """Larger font sizes should produce equal or smaller heading level numbers."""
        sizes = [10, 11, 12, 14, 18, 20, 24]
        levels = [heading_level_from_poppler(s, True) for s in sizes]
        # Filter non-zero levels, check they're non-increasing
        nonzero = [(s, l) for s, l in zip(sizes, levels) if l > 0]
        for i in range(1, len(nonzero)):
            assert nonzero[i][1] <= nonzero[i-1][1], (
                f"Level for {nonzero[i][0]}pt ({nonzero[i][1]}) > "
                f"level for {nonzero[i-1][0]}pt ({nonzero[i-1][1]})"
            )

    def test_poppler_match_never_returns_consumed_index(self):
        """Consumed indices must never be returned."""
        paras = [_make_poppler_para(f"Paragraph {i}") for i in range(10)]
        all_consumed = set(range(10))
        idx, _, _ = _match_poppler_paragraph("Paragraph 5", paras, all_consumed)
        assert idx == -1

    def test_build_lookup_is_pure(self):
        """Calling build_poppler_page_lookup twice yields equal results."""
        data = {
            "pages": [{
                "number": 1,
                "paragraphs": [_make_poppler_para("Hello", top=100)],
            }]
        }
        result1 = build_poppler_page_lookup(data)
        result2 = build_poppler_page_lookup(data)
        assert result1[1][0]["text"] == result2[1][0]["text"]

    def test_apply_styling_idempotent(self):
        """Applying same styling twice produces same result."""
        elem = _make_page_xml(
            '<paragraph><run font-size-pt="11">Text</run></paragraph>'
        ).find("paragraph")
        para = _make_poppler_para("Text", font_size_pts=14.0, bold=True,
                                  color="#ff0000")
        _apply_poppler_styling(elem, para)
        # Get state after first apply
        run = elem.find("run")
        state1 = {
            "font-size-pt": run.get("font-size-pt"),
            "bold": run.get("bold"),
            "color-rgb": run.get("color-rgb"),
        }
        # Apply again
        _apply_poppler_styling(elem, para)
        state2 = {
            "font-size-pt": run.get("font-size-pt"),
            "bold": run.get("bold"),
            "color-rgb": run.get("color-rgb"),
        }
        assert state1 == state2
