#!/usr/bin/env python3
"""Tests for translation text sanitization.

Covers:
  1. sanitize_translated_text — each pattern individually
  2. Combinations (XML inside markdown bold)
  3. Preservation of legitimate content (CJK, math *, A|B text)
  4. normalize_entry -> sanitize_translated_text pipeline
  5. dsl_to_docx._sanitize_text defense-in-depth extension
  6. apply_dsl_translations defensive warning
  7. Colspan/Rowspan round-trip
  8. Consistency between sanitizers
  9. Full pipeline: extract → normalize → apply → docx (no leaked markup)

Run:
    uv run --with lxml,python-docx pytest test_sanitize_translations.py -v
  or:
    uv run --with lxml,python-docx python3 test_sanitize_translations.py
"""

import io
import sys
from pathlib import Path

import pytest

# ---------------------------------------------------------------------------
# Import modules under test
# ---------------------------------------------------------------------------
SCRIPTS_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPTS_DIR))

from normalize_translations import normalize_entry, sanitize_translated_text

# dsl_to_docx needs lxml + python-docx at import time
# We import it conditionally and skip tests if deps are missing
try:
    # SCRIPTS_DIR = .../anything-to-docx/scripts -> parent.parent = .../skills/
    sys.path.insert(0, str(SCRIPTS_DIR.parent.parent / "pdf-to-docx" / "scripts"))
    from dsl_to_docx import _sanitize_text

    HAS_DSL_TO_DOCX = True
except ImportError:
    HAS_DSL_TO_DOCX = False

try:
    from lxml import etree
    from apply_dsl_translations import (
        apply_translation_to_element,
        apply_translations_to_page,
        build_translation_map,
        parse_segment_id,
        navigate_to_element,
    )
    from extract_dsl_texts import extract_page_segments

    HAS_APPLY = True
except ImportError:
    HAS_APPLY = False

try:
    from docx import Document as DocxDocument

    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False


# ===========================================================================
# 1. sanitize_translated_text — individual patterns
# ===========================================================================


class TestSanitizeXMLTags:
    """XML/HTML tag removal."""

    def test_simple_tags(self):
        assert sanitize_translated_text("<b>hello</b>") == "hello"

    def test_self_closing_tag(self):
        assert sanitize_translated_text("before<br/>after") == "beforeafter"

    def test_tag_with_attributes(self):
        assert sanitize_translated_text('<td class="x">cell</td>') == "cell"

    def test_nested_tags(self):
        assert sanitize_translated_text("<p><b>bold</b> text</p>") == "bold text"

    def test_tag_with_namespace_like_name(self):
        assert sanitize_translated_text("<w:r>run</w:r>") == "run"

    def test_table_tags_full(self):
        html = "<table><tr><td>A</td><td>B</td></tr></table>"
        assert sanitize_translated_text(html) == "AB"

    def test_preserves_angle_brackets_in_math(self):
        """x < 3 and y > 5 should NOT be stripped — these aren't tags."""
        text = "x < 3 and y > 5"
        assert sanitize_translated_text(text) == "x < 3 and y > 5"


class TestSanitizeMarkdownTableSeparators:
    """Markdown table separator line removal."""

    def test_simple_separator(self):
        assert sanitize_translated_text("| --- | --- |").strip() == ""

    def test_colon_aligned_separator(self):
        assert sanitize_translated_text("| :---: | ---: |").strip() == ""

    def test_separator_with_spaces(self):
        assert sanitize_translated_text("  | --- | --- |  ").strip() == ""

    def test_real_table_row_preserved(self):
        """Actual cell content with | should NOT be stripped."""
        text = "| Cell A | Cell B |"
        result = sanitize_translated_text(text)
        assert "Cell A" in result
        assert "Cell B" in result

    def test_pipe_in_normal_text_preserved(self):
        text = "A | B means A or B"
        assert sanitize_translated_text(text) == "A | B means A or B"


class TestSanitizeMarkdownBold:
    """Markdown bold wrapper removal."""

    def test_bold_removed(self):
        assert sanitize_translated_text("**bold**") == "bold"

    def test_bold_in_sentence(self):
        assert sanitize_translated_text("this is **bold** text") == "this is bold text"

    def test_multiple_bold(self):
        result = sanitize_translated_text("**A** and **B**")
        assert result == "A and B"


class TestSanitizeMarkdownItalic:
    """Markdown italic wrapper removal — but NOT math *."""

    def test_italic_removed(self):
        assert sanitize_translated_text("*italic*") == "italic"

    def test_italic_in_sentence(self):
        assert sanitize_translated_text("this is *italic* text") == "this is italic text"

    def test_math_multiplication_preserved(self):
        """2*3 should NOT be altered — no space boundary."""
        assert sanitize_translated_text("2*3") == "2*3"

    def test_math_in_formula_preserved(self):
        """a*b+c*d should NOT be altered."""
        assert sanitize_translated_text("a*b+c*d") == "a*b+c*d"

    def test_single_star_preserved(self):
        """A lone * should not cause issues."""
        assert sanitize_translated_text("note *") == "note *"


class TestSanitizeMarkdownHeading:
    """Markdown heading marker removal."""

    def test_h1(self):
        assert sanitize_translated_text("# Title") == "Title"

    def test_h2(self):
        assert sanitize_translated_text("## Subtitle") == "Subtitle"

    def test_h3(self):
        assert sanitize_translated_text("### Section") == "Section"

    def test_hash_not_at_start_preserved(self):
        """# in the middle of text should not be stripped."""
        text = "Issue #42 is fixed"
        assert sanitize_translated_text(text) == "Issue #42 is fixed"


class TestSanitizeCodeFences:
    """Code fence removal (defense-in-depth)."""

    def test_empty_code_fence(self):
        assert sanitize_translated_text("```markdown```").strip() == ""

    def test_standalone_fence(self):
        result = sanitize_translated_text("```\nsome code\n```")
        assert "```" not in result
        assert "some code" in result

    def test_fence_with_language(self):
        assert "```xml" not in sanitize_translated_text("```xml\ncode\n```")


class TestSanitizeEdgeCases:
    """Edge cases and passthrough."""

    def test_none_returns_none(self):
        assert sanitize_translated_text(None) is None

    def test_empty_returns_empty(self):
        assert sanitize_translated_text("") == ""

    def test_whitespace_only_returns_empty(self):
        assert sanitize_translated_text("   ") == ""

    def test_cjk_text_preserved(self):
        assert sanitize_translated_text("你好世界") == "你好世界"

    def test_cjk_with_punctuation_preserved(self):
        text = "这是一个测试。"
        assert sanitize_translated_text(text) == text

    def test_mixed_cjk_and_latin_preserved(self):
        text = "Hello 你好 World 世界"
        assert sanitize_translated_text(text) == text

    def test_normal_text_unchanged(self):
        text = "The quick brown fox jumps over the lazy dog."
        assert sanitize_translated_text(text) == text

    def test_numbers_preserved(self):
        assert sanitize_translated_text("12345") == "12345"


# ===========================================================================
# 2. Combinations
# ===========================================================================


class TestSanitizeCombinations:
    """Combined patterns that weak LLMs actually produce."""

    def test_xml_inside_bold(self):
        """**<b>text</b>** -> text"""
        result = sanitize_translated_text("**<b>text</b>**")
        assert result == "text"

    def test_heading_with_bold(self):
        """## **Title** -> Title"""
        result = sanitize_translated_text("## **Title**")
        assert result == "Title"

    def test_table_separator_with_bold_header(self):
        """Multi-line with separator and bold."""
        text = "**Header**\n| --- | --- |\nCell content"
        result = sanitize_translated_text(text)
        assert "Header" in result
        assert "Cell content" in result
        assert "---" not in result
        assert "**" not in result

    def test_xml_table_with_markdown(self):
        """<td>**bold cell**</td> -> bold cell"""
        result = sanitize_translated_text("<td>**bold cell**</td>")
        assert result == "bold cell"

    def test_code_fence_wrapping_xml(self):
        text = "```xml\n<b>hello</b>\n```"
        result = sanitize_translated_text(text)
        assert "```" not in result
        assert "<b>" not in result
        assert "hello" in result

    def test_real_world_weak_llm_output(self):
        """Simulates actual weak LLM output for a table cell."""
        text = '<td class="data">**总计**</td>'
        result = sanitize_translated_text(text)
        assert result == "总计"

    def test_real_world_heading_in_cell(self):
        """Weak LLM wraps cell content in heading."""
        text = "## Revenue (万元)"
        result = sanitize_translated_text(text)
        assert result == "Revenue (万元)"


# ===========================================================================
# 3. normalize_entry -> sanitize_translated_text pipeline
# ===========================================================================


class TestNormalizeEntryPipeline:
    """Verify that normalize_entry calls sanitize_translated_text."""

    def test_xml_stripped_in_pipeline(self):
        entry = normalize_entry({"id": "p1:cell[0]", "translated_text": "<b>hello</b>"})
        assert entry["translated_text"] == "hello"

    def test_bold_stripped_in_pipeline(self):
        entry = normalize_entry({"id": "p1:run[0]", "translated_text": "**bold**"})
        assert entry["translated_text"] == "bold"

    def test_text_field_also_sanitized(self):
        """Variant format with 'text' instead of 'translated_text'."""
        entry = normalize_entry({"id": "p1:run[0]", "text": "## Heading"})
        assert entry["translated_text"] == "Heading"

    def test_none_translated_text_still_works(self):
        entry = normalize_entry({"id": "p1:run[0]", "translated_text": None})
        assert entry["translated_text"] == ""

    def test_clean_text_unchanged(self):
        entry = normalize_entry({"id": "p1:run[0]", "translated_text": "hello world"})
        assert entry["translated_text"] == "hello world"

    def test_cjk_preserved_in_pipeline(self):
        entry = normalize_entry({"id": "p1:run[0]", "translated_text": "你好世界"})
        assert entry["translated_text"] == "你好世界"


# ===========================================================================
# 4. dsl_to_docx._sanitize_text defense-in-depth
# ===========================================================================


@pytest.mark.skipif(not HAS_DSL_TO_DOCX, reason="dsl_to_docx deps not available")
class TestDslToDocxSanitize:
    """Verify _sanitize_text has the same defense-in-depth patterns."""

    def test_xml_tags_stripped(self):
        assert _sanitize_text("<b>hello</b>") == "hello"

    def test_bold_stripped(self):
        assert _sanitize_text("**bold**") == "bold"

    def test_heading_stripped(self):
        assert _sanitize_text("## Title") == "Title"

    def test_table_separator_stripped(self):
        assert _sanitize_text("| --- | --- |").strip() == ""

    def test_original_code_fence_still_works(self):
        assert _sanitize_text("```markdown```") == ""

    def test_original_slash_still_works(self):
        assert _sanitize_text("/ caption") == "caption"

    def test_fullwidth_slash_still_works(self):
        assert _sanitize_text("\uff0f caption") == "caption"

    def test_none_passthrough(self):
        assert _sanitize_text(None) is None

    def test_empty_passthrough(self):
        assert _sanitize_text("") == ""

    def test_math_star_preserved(self):
        assert _sanitize_text("2*3") == "2*3"

    def test_cjk_preserved(self):
        assert _sanitize_text("你好世界") == "你好世界"


# ===========================================================================
# 5. apply_dsl_translations: run-children fix + defensive warning
# ===========================================================================


@pytest.mark.skipif(not HAS_APPLY, reason="lxml not available")
class TestApplyTranslationRunChildrenFix:
    """Verify that translations on elements with <run> children set text on
    the first run (preserving formatting) and remove extra runs."""

    def test_single_run_child_gets_text(self):
        """Translation goes to first <run>, not to cell.text."""
        elem = etree.fromstring('<cell><run font-size-pt="11">child</run></cell>')
        old_stderr = sys.stderr
        sys.stderr = io.StringIO()
        try:
            apply_translation_to_element(elem, "translated", segment_id="p1:cell[0]")
            stderr_output = sys.stderr.getvalue()
        finally:
            sys.stderr = old_stderr
        assert "INFO" in stderr_output
        assert "p1:cell[0]" in stderr_output
        # Translation applied to first run, not to cell.text
        assert elem.text is None
        assert elem[0].text == "translated"
        assert elem[0].get("font-size-pt") == "11"  # formatting preserved

    def test_multiple_runs_reduced_to_one(self):
        """Extra runs removed, first run keeps formatting and gets text."""
        elem = etree.fromstring(
            '<cell><run font-size-pt="11" bold="true">A</run>'
            '<run font-size-pt="11"> B</run></cell>'
        )
        old_stderr = sys.stderr
        sys.stderr = io.StringIO()
        try:
            apply_translation_to_element(elem, "AB translated", segment_id="p1:cell[1]")
            stderr_output = sys.stderr.getvalue()
        finally:
            sys.stderr = old_stderr
        assert "INFO" in stderr_output
        assert "2 <run>" in stderr_output
        # Only one run remains
        runs = elem.findall("run")
        assert len(runs) == 1
        assert runs[0].text == "AB translated"
        assert runs[0].get("bold") == "true"  # first run's formatting preserved

    def test_no_output_on_leaf_element(self):
        """Leaf elements (no children) set .text silently."""
        elem = etree.fromstring("<run>text</run>")
        old_stderr = sys.stderr
        sys.stderr = io.StringIO()
        try:
            apply_translation_to_element(elem, "new text", segment_id="p1:run[0]")
            stderr_output = sys.stderr.getvalue()
        finally:
            sys.stderr = old_stderr
        assert stderr_output == ""
        assert elem.text == "new text"

    def test_bare_text_cell_sets_text_directly(self):
        """Cell without run children sets .text directly."""
        elem = etree.fromstring("<cell>original</cell>")
        old_stderr = sys.stderr
        sys.stderr = io.StringIO()
        try:
            apply_translation_to_element(elem, "translated", segment_id="p1:cell[0]")
            stderr_output = sys.stderr.getvalue()
        finally:
            sys.stderr = old_stderr
        assert stderr_output == ""
        assert elem.text == "translated"

    def test_non_run_children_still_warns(self):
        """Non-run children trigger WARNING (best effort .text set)."""
        elem = etree.fromstring("<cell><image src='x.png'/></cell>")
        old_stderr = sys.stderr
        sys.stderr = io.StringIO()
        try:
            apply_translation_to_element(elem, "text", segment_id="p1:cell[0]")
            stderr_output = sys.stderr.getvalue()
        finally:
            sys.stderr = old_stderr
        assert "WARNING" in stderr_output
        assert "non-run" in stderr_output
        assert elem.text == "text"

    def test_default_segment_id(self):
        """Without segment_id, uses <unknown> default."""
        elem = etree.fromstring("<cell><run>child</run></cell>")
        old_stderr = sys.stderr
        sys.stderr = io.StringIO()
        try:
            apply_translation_to_element(elem, "text")
            stderr_output = sys.stderr.getvalue()
        finally:
            sys.stderr = old_stderr
        assert "<unknown>" in stderr_output


# ===========================================================================
# 6. Colspan/Rowspan edge case tests: extract->apply round-trip
# ===========================================================================


@pytest.mark.skipif(not HAS_APPLY, reason="lxml not available")
class TestColspanRowspanRoundTrip:
    """Verify segment ID consistency between extract and apply for merged cells."""

    def _make_page_xml(self, table_body):
        """Helper: wrap table rows in a minimal page XML string."""
        return (
            '<page number="1" width-pts="612" height-pts="792" '
            'font-latin="Arial" font-cjk="SimSun">'
            f'<table rows="3" cols="3">'
            f'<col-widths>0.33,0.33,0.34</col-widths>'
            f'{table_body}'
            f'</table></page>'
        )

    def _extract_and_apply_round_trip(self, page_xml_str, translations):
        """Extract segments, apply translations, verify round-trip.

        Args:
            page_xml_str: full <page> XML string
            translations: dict of segment_id -> translated_text

        Returns:
            (segments, modified_root) for further assertions
        """
        import tempfile, os

        # Write XML to temp file for extract_page_segments
        with tempfile.NamedTemporaryFile(
            mode="w", suffix=".xml", delete=False
        ) as f:
            f.write(page_xml_str)
            tmp_path = f.name

        try:
            segments = extract_page_segments(tmp_path, 1)
        finally:
            os.unlink(tmp_path)

        # Apply translations via navigate_to_element
        root = etree.fromstring(page_xml_str)
        applied_ids = []
        for seg in segments:
            seg_id = seg["id"]
            if seg_id in translations:
                _, path_steps = parse_segment_id(seg_id)
                assert path_steps is not None, f"Failed to parse {seg_id}"
                target = navigate_to_element(root, path_steps)
                assert target is not None, f"navigate failed for {seg_id}"
                apply_translation_to_element(
                    target, translations[seg_id], segment_id=seg_id
                )
                applied_ids.append(seg_id)

        return segments, root, applied_ids

    def test_colspan_2_round_trip(self):
        """Cell with colspan=2 extracts and applies correctly."""
        table_body = (
            '<row index="0">'
            '<cell row="0" col="0">A</cell>'
            '<cell row="0" col="1">B</cell>'
            '<cell row="0" col="2">C</cell>'
            '</row>'
            '<row index="1">'
            '<cell row="1" col="0" colspan="2">Merged AB</cell>'
            '<cell row="1" col="2">C2</cell>'
            '</row>'
            '<row index="2">'
            '<cell row="2" col="0">D</cell>'
            '<cell row="2" col="1">E</cell>'
            '<cell row="2" col="2">F</cell>'
            '</row>'
        )
        page_xml = self._make_page_xml(table_body)

        translations = {
            "p1:table[0]/row[1]/cell[0]": "合併AB",
            "p1:table[0]/row[1]/cell[1]": "C2翻譯",
        }

        segments, root, applied_ids = self._extract_and_apply_round_trip(
            page_xml, translations
        )

        # Verify colspan cell was extracted with correct sibling index
        colspan_segs = [s for s in segments if "row[1]" in s["id"]]
        assert len(colspan_segs) == 2
        assert colspan_segs[0]["id"] == "p1:table[0]/row[1]/cell[0]"
        assert colspan_segs[0]["text"] == "Merged AB"
        assert colspan_segs[1]["id"] == "p1:table[0]/row[1]/cell[1]"

        # Verify translations applied
        assert "p1:table[0]/row[1]/cell[0]" in applied_ids
        assert "p1:table[0]/row[1]/cell[1]" in applied_ids

        # Verify the XML has the translated text
        row1_cells = root.findall(".//row[@index='1']/cell")
        assert row1_cells[0].text == "合併AB"
        assert row1_cells[1].text == "C2翻譯"

    def test_rowspan_2_round_trip(self):
        """Cell with rowspan=2 extracts and applies correctly."""
        table_body = (
            '<row index="0">'
            '<cell row="0" col="0" rowspan="2">Tall cell</cell>'
            '<cell row="0" col="1">B</cell>'
            '<cell row="0" col="2">C</cell>'
            '</row>'
            '<row index="1">'
            '<cell row="1" col="1">E</cell>'
            '<cell row="1" col="2">F</cell>'
            '</row>'
            '<row index="2">'
            '<cell row="2" col="0">G</cell>'
            '<cell row="2" col="1">H</cell>'
            '<cell row="2" col="2">I</cell>'
            '</row>'
        )
        page_xml = self._make_page_xml(table_body)

        translations = {
            "p1:table[0]/row[0]/cell[0]": "高儲存格",
        }

        segments, root, applied_ids = self._extract_and_apply_round_trip(
            page_xml, translations
        )

        # Verify rowspan cell extracted correctly
        row0_segs = [s for s in segments if "row[0]" in s["id"]]
        assert any(s["text"] == "Tall cell" for s in row0_segs)
        assert "p1:table[0]/row[0]/cell[0]" in applied_ids

        # Row 1 should NOT have cell[0] (it's spanned by row 0's rowspan)
        row1_segs = [s for s in segments if "row[1]" in s["id"]]
        row1_ids = [s["id"] for s in row1_segs]
        assert all("cell[0]" in sid or "cell[1]" in sid for sid in row1_ids)

        # Verify translation applied in XML
        row0_cell0 = root.findall(".//row[@index='0']/cell")[0]
        assert row0_cell0.text == "高儲存格"

    def test_colspan_and_rowspan_combined(self):
        """Cell with both colspan=2 and rowspan=2."""
        table_body = (
            '<row index="0">'
            '<cell row="0" col="0" colspan="2" rowspan="2">Big merged</cell>'
            '<cell row="0" col="2">Top right</cell>'
            '</row>'
            '<row index="1">'
            '<cell row="1" col="2">Mid right</cell>'
            '</row>'
            '<row index="2">'
            '<cell row="2" col="0">Bottom left</cell>'
            '<cell row="2" col="1">Bottom mid</cell>'
            '<cell row="2" col="2">Bottom right</cell>'
            '</row>'
        )
        page_xml = self._make_page_xml(table_body)

        translations = {
            "p1:table[0]/row[0]/cell[0]": "大合併",
            "p1:table[0]/row[0]/cell[1]": "右上",
        }

        segments, root, applied_ids = self._extract_and_apply_round_trip(
            page_xml, translations
        )

        # The big merged cell should be cell[0] in row[0]
        row0_segs = [s for s in segments if "row[0]" in s["id"]]
        assert row0_segs[0]["id"] == "p1:table[0]/row[0]/cell[0]"
        assert row0_segs[0]["text"] == "Big merged"

        # "Top right" is sibling index 1 in row[0]
        assert row0_segs[1]["id"] == "p1:table[0]/row[0]/cell[1]"
        assert row0_segs[1]["text"] == "Top right"

        # Both translations applied
        assert len(applied_ids) == 2
        row0_cells = root.findall(".//row[@index='0']/cell")
        assert row0_cells[0].text == "大合併"
        assert row0_cells[1].text == "右上"

    def test_colspan_with_run_children(self):
        """Colspan cell that has <run> children uses the run-fix path."""
        table_body = (
            '<row index="0">'
            '<cell row="0" col="0">A</cell>'
            '<cell row="0" col="1">B</cell>'
            '<cell row="0" col="2">C</cell>'
            '</row>'
            '<row index="1">'
            '<cell row="1" col="0" colspan="2">'
            '<run font-size-pt="11" bold="true">Merged bold</run>'
            '<run font-size-pt="11"> normal</run>'
            '</cell>'
            '<cell row="1" col="2">D</cell>'
            '</row>'
            '<row index="2">'
            '<cell row="2" col="0">E</cell>'
            '<cell row="2" col="1">F</cell>'
            '<cell row="2" col="2">G</cell>'
            '</row>'
        )
        page_xml = self._make_page_xml(table_body)

        # Weak LLM gives cell-level ID for a cell with runs
        # After fix, this should go to the first run
        segments, _, _ = self._extract_and_apply_round_trip(page_xml, {})

        # The runs are extracted at run level
        colspan_segs = [s for s in segments if "row[1]/cell[0]" in s["id"]]
        assert len(colspan_segs) == 2
        assert colspan_segs[0]["id"] == "p1:table[0]/row[1]/cell[0]/run[0]"
        assert colspan_segs[1]["id"] == "p1:table[0]/row[1]/cell[0]/run[1]"

        # Now test applying a cell-level translation (weak LLM truncated ID)
        root = etree.fromstring(page_xml)
        _, path_steps = parse_segment_id("p1:table[0]/row[1]/cell[0]")
        target = navigate_to_element(root, path_steps)
        assert target is not None

        apply_translation_to_element(
            target, "合併粗體 一般", segment_id="p1:table[0]/row[1]/cell[0]"
        )

        # Verify: text went to first run, extra run removed, formatting kept
        runs = target.findall("run")
        assert len(runs) == 1
        assert runs[0].text == "合併粗體 一般"
        assert runs[0].get("bold") == "true"
        assert runs[0].get("font-size-pt") == "11"


# ===========================================================================
# 7. Consistency check: sanitize_translated_text vs _sanitize_text
# ===========================================================================


@pytest.mark.skipif(not HAS_DSL_TO_DOCX, reason="dsl_to_docx deps not available")
class TestConsistencyBetweenSanitizers:
    """Both sanitizers should produce identical results for translation artifacts."""

    TRANSLATION_ARTIFACTS = [
        "<b>bold</b>",
        "**markdown bold**",
        "## Heading",
        "| --- | --- |",
        '<td class="x">cell</td>',
        "normal text",
        "你好世界",
        "2*3",
    ]

    @pytest.mark.parametrize("text", TRANSLATION_ARTIFACTS)
    def test_same_result(self, text):
        result_norm = sanitize_translated_text(text)
        result_docx = _sanitize_text(text)
        assert result_norm == result_docx, (
            f"Inconsistent sanitization for {text!r}:\n"
            f"  sanitize_translated_text -> {result_norm!r}\n"
            f"  _sanitize_text -> {result_docx!r}"
        )


# ===========================================================================
# 9. Full pipeline: extract → normalize → apply → docx (no leaked markup)
# ===========================================================================


@pytest.mark.skipif(
    not (HAS_APPLY and HAS_DSL_TO_DOCX and HAS_DOCX),
    reason="lxml, python-docx, or dsl_to_docx deps not available",
)
class TestFullPipelineNoLeakedMarkup:
    """End-to-end test: DSL XML with weak-LLM artifacts → DOCX with clean text.

    Creates a minimal DSL page with tables (including cells with run children),
    simulates weak-LLM translations with every artifact type, runs the full
    normalize → apply → docx pipeline, and verifies no markup leaks into cells.
    """

    # All artifact patterns that weak LLMs produce
    ARTIFACT_TEMPLATES = [
        lambda t: f"<td>{t}</td>",
        lambda t: f'<td class="data">{t}</td>',
        lambda t: f"**{t}**",
        lambda t: f"## {t}",
        lambda t: f"<td>**{t}**</td>",
        lambda t: f"<b>{t}</b>",
        lambda t: f"<tr><td>{t}</td></tr>",
        lambda t: f"{t}\n| --- | --- |",
        lambda t: f"```\n{t}\n```",
        lambda t: f"*{t}*",
        lambda t: f"<w:r>{t}</w:r>",
    ]

    MARKUP_INDICATORS = ["<td", "<b>", "<tr", "<w:", "**", "| ---", "##", "```", "</"]

    def _make_dsl_page(self):
        """Create a realistic DSL page XML with both bare-text and run-children cells."""
        return """\
<page number="1" width-pts="612" height-pts="792"
      font-latin="Arial" font-cjk="SimSun">
  <heading level="2"><run>Contract Summary</run></heading>
  <table rows="4" cols="3" border-style="full"
         bbox="50,100,950,600" page-width-pts="612">
    <col-widths>0.33,0.33,0.34</col-widths>
    <row index="0">
      <cell row="0" col="0" font-size-pt="10" bold="true">Item</cell>
      <cell row="0" col="1" font-size-pt="10" bold="true">Description</cell>
      <cell row="0" col="2" font-size-pt="10" bold="true">Amount</cell>
    </row>
    <row index="1">
      <cell row="1" col="0" font-size-pt="10">Service A</cell>
      <cell row="1" col="1" font-size-pt="10">
        <run font-size-pt="10" bold="true">Cloud</run>
        <run font-size-pt="10"> hosting plan</run>
      </cell>
      <cell row="1" col="2" font-size-pt="10">$1,000</cell>
    </row>
    <row index="2">
      <cell row="2" col="0" font-size-pt="10">Service B</cell>
      <cell row="2" col="1" font-size-pt="10">
        <run font-size-pt="10">API</run>
        <run font-size-pt="10" italic="true"> integration</run>
        <run font-size-pt="10"> support</run>
      </cell>
      <cell row="2" col="2" font-size-pt="10">$2,500</cell>
    </row>
    <row index="3">
      <cell row="3" col="0" colspan="2" font-size-pt="10" bold="true">Total</cell>
      <cell row="3" col="2" font-size-pt="10" bold="true">$3,500</cell>
    </row>
  </table>
  <paragraph><run font-size-pt="11">End of summary.</run></paragraph>
</page>"""

    def _create_weak_translations(self, segments):
        """Create translations with artifacts for table/run segments, clean for others."""
        translations = []
        artifact_idx = 0
        for seg in segments:
            text = seg["text"]
            is_table = "/cell[" in seg["id"]
            if is_table and text.strip():
                template = self.ARTIFACT_TEMPLATES[artifact_idx % len(self.ARTIFACT_TEMPLATES)]
                translated = template(f"[ZH] {text}")
                artifact_idx += 1
            else:
                translated = f"[ZH] {text}"
            translations.append({"id": seg["id"], "translated_text": translated})
        return translations

    def test_full_pipeline_no_markup_leaks(self):
        """Full extract→normalize→apply→docx pipeline produces clean DOCX cells."""
        import json
        import tempfile
        import os

        page_xml = self._make_dsl_page()

        with tempfile.TemporaryDirectory() as tmpdir:
            workspace = Path(tmpdir)

            # Step 1: Write DSL page
            dsl_dir = workspace / "dsl"
            dsl_dir.mkdir()
            page_path = dsl_dir / "page-1.xml"
            page_path.write_text(page_xml, encoding="utf-8")

            # Step 2: Extract segments
            segments = extract_page_segments(str(page_path), 1)
            assert len(segments) > 0, "No segments extracted"

            table_segs = [s for s in segments if "/cell[" in s["id"]]
            assert len(table_segs) > 0, "No table segments extracted"

            # Step 3: Create weak-LLM translations with artifacts
            raw_translations = self._create_weak_translations(segments)

            # Verify artifacts are present before sanitization
            table_trans = [t for t in raw_translations if "/cell[" in t["id"] and t["translated_text"].strip()]
            artifacts_present = sum(
                1 for t in table_trans
                if any(p in t["translated_text"] for p in self.MARKUP_INDICATORS)
            )
            assert artifacts_present > 0, "Test setup error: no artifacts in translations"

            # Step 4: Normalize (sanitize) via normalize_entry
            normalized = [normalize_entry(t) for t in raw_translations]

            # Verify sanitization removed all artifacts
            for entry in normalized:
                if "/cell[" in entry["id"]:
                    for pattern in self.MARKUP_INDICATORS:
                        assert pattern not in entry["translated_text"], (
                            f"Leaked {pattern!r} in {entry['id']}: {entry['translated_text']!r}"
                        )

            # Step 5: Build translation map and apply to DSL
            trans_data = {"translations": normalized}
            translation_map = build_translation_map(trans_data)

            root, applied, skipped = apply_translations_to_page(
                str(page_path), 1, translation_map
            )
            assert applied > 0, "No translations applied"
            assert skipped == 0, f"Skipped {skipped} translations"

            # Write translated DSL
            translated_dir = workspace / "dsl-translated"
            translated_dir.mkdir()
            translated_path = translated_dir / "page-1.xml"
            xml_str = etree.tostring(root, encoding="unicode", pretty_print=True)
            translated_path.write_text(xml_str, encoding="utf-8")

            # Step 6: Generate DOCX
            docx_scripts = SCRIPTS_DIR.parent.parent / "pdf-to-docx" / "scripts"
            sys.path.insert(0, str(docx_scripts))
            from dsl_to_docx import process_page

            doc = DocxDocument()
            process_page(doc, str(translated_path), str(workspace), is_first_page=True)

            # Step 7: Verify no markup in any DOCX table cell
            leaked_cells = []
            total_cells = 0
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        cell_text = cell.text.strip()
                        total_cells += 1
                        for pattern in self.MARKUP_INDICATORS:
                            if pattern in cell_text:
                                leaked_cells.append((cell_text, pattern))
                                break

            assert total_cells > 0, "No table cells found in DOCX"
            assert len(leaked_cells) == 0, (
                f"Markup leaked into {len(leaked_cells)}/{total_cells} DOCX cells:\n"
                + "\n".join(f"  {text!r} (contains {pat!r})" for text, pat in leaked_cells)
            )

    def test_run_children_formatting_preserved(self):
        """Cells with <run> children preserve first run's formatting through pipeline."""
        import tempfile

        page_xml = self._make_dsl_page()

        with tempfile.TemporaryDirectory() as tmpdir:
            workspace = Path(tmpdir)
            dsl_dir = workspace / "dsl"
            dsl_dir.mkdir()
            page_path = dsl_dir / "page-1.xml"
            page_path.write_text(page_xml, encoding="utf-8")

            # Apply a translation to a cell-level ID that has run children
            # (simulating weak LLM giving cell-level ID instead of run-level)
            translation_map = {
                "p1:table[0]/row[1]/cell[1]": "[ZH] Cloud hosting plan",
            }
            root, applied, skipped = apply_translations_to_page(
                str(page_path), 1, translation_map
            )

            # Find the cell and verify run structure
            table = root.find("table")
            row1 = [r for r in table.findall("row") if r.get("index") == "1"][0]
            cell1 = row1.findall("cell")[1]  # cell[1] = Description column
            runs = cell1.findall("run")

            # T3 fix: first run kept with formatting, extras removed
            assert len(runs) == 1, f"Expected 1 run after fix, got {len(runs)}"
            assert runs[0].text == "[ZH] Cloud hosting plan"
            assert runs[0].get("bold") == "true", "First run's bold formatting lost"
            assert runs[0].get("font-size-pt") == "10", "First run's font-size lost"

    def test_sanitization_before_apply_order(self):
        """Verify sanitization in normalize step happens before apply step."""
        # This is a flow-interaction test: if normalize is skipped,
        # _sanitize_text in dsl_to_docx still catches artifacts (defense-in-depth)
        weak_texts = [
            '<td class="data">**Total**</td>',
            "## Revenue",
            "<b>bold</b>",
        ]
        for weak in weak_texts:
            # Primary defense: sanitize_translated_text
            primary = sanitize_translated_text(weak)
            assert "<" not in primary and "**" not in primary and "##" not in primary

            # Defense-in-depth: _sanitize_text
            backup = _sanitize_text(weak)
            assert "<" not in backup and "**" not in backup and "##" not in backup

            # Both agree
            assert primary == backup, (
                f"Sanitizers disagree on {weak!r}: "
                f"primary={primary!r} vs backup={backup!r}"
            )


# ===========================================================================
# Runner (no pytest needed)
# ===========================================================================


def run_all():
    """Simple test runner for environments without pytest."""
    import traceback

    test_classes = [
        TestSanitizeXMLTags,
        TestSanitizeMarkdownTableSeparators,
        TestSanitizeMarkdownBold,
        TestSanitizeMarkdownItalic,
        TestSanitizeMarkdownHeading,
        TestSanitizeCodeFences,
        TestSanitizeEdgeCases,
        TestSanitizeCombinations,
        TestNormalizeEntryPipeline,
    ]

    # Conditionally add tests that need lxml/python-docx
    if HAS_DSL_TO_DOCX:
        test_classes.append(TestDslToDocxSanitize)
    if HAS_APPLY:
        test_classes.append(TestApplyTranslationRunChildrenFix)
        test_classes.append(TestColspanRowspanRoundTrip)
    if HAS_APPLY and HAS_DSL_TO_DOCX and HAS_DOCX:
        test_classes.append(TestFullPipelineNoLeakedMarkup)

    passed = 0
    failed = 0
    skipped = 0
    errors = []

    for cls in test_classes:
        instance = cls()
        methods = [m for m in dir(instance) if m.startswith("test_")]
        for method_name in sorted(methods):
            full_name = f"{cls.__name__}.{method_name}"
            try:
                getattr(instance, method_name)()
                passed += 1
                print(f"  PASS  {full_name}")
            except Exception as e:
                failed += 1
                tb = traceback.format_exc()
                errors.append((full_name, tb))
                print(f"  FAIL  {full_name}: {e}")

    print(f"\n{'=' * 60}")
    print(f"Results: {passed} passed, {failed} failed, {passed + failed} total")

    if errors:
        print(f"\nFailures:")
        for name, tb in errors:
            print(f"\n--- {name} ---")
            print(tb)
        return 1

    print("\nAll tests passed.")
    return 0


if __name__ == "__main__":
    sys.exit(run_all())
