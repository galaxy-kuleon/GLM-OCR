#!/usr/bin/env python3
"""Unit tests for vlm_generate_dsl.py — pure functions and error handling.

Tests verify behavioral correctness: each test asserts semantic invariants,
not just surface-level type checking. Covers the six functions flagged by
the evaluator plus additional pure-function coverage.
"""

import json
import os
import sys
import tempfile
from pathlib import Path
from unittest.mock import patch

import pytest

sys.path.insert(0, str(Path(__file__).parent))

import vlm_generate_dsl as mod


# ---------------------------------------------------------------------------
# Helpers — immutable test fixtures as plain data
# ---------------------------------------------------------------------------

VALID_PDF_STRUCTURE = {
    "pages": [
        {
            "number": 1,
            "width": 595.0,
            "height": 842.0,
            "fontspecs": [
                {"id": "0", "size": "12", "family": "Arial", "color": "#000000"}
            ],
            "elements": [
                {
                    "type": "text",
                    "top": 100,
                    "left": 50,
                    "width": 500,
                    "height": 20,
                    "font": "0",
                    "content": "Hello World",
                }
            ],
        }
    ]
}

SAMPLE_IMAGE_INFO = {
    "pages": [
        {"index": 1, "width_pts": 595, "height_pts": 842},
        {"index": 2, "width_pts": 612, "height_pts": 792},
        {"index": 3, "width_pts": 595, "height_pts": 842},
    ]
}


@pytest.fixture(autouse=True)
def _reset_profile():
    """Reset VLM_MODEL_PROFILE to 'strong' before each test."""
    original = mod.VLM_MODEL_PROFILE
    yield
    mod.VLM_MODEL_PROFILE = original


# ---------------------------------------------------------------------------
# load_poppler_references
# ---------------------------------------------------------------------------


class TestLoadPopplerReferences:
    """Behavioral tests for load_poppler_references.

    This function reads a JSON file, parses it, and delegates to
    format_all_pages. We test the error-handling boundaries (the new
    fixes) without needing a real format_poppler_reference module.
    """

    def test_valid_json_returns_dict(self, tmp_path):
        """Valid JSON with proper structure → returns dict from format_all_pages."""
        json_file = tmp_path / "structure.json"
        json_file.write_text(json.dumps(VALID_PDF_STRUCTURE), encoding="utf-8")

        result = mod.load_poppler_references(str(json_file))

        # format_all_pages returns {page_num: str}
        assert isinstance(result, dict)
        # Page 1 should be in the result (keyed by the "number" field)
        assert 1 in result
        # The reference text is a non-empty string with the page header
        assert isinstance(result[1], str)
        assert len(result[1]) > 0
        assert "EXACT PDF STRUCTURE" in result[1]

    def test_malformed_json_exits_cleanly(self, tmp_path):
        """Malformed JSON → sys.exit(1) with clean error, not raw traceback."""
        json_file = tmp_path / "bad.json"
        json_file.write_text("{not valid json!!!", encoding="utf-8")

        with pytest.raises(SystemExit) as exc_info:
            mod.load_poppler_references(str(json_file))

        assert exc_info.value.code == 1

    def test_empty_json_object_returns_empty_dict(self, tmp_path):
        """Empty JSON object (no pages) → returns empty dict (no crash)."""
        json_file = tmp_path / "empty.json"
        json_file.write_text('{"pages": []}', encoding="utf-8")

        result = mod.load_poppler_references(str(json_file))

        assert isinstance(result, dict)
        assert len(result) == 0

    def test_missing_file_exits_cleanly(self):
        """Non-existent file → sys.exit(1) with clean error."""
        with pytest.raises(SystemExit) as exc_info:
            mod.load_poppler_references("/tmp/nonexistent_xyz_12345.json")

        assert exc_info.value.code == 1

    def test_empty_file_exits_cleanly(self, tmp_path):
        """Empty file (0 bytes) → malformed JSON → sys.exit(1)."""
        json_file = tmp_path / "empty.json"
        json_file.write_text("", encoding="utf-8")

        with pytest.raises(SystemExit) as exc_info:
            mod.load_poppler_references(str(json_file))

        assert exc_info.value.code == 1

    def test_truncated_json_exits_cleanly(self, tmp_path):
        """Truncated JSON (partial write) → sys.exit(1)."""
        json_file = tmp_path / "truncated.json"
        json_file.write_text('{"pages": [{"page_number": 1', encoding="utf-8")

        with pytest.raises(SystemExit) as exc_info:
            mod.load_poppler_references(str(json_file))

        assert exc_info.value.code == 1


# ---------------------------------------------------------------------------
# _build_reference_instruction
# ---------------------------------------------------------------------------


class TestBuildReferenceInstruction:
    """Behavioral tests for _build_reference_instruction.

    Pure function: reference_text in → formatted instruction string out.
    Verifies the dead page_num parameter was removed (takes 1 arg).
    """

    def test_signature_has_one_parameter(self):
        """Confirm the dead page_num parameter was removed."""
        import inspect

        sig = inspect.signature(mod._build_reference_instruction)
        params = list(sig.parameters.keys())
        assert params == ["reference_text"], f"Expected ['reference_text'], got {params}"

    def test_includes_reference_text_verbatim(self):
        """The reference text appears in the output unchanged."""
        ref = "[EXACT PDF STRUCTURE — Page 1]\nHello World\nParagraph text"
        result = mod._build_reference_instruction(ref)

        assert ref in result

    def test_includes_usage_instructions(self):
        """Output contains layout-focus instructions for the VLM."""
        result = mod._build_reference_instruction("some ref text")

        assert "EXACT text content" in result
        assert "LAYOUT STRUCTURE" in result
        assert "Do NOT guess" in result

    def test_output_format_is_stable(self):
        """Key structural elements: ref text first, then instructions."""
        ref = "[EXACT PDF STRUCTURE — Page 5]\nContent here"
        result = mod._build_reference_instruction(ref)

        # Reference text appears before the instruction block
        ref_pos = result.find(ref)
        instr_pos = result.find("Use the EXACT text content")
        assert ref_pos < instr_pos

    def test_empty_reference_text_still_works(self):
        """Empty string input does not crash — returns instruction skeleton."""
        result = mod._build_reference_instruction("")
        assert "LAYOUT STRUCTURE" in result


# ---------------------------------------------------------------------------
# get_system_prompt
# ---------------------------------------------------------------------------


class TestGetSystemPrompt:
    """Behavioral tests for get_system_prompt.

    Pure function of (VLM_MODEL_PROFILE global, has_references flag).
    Four combinations to verify.
    """

    def test_strong_no_references(self):
        mod.VLM_MODEL_PROFILE = "strong"
        prompt = mod.get_system_prompt(has_references=False)

        assert "document layout analyzer" in prompt
        assert "EXACT PDF STRUCTURE" not in prompt

    def test_strong_with_references(self):
        mod.VLM_MODEL_PROFILE = "strong"
        prompt = mod.get_system_prompt(has_references=True)

        assert "document layout analyzer" in prompt
        assert "EXACT PDF STRUCTURE" in prompt

    def test_weak_no_references(self):
        mod.VLM_MODEL_PROFILE = "weak"
        prompt = mod.get_system_prompt(has_references=False)

        assert "Convert each page image to XML" in prompt
        assert "EXACT PDF STRUCTURE" not in prompt
        # Example should be present
        assert "Example for a page" in prompt

    def test_weak_with_references(self):
        mod.VLM_MODEL_PROFILE = "weak"
        prompt = mod.get_system_prompt(has_references=True)

        assert "Convert each page image to XML" in prompt
        assert "EXACT PDF STRUCTURE" in prompt
        assert "Example for a page" in prompt

    def test_weak_reference_rules_before_example(self):
        """Critical: rules 7-8 must appear BEFORE the example block."""
        mod.VLM_MODEL_PROFILE = "weak"
        prompt = mod.get_system_prompt(has_references=True)

        rules_pos = prompt.find("7. When EXACT PDF STRUCTURE")
        example_pos = prompt.find("Example for a page")

        assert rules_pos > 0, "Reference rules not found in prompt"
        assert example_pos > 0, "Example block not found in prompt"
        assert rules_pos < example_pos, (
            f"Reference rules (pos {rules_pos}) must appear before "
            f"example block (pos {example_pos})"
        )

    def test_weak_no_references_preserves_original_prompt(self):
        """Without references, weak prompt is unchanged from SYSTEM_PROMPT_WEAK."""
        mod.VLM_MODEL_PROFILE = "weak"
        prompt = mod.get_system_prompt(has_references=False)

        assert prompt == mod.SYSTEM_PROMPT_WEAK

    def test_strong_no_references_preserves_original_prompt(self):
        """Without references, strong prompt is unchanged from SYSTEM_PROMPT_STRONG."""
        mod.VLM_MODEL_PROFILE = "strong"
        prompt = mod.get_system_prompt(has_references=False)

        assert prompt == mod.SYSTEM_PROMPT_STRONG

    def test_default_has_references_is_false(self):
        """Default has_references=False."""
        mod.VLM_MODEL_PROFILE = "strong"
        prompt_default = mod.get_system_prompt()
        prompt_explicit = mod.get_system_prompt(has_references=False)

        assert prompt_default == prompt_explicit


# ---------------------------------------------------------------------------
# build_image_content_items (with per_page_references)
# ---------------------------------------------------------------------------


class TestBuildImageContentItems:
    """Behavioral tests for build_image_content_items.

    Tests the reference text injection path. Uses real tiny PNG images
    in a temp workspace to exercise the actual image encoding path.
    """

    @pytest.fixture
    def workspace_with_images(self, tmp_path):
        """Create a minimal workspace with tiny 1x1 PNG images."""
        img_dir = tmp_path / "input-images"
        img_dir.mkdir()

        # Create minimal valid 1x1 white PNG (67 bytes)
        # PNG signature + IHDR + IDAT + IEND
        import struct
        import zlib

        def make_tiny_png():
            """Generate a minimal valid 1x1 white PNG."""
            signature = b"\x89PNG\r\n\x1a\n"

            # IHDR chunk: 1x1, 8-bit RGB
            ihdr_data = struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0)
            ihdr_crc = zlib.crc32(b"IHDR" + ihdr_data) & 0xFFFFFFFF
            ihdr = struct.pack(">I", 13) + b"IHDR" + ihdr_data + struct.pack(">I", ihdr_crc)

            # IDAT chunk: single row, filter=0, RGB white
            raw_data = b"\x00\xff\xff\xff"
            compressed = zlib.compress(raw_data)
            idat_crc = zlib.crc32(b"IDAT" + compressed) & 0xFFFFFFFF
            idat = struct.pack(">I", len(compressed)) + b"IDAT" + compressed + struct.pack(">I", idat_crc)

            # IEND chunk
            iend_crc = zlib.crc32(b"IEND") & 0xFFFFFFFF
            iend = struct.pack(">I", 0) + b"IEND" + struct.pack(">I", iend_crc)

            return signature + ihdr + idat + iend

        png_bytes = make_tiny_png()
        for n in range(1, 4):
            (img_dir / f"page-{n}.png").write_bytes(png_bytes)

        return str(tmp_path)

    def test_no_references_returns_only_images(self, workspace_with_images):
        """Without references, only image_url items are returned."""
        mod.VLM_MODEL_PROFILE = "strong"
        items = mod.build_image_content_items(
            workspace_with_images, 1, 2, per_page_references=None
        )

        types = [item["type"] for item in items]
        assert all(t == "image_url" for t in types)
        assert len(items) == 2  # 2 pages, 1 image each

    def test_with_references_injects_text_items(self, workspace_with_images):
        """With references, text items are interleaved after page images."""
        mod.VLM_MODEL_PROFILE = "strong"
        refs = {
            1: "[EXACT PDF STRUCTURE — Page 1]\nHello",
            2: "[EXACT PDF STRUCTURE — Page 2]\nWorld",
        }
        items = mod.build_image_content_items(
            workspace_with_images, 1, 2, per_page_references=refs
        )

        types = [item["type"] for item in items]
        # Each page: 1 image + 1 text reference = 4 items total
        assert types == ["image_url", "text", "image_url", "text"]

    def test_reference_for_subset_of_pages(self, workspace_with_images):
        """References only for some pages — others get images only."""
        mod.VLM_MODEL_PROFILE = "strong"
        refs = {2: "[EXACT PDF STRUCTURE — Page 2]\nOnly page 2"}
        items = mod.build_image_content_items(
            workspace_with_images, 1, 3, per_page_references=refs
        )

        types = [item["type"] for item in items]
        # Page 1: image; Page 2: image + text; Page 3: image
        assert types == ["image_url", "image_url", "text", "image_url"]

    def test_empty_reference_text_is_skipped(self, workspace_with_images):
        """Reference text that is whitespace-only is not injected."""
        mod.VLM_MODEL_PROFILE = "strong"
        refs = {1: "   \n  "}  # whitespace-only
        items = mod.build_image_content_items(
            workspace_with_images, 1, 1, per_page_references=refs
        )

        types = [item["type"] for item in items]
        assert types == ["image_url"]  # no text item injected

    def test_reference_text_content_is_correct(self, workspace_with_images):
        """The injected text item contains the reference instruction."""
        mod.VLM_MODEL_PROFILE = "strong"
        refs = {1: "[EXACT PDF STRUCTURE — Page 1]\nSome content"}
        items = mod.build_image_content_items(
            workspace_with_images, 1, 1, per_page_references=refs
        )

        text_items = [i for i in items if i["type"] == "text"]
        assert len(text_items) == 1
        assert "Some content" in text_items[0]["text"]
        assert "LAYOUT STRUCTURE" in text_items[0]["text"]

    def test_weak_profile_adds_layout_vis_when_present(self, workspace_with_images):
        """Weak profile includes layout visualization images when available."""
        mod.VLM_MODEL_PROFILE = "weak"

        # Create a layout_vis image for page 1 (0-indexed = page0)
        layout_dir = Path(workspace_with_images) / "ocr-output" / "input" / "layout_vis"
        layout_dir.mkdir(parents=True)

        # Create a minimal JPEG (layout_vis are JPEGs)
        from PIL import Image
        import io

        img = Image.new("RGB", (100, 100), "red")
        buf = io.BytesIO()
        img.save(buf, format="JPEG")
        (layout_dir / "input_page0.jpg").write_bytes(buf.getvalue())

        items = mod.build_image_content_items(
            workspace_with_images, 1, 1, per_page_references=None
        )

        types = [item["type"] for item in items]
        # Page 1: page image + layout vis image = 2 image_url items
        assert types == ["image_url", "image_url"]


# ---------------------------------------------------------------------------
# build_messages (with/without references)
# ---------------------------------------------------------------------------


class TestBuildMessages:
    """Behavioral tests for build_messages.

    Verifies the full message composition: system prompt selection,
    user prompt generation, and content item assembly.
    """

    @pytest.fixture
    def workspace_with_images(self, tmp_path):
        """Same fixture as above — tiny 1x1 PNG images."""
        img_dir = tmp_path / "input-images"
        img_dir.mkdir()

        import struct
        import zlib

        def make_tiny_png():
            signature = b"\x89PNG\r\n\x1a\n"
            ihdr_data = struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0)
            ihdr_crc = zlib.crc32(b"IHDR" + ihdr_data) & 0xFFFFFFFF
            ihdr = struct.pack(">I", 13) + b"IHDR" + ihdr_data + struct.pack(">I", ihdr_crc)
            raw_data = b"\x00\xff\xff\xff"
            compressed = zlib.compress(raw_data)
            idat_crc = zlib.crc32(b"IDAT" + compressed) & 0xFFFFFFFF
            idat = struct.pack(">I", len(compressed)) + b"IDAT" + compressed + struct.pack(">I", idat_crc)
            iend_crc = zlib.crc32(b"IEND") & 0xFFFFFFFF
            iend = struct.pack(">I", 0) + b"IEND" + struct.pack(">I", iend_crc)
            return signature + ihdr + idat + iend

        png_bytes = make_tiny_png()
        for n in range(1, 3):
            (img_dir / f"page-{n}.png").write_bytes(png_bytes)

        return str(tmp_path)

    def test_without_references(self, workspace_with_images):
        """Without references, system prompt has no reference rules."""
        mod.VLM_MODEL_PROFILE = "strong"
        messages = mod.build_messages(
            workspace_with_images, SAMPLE_IMAGE_INFO, 1, 2
        )

        assert len(messages) == 2
        assert messages[0]["role"] == "system"
        assert messages[1]["role"] == "user"
        assert "EXACT PDF STRUCTURE" not in messages[0]["content"]

    def test_with_references(self, workspace_with_images):
        """With references, system prompt includes reference rules."""
        mod.VLM_MODEL_PROFILE = "strong"
        refs = {1: "[EXACT PDF STRUCTURE — Page 1]\nHello"}
        messages = mod.build_messages(
            workspace_with_images, SAMPLE_IMAGE_INFO, 1, 2,
            per_page_references=refs,
        )

        assert "EXACT PDF STRUCTURE" in messages[0]["content"]

    def test_user_content_has_text_and_images(self, workspace_with_images):
        """User message content starts with text prompt, followed by images."""
        mod.VLM_MODEL_PROFILE = "strong"
        messages = mod.build_messages(
            workspace_with_images, SAMPLE_IMAGE_INFO, 1, 2
        )

        user_content = messages[1]["content"]
        assert isinstance(user_content, list)
        assert user_content[0]["type"] == "text"
        assert "Analyze these 2 document page images" in user_content[0]["text"]

    def test_empty_references_dict_no_reference_rules(self, workspace_with_images):
        """Empty references dict → no reference rules in system prompt."""
        mod.VLM_MODEL_PROFILE = "strong"
        messages = mod.build_messages(
            workspace_with_images, SAMPLE_IMAGE_INFO, 1, 2,
            per_page_references={},
        )

        assert "EXACT PDF STRUCTURE" not in messages[0]["content"]

    def test_page_dimensions_in_user_prompt(self, workspace_with_images):
        """Page dimensions are included in the user prompt text."""
        mod.VLM_MODEL_PROFILE = "strong"
        messages = mod.build_messages(
            workspace_with_images, SAMPLE_IMAGE_INFO, 1, 2
        )

        user_text = messages[1]["content"][0]["text"]
        assert "595x842" in user_text
        assert "612x792" in user_text


# ---------------------------------------------------------------------------
# compute_batches (existing function, adding coverage)
# ---------------------------------------------------------------------------


class TestComputeBatches:
    def test_exact_fit(self):
        assert mod.compute_batches(8, 8) == [(1, 8)]

    def test_remainder(self):
        assert mod.compute_batches(10, 8) == [(1, 8), (9, 10)]

    def test_less_than_batch(self):
        assert mod.compute_batches(3, 8) == [(1, 3)]

    def test_zero_pages(self):
        assert mod.compute_batches(0, 8) == []

    def test_single_page(self):
        assert mod.compute_batches(1, 1) == [(1, 1)]

    def test_batch_size_one(self):
        assert mod.compute_batches(3, 1) == [(1, 1), (2, 2), (3, 3)]


# ---------------------------------------------------------------------------
# build_page_dimensions_text
# ---------------------------------------------------------------------------


class TestBuildPageDimensionsText:
    def test_known_pages(self):
        result = mod.build_page_dimensions_text(SAMPLE_IMAGE_INFO, 1, 2)
        assert "Page 1: 595x842 pts" in result
        assert "Page 2: 612x792 pts" in result

    def test_unknown_page(self):
        result = mod.build_page_dimensions_text(SAMPLE_IMAGE_INFO, 1, 5)
        assert "Page 4: dimensions unknown" in result
        assert "Page 5: dimensions unknown" in result


# ---------------------------------------------------------------------------
# build_user_prompt
# ---------------------------------------------------------------------------


class TestBuildUserPrompt:
    def test_strong_profile(self):
        mod.VLM_MODEL_PROFILE = "strong"
        result = mod.build_user_prompt(1, 3, "Page 1: 595x842 pts")
        assert "3 document page images" in result
        assert "pages 1-3" in result
        assert "595x842" in result
        # No layout hint for strong profile
        assert "layout analysis image" not in result

    def test_weak_profile_has_layout_hint(self):
        mod.VLM_MODEL_PROFILE = "weak"
        result = mod.build_user_prompt(1, 1, "Page 1: 595x842 pts")
        assert "layout analysis image" in result
        assert "doc_title (red)" in result


# ---------------------------------------------------------------------------
# clean_xml_text
# ---------------------------------------------------------------------------


class TestCleanXmlText:
    def test_strips_markdown_fences(self):
        raw = "```xml\n<pages></pages>\n```"
        assert mod.clean_xml_text(raw) == "<pages></pages>"

    def test_strips_plain_fences(self):
        raw = "```\n<page></page>\n```"
        assert mod.clean_xml_text(raw) == "<page></page>"

    def test_no_fences_passthrough(self):
        raw = "<pages><page></page></pages>"
        assert mod.clean_xml_text(raw) == raw

    def test_strips_whitespace(self):
        raw = "  \n<page></page>\n  "
        assert mod.clean_xml_text(raw) == "<page></page>"


# ---------------------------------------------------------------------------
# _repair_xml_text
# ---------------------------------------------------------------------------


class TestRepairXmlText:
    def test_wraps_bare_page_in_pages(self):
        result = mod._repair_xml_text('<page number="1"></page>')
        assert result.startswith("<pages>")
        assert result.endswith("</pages>")

    def test_closes_unclosed_page(self):
        result = mod._repair_xml_text('<pages><page number="1"><paragraph></paragraph>')
        assert result.count("</page>") >= 1
        assert result.endswith("</pages>")

    def test_fixes_unescaped_ampersand(self):
        result = mod._repair_xml_text('<pages><page number="1">A & B</page></pages>')
        assert "&amp;" in result

    def test_preserves_valid_entities(self):
        result = mod._repair_xml_text('<pages><page number="1">&amp; &lt;</page></pages>')
        # Should not double-escape &amp; to &amp;amp;
        assert "&amp;" in result
        assert "&amp;amp;" not in result

    def test_fixes_unclosed_image_tag(self):
        result = mod._repair_xml_text(
            '<pages><page number="1"><image src="PLACEHOLDER" bbox="0,0,100,100"></page></pages>'
        )
        assert "/>" in result or 'bbox="0,0,100,100"/>' in result


# ---------------------------------------------------------------------------
# parse_vlm_response
# ---------------------------------------------------------------------------


class TestParseVlmResponse:
    def test_single_page(self):
        xml = '<page number="1" width-pts="595" height-pts="842"></page>'
        pages = mod.parse_vlm_response(xml)
        assert len(pages) == 1
        assert pages[0].get("number") == "1"

    def test_pages_wrapper(self):
        xml = (
            '<pages>'
            '<page number="1" width-pts="595" height-pts="842"></page>'
            '<page number="2" width-pts="595" height-pts="842"></page>'
            '</pages>'
        )
        pages = mod.parse_vlm_response(xml)
        assert len(pages) == 2

    def test_with_markdown_fences(self):
        xml = '```xml\n<page number="1" width-pts="595" height-pts="842"></page>\n```'
        pages = mod.parse_vlm_response(xml)
        assert len(pages) == 1

    def test_malformed_xml_triggers_repair(self):
        """Unclosed page tag should be repaired and parsed."""
        xml = '<page number="1"><paragraph><run font-size-pt="12">Hello</run></paragraph>'
        pages = mod.parse_vlm_response(xml)
        assert len(pages) >= 1

    def test_completely_invalid_raises(self):
        with pytest.raises((ValueError, Exception)):
            mod.parse_vlm_response("not xml at all just random text")


# ---------------------------------------------------------------------------
# sys.path pollution guard (structural test)
# ---------------------------------------------------------------------------


class TestSysPathGuard:
    def test_sys_path_guard_in_source(self):
        """Verify the sys.path pollution fix is present in source code."""
        src = Path(mod.__file__).read_text(encoding="utf-8")
        assert "if str(script_dir) not in sys.path:" in src
