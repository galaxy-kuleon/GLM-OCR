#!/usr/bin/env python3
"""Comprehensive tests for normalize_translations.py.

Covers all 8 pure functions with adversarial inputs matching evaluator findings:
  1. None/null translated_text coercion
  2. Missing/empty id warnings
  3. Malformed JSON handling
  4. Duplicate ID detection
  5. Exit code on validation mismatch
  6. (Integration - tested via file existence check)
  7. detect_format edge cases
  8. All pure function coverage

Run: python3 test_normalize_translations.py
  or: python3 -m pytest test_normalize_translations.py -v
"""

import io
import json
import subprocess
import sys
import tempfile
from pathlib import Path

# ---------------------------------------------------------------------------
# Import the module under test
# ---------------------------------------------------------------------------
sys.path.insert(0, str(Path(__file__).resolve().parent))

from normalize_translations import (
    detect_format,
    extract_entries,
    find_batch_files,
    load_and_normalize_batch,
    load_expected_total,
    merge_all_batches,
    normalize_batch,
    normalize_entry,
    validate_count,
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def capture_stderr(fn, *args, **kwargs):
    """Run fn capturing stderr. Returns (result, stderr_text)."""
    old = sys.stderr
    sys.stderr = io.StringIO()
    try:
        result = fn(*args, **kwargs)
        return result, sys.stderr.getvalue()
    finally:
        sys.stderr = old


class ManagedWorkspace:
    """Context manager that creates a temp workspace and cleans up after use."""

    def __init__(self, files: dict[str, object]):
        self._files = files
        self._tmpdir = None

    def __enter__(self) -> Path:
        self._tmpdir = tempfile.TemporaryDirectory()
        ws = Path(self._tmpdir.name)
        for name, data in self._files.items():
            path = ws / name
            if isinstance(data, str):
                path.write_text(data, encoding="utf-8")
            elif isinstance(data, bytes):
                path.write_bytes(data)
            else:
                with open(path, "w", encoding="utf-8") as f:
                    json.dump(data, f, ensure_ascii=False)
        return ws

    def __exit__(self, *exc):
        if self._tmpdir:
            self._tmpdir.cleanup()


def make_workspace(files: dict[str, object]) -> Path:
    """Create a temp workspace with given filename -> data mappings.

    Data is JSON-serialized unless it's a string (written raw) or bytes (written raw).
    Returns Path to the workspace directory.

    NOTE: For tests that need cleanup, prefer ManagedWorkspace context manager.
    This function is kept for backward compat but callers should migrate.
    """
    tmpdir = Path(tempfile.mkdtemp())
    for name, data in files.items():
        path = tmpdir / name
        if isinstance(data, str):
            path.write_text(data, encoding="utf-8")
        elif isinstance(data, bytes):
            path.write_bytes(data)
        else:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False)
    return tmpdir


# ===========================================================================
# 1. normalize_entry
# ===========================================================================

class TestNormalizeEntry:

    def test_canonical_passthrough(self):
        entry = normalize_entry({"id": "p1:r0", "translated_text": "hello"})
        assert entry == {"id": "p1:r0", "translated_text": "hello"}

    def test_text_field_renamed(self):
        entry = normalize_entry({"id": "p1:r0", "text": "hello"})
        assert entry == {"id": "p1:r0", "translated_text": "hello"}

    def test_translated_text_none_coerced_to_empty(self):
        """Finding 1: JSON null -> empty string."""
        entry = normalize_entry({"id": "x", "translated_text": None})
        assert entry["translated_text"] == ""

    def test_text_field_none_coerced_to_empty(self):
        """Finding 1 variant: text field is None."""
        entry = normalize_entry({"id": "x", "text": None})
        assert entry["translated_text"] == ""

    def test_missing_both_text_fields(self):
        entry = normalize_entry({"id": "x"})
        assert entry["translated_text"] == ""

    def test_empty_string_preserved(self):
        entry = normalize_entry({"id": "x", "translated_text": ""})
        assert entry["translated_text"] == ""

    def test_missing_id_warns(self):
        """Finding 2: missing id emits warning."""
        entry, stderr = capture_stderr(normalize_entry, {"translated_text": "hi"})
        assert entry["id"] == ""
        assert "WARNING" in stderr
        assert "missing or empty" in stderr

    def test_empty_id_warns(self):
        """Finding 2: empty string id emits warning."""
        entry, stderr = capture_stderr(normalize_entry, {"id": "", "translated_text": "hi"})
        assert entry["id"] == ""
        assert "WARNING" in stderr

    def test_whitespace_only_id_warns_and_coerced_to_empty(self):
        """REC 6: whitespace-only id emits warning and is coerced to ""."""
        entry, stderr = capture_stderr(normalize_entry, {"id": "   ", "translated_text": "hi"})
        assert "WARNING" in stderr
        assert entry["id"] == "", f"Expected empty string, got {entry['id']!r}"

    def test_translated_text_takes_priority_over_text(self):
        entry = normalize_entry({"id": "x", "translated_text": "A", "text": "B"})
        assert entry["translated_text"] == "A"

    def test_extra_fields_stripped(self):
        entry = normalize_entry({"id": "x", "translated_text": "hi", "extra": 123, "batch": 1})
        assert set(entry.keys()) == {"id", "translated_text"}

    # RED 1: non-dict entry type guards
    def test_int_entry_returns_empty_defaults(self):
        """RED 1: int entry must not crash, returns empty defaults."""
        entry, stderr = capture_stderr(normalize_entry, 42)
        assert entry == {"id": "", "translated_text": ""}
        assert "WARNING" in stderr
        assert "non-dict" in stderr
        assert "int" in stderr

    def test_string_entry_returns_empty_defaults(self):
        """RED 1: string entry must not crash, returns empty defaults."""
        entry, stderr = capture_stderr(normalize_entry, "hello")
        assert entry == {"id": "", "translated_text": ""}
        assert "non-dict" in stderr

    def test_list_entry_returns_empty_defaults(self):
        """RED 1: list entry must not crash, returns empty defaults."""
        entry, stderr = capture_stderr(normalize_entry, ["a", "b"])
        assert entry == {"id": "", "translated_text": ""}
        assert "non-dict" in stderr

    def test_none_entry_returns_empty_defaults(self):
        """RED 1: None entry must not crash, returns empty defaults."""
        entry, stderr = capture_stderr(normalize_entry, None)
        assert entry == {"id": "", "translated_text": ""}
        assert "non-dict" in stderr

    def test_bool_entry_returns_empty_defaults(self):
        """RED 1: bool entry must not crash, returns empty defaults."""
        entry, stderr = capture_stderr(normalize_entry, True)
        assert entry == {"id": "", "translated_text": ""}
        assert "non-dict" in stderr

    def test_float_entry_returns_empty_defaults(self):
        """RED 1: float entry must not crash."""
        entry, stderr = capture_stderr(normalize_entry, 3.14)
        assert entry == {"id": "", "translated_text": ""}
        assert "non-dict" in stderr


# ===========================================================================
# 2. extract_entries
# ===========================================================================

class TestExtractEntries:

    def test_canonical_dict(self):
        data = {"translations": [{"id": "a", "translated_text": "x"}]}
        assert extract_entries(data) == [{"id": "a", "translated_text": "x"}]

    def test_segments_dict(self):
        data = {"segments": [{"id": "a", "text": "x"}]}
        assert extract_entries(data) == [{"id": "a", "text": "x"}]

    def test_batch_segments_dict(self):
        data = {"batch": 1, "segments": [{"id": "a", "text": "x"}]}
        assert extract_entries(data) == [{"id": "a", "text": "x"}]

    def test_bare_array(self):
        data = [{"id": "a", "translated_text": "x"}]
        assert extract_entries(data) == [{"id": "a", "translated_text": "x"}]

    def test_empty_bare_array(self):
        assert extract_entries([]) == []

    def test_translations_with_text_field(self):
        data = {"translations": [{"id": "a", "text": "x"}]}
        assert extract_entries(data) == [{"id": "a", "text": "x"}]

    def test_unrecognizable_format_raises(self):
        try:
            extract_entries({"bogus_key": []})
            assert False, "Should have raised ValueError"
        except ValueError as e:
            assert "Unrecognizable" in str(e)

    def test_non_dict_non_list_raises(self):
        try:
            extract_entries("just a string")
            assert False, "Should have raised ValueError"
        except ValueError:
            pass

    def test_translations_key_not_a_list_falls_through(self):
        """translations key with non-list value should fall through to segments or raise."""
        try:
            extract_entries({"translations": "not a list"})
            assert False, "Should have raised ValueError"
        except ValueError:
            pass


# ===========================================================================
# 3. detect_format
# ===========================================================================

class TestDetectFormat:

    def test_canonical(self):
        data = {"translations": [{"id": "a", "translated_text": "x"}]}
        assert detect_format(data) == "canonical"

    def test_translations_with_text_field(self):
        data = {"translations": [{"id": "a", "text": "x"}]}
        assert detect_format(data) == "translations_with_text_field"

    def test_batch_segments(self):
        data = {"batch": 1, "segments": [{"id": "a", "text": "x"}]}
        assert detect_format(data) == "batch_segments"

    def test_segments_only(self):
        data = {"segments": [{"id": "a", "text": "x"}]}
        assert detect_format(data) == "segments_only"

    def test_bare_array(self):
        data = [{"id": "a", "translated_text": "x"}]
        assert detect_format(data) == "bare_array"

    def test_unknown(self):
        assert detect_format({"random_key": 123}) == "unknown"

    # Finding 7: edge cases
    def test_bare_array_empty(self):
        assert detect_format([]) == "bare_array_empty"

    def test_bare_array_no_text_fields(self):
        """Finding 7: entries without text fields."""
        data = [{"id": "x", "random": "y"}]
        assert detect_format(data) == "bare_array_no_text_fields"

    def test_translations_empty(self):
        """Finding 7: translations key with empty list."""
        assert detect_format({"translations": []}) == "translations_empty"

    def test_translations_no_text_fields(self):
        """Finding 7: translations entries missing both text fields."""
        data = {"translations": [{"id": "x", "bogus": "y"}]}
        assert detect_format(data) == "translations_no_text_fields"

    def test_segments_empty(self):
        assert detect_format({"segments": []}) == "segments_only_empty"

    def test_batch_segments_no_text_fields(self):
        data = {"batch": 1, "segments": [{"id": "x"}]}
        assert detect_format(data) == "batch_segments_no_text_fields"


# ===========================================================================
# 4. normalize_batch
# ===========================================================================

class TestNormalizeBatch:

    def test_canonical_passthrough(self):
        data = {"translations": [{"id": "a", "translated_text": "x"}]}
        result = normalize_batch(data)
        assert result == [{"id": "a", "translated_text": "x"}]

    def test_segments_normalized(self):
        data = {"segments": [{"id": "a", "text": "hello"}]}
        result = normalize_batch(data)
        assert result == [{"id": "a", "translated_text": "hello"}]

    def test_bare_array_with_nulls(self):
        """Combined Finding 1 + variant d."""
        data = [{"id": "a", "translated_text": None}, {"id": "b", "text": None}]
        result = normalize_batch(data)
        assert all(e["translated_text"] == "" for e in result)

    def test_empty_input(self):
        data = {"translations": []}
        result = normalize_batch(data)
        assert result == []

    def test_mixed_dict_and_non_dict_entries(self):
        """RED 1: batch with mixed types must not crash, non-dicts get defaults."""
        data = {"translations": [42, "string", {"id": "a", "translated_text": "x"}, ["a", "b"]]}
        result, stderr = capture_stderr(normalize_batch, data)
        assert len(result) == 4
        # First two are non-dict -> empty defaults
        assert result[0] == {"id": "", "translated_text": ""}
        assert result[1] == {"id": "", "translated_text": ""}
        # Third is a real entry
        assert result[2] == {"id": "a", "translated_text": "x"}
        # Fourth is list -> empty defaults
        assert result[3] == {"id": "", "translated_text": ""}
        assert stderr.count("non-dict") == 3


# ===========================================================================
# 5. find_batch_files
# ===========================================================================

class TestFindBatchFiles:

    def test_finds_numbered_batches(self):
        with ManagedWorkspace({
            "batch-1.json": {"translations": []},
            "batch-2.json": {"translations": []},
            "batch-10.json": {"translations": []},
        }) as ws:
            result = find_batch_files(ws)
            assert [num for num, _ in result] == [1, 2, 10]

    def test_excludes_input_and_recovery(self):
        with ManagedWorkspace({
            "batch-1.json": {"translations": []},
            "batch-input-1.json": [{"id": "a", "text": "hi"}],
            "batch-recovery-1.json": {"translations": []},
        }) as ws:
            result = find_batch_files(ws)
            assert len(result) == 1
            assert result[0][0] == 1

    def test_empty_workspace(self):
        with ManagedWorkspace({}) as ws:
            assert find_batch_files(ws) == []

    def test_sorted_order(self):
        with ManagedWorkspace({
            "batch-3.json": {"translations": []},
            "batch-1.json": {"translations": []},
            "batch-2.json": {"translations": []},
        }) as ws:
            result = find_batch_files(ws)
            assert [num for num, _ in result] == [1, 2, 3]


# ===========================================================================
# 6. load_and_normalize_batch
# ===========================================================================

class TestLoadAndNormalizeBatch:

    def test_loads_canonical(self):
        with ManagedWorkspace({
            "batch-1.json": {"translations": [{"id": "a", "translated_text": "hello"}]},
        }) as ws:
            entries, fmt = load_and_normalize_batch(ws / "batch-1.json")
            assert len(entries) == 1
            assert fmt == "canonical"

    def test_malformed_json_raises_friendly(self):
        """Finding 3: user-friendly error on bad JSON."""
        with ManagedWorkspace({"batch-1.json": "{broken!!!"}) as ws:
            try:
                load_and_normalize_batch(ws / "batch-1.json")
                assert False, "Should have raised JSONDecodeError"
            except json.JSONDecodeError as e:
                assert "Malformed JSON" in e.msg
                assert "batch-1.json" in e.msg

    def test_loads_segments_variant(self):
        with ManagedWorkspace({
            "batch-1.json": {"segments": [{"id": "a", "text": "x"}]},
        }) as ws:
            entries, fmt = load_and_normalize_batch(ws / "batch-1.json")
            assert entries == [{"id": "a", "translated_text": "x"}]
            assert fmt == "segments_only"

    def test_unrecognizable_format_raises_valueerror(self):
        """RED 2: unrecognizable format must raise ValueError, not crash."""
        with ManagedWorkspace({
            "batch-1.json": {"bogus": "data"},
        }) as ws:
            try:
                load_and_normalize_batch(ws / "batch-1.json")
                assert False, "Should have raised ValueError"
            except ValueError as e:
                assert "Unrecognizable" in str(e)

    def test_utf8_bom_file_loads(self):
        """REC 5: BOM-prefixed JSON file parses correctly."""
        bom_content = b'\xef\xbb\xbf{"translations": [{"id": "x", "translated_text": "Y"}]}'
        with ManagedWorkspace({"batch-1.json": bom_content}) as ws:
            entries, fmt = load_and_normalize_batch(ws / "batch-1.json")
            assert len(entries) == 1
            assert entries[0] == {"id": "x", "translated_text": "Y"}


# ===========================================================================
# 7. merge_all_batches
# ===========================================================================

class TestMergeAllBatches:

    def test_merge_two_batches(self):
        b1 = [{"id": "a", "translated_text": "x"}]
        b2 = [{"id": "b", "translated_text": "y"}]
        result = merge_all_batches([b1, b2])
        assert len(result["translations"]) == 2

    def test_empty_batches(self):
        result = merge_all_batches([[], []])
        assert result == {"translations": []}

    def test_duplicate_ids_warned(self):
        """Finding 4: duplicate IDs detected and warned."""
        b1 = [{"id": "dup", "translated_text": "a"}]
        b2 = [{"id": "dup", "translated_text": "b"}]
        result, stderr = capture_stderr(merge_all_batches, [b1, b2])
        assert "duplicate" in stderr.lower()
        # All entries kept (no silent dedup)
        assert len(result["translations"]) == 2

    def test_no_warning_without_duplicates(self):
        b1 = [{"id": "a", "translated_text": "x"}]
        b2 = [{"id": "b", "translated_text": "y"}]
        result, stderr = capture_stderr(merge_all_batches, [b1, b2])
        assert "duplicate" not in stderr.lower()

    def test_duplicate_within_same_batch(self):
        b1 = [{"id": "x", "translated_text": "a"}, {"id": "x", "translated_text": "b"}]
        result, stderr = capture_stderr(merge_all_batches, [b1])
        assert "duplicate" in stderr.lower()


# ===========================================================================
# 8. validate_count
# ===========================================================================

class TestValidateCount:

    def test_match(self):
        trans = {"translations": [{"id": "a"}, {"id": "b"}]}
        is_valid, msg = validate_count(trans, 2)
        assert is_valid is True
        assert "OK" in msg

    def test_mismatch(self):
        trans = {"translations": [{"id": "a"}]}
        is_valid, msg = validate_count(trans, 3)
        assert is_valid is False
        assert "MISMATCH" in msg

    def test_zero_expected_zero_actual(self):
        trans = {"translations": []}
        is_valid, msg = validate_count(trans, 0)
        assert is_valid is True


# ===========================================================================
# 9. load_expected_total
# ===========================================================================

class TestLoadExpectedTotal:

    def test_returns_total(self):
        with ManagedWorkspace({
            "texts.json": {"total_segments": 42, "segments": []},
        }) as ws:
            assert load_expected_total(ws) == 42

    def test_missing_texts_json(self):
        with ManagedWorkspace({}) as ws:
            assert load_expected_total(ws) is None

    def test_missing_total_segments_key(self):
        with ManagedWorkspace({"texts.json": {"segments": []}}) as ws:
            assert load_expected_total(ws) is None


# ===========================================================================
# 10. End-to-end: main() exit codes
# ===========================================================================

class TestMainExitCodes:

    SCRIPT = str(Path(__file__).resolve().parent / "normalize_translations.py")

    def test_exit_0_on_match(self):
        ws = make_workspace({
            "texts.json": {"total_segments": 2, "segments": [
                {"id": "a", "text": "x"},
                {"id": "b", "text": "y"},
            ]},
            "batch-1.json": {"translations": [
                {"id": "a", "translated_text": "X"},
                {"id": "b", "translated_text": "Y"},
            ]},
        })
        result = subprocess.run(
            [sys.executable, self.SCRIPT, "--workspace", str(ws)],
            capture_output=True, text=True,
        )
        assert result.returncode == 0, f"Expected 0, got {result.returncode}\nstderr: {result.stderr}"

    def test_exit_nonzero_on_mismatch(self):
        """Finding 5: non-zero exit on validation mismatch."""
        ws = make_workspace({
            "texts.json": {"total_segments": 3, "segments": [
                {"id": "a", "text": "x"},
                {"id": "b", "text": "y"},
                {"id": "c", "text": "z"},
            ]},
            "batch-1.json": {"translations": [
                {"id": "a", "translated_text": "X"},
            ]},
        })
        result = subprocess.run(
            [sys.executable, self.SCRIPT, "--workspace", str(ws)],
            capture_output=True, text=True,
        )
        assert result.returncode != 0, f"Expected non-zero exit on mismatch"

    def test_exit_nonzero_on_malformed_json(self):
        """Finding 3+5: malformed JSON -> non-zero exit."""
        ws = make_workspace({
            "texts.json": {"total_segments": 1, "segments": [{"id": "a", "text": "x"}]},
            "batch-1.json": "NOT JSON AT ALL {{{{",
        })
        result = subprocess.run(
            [sys.executable, self.SCRIPT, "--workspace", str(ws)],
            capture_output=True, text=True,
        )
        assert result.returncode != 0

    def test_exit_0_without_texts_json(self):
        """No texts.json -> validation skipped -> exit 0 if no other errors."""
        ws = make_workspace({
            "batch-1.json": {"translations": [{"id": "a", "translated_text": "X"}]},
        })
        result = subprocess.run(
            [sys.executable, self.SCRIPT, "--workspace", str(ws)],
            capture_output=True, text=True,
        )
        assert result.returncode == 0

    def test_no_batch_files_exits_nonzero(self):
        ws = make_workspace({"texts.json": {"total_segments": 0, "segments": []}})
        result = subprocess.run(
            [sys.executable, self.SCRIPT, "--workspace", str(ws)],
            capture_output=True, text=True,
        )
        assert result.returncode != 0

    def test_output_written_even_on_mismatch(self):
        """Output file should exist even when validation fails."""
        ws = make_workspace({
            "texts.json": {"total_segments": 5, "segments": []},
            "batch-1.json": {"translations": [{"id": "a", "translated_text": "X"}]},
        })
        subprocess.run(
            [sys.executable, self.SCRIPT, "--workspace", str(ws)],
            capture_output=True, text=True,
        )
        out_path = ws / "translations.json"
        assert out_path.exists(), "translations.json should be written even on mismatch"
        data = json.loads(out_path.read_text(encoding="utf-8"))
        assert len(data["translations"]) == 1

    def test_variant_segments_format_works_e2e(self):
        """RED 3: variant format that used to crash old T4 now works via normalize_translations.py."""
        ws = make_workspace({
            "texts.json": {"total_segments": 2, "segments": [
                {"id": "a", "text": "x"},
                {"id": "b", "text": "y"},
            ]},
            "batch-1.json": {"segments": [
                {"id": "a", "text": "Translated X"},
                {"id": "b", "text": "Translated Y"},
            ]},
        })
        result = subprocess.run(
            [sys.executable, self.SCRIPT, "--workspace", str(ws)],
            capture_output=True, text=True,
        )
        assert result.returncode == 0, f"Expected 0, got {result.returncode}\nstderr: {result.stderr}"
        data = json.loads((ws / "translations.json").read_text(encoding="utf-8"))
        assert len(data["translations"]) == 2
        assert data["translations"][0]["translated_text"] == "Translated X"

    def test_batch_segments_variant_works_e2e(self):
        """RED 3: batch+segments variant that old T4 couldn't handle."""
        ws = make_workspace({
            "batch-1.json": {"batch": 1, "segments": [
                {"id": "a", "text": "hello"},
            ]},
        })
        result = subprocess.run(
            [sys.executable, self.SCRIPT, "--workspace", str(ws)],
            capture_output=True, text=True,
        )
        assert result.returncode == 0, f"Expected 0, got {result.returncode}\nstderr: {result.stderr}"
        data = json.loads((ws / "translations.json").read_text(encoding="utf-8"))
        assert data["translations"][0] == {"id": "a", "translated_text": "hello"}

    def test_bare_array_variant_works_e2e(self):
        """RED 3: bare array variant that old T4 couldn't handle."""
        ws = make_workspace({
            "batch-1.json": [{"id": "a", "translated_text": "X"}],
        })
        result = subprocess.run(
            [sys.executable, self.SCRIPT, "--workspace", str(ws)],
            capture_output=True, text=True,
        )
        assert result.returncode == 0, f"Expected 0\nstderr: {result.stderr}"

    def test_unrecognizable_format_handled_gracefully(self):
        """RED 2: ValueError from unrecognizable format should not crash main()."""
        ws = make_workspace({
            "batch-1.json": {"bogus_key": "no translations here"},
        })
        result = subprocess.run(
            [sys.executable, self.SCRIPT, "--workspace", str(ws)],
            capture_output=True, text=True,
        )
        # Should exit non-zero but NOT with a traceback
        assert result.returncode != 0
        assert "Unrecognizable" in result.stderr or "Error" in result.stderr
        # No Python traceback
        assert "Traceback" not in result.stderr, f"Got traceback:\n{result.stderr}"

    def test_non_dict_entries_handled_gracefully_e2e(self):
        """RED 1+2: batch with non-dict entries doesn't crash, produces output."""
        ws = make_workspace({
            "batch-1.json": {"translations": [42, "string", {"id": "a", "translated_text": "X"}]},
        })
        result = subprocess.run(
            [sys.executable, self.SCRIPT, "--workspace", str(ws)],
            capture_output=True, text=True,
        )
        assert result.returncode == 0, f"Expected 0\nstderr: {result.stderr}"
        data = json.loads((ws / "translations.json").read_text(encoding="utf-8"))
        assert len(data["translations"]) == 3
        # Non-dict entries get empty defaults
        assert data["translations"][0] == {"id": "", "translated_text": ""}
        # Real entry preserved
        assert data["translations"][2] == {"id": "a", "translated_text": "X"}

    def test_utf8_bom_handled(self):
        """REC 5: BOM-prefixed batch file should parse correctly."""
        bom_json = b'\xef\xbb\xbf{"translations": [{"id": "a", "translated_text": "X"}]}'
        ws = make_workspace({
            "batch-1.json": bom_json,
        })
        result = subprocess.run(
            [sys.executable, self.SCRIPT, "--workspace", str(ws)],
            capture_output=True, text=True,
        )
        assert result.returncode == 0, f"Expected 0 on BOM file\nstderr: {result.stderr}"
        data = json.loads((ws / "translations.json").read_text(encoding="utf-8"))
        assert data["translations"][0]["id"] == "a"


# ===========================================================================
# 11. Integration: normalize_translations.py referenced in pipeline
# ===========================================================================

class TestPipelineIntegration:

    def test_script_referenced_in_translation_prompt(self):
        """Finding 6: script must be referenced in translation-prompt.md."""
        prompt_path = Path(__file__).resolve().parent.parent / "translation-prompt.md"
        assert prompt_path.exists(), f"translation-prompt.md not found at {prompt_path}"
        content = prompt_path.read_text(encoding="utf-8")
        assert "normalize_translations.py" in content, \
            "normalize_translations.py not referenced in translation-prompt.md"

    def test_script_referenced_in_skill_md(self):
        """Finding 6: script must be in SKILL.md validation block."""
        skill_path = Path(__file__).resolve().parent.parent / "SKILL.md"
        assert skill_path.exists(), f"SKILL.md not found at {skill_path}"
        content = skill_path.read_text(encoding="utf-8")
        assert "normalize_translations.py" in content, \
            "normalize_translations.py not referenced in SKILL.md"

    def test_t4_uses_normalize_script_not_inline_python(self):
        """RED 3: T4 must call normalize_translations.py, not inline data['translations']."""
        prompt_path = Path(__file__).resolve().parent.parent / "translation-prompt.md"
        content = prompt_path.read_text(encoding="utf-8")
        # T4 section should invoke the script
        assert "uv run" in content and "normalize_translations.py" in content
        # The old fragile inline merge must be gone
        assert "items.extend(data[\"translations\"])" not in content, \
            "Old fragile T4 inline merge code still present"
        # No T4.5 section should exist
        assert "T4.5" not in content, "T4.5 should not exist — normalize IS T4 now"

    def test_translation_prompt_steps_are_t1_through_t5(self):
        """RED 3: Steps should be T1, T2, T3, T4, T5 with no gaps or halves."""
        prompt_path = Path(__file__).resolve().parent.parent / "translation-prompt.md"
        content = prompt_path.read_text(encoding="utf-8")
        import re
        headings = re.findall(r"^## (T\S+)", content, re.MULTILINE)
        # Extract just the step identifiers
        step_ids = [h.rstrip(":") for h in headings]
        # Should start with T1-T5, no T4.5
        assert "T1" in step_ids[0], f"First step should be T1, got {step_ids}"
        for h in step_ids:
            assert "." not in h, f"No fractional steps allowed, found {h}"


# ===========================================================================
# Runner
# ===========================================================================

def run_all():
    """Simple test runner that doesn't need pytest."""
    import traceback

    test_classes = [
        TestNormalizeEntry,
        TestExtractEntries,
        TestDetectFormat,
        TestNormalizeBatch,
        TestFindBatchFiles,
        TestLoadAndNormalizeBatch,
        TestMergeAllBatches,
        TestValidateCount,
        TestLoadExpectedTotal,
        TestMainExitCodes,
        TestPipelineIntegration,
    ]

    passed = 0
    failed = 0
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

    print(f"\n{'='*60}")
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
