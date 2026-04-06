#!/usr/bin/env python3
"""normalize_translations.py - Normalize variant batch translation formats to canonical form.

Weak AI models (e.g., qwen3.5-122b-a10b) produce batch files in non-standard formats.
This script normalizes all known variants into the canonical format expected by
apply_dsl_translations.py:

    {"translations": [{"id": "...", "translated_text": "..."}]}

Known variant formats handled:
  (a) {"translations": [{"id": "...", "translated_text": "..."}]}  -- canonical, pass through
  (b) {"segments": [{"id": "...", "text": "..."}]}                 -- segments key, text field
  (c) {"batch": N, "segments": [{"id": "...", "text": "..."}]}     -- batch number + segments
  (d) [{"id": "...", "translated_text": "..."}]                    -- bare array at top level
  (e) {"translations": [{"id": "...", "text": "..."}]}             -- correct key, wrong field

Usage:
    uv run normalize_translations.py --workspace PATH

Reads batch-*.json and texts.json from workspace, writes translations.json.
"""

import argparse
import json
import re
import sys
from pathlib import Path


# ---------------------------------------------------------------------------
# Pure functions: format detection and normalization
# ---------------------------------------------------------------------------

# Matches batch-N.json but NOT batch-input-N.json or batch-recovery-N.json
BATCH_FILE_PATTERN = re.compile(r"^batch-(\d+)\.json$")


def sanitize_translated_text(text: str) -> str:
    """Strip markup artifacts that weak LLMs inject into translated text.

    Removes:
    - XML/HTML tags (preserving inner text)
    - Markdown table separator lines (e.g., | --- | --- |)
    - Markdown bold/italic wrappers (**text** -> text, *text* -> text)
      but NOT standalone * in math contexts (e.g., 2*3)
    - Markdown heading markers (## Heading -> Heading)
    - Code fences (defense-in-depth, also handled by extract_dsl_texts.py)

    Pure function: (str) -> str, no side effects.
    """
    if not text:
        return text

    # 1. Remove code fences (defense-in-depth)
    cleaned = re.sub(r"```(?:markdown|xml|json|html)?\s*```", "", text)
    cleaned = re.sub(r"```(?:markdown|xml|json|html)?\s*\n?\s*```", "", cleaned)
    cleaned = re.sub(r"^```(?:markdown|xml|json|html)?$", "", cleaned, flags=re.MULTILINE)

    # 2. Remove XML/HTML tags, keep inner text
    # Matches <tag>, <tag attr="val">, </tag>, <tag/>, <ns:tag>
    cleaned = re.sub(r"</?[a-zA-Z][a-zA-Z0-9:]*(?:\s[^>]*)?\s*/?>", "", cleaned)

    # 3. Remove markdown table separator lines: | --- | --- | or |:---:|
    cleaned = re.sub(r"^\s*\|[\s:|-]+\|\s*$", "", cleaned, flags=re.MULTILINE)

    # 4. Remove markdown bold wrappers: **text** -> text
    cleaned = re.sub(r"\*\*(.+?)\*\*", r"\1", cleaned)

    # 5. Remove markdown italic wrappers: *text* -> text
    # But NOT standalone * in math (e.g., 2*3) — require space or start/end boundary
    cleaned = re.sub(r"(?<!\w)\*([^\s*][^*]*?)\*(?!\w)", r"\1", cleaned)

    # 6. Remove markdown heading markers at start of text
    cleaned = re.sub(r"^#+\s+", "", cleaned, flags=re.MULTILINE)

    # 7. Collapse multiple blank lines left by removals
    cleaned = re.sub(r"\n{3,}", "\n\n", cleaned)

    return cleaned.strip()


def find_batch_files(workspace: Path) -> list[tuple[int, Path]]:
    """Find and sort batch-N.json files, excluding batch-input-* and batch-recovery-*.

    Returns list of (batch_number, path) tuples sorted by batch number.
    """
    results = []
    for path in workspace.iterdir():
        match = BATCH_FILE_PATTERN.match(path.name)
        if match:
            batch_num = int(match.group(1))
            results.append((batch_num, path))
    return sorted(results, key=lambda pair: pair[0])


def normalize_entry(entry, source_file: str = "<unknown>") -> dict:
    """Normalize a single translation entry to canonical form.

    Canonical: {"id": "...", "translated_text": "..."}

    Handles:
    - entries with "text" instead of "translated_text"
    - None/null coerced to empty string
    - non-dict entries (int, str, list) -> type guard warning + empty defaults
    - whitespace-only id coerced to ""
    """
    # Type guard: non-dict entries from weak models
    if not isinstance(entry, dict):
        print(
            f"WARNING: non-dict entry in {source_file}: {entry!r} (type={type(entry).__name__})",
            file=sys.stderr,
        )
        return {"id": "", "translated_text": ""}

    seg_id = entry.get("id")

    # Warn on missing, empty, or whitespace-only id; coerce to ""
    if seg_id is None or (isinstance(seg_id, str) and seg_id.strip() == ""):
        print(
            f"WARNING: entry with missing or empty 'id' in {source_file}: {entry!r}",
            file=sys.stderr,
        )
        # Coerce None and whitespace-only to ""
        seg_id = ""

    # Prefer "translated_text", fall back to "text"
    if "translated_text" in entry:
        translated = entry["translated_text"]
    else:
        translated = entry.get("text", "")

    # Coerce None (JSON null) to empty string
    if translated is None:
        translated = ""

    # Sanitize: strip XML tags, markdown artifacts from weak LLM output
    translated = sanitize_translated_text(translated)

    return {"id": seg_id, "translated_text": translated}


def extract_entries(data) -> list[dict]:
    """Extract the list of translation entries from any known format variant.

    Handles:
      - dict with "translations" key (canonical or variant e)
      - dict with "segments" key (variants b, c)
      - bare list (variant d)

    Returns list of raw entry dicts (not yet normalized).
    Raises ValueError if format is unrecognizable.
    """
    if isinstance(data, list):
        # Variant (d): bare array
        return data

    if isinstance(data, dict):
        if "translations" in data:
            # Variant (a) or (e)
            entries = data["translations"]
            if isinstance(entries, list):
                return entries
        if "segments" in data:
            # Variant (b) or (c)
            entries = data["segments"]
            if isinstance(entries, list):
                return entries

    raise ValueError(
        f"Unrecognizable batch format. Top-level keys: "
        f"{list(data.keys()) if isinstance(data, dict) else type(data).__name__}"
    )


def detect_format(data) -> str:
    """Detect which variant format a batch file uses. For diagnostics only.

    Returns a descriptive format name. Handles edge cases:
    - Empty lists -> "bare_array_empty"
    - Entries without text fields -> appends "_no_text_fields"
    """
    if isinstance(data, list):
        if len(data) == 0:
            return "bare_array_empty"
        first = data[0]
        if not isinstance(first, dict):
            return "bare_array_non_dict_entries"
        has_text_field = "translated_text" in first or "text" in first
        if not has_text_field:
            return "bare_array_no_text_fields"
        return "bare_array"

    if isinstance(data, dict):
        has_translations = "translations" in data
        has_segments = "segments" in data
        has_batch = "batch" in data

        if has_translations:
            entries = data.get("translations", [])
            if not isinstance(entries, list) or len(entries) == 0:
                return "translations_empty"
            first = entries[0]
            if not isinstance(first, dict):
                return "translations_non_dict_entries"
            if "translated_text" in first:
                return "canonical"
            if "text" in first:
                return "translations_with_text_field"
            return "translations_no_text_fields"

        if has_segments:
            entries = data.get("segments", [])
            suffix = ""
            if not isinstance(entries, list) or len(entries) == 0:
                suffix = "_empty"
            elif not isinstance(entries[0], dict):
                suffix = "_non_dict_entries"
            elif "text" not in entries[0] and "translated_text" not in entries[0]:
                suffix = "_no_text_fields"

            if has_batch:
                return f"batch_segments{suffix}"
            return f"segments_only{suffix}"

    return "unknown"


def normalize_batch(data, source_file: str = "<unknown>") -> list[dict]:
    """Normalize a single batch's data to a list of canonical entries.

    Pipeline: extract_entries -> map normalize_entry
    """
    raw_entries = extract_entries(data)
    return [normalize_entry(entry, source_file=source_file) for entry in raw_entries]


def load_and_normalize_batch(path: Path) -> tuple[list[dict], str]:
    """Load a batch file and normalize it.

    Returns (normalized_entries, detected_format).
    Raises json.JSONDecodeError with a user-friendly wrapper on malformed JSON.
    """
    # Finding 3: Catch JSONDecodeError and provide user-friendly message
    try:
        with open(path, encoding="utf-8-sig") as f:
            data = json.load(f)
    except json.JSONDecodeError as exc:
        raise json.JSONDecodeError(
            f"Malformed JSON in {path.name}: {exc.msg}",
            exc.doc,
            exc.pos,
        ) from exc

    fmt = detect_format(data)
    entries = normalize_batch(data, source_file=str(path))
    return entries, fmt


def merge_all_batches(batch_entries: list[list[dict]]) -> dict:
    """Merge normalized batch entry lists into canonical translations dict.

    Finding 4: Detects and warns about duplicate IDs across batches.

    Returns {"translations": [...all entries...]}
    """
    all_entries: list[dict] = []
    seen_ids: dict[str, int] = {}  # id -> batch index (0-based)
    duplicate_count = 0

    for batch_idx, entries in enumerate(batch_entries):
        for entry in entries:
            entry_id = entry["id"]
            if entry_id in seen_ids:
                duplicate_count += 1
                print(
                    f"WARNING: duplicate id '{entry_id}' in batch {batch_idx + 1} "
                    f"(first seen in batch {seen_ids[entry_id] + 1})",
                    file=sys.stderr,
                )
            else:
                seen_ids[entry_id] = batch_idx
            all_entries.append(entry)

    if duplicate_count > 0:
        print(
            f"WARNING: {duplicate_count} duplicate ID(s) found across batches",
            file=sys.stderr,
        )

    return {"translations": all_entries}


def validate_count(translations: dict, expected_total: int) -> tuple[bool, str]:
    """Validate that translation count matches expected total.

    Returns (is_valid, message).
    """
    actual = len(translations["translations"])
    if actual == expected_total:
        return True, f"OK: {actual}/{expected_total} segments"
    return False, f"MISMATCH: {actual}/{expected_total} segments"


def load_expected_total(workspace: Path) -> int | None:
    """Load total_segments from texts.json if it exists."""
    texts_path = workspace / "texts.json"
    if not texts_path.exists():
        return None
    with open(texts_path, encoding="utf-8-sig") as f:
        data = json.load(f)
    return data.get("total_segments")


# ---------------------------------------------------------------------------
# Main: thin CLI shell over pure functions
# ---------------------------------------------------------------------------


def main():
    parser = argparse.ArgumentParser(
        description="Normalize variant batch translation formats to canonical form"
    )
    parser.add_argument(
        "--workspace", required=True,
        help="Workspace directory containing batch-*.json and texts.json"
    )
    args = parser.parse_args()

    workspace = Path(args.workspace)
    if not workspace.is_dir():
        print(f"Error: workspace not found: {workspace}", file=sys.stderr)
        sys.exit(1)

    # Find batch files
    batch_files = find_batch_files(workspace)
    if not batch_files:
        print("Error: no batch-N.json files found in workspace", file=sys.stderr)
        sys.exit(1)

    print(f"Found {len(batch_files)} batch file(s)")

    # Load and normalize each batch
    all_batch_entries = []
    has_errors = False
    for batch_num, path in batch_files:
        try:
            entries, fmt = load_and_normalize_batch(path)
        except json.JSONDecodeError as exc:
            print(f"Error: {exc.msg}", file=sys.stderr)
            has_errors = True
            continue
        except (ValueError, AttributeError, TypeError) as exc:
            print(f"Error in {path.name}: {exc}", file=sys.stderr)
            has_errors = True
            continue
        print(f"  batch-{batch_num}.json: {len(entries)} entries, format={fmt}")
        all_batch_entries.append(entries)

    if has_errors and not all_batch_entries:
        print("Error: all batch files had JSON errors, cannot proceed", file=sys.stderr)
        sys.exit(1)

    # Merge
    translations = merge_all_batches(all_batch_entries)
    total = len(translations["translations"])
    print(f"Total normalized entries: {total}")

    # Validate against texts.json if available
    exit_code = 0
    expected = load_expected_total(workspace)
    if expected is not None:
        is_valid, msg = validate_count(translations, expected)
        print(f"Validation: {msg}")
        if not is_valid:
            # Finding 5: Exit non-zero on validation mismatch
            print("Error: count mismatch — output written but exit code will be non-zero", file=sys.stderr)
            exit_code = 1
    else:
        print("Validation: skipped (no texts.json found)")

    if has_errors:
        # Finding 5: Also non-zero if any batch files had JSON errors
        exit_code = 1

    # Write output
    out_path = workspace / "translations.json"
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(translations, f, ensure_ascii=False, indent=2)

    print(f"Wrote {total} translations to {out_path}")
    sys.exit(exit_code)


if __name__ == "__main__":
    main()
