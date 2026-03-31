#!/usr/bin/env python3
"""apply_dsl_translations.py - Apply translated text back into XML DSL files.

Usage:
    uv run --with lxml apply_dsl_translations.py --workspace PATH --pages N \
        --translations translations.json --output-dir dsl-translated

Reads translations JSON, deep-copies each page XML, applies text replacements
by matching segment IDs, and writes to $WORKSPACE/{output-dir}/page-{N}.xml.
"""

import argparse
import copy
import json
import re
import sys
from pathlib import Path

from lxml import etree


# ---------------------------------------------------------------------------
# ID parsing: pure functions
# ---------------------------------------------------------------------------

# Pattern for path steps like "heading[0]", "run[2]", "cell[1]"
STEP_PATTERN = re.compile(r"^([a-zA-Z][a-zA-Z0-9_-]*)\[(\d+)\]$")


def parse_segment_id(segment_id):
    """Parse a segment ID into (page_num, path_steps).

    Example: "p1:heading[0]/run[0]" -> (1, [("heading", 0), ("run", 0)])
    Example: "p2:page-header/paragraph[0]/run[0]" -> (2, [("page-header", None), ("paragraph", 0), ("run", 0)])

    Returns:
        (page_num, path_steps) where path_steps is a list of (tag, index_or_None)
        Returns (None, None) on parse failure.
    """
    if ":" not in segment_id:
        return None, None

    page_part, path_part = segment_id.split(":", 1)

    # Parse page number
    if not page_part.startswith("p"):
        return None, None
    try:
        page_num = int(page_part[1:])
    except ValueError:
        return None, None

    # Parse path steps
    raw_steps = path_part.split("/")
    path_steps = []
    for raw in raw_steps:
        match = STEP_PATTERN.match(raw)
        if match:
            tag = match.group(1)
            index = int(match.group(2))
            path_steps.append((tag, index))
        elif raw in ("page-header", "page-footer"):
            # Singleton elements without index
            path_steps.append((raw, None))
        else:
            return None, None

    if not path_steps:
        return None, None

    return page_num, path_steps


def navigate_to_element(root, path_steps):
    """Navigate from root element to the target element via path_steps.

    Args:
        root: lxml root element (<page>)
        path_steps: list of (tag, index_or_None) tuples

    Returns:
        The target lxml element, or None if not found.
    """
    current = root
    for tag, index in path_steps:
        if index is None:
            # Singleton: find first child with this tag
            found = None
            for child in current:
                if child.tag == tag:
                    found = child
                    break
            if found is None:
                return None
            current = found
        else:
            # Find the Nth child with this tag
            count = 0
            found = None
            for child in current:
                if child.tag == tag:
                    if count == index:
                        found = child
                        break
                    count += 1
            if found is None:
                return None
            current = found
    return current


def apply_translation_to_element(elem, translated_text):
    """Apply translated text to an element.

    For <run> elements: sets .text
    For <cell> elements (no run children): sets .text
    """
    elem.text = translated_text


# ---------------------------------------------------------------------------
# Page-level processing
# ---------------------------------------------------------------------------


def build_translation_map(translations_data):
    """Build a dict from segment ID to translated text.

    Args:
        translations_data: dict with "translations" key containing list of
            {"id": "...", "translated_text": "..."} dicts

    Returns:
        dict mapping id -> translated_text
    """
    result = {}
    for entry in translations_data.get("translations", []):
        seg_id = entry.get("id", "")
        translated = entry.get("translated_text", "")
        if seg_id and translated:
            result[seg_id] = translated
    return result


def apply_translations_to_page(xml_path, page_num, translation_map):
    """Apply translations to a single page XML file.

    Deep-copies the XML tree, applies matching translations, returns
    the modified tree and stats.

    Args:
        xml_path: Path to the source page XML
        page_num: 1-based page number
        translation_map: dict mapping segment_id -> translated_text

    Returns:
        (modified_tree, applied_count, skipped_count)
    """
    tree = etree.parse(str(xml_path))
    root = copy.deepcopy(tree.getroot())

    applied = 0
    skipped = 0

    # Filter translations for this page
    page_prefix = f"p{page_num}:"
    page_translations = {
        k: v for k, v in translation_map.items() if k.startswith(page_prefix)
    }

    for seg_id, translated_text in page_translations.items():
        _, path_steps = parse_segment_id(seg_id)
        if path_steps is None:
            print(f"  Warning: cannot parse ID '{seg_id}', skipping",
                  file=sys.stderr)
            skipped += 1
            continue

        target = navigate_to_element(root, path_steps)
        if target is None:
            print(f"  Warning: element not found for '{seg_id}', skipping",
                  file=sys.stderr)
            skipped += 1
            continue

        apply_translation_to_element(target, translated_text)
        applied += 1

    return root, applied, skipped


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def main():
    parser = argparse.ArgumentParser(
        description="Apply translated text back into XML DSL files"
    )
    parser.add_argument("--workspace", required=True, help="Workspace directory path")
    parser.add_argument("--pages", required=True, type=int, help="Total page count")
    parser.add_argument("--translations", required=True,
                        help="Path to translations JSON file")
    parser.add_argument("--output-dir", required=True,
                        help="Output directory name under workspace")
    args = parser.parse_args()

    workspace = Path(args.workspace)
    total_pages = args.pages

    # Read translations
    translations_path = Path(args.translations)
    if not translations_path.exists():
        print(f"Error: translations file not found: {translations_path}",
              file=sys.stderr)
        sys.exit(1)

    with open(translations_path) as f:
        translations_data = json.load(f)

    translation_map = build_translation_map(translations_data)
    print(f"Loaded {len(translation_map)} translations")

    # Create output directory
    output_dir = workspace / args.output_dir
    output_dir.mkdir(parents=True, exist_ok=True)

    total_applied = 0
    total_skipped = 0

    for page_num in range(1, total_pages + 1):
        xml_path = workspace / "dsl" / f"page-{page_num}.xml"
        if not xml_path.exists():
            print(f"Warning: {xml_path} not found, skipping page {page_num}",
                  file=sys.stderr)
            continue

        root, applied, skipped = apply_translations_to_page(
            xml_path, page_num, translation_map
        )

        # Write output XML with proper formatting
        output_path = output_dir / f"page-{page_num}.xml"
        xml_str = etree.tostring(root, encoding="unicode", pretty_print=True)
        output_path.write_text(xml_str, encoding="utf-8")

        total_applied += applied
        total_skipped += skipped
        page_total = applied + skipped
        print(f"Applied {applied}/{page_total} translations for page {page_num}")

    print(f"Done. Total: {total_applied} applied, {total_skipped} skipped")


if __name__ == "__main__":
    main()
