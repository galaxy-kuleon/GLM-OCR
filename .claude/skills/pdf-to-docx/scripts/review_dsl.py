#!/usr/bin/env python3
"""review_dsl.py - Review XML DSL against PDF page images using Poe AI.

Usage:
    python review_dsl.py --workspace PATH --pages N [--poe-api-key KEY]

For documents with <=5 pages, sends ALL page images + ALL XML DSLs in a single
API call for cross-page consistency checking. For >5 pages or if multi-page mode
fails, falls back to per-page review.

Output: $WORKSPACE/dsl/review-page-{N}.json (per page)

If no issues found, the review file contains [].
API key is read from --poe-api-key argument or POE_API_KEY environment variable
or .env file.
"""

import argparse
import base64
import json
import os
import sys
from pathlib import Path

try:
    import requests
except ImportError:
    requests = None


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

POE_API_URL = "https://api.poe.com/v1/chat/completions"
POE_MODEL = "gemini-3-flash"

# Single-page review prompt (fallback)
SYSTEM_PROMPT = """You compare a PDF page image against an XML description of that page.
Find issues: missing text, wrong element order, incorrect styles, missing or incorrect text boxes/frames.
Also check: paragraph/heading background shading colors, table border styles (single vs double line),
text-level highlight/background colors, and font family (serif vs sans-serif).
Also check for: LaTeX/math artifacts that should have been cleaned ($...$, \\mathrm, etc.),
incorrect font family detection (serif text marked as sans or vice versa),
and incorrect paragraph spacing (space-before, space-after, line-spacing).

CRITICAL — BACKGROUND COLOR CHECKING:
1. Heading/paragraph background: If the image shows a heading or paragraph with a gray/colored background strip, the XML MUST have a bg-color attribute on that <paragraph>. Report "wrong_background" if missing.
2. Table row background: If the image shows alternating gray/white rows (zebra striping) in a table, EACH gray row's <cell> elements MUST have bg-color. Report "wrong_background" if missing.
3. Text-level highlight: If specific words have a small colored background behind them, the corresponding <run> MUST have text-bg-color. Report "missing_text_highlight" if missing.

Respond with JSON array of fixes. If no issues, respond with []."""

USER_PROMPT_TEMPLATE = """PDF page image is attached. Here is the XML DSL for this page:

{xml_content}

Compare the image to the XML. For each issue found, respond:
[
  {{"type": "missing_text", "description": "...", "after_region": N}},
  {{"type": "wrong_style", "region": N, "field": "bold", "expected": true}},
  {{"type": "wrong_order", "description": "region 3 should come before region 2"}},
  {{"type": "latex_artifact", "region": N, "description": "LaTeX markup not cleaned: $...$"}},
  {{"type": "font_mismatch", "region": N, "field": "font-family", "expected": "serif", "actual": "sans"}},
  {{"type": "spacing_issue", "region": N, "field": "space-before-pt", "expected": 12, "actual": 0}}
]
Types: missing_text, wrong_style, wrong_order, missing_image, extra_content,
missing_textframe (text in image has a bordered box but XML uses plain paragraph instead of text-frame),
wrong_textframe (text-frame exists but has wrong attributes, e.g. missing border),
wrong_background (paragraph/heading background shading color is incorrect or missing),
wrong_border_style (table border style is incorrect, e.g. should be double but is single),
missing_text_highlight (text should have a colored highlight/background behind it but doesn't),
latex_artifact (LaTeX/math markup like $...$, \\mathrm, ^{...} not properly cleaned from text),
font_mismatch (font-family attribute doesn't match actual font in image: serif vs sans-serif),
spacing_issue (paragraph spacing like space-before, space-after, or line-spacing is incorrect).
Also check: if headings/paragraphs have background shading that matches the original image,
if table border styles match (single line vs double line), if any text has a colored highlight/background,
if LaTeX artifacts have been properly cleaned from text content,
if font-family (serif/sans) matches the actual font in the image,
and if paragraph spacing values (space-before-pt, space-after-pt, line-spacing) are reasonable.

BACKGROUND COLOR CHECKLIST — compare the image against the XML carefully:
1. For each heading/paragraph in the image that has a gray or colored background strip: verify the corresponding <paragraph> in XML has bg-color="R,G,B". If missing, report type="wrong_background".
2. For each table with alternating gray/white rows: verify that gray rows have bg-color on their <cell> elements. If missing, report type="wrong_background".
3. For each word/character in the image that has a small colored background highlight: verify the corresponding <run> has text-bg-color="R,G,B". If missing, report type="missing_text_highlight".
Only output JSON array, nothing else. If everything matches, output []."""

# Multi-page review prompt
MULTI_PAGE_SYSTEM_PROMPT = """You compare multiple PDF page images against their XML descriptions.
Find issues on EACH page: missing text, wrong element order, incorrect styles, missing or incorrect text boxes/frames.
Also check: paragraph/heading background shading colors, table border styles (single vs double line),
text-level highlight/background colors, and font family (serif vs sans-serif).
Also check for: LaTeX/math artifacts that should have been cleaned ($...$, \\mathrm, etc.),
incorrect font family detection, and incorrect paragraph spacing.
Also check cross-page consistency: font sizes, colors, column widths, font families, and styles should be consistent across pages unless intentionally different.

CRITICAL — BACKGROUND COLOR CHECKING:
1. Heading/paragraph background: If the image shows a heading or paragraph with a gray/colored background strip, the XML MUST have a bg-color attribute on that <paragraph>. Report "wrong_background" if missing.
2. Table row background: If the image shows alternating gray/white rows (zebra striping) in a table, EACH gray row's <cell> elements MUST have bg-color. Report "wrong_background" if missing.
3. Text-level highlight: If specific words have a small colored background behind them, the corresponding <run> MUST have text-bg-color. Report "missing_text_highlight" if missing.

Respond with a JSON object keyed by page number. If no issues on a page, its value is []."""

MULTI_PAGE_USER_PROMPT_TEMPLATE = """All PDF page images are attached (in order). Here are the XML DSLs:

{all_xml_content}

Compare each page image to its XML. Also check cross-page consistency:
- Font sizes should be consistent for the same type of content across pages
- Font families (serif/sans) should be consistent for the same type of content
- Colors should be consistent (e.g., if header text is red on page 1, it should be red on page 2)
- Column widths in tables should be consistent across pages if they represent the same table
- Alignment and spacing should be consistent
- Paragraph spacing (space-before, space-after, line-spacing) should be consistent for similar content

For each page, list issues found. Also check:
- If headings/paragraphs have background shading that matches the original image
- If table border styles match (single line vs double line)
- If any text has a colored highlight/background behind it
- If LaTeX artifacts have been properly cleaned from text content
- If font-family (serif/sans) matches the actual font in the image
- If paragraph spacing values are reasonable and consistent

Respond with JSON object:
{{
  "page_1": [
    {{"type": "missing_text", "description": "...", "after_region": N}},
    {{"type": "wrong_style", "region": N, "field": "font-size-pt", "expected": 11, "description": "inconsistent with page 2"}},
    {{"type": "latex_artifact", "region": N, "description": "LaTeX markup not cleaned"}},
    {{"type": "font_mismatch", "region": N, "field": "font-family", "expected": "serif", "actual": "sans"}},
    {{"type": "spacing_issue", "region": N, "field": "line-spacing", "expected": 1.5, "actual": 1.0}}
  ],
  "page_2": []
}}
Types: missing_text, wrong_style, wrong_order, missing_image, extra_content,
missing_textframe, wrong_textframe, cross_page_inconsistency,
wrong_background, wrong_border_style, missing_text_highlight,
latex_artifact (LaTeX/math markup not properly cleaned),
font_mismatch (font-family doesn't match actual font),
spacing_issue (paragraph spacing is incorrect or inconsistent).

BACKGROUND COLOR CHECKLIST — compare each page image against its XML carefully:
1. For each heading/paragraph in the image that has a gray or colored background strip: verify the corresponding <paragraph> in XML has bg-color="R,G,B". If missing, report type="wrong_background".
2. For each table with alternating gray/white rows: verify that gray rows have bg-color on their <cell> elements. If missing, report type="wrong_background".
3. For each word/character in the image that has a small colored background highlight: verify the corresponding <run> has text-bg-color="R,G,B". If missing, report type="missing_text_highlight".
Only output JSON object, nothing else."""


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def load_api_key(args_key, workspace):
    """Load Poe API key from args, env, or .env file."""
    if args_key:
        return args_key
    if os.environ.get("POE_API_KEY"):
        return os.environ["POE_API_KEY"]
    env_file = Path(workspace) / ".env"
    if env_file.exists():
        for line in env_file.read_text().splitlines():
            line = line.strip()
            if line.startswith("POE_API_KEY="):
                return line.split("=", 1)[1].strip().strip("\"'")
    cwd_env = Path(".env")
    if cwd_env.exists():
        for line in cwd_env.read_text().splitlines():
            line = line.strip()
            if line.startswith("POE_API_KEY="):
                return line.split("=", 1)[1].strip().strip("\"'")
    return None


def encode_image(path):
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")


def call_poe_review(api_key, page_image_b64, xml_content):
    """Call Poe AI API for single-page DSL review and return parsed JSON array."""
    if requests is None:
        raise RuntimeError("requests library not available")

    user_prompt = USER_PROMPT_TEMPLATE.format(xml_content=xml_content)

    payload = {
        "model": POE_MODEL,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {
                "role": "user",
                "content": [
                    {
                        "type": "image_url",
                        "image_url": {"url": f"data:image/png;base64,{page_image_b64}"},
                    },
                    {"type": "text", "text": user_prompt},
                ],
            },
        ],
        "temperature": 0.1,
        "max_tokens": 65536,
    }

    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }

    resp = requests.post(POE_API_URL, json=payload, headers=headers, timeout=120)
    resp.raise_for_status()

    data = resp.json()
    content = data["choices"][0]["message"]["content"]

    # Strip markdown code fences if present
    content = content.strip()
    if content.startswith("```"):
        content = content.split("\n", 1)[1].rsplit("```", 1)[0]
    content = content.strip()

    result = json.loads(content)
    if isinstance(result, list):
        return result

    raise ValueError(f"Unexpected response format: {type(result)}")


def call_poe_review_multi(api_key, page_images_b64, xml_contents, page_numbers):
    """Call Poe AI API for multi-page DSL review.

    Args:
        api_key: Poe API key
        page_images_b64: list of base64-encoded page images
        xml_contents: list of XML DSL strings
        page_numbers: list of page numbers (1-based)

    Returns:
        dict mapping page_num -> list of issues
    """
    if requests is None:
        raise RuntimeError("requests library not available")

    # Build combined XML content with page separators
    all_xml_parts = []
    for pn, xml in zip(page_numbers, xml_contents):
        all_xml_parts.append(f"--- Page {pn} XML ---\n{xml}")
    all_xml_content = "\n\n".join(all_xml_parts)

    user_prompt = MULTI_PAGE_USER_PROMPT_TEMPLATE.format(
        all_xml_content=all_xml_content
    )

    # Build content array: all images first, then text prompt
    content_parts = []
    for img_b64 in page_images_b64:
        content_parts.append(
            {
                "type": "image_url",
                "image_url": {"url": f"data:image/png;base64,{img_b64}"},
            }
        )
    content_parts.append({"type": "text", "text": user_prompt})

    payload = {
        "model": POE_MODEL,
        "messages": [
            {"role": "system", "content": MULTI_PAGE_SYSTEM_PROMPT},
            {
                "role": "user",
                "content": content_parts,
            },
        ],
        "temperature": 0.1,
        "max_tokens": 65536,
    }

    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }

    resp = requests.post(POE_API_URL, json=payload, headers=headers, timeout=180)
    resp.raise_for_status()

    data = resp.json()
    content = data["choices"][0]["message"]["content"]

    # Strip markdown code fences if present
    content = content.strip()
    if content.startswith("```"):
        content = content.split("\n", 1)[1].rsplit("```", 1)[0]
    content = content.strip()

    result = json.loads(content)
    if isinstance(result, dict):
        # Normalize keys: "page_1" -> 1, "1" -> 1, etc.
        normalized = {}
        for key, issues in result.items():
            if isinstance(key, str):
                # Extract number from "page_1" or "1"
                num_str = key.replace("page_", "").strip()
                try:
                    page_num = int(num_str)
                except ValueError:
                    continue
            else:
                page_num = int(key)
            if isinstance(issues, list):
                normalized[page_num] = issues
        return normalized

    raise ValueError(f"Unexpected response format: {type(result)}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def main():
    parser = argparse.ArgumentParser(
        description="Review XML DSL against page images using Poe AI"
    )
    parser.add_argument("--workspace", required=True, help="Workspace directory path")
    parser.add_argument("--pages", required=True, type=int, help="Total page count")
    parser.add_argument(
        "--poe-api-key", default=None, help="Poe API key (or set POE_API_KEY env)"
    )
    args = parser.parse_args()

    workspace = Path(args.workspace)
    total_pages = args.pages

    # Load API key
    api_key = load_api_key(args.poe_api_key, workspace)
    if not api_key:
        print("Warning: POE_API_KEY not found. Skipping VLM review.", file=sys.stderr)
        # Write empty reviews
        dsl_dir = workspace / "dsl"
        for page_num in range(1, total_pages + 1):
            out_path = dsl_dir / f"review-page-{page_num}.json"
            out_path.write_text("[]")
        print("Wrote empty review files (no API key).")
        return

    if requests is None:
        print(
            "Warning: requests library not available. Skipping VLM review.",
            file=sys.stderr,
        )
        return

    dsl_dir = workspace / "dsl"
    total_issues = 0

    # Try multi-page mode for <=5 pages
    if total_pages <= 5:
        print(f"Using multi-page review mode ({total_pages} pages)...")
        try:
            page_images_b64 = []
            xml_contents = []
            page_numbers = []

            for page_num in range(1, total_pages + 1):
                # Load XML DSL
                xml_path = dsl_dir / f"page-{page_num}.xml"
                if not xml_path.exists():
                    raise FileNotFoundError(f"XML not found: {xml_path}")
                xml_contents.append(xml_path.read_text())

                # Load page image
                img_path = (
                    workspace / "input-pdf-rendered-pngs" / f"page-{page_num}.png"
                )
                if not img_path.exists():
                    img_path = (
                        workspace
                        / "input-pdf-rendered-pngs"
                        / f"page-{page_num:02d}.png"
                    )
                if not img_path.exists():
                    raise FileNotFoundError(f"Image not found for page {page_num}")
                page_images_b64.append(encode_image(img_path))
                page_numbers.append(page_num)

            # Call multi-page review API
            results = call_poe_review_multi(
                api_key, page_images_b64, xml_contents, page_numbers
            )

            # Write results for each page
            for page_num in range(1, total_pages + 1):
                issues = results.get(page_num, [])
                out_path = dsl_dir / f"review-page-{page_num}.json"
                out_path.write_text(json.dumps(issues, ensure_ascii=False, indent=2))
                n_issues = len(issues)
                total_issues += n_issues
                if n_issues == 0:
                    print(f"  Page {page_num}: PASS (no issues)")
                else:
                    print(f"  Page {page_num}: {n_issues} issue(s) found")

            print(f"\nMulti-page review complete. Total issues: {total_issues}")
            if total_issues > 0:
                print(
                    "Agent should read review-page-{N}.json files and edit page-{N}.xml accordingly."
                )
            return

        except Exception as e:
            print(
                f"Multi-page review failed: {e}. Falling back to per-page review.",
                file=sys.stderr,
            )
            total_issues = 0

    # Per-page review (fallback or >5 pages)
    print("Using per-page review mode...")
    for page_num in range(1, total_pages + 1):
        print(f"Reviewing page {page_num}/{total_pages}...", end=" ")

        # Load XML DSL
        xml_path = dsl_dir / f"page-{page_num}.xml"
        if not xml_path.exists():
            print("XML not found, skipped")
            continue

        xml_content = xml_path.read_text()

        # Load page image
        img_path = workspace / "input-pdf-rendered-pngs" / f"page-{page_num}.png"
        if not img_path.exists():
            img_path = (
                workspace / "input-pdf-rendered-pngs" / f"page-{page_num:02d}.png"
            )
        if not img_path.exists():
            print("image not found, skipped")
            continue

        try:
            page_image_b64 = encode_image(img_path)
            issues = call_poe_review(api_key, page_image_b64, xml_content)

            out_path = dsl_dir / f"review-page-{page_num}.json"
            out_path.write_text(json.dumps(issues, ensure_ascii=False, indent=2))

            n_issues = len(issues)
            total_issues += n_issues
            if n_issues == 0:
                print("PASS (no issues)")
            else:
                print(f"{n_issues} issue(s) found")

        except Exception as e:
            print(f"API failed: {e}", file=sys.stderr)
            out_path = dsl_dir / f"review-page-{page_num}.json"
            out_path.write_text("[]")

    print(f"\nReview complete. Total issues across all pages: {total_issues}")
    if total_issues > 0:
        print(
            "Agent should read review-page-{N}.json files and edit page-{N}.xml accordingly."
        )


if __name__ == "__main__":
    main()
