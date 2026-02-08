#!/usr/bin/env python3
"""verify_docx_visual.py - Visual comparison of DOCX output against original PDF.

Usage:
    python verify_docx_visual.py --workspace PATH --pages N --docx PATH [--poe-api-key KEY]

Renders the DOCX to PNGs via soffice+pdftocairo, then uses VLM to compare
each page against the original PDF page images.

Output: $WORKSPACE/dsl/visual-review-page-{N}.json (N = input page number)
"""

import argparse
import base64
import json
import math
import os
import shutil
import subprocess
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

SYSTEM_PROMPT_MATCHED = """You compare two document page images: the Original PDF (first image) and a DOCX rendering (second image), along with the XML DSL that produced the DOCX.
Find visual differences and provide QUANTITATIVE, ACTIONABLE fixes referencing specific XML DSL attributes.

CRITICAL — Your descriptions MUST be specific and measurable. For every issue:
- Reference the exact XML element (by tag name, text content snippet, or position in document order)
- State the CURRENT value from the XML DSL (or "missing" if the attribute is absent)
- State the SUGGESTED new value (a concrete number, color, or setting)
- Use the exact XML attribute names: font-size-pt, bbox, x-twips, y-twips, width-twips, height-twips, has-border, border-color, col-widths, space-before-pt, space-after-pt, line-spacing, margin-top-cm, margin-bottom-cm, color-rgb, bg-color, bold, italic, alignment, border-style, font-family

BAD example: "The font size appears smaller in the DOCX version"
GOOD example: "The heading <run> 'This is Title' has font-size-pt=28 but visually needs ~32pt to match the Original PDF"

BAD example: "The image is smaller and positioned higher"
GOOD example: "The <image> with bbox='55,528,442,799' should be changed to bbox='55,500,470,810' to match Original size and position"

BAD example: "The table rows are shorter"
GOOD example: "Table row heights are too small; the table bbox height (321-489=168 normalized units) should be increased to approximately 321-530 (209 normalized units) to match Original"

Pay special attention to:
1. Missing heading/paragraph background colors — gray or colored strips behind headings in the Original that are missing in the DOCX.
2. Missing table row shading — alternating gray/white rows (zebra striping) in the Original that appear as all-white in the DOCX.
3. Missing text-level highlights — small colored backgrounds behind specific words/characters in the Original that are missing in the DOCX.
Respond ONLY with a JSON array of issues. If no differences, respond with []."""

SYSTEM_PROMPT_MISMATCHED = """You compare a multi-page Original PDF with a DOCX rendering that has a different page count. The Original PDF has {input_pages} pages and the DOCX rendering has {docx_pages} pages. Images are labeled. XML DSL for each page is also provided.
Identify content mapping across pages and find visual differences.
Report differences in: text content, positioning, colors, fonts, borders, backgrounds, images, tables, alignment.
Also report why page counts differ (font size, table reflow, margins, etc.).

CRITICAL — Your descriptions MUST be specific and measurable. For every issue:
- Reference the exact XML element (by tag name, text content snippet, or position in document order)
- State the CURRENT value from the XML DSL (or "missing" if the attribute is absent)
- State the SUGGESTED new value (a concrete number, color, or setting)
- Use the exact XML attribute names: font-size-pt, bbox, x-twips, y-twips, width-twips, height-twips, has-border, border-color, col-widths, space-before-pt, space-after-pt, line-spacing, margin-top-cm, margin-bottom-cm, color-rgb, bg-color, bold, italic, alignment, border-style, font-family

BAD example: "The font size appears smaller in the DOCX version"
GOOD example: "The heading <run> 'This is Title' has font-size-pt=28 but visually needs ~32pt to match the Original PDF"

Pay special attention to:
1. Missing heading/paragraph background colors — gray or colored strips behind headings in the Original that are missing in the DOCX.
2. Missing table row shading — alternating gray/white rows (zebra striping) in the Original that appear as all-white in the DOCX.
3. Missing text-level highlights — small colored backgrounds behind specific words/characters in the Original that are missing in the DOCX.
Respond ONLY with a JSON object keyed by original page number. If no differences for a page, its value is []."""

USER_PROMPT_MATCHED = """Compare these two page images and report any visual differences.

Original PDF page (first image) vs DOCX rendering (second image).

Here is the XML DSL that produced the DOCX:
{xml_content}

For each difference found, respond with JSON array. Each issue MUST include:
- "type": issue category
- "element": which XML element is affected (e.g. "heading 'This is Title'", "table row 0", "image bbox='55,528,442,799'", "text-frame at x-twips=4857")
- "field": the XML attribute to change (e.g. "font-size-pt", "bbox", "width-twips", "border-color", "col-widths", "has-border")
- "current_value": the current value in the XML DSL (or "missing" if attribute absent)
- "suggested_value": the specific value to change to
- "description": brief explanation of what's visually wrong

Example output:
[
  {{"type": "font_difference", "element": "heading run 'This is Title'", "field": "font-size-pt", "current_value": "28", "suggested_value": "32", "description": "Title text is visually smaller in DOCX than Original PDF"}},
  {{"type": "wrong_positioning", "element": "image src='cropped_page0_idx0.jpg'", "field": "bbox", "current_value": "55,528,442,799", "suggested_value": "55,500,470,810", "description": "Image is smaller and positioned lower in DOCX vs Original"}},
  {{"type": "table_issue", "element": "table rows=4 cols=5", "field": "bbox", "current_value": "94,321,903,489", "suggested_value": "94,321,903,540", "description": "Table row heights are too short, text appears cramped"}},
  {{"type": "border_missing", "element": "text-frame at x-twips=4857 y-twips=14093", "field": "has-border", "current_value": "true", "suggested_value": "true with border-size=12 (1.5pt)", "description": "Text box border is much thinner than Original; increase border width"}}
]

Issue types: missing_text, extra_text, wrong_positioning, layout_mismatch,
color_difference, font_difference, border_missing, table_issue,
background_missing (heading/paragraph/row background color present in Original but missing in DOCX).
If no differences found, respond with [].
Only output JSON array."""

USER_PROMPT_MISMATCHED = """Compare the Original PDF pages with the DOCX rendering pages.
The Original PDF has {input_pages} pages, but the DOCX rendered as {docx_pages} pages.

Images are provided in order: first all Original PDF pages (labeled), then all DOCX pages (labeled).

Here are the XML DSLs for each page:
{all_xml_content}

For each original page, find the corresponding DOCX content and report differences.
Also report why the page count differs with specific attribute fixes.

Each issue MUST include "element", "field", "current_value", "suggested_value" where applicable.

Respond with JSON object:
{{
  "page_1": [
    {{"type": "page_count_mismatch", "description": "...", "input_pages": {input_pages}, "docx_pages": {docx_pages}, "element": "page", "field": "margin-top-cm", "current_value": "1.27", "suggested_value": "1.0"}},
    {{"type": "font_difference", "element": "paragraph run 'body text'", "field": "font-size-pt", "current_value": "11", "suggested_value": "10", "description": "Reducing font size to prevent page overflow"}}
  ],
  "page_2": []
}}

Issue types: page_count_mismatch, missing_text, extra_text, wrong_positioning,
layout_mismatch, color_difference, font_difference, border_missing, table_issue,
background_missing (heading/paragraph/row background color present in Original but missing in DOCX).
Only output JSON object."""

USER_PROMPT_MULTI_MATCHED = """Compare all page images. For each page pair (Original PDF vs DOCX rendering), report visual differences.

Images are in order: Original PDF page 1, DOCX page 1, Original PDF page 2, DOCX page 2, ...

Here are the XML DSLs for each page:
{all_xml_content}

Each issue MUST include quantitative details:
- "element": which XML element is affected
- "field": the XML attribute to change
- "current_value": current value in XML DSL
- "suggested_value": specific value to change to
- "description": brief explanation

Respond with JSON object:
{{
  "page_1": [],
  "page_2": [{{"type": "color_difference", "element": "heading run 'Section Title'", "field": "color-rgb", "current_value": "0,0,0", "suggested_value": "0,0,255", "description": "Heading color should be blue to match Original"}}]
}}

Issue types: missing_text, extra_text, wrong_positioning, layout_mismatch,
color_difference, font_difference, border_missing, table_issue,
background_missing (heading/paragraph/row background color present in Original but missing in DOCX).
Only output JSON object."""


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


def check_soffice():
    """Check if soffice (LibreOffice) is available."""
    return shutil.which("soffice") is not None


def docx_to_pdf(docx_path, output_dir):
    """Convert DOCX to PDF using soffice."""
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    result = subprocess.run(
        ["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(output_dir), str(docx_path)],
        capture_output=True, text=True, timeout=120
    )
    if result.returncode != 0:
        raise RuntimeError(f"soffice conversion failed: {result.stderr}")
    # Find the output PDF
    docx_stem = Path(docx_path).stem
    pdf_path = output_dir / f"{docx_stem}.pdf"
    if not pdf_path.exists():
        raise RuntimeError(f"Expected PDF not found at {pdf_path}")
    return pdf_path


def pdf_to_pngs(pdf_path, output_dir, prefix="page"):
    """Convert PDF to PNGs using pdftocairo."""
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    result = subprocess.run(
        ["pdftocairo", "-png", "-r", "200", str(pdf_path), str(output_dir / prefix)],
        capture_output=True, text=True, timeout=120
    )
    if result.returncode != 0:
        raise RuntimeError(f"pdftocairo failed: {result.stderr}")
    # Collect output PNGs
    pngs = sorted(output_dir.glob(f"{prefix}-*.png"))
    return pngs


def find_input_page_png(workspace, page_num):
    """Find the input PDF page PNG."""
    img_path = Path(workspace) / "input-pdf-rendered-pngs" / f"page-{page_num}.png"
    if img_path.exists():
        return img_path
    img_path = Path(workspace) / "input-pdf-rendered-pngs" / f"page-{page_num:02d}.png"
    if img_path.exists():
        return img_path
    return None


def call_poe_visual_compare(api_key, content_parts, system_prompt, timeout=180):
    """Call Poe AI for visual comparison."""
    if requests is None:
        raise RuntimeError("requests library not available")

    payload = {
        "model": POE_MODEL,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": content_parts},
        ],
        "temperature": 0.1,
        "max_tokens": 65536,
    }

    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }

    resp = requests.post(POE_API_URL, json=payload, headers=headers, timeout=timeout)
    resp.raise_for_status()

    data = resp.json()
    content = data["choices"][0]["message"]["content"]

    # Strip markdown code fences
    content = content.strip()
    if content.startswith("```"):
        content = content.split("\n", 1)[1].rsplit("```", 1)[0]
    content = content.strip()

    return json.loads(content)


# ---------------------------------------------------------------------------
# Comparison strategies
# ---------------------------------------------------------------------------

def compare_per_page(api_key, input_png, docx_png, page_num, xml_content=""):
    """Compare a single input page against its DOCX rendering."""
    input_b64 = encode_image(input_png)
    docx_b64 = encode_image(docx_png)

    prompt = USER_PROMPT_MATCHED.format(xml_content=xml_content if xml_content else "(XML DSL not available)")

    content_parts = [
        {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{input_b64}"}},
        {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{docx_b64}"}},
        {"type": "text", "text": prompt},
    ]

    result = call_poe_visual_compare(api_key, content_parts, SYSTEM_PROMPT_MATCHED)
    if isinstance(result, list):
        return result
    return []


def compare_multi_page_matched(api_key, input_pngs, docx_pngs, page_numbers, xml_contents=None):
    """Compare multiple pages in a single call (pages match 1:1)."""
    content_parts = []
    for i, (inp, doc) in enumerate(zip(input_pngs, docx_pngs)):
        content_parts.append({"type": "image_url", "image_url": {"url": f"data:image/png;base64,{encode_image(inp)}"}})
        content_parts.append({"type": "image_url", "image_url": {"url": f"data:image/png;base64,{encode_image(doc)}"}})

    # Build combined XML content
    all_xml_parts = []
    if xml_contents:
        for pn, xml in zip(page_numbers, xml_contents):
            all_xml_parts.append(f"--- Page {pn} XML ---\n{xml}")
    all_xml_content = "\n\n".join(all_xml_parts) if all_xml_parts else "(XML DSL not available)"

    content_parts.append({"type": "text", "text": USER_PROMPT_MULTI_MATCHED.format(all_xml_content=all_xml_content)})

    result = call_poe_visual_compare(api_key, content_parts, SYSTEM_PROMPT_MATCHED)

    if isinstance(result, dict):
        normalized = {}
        for key, issues in result.items():
            num_str = str(key).replace("page_", "").strip()
            try:
                normalized[int(num_str)] = issues if isinstance(issues, list) else []
            except ValueError:
                continue
        return normalized

    return {pn: [] for pn in page_numbers}


def compare_mismatched(api_key, input_pngs, docx_pngs, input_pages, docx_pages, page_numbers, xml_contents=None):
    """Compare when page counts don't match."""
    total_images = len(input_pngs) + len(docx_pngs)

    # Build combined XML content
    all_xml_parts = []
    if xml_contents:
        for pn, xml in zip(page_numbers, xml_contents):
            all_xml_parts.append(f"--- Page {pn} XML ---\n{xml}")
    all_xml_content = "\n\n".join(all_xml_parts) if all_xml_parts else "(XML DSL not available)"

    if total_images <= 10:
        # Send all images together
        content_parts = []
        for i, inp in enumerate(input_pngs):
            content_parts.append({"type": "text", "text": f"Original PDF Page {i+1}:"})
            content_parts.append({"type": "image_url", "image_url": {"url": f"data:image/png;base64,{encode_image(inp)}"}})
        for i, doc in enumerate(docx_pngs):
            content_parts.append({"type": "text", "text": f"DOCX Rendered Page {i+1}:"})
            content_parts.append({"type": "image_url", "image_url": {"url": f"data:image/png;base64,{encode_image(doc)}"}})

        prompt = USER_PROMPT_MISMATCHED.format(input_pages=input_pages, docx_pages=docx_pages, all_xml_content=all_xml_content)
        content_parts.append({"type": "text", "text": prompt})

        system = SYSTEM_PROMPT_MISMATCHED.format(input_pages=input_pages, docx_pages=docx_pages)
        result = call_poe_visual_compare(api_key, content_parts, system)

        if isinstance(result, dict):
            normalized = {}
            for key, issues in result.items():
                num_str = str(key).replace("page_", "").strip()
                try:
                    normalized[int(num_str)] = issues if isinstance(issues, list) else []
                except ValueError:
                    continue
            return normalized
        return {pn: [] for pn in page_numbers}
    else:
        # Batch by ratio
        ratio = docx_pages / input_pages if input_pages > 0 else 1
        results = {}

        for i, pn in enumerate(page_numbers):
            docx_start = int(i * ratio)
            docx_end = min(int((i + 1) * ratio), docx_pages)
            if docx_end <= docx_start:
                docx_end = min(docx_start + 1, docx_pages)

            content_parts = []
            content_parts.append({"type": "text", "text": f"Original PDF Page {pn}:"})
            content_parts.append({"type": "image_url", "image_url": {"url": f"data:image/png;base64,{encode_image(input_pngs[i])}"}})

            for j in range(docx_start, docx_end):
                if j < len(docx_pngs):
                    content_parts.append({"type": "text", "text": f"DOCX Rendered Page {j+1}:"})
                    content_parts.append({"type": "image_url", "image_url": {"url": f"data:image/png;base64,{encode_image(docx_pngs[j])}"}})

            # Build per-page XML content
            per_page_xml = "(XML DSL not available)"
            if xml_contents and i < len(xml_contents):
                per_page_xml = f"--- Page {pn} XML ---\n{xml_contents[i]}"

            prompt = USER_PROMPT_MISMATCHED.format(input_pages=input_pages, docx_pages=docx_pages, all_xml_content=per_page_xml)
            content_parts.append({"type": "text", "text": prompt})

            system = SYSTEM_PROMPT_MISMATCHED.format(input_pages=input_pages, docx_pages=docx_pages)

            try:
                result = call_poe_visual_compare(api_key, content_parts, system)
                if isinstance(result, dict):
                    issues = result.get(pn, result.get(f"page_{pn}", []))
                    results[pn] = issues if isinstance(issues, list) else []
                elif isinstance(result, list):
                    results[pn] = result
                else:
                    results[pn] = []
            except Exception as e:
                print(f"  Page {pn} comparison failed: {e}", file=sys.stderr)
                results[pn] = []

        return results


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Visual comparison of DOCX vs original PDF")
    parser.add_argument("--workspace", required=True, help="Workspace directory path")
    parser.add_argument("--pages", required=True, type=int, help="Input PDF page count")
    parser.add_argument("--docx", required=True, help="Path to output DOCX file")
    parser.add_argument("--poe-api-key", default=None, help="Poe API key")
    args = parser.parse_args()

    workspace = Path(args.workspace)
    input_pages = args.pages
    docx_path = Path(args.docx)
    dsl_dir = workspace / "dsl"

    # Check prerequisites
    if not check_soffice():
        print("WARNING: soffice not available. Skipping visual verification.", file=sys.stderr)
        for pn in range(1, input_pages + 1):
            out_path = dsl_dir / f"visual-review-page-{pn}.json"
            out_path.write_text("[]")
        print("Wrote empty visual review files (soffice not available).")
        return

    if not docx_path.exists():
        print(f"Error: DOCX file not found: {docx_path}", file=sys.stderr)
        sys.exit(1)

    api_key = load_api_key(args.poe_api_key, workspace)
    if not api_key:
        print("WARNING: POE_API_KEY not found. Skipping visual verification.", file=sys.stderr)
        for pn in range(1, input_pages + 1):
            out_path = dsl_dir / f"visual-review-page-{pn}.json"
            out_path.write_text("[]")
        print("Wrote empty visual review files (no API key).")
        return

    if requests is None:
        print("WARNING: requests library not available. Skipping visual verification.", file=sys.stderr)
        for pn in range(1, input_pages + 1):
            out_path = dsl_dir / f"visual-review-page-{pn}.json"
            out_path.write_text("[]")
        return

    # Step 1: DOCX → PDF
    print("Converting DOCX to PDF...")
    docx_render_dir = workspace / "docx-rendered-pngs"
    try:
        pdf_path = docx_to_pdf(docx_path, workspace)
    except Exception as e:
        print(f"Error converting DOCX to PDF: {e}", file=sys.stderr)
        for pn in range(1, input_pages + 1):
            (dsl_dir / f"visual-review-page-{pn}.json").write_text("[]")
        return

    # Step 2: PDF → PNGs
    print("Rendering PDF to PNGs...")
    try:
        docx_pngs = pdf_to_pngs(pdf_path, docx_render_dir)
    except Exception as e:
        print(f"Error rendering PDF to PNGs: {e}", file=sys.stderr)
        for pn in range(1, input_pages + 1):
            (dsl_dir / f"visual-review-page-{pn}.json").write_text("[]")
        return

    docx_pages = len(docx_pngs)
    print(f"Input PDF: {input_pages} pages, DOCX rendered: {docx_pages} pages")

    # Collect input page PNGs
    input_pngs = []
    page_numbers = []
    for pn in range(1, input_pages + 1):
        png = find_input_page_png(workspace, pn)
        if png:
            input_pngs.append(png)
            page_numbers.append(pn)
        else:
            print(f"Warning: Input page {pn} PNG not found", file=sys.stderr)

    if not input_pngs:
        print("Error: No input page PNGs found", file=sys.stderr)
        for pn in range(1, input_pages + 1):
            (dsl_dir / f"visual-review-page-{pn}.json").write_text("[]")
        return

    # Load XML DSL files for quantitative review
    xml_contents = []
    for pn in page_numbers:
        xml_path = dsl_dir / f"page-{pn}.xml"
        if xml_path.exists():
            xml_contents.append(xml_path.read_text())
        else:
            xml_contents.append("")

    # Step 3: Choose comparison strategy
    dsl_dir.mkdir(parents=True, exist_ok=True)

    if input_pages == docx_pages:
        # Pages match
        if input_pages <= 3:
            # Multi-page mode: send all together
            print("Using multi-page matched comparison mode...")
            try:
                results = compare_multi_page_matched(api_key, input_pngs, docx_pngs, page_numbers, xml_contents)
                for pn in page_numbers:
                    issues = results.get(pn, [])
                    out_path = dsl_dir / f"visual-review-page-{pn}.json"
                    out_path.write_text(json.dumps(issues, ensure_ascii=False, indent=2))
                    n = len(issues)
                    if n == 0:
                        print(f"  Page {pn}: PASS")
                    else:
                        print(f"  Page {pn}: {n} difference(s)")
            except Exception as e:
                print(f"Multi-page comparison failed: {e}. Falling back to per-page.", file=sys.stderr)
                for i, pn in enumerate(page_numbers):
                    try:
                        if i < len(docx_pngs):
                            xml_c = xml_contents[i] if i < len(xml_contents) else ""
                            issues = compare_per_page(api_key, input_pngs[i], docx_pngs[i], pn, xml_c)
                        else:
                            issues = []
                        out_path = dsl_dir / f"visual-review-page-{pn}.json"
                        out_path.write_text(json.dumps(issues, ensure_ascii=False, indent=2))
                        n = len(issues)
                        if n == 0:
                            print(f"  Page {pn}: PASS")
                        else:
                            print(f"  Page {pn}: {n} difference(s)")
                    except Exception as pe:
                        print(f"  Page {pn}: comparison failed: {pe}", file=sys.stderr)
                        (dsl_dir / f"visual-review-page-{pn}.json").write_text("[]")
        else:
            # Per-page mode
            print("Using per-page matched comparison mode...")
            for i, pn in enumerate(page_numbers):
                try:
                    if i < len(docx_pngs):
                        xml_c = xml_contents[i] if i < len(xml_contents) else ""
                        issues = compare_per_page(api_key, input_pngs[i], docx_pngs[i], pn, xml_c)
                    else:
                        issues = []
                    out_path = dsl_dir / f"visual-review-page-{pn}.json"
                    out_path.write_text(json.dumps(issues, ensure_ascii=False, indent=2))
                    n = len(issues)
                    if n == 0:
                        print(f"  Page {pn}: PASS")
                    else:
                        print(f"  Page {pn}: {n} difference(s)")
                except Exception as e:
                    print(f"  Page {pn}: comparison failed: {e}", file=sys.stderr)
                    (dsl_dir / f"visual-review-page-{pn}.json").write_text("[]")
    else:
        # Pages don't match
        print(f"Page count mismatch! Input: {input_pages}, DOCX: {docx_pages}")
        print("Using mismatched comparison mode...")
        try:
            results = compare_mismatched(api_key, input_pngs, docx_pngs, input_pages, docx_pages, page_numbers, xml_contents)

            # Add page_count_mismatch issue to first page if not already present
            first_pn = page_numbers[0] if page_numbers else 1
            first_issues = results.get(first_pn, [])
            has_mismatch = any(i.get("type") == "page_count_mismatch" for i in first_issues)
            if not has_mismatch:
                first_issues.insert(0, {
                    "type": "page_count_mismatch",
                    "description": f"Original PDF has {input_pages} pages but DOCX rendered as {docx_pages} pages",
                    "input_pages": input_pages,
                    "docx_pages": docx_pages,
                })
                results[first_pn] = first_issues

            for pn in page_numbers:
                issues = results.get(pn, [])
                out_path = dsl_dir / f"visual-review-page-{pn}.json"
                out_path.write_text(json.dumps(issues, ensure_ascii=False, indent=2))
                n = len(issues)
                if n == 0:
                    print(f"  Page {pn}: PASS")
                else:
                    print(f"  Page {pn}: {n} difference(s)")
        except Exception as e:
            print(f"Mismatched comparison failed: {e}", file=sys.stderr)
            for pn in page_numbers:
                issues = [{
                    "type": "page_count_mismatch",
                    "description": f"Original PDF has {input_pages} pages but DOCX rendered as {docx_pages} pages. Comparison failed: {e}",
                    "input_pages": input_pages,
                    "docx_pages": docx_pages,
                }]
                (dsl_dir / f"visual-review-page-{pn}.json").write_text(json.dumps(issues, ensure_ascii=False, indent=2))

    # Write empty reviews for any missing pages
    for pn in range(1, input_pages + 1):
        out_path = dsl_dir / f"visual-review-page-{pn}.json"
        if not out_path.exists():
            out_path.write_text("[]")

    total_issues = 0
    for pn in range(1, input_pages + 1):
        out_path = dsl_dir / f"visual-review-page-{pn}.json"
        issues = json.loads(out_path.read_text())
        total_issues += len(issues)

    print(f"\nVisual verification complete. Total differences: {total_issues}")
    if total_issues > 0:
        print("Agent should read visual-review-page-{N}.json files and edit page-{N}.xml accordingly.")


if __name__ == "__main__":
    main()
