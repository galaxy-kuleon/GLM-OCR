#!/usr/bin/env python3
"""extract_styles.py - Extract text styling from PDF page images using Poe AI.

Usage:
    python extract_styles.py --workspace PATH --pages N [--poe-api-key KEY]

For each page, calls Poe AI (gemini-3-flash) with a simplified prompt to
extract font size, bold, color, and alignment per OCR region.

Output: $WORKSPACE/ocr-output/input/style-page-{N}.json (1-based, per page)

If Poe API is unavailable, falls back to pure defaults (pipeline continues).
API key is read from --poe-api-key argument or POE_API_KEY environment variable
or .env file in workspace.
"""

import argparse
import base64
import json
import os
import re
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

SYSTEM_PROMPT = """You analyze document page images and describe text styling.
For each numbered text region, report: font size (pt), bold (true/false),
text color as [R,G,B], alignment (left/center/right),
background shading color bg=[R,G,B] if the paragraph/heading has ANY visible colored or gray background fill (even light gray like [230,230,230]),
and font-family: "serif" or "sans" (detect if text uses serif like Times/SimSun or sans-serif like Arial/Helvetica).

Also detect spacing: space-before (pt), space-after (pt), and line-spacing (1.0=single, 1.5=1.5, 2.0=double).

CRITICAL — BACKGROUND COLOR DETECTION:
You MUST report bg=[R,G,B] for ANY region that has a colored or gray background fill behind the text. This includes:
- Title bars / heading strips with gray or colored backgrounds (common: [217,217,217], [230,230,230], [242,242,242], [191,191,191])
- Paragraphs with light gray shading
- Any region where the background is NOT pure white [255,255,255]
Do NOT omit bg for light gray backgrounds — these are extremely important for document fidelity.
If a region has a white background, simply omit the bg field.

For tables: report if the first row is a header (bold/shaded background).
For tables: also report the border line style: single (one line), double (two thin lines), or none.
For text regions inside a visible text box or bordered frame: report tb=true and bd=true if it has a visible border.
NEVER set tb or bd for image regions (label=image).
Respond ONLY with a JSON array. No explanation."""

USER_PROMPT_TEMPLATE = """Page image is attached. OCR detected these regions:

{regions_summary}

Respond with JSON array:
[
  {{"i": 0, "fs": 14, "b": true, "c": [0,0,0], "a": "center", "bg": [217,217,217], "ff": "serif", "sb": 12, "sa": 6, "ls": 1.0}},
  {{"i": 1, "fs": 11, "b": false, "c": [0,0,0], "a": "left", "bg": [242,242,242], "ff": "sans", "sb": 0, "sa": 6, "ls": 1.0}},
  {{"i": 2, "fs": 11, "b": false, "c": [0,0,0], "a": "left", "th": true, "bs": "double", "ff": "sans", "sb": 0, "sa": 0, "ls": 1.0}},
  {{"i": 3, "fs": 10, "b": false, "c": [0,0,0], "a": "left", "tb": true, "bd": true, "ff": "serif", "sb": 0, "sa": 0, "ls": 1.0}}
]

i=region index, fs=font size pt, b=bold, c=color RGB, a=alignment,
bg=[R,G,B] paragraph/heading background fill color — report for ANY colored or gray fill behind text (e.g. gray title bar, colored header). Do NOT omit light gray backgrounds,
ff=font family: "serif", "sans", or "mono" (serif=Times/SimSun, sans=Arial/Helvetica, mono=Courier),
sb=space before paragraph in pt, sa=space after paragraph in pt,
ls=line spacing multiplier (1.0=single, 1.5=one-half, 2.0=double),
th=table has header row (only for tables),
bs=table border style: "single", "double", or "none" (only for tables),
tb=text region inside a text box (NEVER for image regions),
bd=text box has visible border (NEVER for image regions).

IMPORTANT — bg=[R,G,B]: Look carefully at EVERY region. If a heading or paragraph has a gray or colored strip/bar/fill behind the text, you MUST report bg. Common values: [217,217,217], [230,230,230], [242,242,242], [191,191,191]. Do NOT omit bg for light gray backgrounds.
Only output JSON array."""


# ---------------------------------------------------------------------------
# Table cell-level style extraction prompts
# ---------------------------------------------------------------------------

TABLE_CELL_STYLE_SYSTEM = """You analyze table images to identify cells with non-default styling. You MUST detect:
1. ALTERNATING ROW SHADING — many tables have alternating gray/white rows (zebra striping). Scan EVERY row for background color.
2. HEADER ROW BACKGROUND — the first row often has a darker gray or colored background.
3. TEXT-LEVEL HIGHLIGHTS — specific words/characters may have a small colored background behind them (not the whole cell).
4. CELL BACKGROUND COLORS — individual cells may have non-white backgrounds.
Report ONLY cells/columns/rows that differ from the default (black text, white/no background).
Respond ONLY with a JSON object. No explanation."""

TABLE_CELL_STYLE_PROMPT = """This image shows a table from a document page. The table has {num_rows} rows (row 0 to row {last_row}) and {num_cols} columns.

Follow these steps to analyze the table:

STEP 1 — ALTERNATING ROW SHADING: Scan every row from row 0 to row {last_row}. Many tables use alternating gray/white rows (zebra striping). For EACH row that has a non-white background, add it to "row_colors". Common shading colors: [217,217,217], [230,230,230], [242,242,242], [245,245,245].

STEP 2 — TEXT COLOR OVERRIDES: Check if any column has ALL cells in a non-black text color (e.g., red, blue). Add to "col_colors". Check individual cells for non-black text color; add to "cell_colors".

STEP 3 — TEXT-LEVEL HIGHLIGHTS: Look for specific words or characters that have a small colored background behind them (not the whole cell). These are text-level highlights. Report in "keyword_styles".

STEP 4 — INDIVIDUAL CELL BACKGROUNDS: Check if any individual cell has a unique background color different from its row. Add to "cell_colors" with "bg" field.

Respond with this JSON format:
{{
  "col_colors": [{{"col": 5, "c": [204,0,0], "type": "text"}}],
  "row_colors": [
    {{"row": 0, "bg": [217,217,217], "type": "bg"}},
    {{"row": 2, "bg": [242,242,242], "type": "bg"}},
    {{"row": 4, "bg": [242,242,242], "type": "bg"}}
  ],
  "cell_colors": [{{"row": 2, "col": 3, "c": [0,0,255], "type": "text", "text_bg": [200,200,200]}}],
  "keyword_styles": [
    {{"row": 2, "col": 5, "keyword": "example", "c": [255,0,0], "bold": true, "text_bg": [200,200,200]}}
  ]
}}

- col_colors: columns where ALL cells in that column have non-default color
- row_colors: rows where ALL cells in that row have non-default background
- cell_colors: individual cells with non-default color
- "c" = [R,G,B] for non-default text color
- "bg" = [R,G,B] for cell/row background shading color
- "text_bg" = [R,G,B] for text highlight/background color (text-level, not cell-level)
- "type" = "text" for text color override, "bg" for background color override
- Use 0-based indices for rows and columns
- "keyword" = exact text substring with different styling
- "c" = keyword text color (only if different from rest of cell)
- "bold" = true if keyword is bold but rest is not
- "text_bg" = [R,G,B] highlight color behind keyword
- If no keyword-level styling, set "keyword_styles": []

IMPORTANT: If this table has alternating row shading (zebra striping), you MUST list EVERY shaded row in "row_colors". Do NOT omit rows just because there are many of them.
- If no non-default colors found, return {{"col_colors": [], "row_colors": [], "cell_colors": [], "keyword_styles": []}}
Only output JSON object."""

# Style defaults by native_label
STYLE_DEFAULTS = {
    "doc_title": {"fs": 18, "b": True, "a": "center", "c": [0, 0, 0]},
    "paragraph_title": {"fs": 14, "b": True, "a": "left", "c": [0, 0, 0]},
    "text": {"fs": 11, "b": False, "a": "left", "c": [0, 0, 0]},
    "figure_title": {"fs": 10, "b": False, "a": "center", "c": [0, 0, 0]},
    "vision_footnote": {"fs": 9, "b": False, "a": "left", "c": [0, 0, 0]},
    "table": {"fs": 9, "b": False, "a": "left", "c": [0, 0, 0], "th": True},
    "display_formula": {"fs": 11, "b": False, "a": "center", "c": [0, 0, 0]},
}


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def load_api_key(args_key, workspace):
    """Load Poe API key from args, env, or .env file."""
    if args_key:
        return args_key
    if os.environ.get("POE_API_KEY"):
        return os.environ["POE_API_KEY"]
    # Try .env file in workspace
    env_file = Path(workspace) / ".env"
    if env_file.exists():
        for line in env_file.read_text().splitlines():
            line = line.strip()
            if line.startswith("POE_API_KEY="):
                return line.split("=", 1)[1].strip().strip("\"'")
    # Try .env in current directory
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


def build_regions_summary(regions):
    """Build a compact text summary of regions for the prompt."""
    lines = []
    for r in regions:
        idx = r.get("index", 0)
        label = r.get("label", "text")
        native = r.get("native_label", "text")
        content = (r.get("content") or "")[:60].replace("\n", " ")
        bbox = r.get("bbox_2d")
        bbox_str = f" bbox={bbox}" if bbox else ""
        lines.append(f'  [{idx}] {native} ({label}){bbox_str}: "{content}"')
    return "\n".join(lines)


def expand_style_entry(entry, regions):
    """Expand short-form VLM response to full style-page JSON format."""
    idx = entry.get("i", -1)

    # Find matching region for native_label defaults and label
    native_label = "text"
    region_label = "text"
    for r in regions:
        if r.get("index") == idx:
            native_label = r.get("native_label", "text")
            region_label = r.get("label", "text")
            break

    defaults = STYLE_DEFAULTS.get(native_label, STYLE_DEFAULTS["text"])

    result = {
        "region_index": idx,
        "font_size_pt": entry.get("fs", defaults["fs"]),
        "bold": entry.get("b", defaults["b"]),
        "italic": False,
        "underline": False,
        "color_rgb": entry.get("c", defaults["c"]),
        "alignment": entry.get("a", defaults["a"]),
    }

    # Validate and clamp values
    if isinstance(result["font_size_pt"], (int, float)):
        result["font_size_pt"] = max(6, min(72, result["font_size_pt"]))
    else:
        result["font_size_pt"] = defaults["fs"]

    if isinstance(result["color_rgb"], list) and len(result["color_rgb"]) == 3:
        result["color_rgb"] = [max(0, min(255, v)) for v in result["color_rgb"]]
    else:
        result["color_rgb"] = defaults["c"]

    if result["alignment"] not in ("left", "center", "right"):
        result["alignment"] = defaults["a"]

    # Background shading color (paragraph/heading level)
    if "bg" in entry and isinstance(entry["bg"], list) and len(entry["bg"]) == 3:
        bg = [max(0, min(255, v)) for v in entry["bg"]]
        # Ignore white (no shading)
        if bg != [255, 255, 255]:
            result["bg_rgb"] = bg

    # Table-specific
    if "th" in entry:
        result["th"] = entry["th"]

    # Table border style
    if "bs" in entry:
        bs = entry["bs"]
        if bs in ("single", "double", "none"):
            result["border_style"] = bs

    # Text box fields — strip tb/bd from non-text regions (e.g. image)
    if "tb" in entry and region_label == "text":
        result["tb"] = entry["tb"]
    if "bd" in entry and region_label == "text":
        result["bd"] = entry["bd"]

    return result


def generate_defaults_for_page(regions):
    """Generate default style entries when API is unavailable."""
    styles = []
    for r in regions:
        native = r.get("native_label", "text")
        defaults = STYLE_DEFAULTS.get(native, STYLE_DEFAULTS["text"])
        styles.append(
            {
                "region_index": r.get("index", 0),
                "font_size_pt": defaults["fs"],
                "bold": defaults["b"],
                "italic": False,
                "underline": False,
                "color_rgb": defaults["c"],
                "alignment": defaults["a"],
            }
        )
    return styles


def extract_table_metadata(html_content):
    """Extract row/column counts from HTML table content using regex.

    Returns (num_rows, num_cols) or (0, 0) if parsing fails.
    """
    # Count <tr> tags for rows
    rows = re.findall(r"<tr[^>]*>", html_content, re.IGNORECASE)
    num_rows = len(rows)

    if num_rows == 0:
        return (0, 0)

    # Find max columns by counting <td>/<th> in each row (accounting for colspan)
    max_cols = 0
    # Split by <tr> to process each row
    row_chunks = re.split(r"<tr[^>]*>", html_content, flags=re.IGNORECASE)[1:]
    for chunk in row_chunks:
        col_count = 0
        cells = re.findall(r"<(?:td|th)([^>]*)>", chunk, re.IGNORECASE)
        for cell_attrs in cells:
            colspan_match = re.search(r'colspan\s*=\s*["\']?(\d+)', cell_attrs)
            colspan = int(colspan_match.group(1)) if colspan_match else 1
            col_count += colspan
        max_cols = max(max_cols, col_count)

    return (num_rows, max_cols)


def call_poe_table_cell_styles(api_key, page_image_b64, num_rows, num_cols):
    """Call Poe AI to extract cell-level color overrides from a table image.

    Returns dict with col_colors, row_colors, cell_colors or empty dict on failure.
    """
    if requests is None:
        raise RuntimeError("requests library not available")

    user_prompt = TABLE_CELL_STYLE_PROMPT.format(num_rows=num_rows, num_cols=num_cols, last_row=num_rows - 1)

    payload = {
        "model": POE_MODEL,
        "messages": [
            {"role": "system", "content": TABLE_CELL_STYLE_SYSTEM},
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
    if isinstance(result, dict):
        # Validate structure
        validated = {
            "col_colors": result.get("col_colors", []),
            "row_colors": result.get("row_colors", []),
            "cell_colors": result.get("cell_colors", []),
            "keyword_styles": result.get("keyword_styles", []),
        }
        # Validate and clamp RGB values
        for key in ("col_colors", "row_colors", "cell_colors"):
            for entry in validated[key]:
                if (
                    "c" in entry
                    and isinstance(entry["c"], list)
                    and len(entry["c"]) == 3
                ):
                    entry["c"] = [max(0, min(255, v)) for v in entry["c"]]
                if (
                    "bg" in entry
                    and isinstance(entry["bg"], list)
                    and len(entry["bg"]) == 3
                ):
                    entry["bg"] = [max(0, min(255, v)) for v in entry["bg"]]
                if (
                    "text_bg" in entry
                    and isinstance(entry["text_bg"], list)
                    and len(entry["text_bg"]) == 3
                ):
                    entry["text_bg"] = [max(0, min(255, v)) for v in entry["text_bg"]]
        # Validate keyword_styles entries
        for entry in validated["keyword_styles"]:
            if "c" in entry and isinstance(entry["c"], list) and len(entry["c"]) == 3:
                entry["c"] = [max(0, min(255, v)) for v in entry["c"]]
            if (
                "text_bg" in entry
                and isinstance(entry["text_bg"], list)
                and len(entry["text_bg"]) == 3
            ):
                entry["text_bg"] = [max(0, min(255, v)) for v in entry["text_bg"]]
        return validated

    return {"col_colors": [], "row_colors": [], "cell_colors": [], "keyword_styles": []}


def call_poe_api(api_key, page_image_b64, regions_summary):
    """Call Poe AI API and return parsed JSON array."""
    if requests is None:
        raise RuntimeError("requests library not available")

    user_prompt = USER_PROMPT_TEMPLATE.format(regions_summary=regions_summary)

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

    # Parse JSON array
    result = json.loads(content)
    if isinstance(result, dict):
        # Might be wrapped in an object
        for key in result:
            if isinstance(result[key], list):
                return result[key]
    if isinstance(result, list):
        return result

    raise ValueError(f"Unexpected response format: {type(result)}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def main():
    parser = argparse.ArgumentParser(description="Extract text styles using Poe AI")
    parser.add_argument("--workspace", required=True, help="Workspace directory path")
    parser.add_argument("--pages", required=True, type=int, help="Total page count")
    parser.add_argument(
        "--poe-api-key", default=None, help="Poe API key (or set POE_API_KEY env)"
    )
    args = parser.parse_args()

    workspace = Path(args.workspace)
    total_pages = args.pages

    # Load OCR data
    ocr_path = workspace / "ocr-output" / "input" / "input.json"
    if not ocr_path.exists():
        print(f"Error: OCR JSON not found at {ocr_path}", file=sys.stderr)
        sys.exit(1)

    with open(ocr_path) as f:
        ocr_data = json.load(f)

    # Load API key
    api_key = load_api_key(args.poe_api_key, workspace)
    api_available = api_key is not None and requests is not None

    if not api_available:
        if not api_key:
            print(
                "Warning: POE_API_KEY not found. Using default styles only.",
                file=sys.stderr,
            )
        if requests is None:
            print(
                "Warning: requests library not available. Using default styles only.",
                file=sys.stderr,
            )

    output_dir = workspace / "ocr-output" / "input"
    output_dir.mkdir(parents=True, exist_ok=True)

    for page_num in range(1, total_pages + 1):
        print(f"Processing page {page_num}/{total_pages}...", end=" ")

        if page_num - 1 >= len(ocr_data):
            print("skipped (no OCR data)")
            continue

        page_regions = ocr_data[page_num - 1]

        if not api_available:
            styles = generate_defaults_for_page(page_regions)
            out_path = output_dir / f"style-page-{page_num}.json"
            out_path.write_text(json.dumps(styles, ensure_ascii=False, indent=2))
            print(f"defaults ({len(styles)} regions)")
            continue

        # Load page image
        # Handle zero-padded filenames for ≥10 pages
        img_path = workspace / "input-pdf-rendered-pngs" / f"page-{page_num}.png"
        if not img_path.exists():
            # Try zero-padded
            img_path = (
                workspace / "input-pdf-rendered-pngs" / f"page-{page_num:02d}.png"
            )
        if not img_path.exists():
            print(f"image not found, using defaults")
            styles = generate_defaults_for_page(page_regions)
            out_path = output_dir / f"style-page-{page_num}.json"
            out_path.write_text(json.dumps(styles, ensure_ascii=False, indent=2))
            continue

        try:
            page_image_b64 = encode_image(img_path)
            regions_summary = build_regions_summary(page_regions)
            vlm_result = call_poe_api(api_key, page_image_b64, regions_summary)

            # Expand and validate entries
            styles = [expand_style_entry(entry, page_regions) for entry in vlm_result]

            # Fill in any missing regions with defaults
            seen_indices = {s["region_index"] for s in styles}
            for r in page_regions:
                if r["index"] not in seen_indices:
                    native = r.get("native_label", "text")
                    defaults = STYLE_DEFAULTS.get(native, STYLE_DEFAULTS["text"])
                    styles.append(
                        {
                            "region_index": r["index"],
                            "font_size_pt": defaults["fs"],
                            "bold": defaults["b"],
                            "italic": False,
                            "underline": False,
                            "color_rgb": defaults["c"],
                            "alignment": defaults["a"],
                        }
                    )

            # Second pass: extract cell-level styles for table regions
            table_regions = [r for r in page_regions if r.get("label") == "table"]
            for t_region in table_regions:
                t_idx = t_region.get("index", -1)
                content = t_region.get("content", "") or ""
                num_rows, num_cols = extract_table_metadata(content)
                if num_rows == 0 or num_cols == 0:
                    continue
                try:
                    cell_overrides = call_poe_table_cell_styles(
                        api_key, page_image_b64, num_rows, num_cols
                    )
                    # Attach cell_overrides to the matching style entry
                    has_data = (
                        cell_overrides.get("col_colors")
                        or cell_overrides.get("row_colors")
                        or cell_overrides.get("cell_colors")
                        or cell_overrides.get("keyword_styles")
                    )
                    if has_data:
                        for s in styles:
                            if s["region_index"] == t_idx:
                                s["cell_overrides"] = cell_overrides
                                break
                        print(
                            f"    table region {t_idx}: cell overrides extracted",
                            end="",
                        )
                except Exception as te:
                    print(
                        f"    table region {t_idx}: cell style extraction failed: {te}",
                        file=sys.stderr,
                        end="",
                    )

            out_path = output_dir / f"style-page-{page_num}.json"
            out_path.write_text(json.dumps(styles, ensure_ascii=False, indent=2))
            print(f"API OK ({len(styles)} regions)")

        except Exception as e:
            print(f"API failed: {e}. Using defaults.", file=sys.stderr)
            styles = generate_defaults_for_page(page_regions)
            out_path = output_dir / f"style-page-{page_num}.json"
            out_path.write_text(json.dumps(styles, ensure_ascii=False, indent=2))

    print("Style extraction complete.")


if __name__ == "__main__":
    main()
