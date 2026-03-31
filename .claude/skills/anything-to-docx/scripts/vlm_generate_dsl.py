#!/usr/bin/env python3
"""vlm_generate_dsl.py - Send page images to VLM to generate XML DSL.

Reads page images from $WORKSPACE/input-images/page-{N}.png, batches them,
sends each batch to a VLM (LMStudio) API, parses the XML response, and
saves individual page XML files to $WORKSPACE/dsl-vlm/page-{N}.xml.

Usage:
    uv run --with requests,Pillow,lxml \
        .claude/skills/anything-to-docx/scripts/vlm_generate_dsl.py \
        --workspace /path/to/workspace --pages 10 --batch-size 8
"""

import argparse
import base64
import json
import os
import re
import sys
import time

import requests
from lxml import etree

# ---------------------------------------------------------------------------
# VLM Configuration — override any value via environment variable
# ---------------------------------------------------------------------------

# Model profile: "strong" (default) or "weak" (for small models like Qwen3.5-35B-A3B)
# Weak profile: simplified prompt, smaller defaults, more repair
VLM_MODEL_PROFILE = os.environ.get("VLM_MODEL_PROFILE", "strong")

# LMStudio local endpoint; override with VLM_ENDPOINT env var for remote/cloud VLM
VLM_ENDPOINT = os.environ.get("VLM_ENDPOINT", "http://localhost:1234/v1/chat/completions")

# Qwen3.5-122B-A10B: native multimodal VLM with 256K context, 65K output
VLM_MODEL = os.environ.get("VLM_MODEL", "qwen3.5-35b-a3b")

# LMStudio ignores API keys, but OpenAI-compatible API requires the header
VLM_API_KEY = os.environ.get("VLM_API_KEY", "lm-studio")

# 10 min — an 8-page batch generating detailed XML takes ~3-5 min on consumer GPUs
VLM_TIMEOUT = int(os.environ.get("VLM_TIMEOUT", "600"))

# Qwen3.5 supports up to 128K output tokens; large pages need ~3-5K tokens each
VLM_MAX_TOKENS = int(os.environ.get("VLM_MAX_TOKENS", "131072"))

# 2 min between retries — gives GPU time to recover from OOM or thermal throttle
RETRY_DELAY = int(os.environ.get("VLM_RETRY_DELAY", "120"))

# Low temperature for deterministic structured XML output (not creative text)
VLM_TEMPERATURE = float(os.environ.get("VLM_TEMPERATURE", "0.6"))

# Profile-dependent defaults (overridden by explicit env vars)
_PROFILE_DEFAULTS = {
    "strong": {
        "batch_size": 8,
        "max_tokens": 131072,
        "temperature": 0.6,
    },
    "weak": {
        "batch_size": 1,      # one page at a time to ensure complete coverage
        "max_tokens": 131072,  # qwen3.5 supports 128k output
        "temperature": 0.6,    # qwen3.5 35b-a3b needs 0.6 for quality output
    },
}


def _get_profile_default(key):
    """Get profile-specific default for a config key."""
    profile = VLM_MODEL_PROFILE if VLM_MODEL_PROFILE in _PROFILE_DEFAULTS else "strong"
    return _PROFILE_DEFAULTS[profile].get(key)


# ---------------------------------------------------------------------------
# System prompts: strong vs weak
# ---------------------------------------------------------------------------

SYSTEM_PROMPT_STRONG = """\
You are a document layout analyzer. For each page image, produce an XML DSL that precisely describes the visual layout, text content, and styling.

Use this exact XML schema:

<page number="N" width-pts="W" height-pts="H"
      margin-top-cm="M" margin-bottom-cm="M" margin-left-cm="M" margin-right-cm="M"
      font-latin="FONT" font-cjk="FONT">

Elements (in document order):
- <heading level="1|2|3" alignment="left|center|right" font-family="sans|serif|mono" space-before-pt="N" space-after-pt="N">
    <run font-size-pt="N" bold="true|false" italic="true|false" color-rgb="R,G,B" underline="true|false" superscript="true|false" subscript="true|false" strikethrough="true|false" highlight-color="HEXCOLOR">text</run>
  </heading>
- <paragraph alignment="left|center|right|justify" space-before-pt="N" space-after-pt="N" line-spacing="F" indent-left-cm="F" indent-right-cm="F" indent-first-line-cm="F" list-level="N" list-type="bullet|number" font-family="sans|serif|mono" bg-color="HEXCOLOR">
    <run ...>text</run>
  </paragraph>
- <table rows="N" cols="N" border-style="single|double|none" bbox="x1,y1,x2,y2" page-width-pts="W">
    <col-widths>0.25,0.25,0.25,0.25</col-widths>
    <row index="N" is-header="true|false">
      <cell row="N" col="N" colspan="N" rowspan="N" font-size-pt="N" bold="true|false" italic="true|false" alignment="left|center|right" color-rgb="R,G,B" bg-color="HEXCOLOR">text</cell>
    </row>
  </table>
- <image src="PLACEHOLDER" bbox="x1,y1,x2,y2" page-width-pts="W" alignment="left|center|right" />
- <text-frame bbox="x1,y1,x2,y2" page-width-pts="W" page-height-pts="H" has-border="true|false">
    <paragraph ...><run ...>text</run></paragraph>
  </text-frame>
- <horizontal-rule />
- <page-header><paragraph ...><run ...>text</run></paragraph></page-header>
- <page-footer><paragraph ...><run ...>text</run></paragraph></page-footer>
- <side-by-side cols="N">
    <column><paragraph ...><run ...>text</run></paragraph></column>
  </side-by-side>

Rules:
1. bbox values are normalized 0-1000 (not pixels).
2. For images, set src="PLACEHOLDER" — the pipeline will resolve actual paths later.
3. Detect font families visually: sans-serif fonts → "sans", serif fonts → "serif", monospace → "mono".
4. Estimate font sizes in points. Common body text is 10-12pt, headings 14-24pt.
5. Detect text colors as R,G,B values (0-255 each).
6. Detect background colors on paragraphs and table cells as hex (e.g., "F0F0F0").
7. For tables, estimate column width ratios (must sum to 1.0).
8. Detect bullet/numbered lists: set list-level (1=top, 2=nested, etc.) and list-type.
9. Detect indentation visually and estimate in cm.
10. For each page, wrap output in <page number="N" ...> tags.
11. Output ONLY valid XML. No explanations, no markdown code fences."""

# Simplified prompt for weak VLMs (fewer rules, focus on essentials)
SYSTEM_PROMPT_WEAK = """\
Convert each page image to XML. Output ONLY valid XML, no explanations.

Schema:

<page number="N" width-pts="W" height-pts="H" margin-top-cm="1.27" margin-bottom-cm="1.27" margin-left-cm="1.27" margin-right-cm="1.27" font-latin="Arial" font-cjk="SimSun">
  <heading level="1|2|3" alignment="left|center"><run font-size-pt="N" bold="true">text</run></heading>
  <paragraph alignment="left|justify"><run font-size-pt="N">text</run></paragraph>
  <table rows="N" cols="N" bbox="x1,y1,x2,y2" page-width-pts="W">
    <col-widths>0.5,0.5</col-widths>
    <row index="N"><cell row="N" col="N" font-size-pt="9">text</cell></row>
  </table>
  <image src="PLACEHOLDER" bbox="x1,y1,x2,y2" page-width-pts="W" />
</page>

Rules:
1. bbox values: 0-1000 normalized coordinates.
2. Images: always src="PLACEHOLDER".
3. Body text: 10-12pt. Headings: 14-24pt.
4. Tables: estimate column width ratios (sum to 1.0).
5. Wrap each page in <page number="N"> tags.
6. Multiple pages: wrap all in <pages>...</pages>.

Example for a page with a title, paragraph, and table:

<pages>
<page number="1" width-pts="595" height-pts="842" margin-top-cm="1.27" margin-bottom-cm="1.27" margin-left-cm="1.27" margin-right-cm="1.27" font-latin="Arial" font-cjk="SimSun">
  <heading level="1" alignment="center"><run font-size-pt="18" bold="true">Document Title</run></heading>
  <paragraph alignment="left"><run font-size-pt="11">This is body text.</run></paragraph>
  <table rows="2" cols="2" bbox="100,300,900,500" page-width-pts="595">
    <col-widths>0.4,0.6</col-widths>
    <row index="0"><cell row="0" col="0" font-size-pt="9" bold="true">Header 1</cell><cell row="0" col="1" font-size-pt="9" bold="true">Header 2</cell></row>
    <row index="1"><cell row="1" col="0" font-size-pt="9">Data 1</cell><cell row="1" col="1" font-size-pt="9">Data 2</cell></row>
  </table>
</page>
</pages>"""


def get_system_prompt():
    """Return the appropriate system prompt based on model profile."""
    if VLM_MODEL_PROFILE == "weak":
        return SYSTEM_PROMPT_WEAK
    return SYSTEM_PROMPT_STRONG


# Back-compat alias
SYSTEM_PROMPT = SYSTEM_PROMPT_STRONG


# ---------------------------------------------------------------------------
# Pure functions
# ---------------------------------------------------------------------------


def load_image_info(workspace):
    """Load page dimension info from image-info.json. Returns dict."""
    path = os.path.join(workspace, "image-info.json")
    with open(path, "r") as f:
        return json.load(f)


def encode_image_to_base64(image_path):
    """Read a PNG file and return its base64-encoded string."""
    with open(image_path, "rb") as f:
        return base64.b64encode(f.read()).decode("ascii")


def _encode_resized_image(image_path, max_width=768):
    """Resize image to max_width and return base64-encoded JPEG string.

    Used for layout visualization images to reduce VLM inference time.
    """
    from PIL import Image
    import io

    img = Image.open(image_path)
    if img.width > max_width:
        ratio = max_width / img.width
        new_size = (max_width, int(img.height * ratio))
        img = img.resize(new_size, Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=75)
    return base64.b64encode(buf.getvalue()).decode("ascii")


def compute_batches(total_pages, batch_size):
    """Return list of (start, end) tuples, 1-indexed inclusive.

    >>> compute_batches(10, 8)
    [(1, 8), (9, 10)]
    >>> compute_batches(3, 8)
    [(1, 3)]
    >>> compute_batches(0, 8)
    []
    """
    batches = []
    for start in range(1, total_pages + 1, batch_size):
        end = min(start + batch_size - 1, total_pages)
        batches.append((start, end))
    return batches


def build_page_dimensions_text(image_info, start, end):
    """Build the page dimensions section for the user prompt.

    Returns a string like:
        Page 1: 595x842 pts
        Page 2: 595x842 pts
    """
    lines = []
    pages_by_index = {p["index"]: p for p in image_info["pages"]}
    for n in range(start, end + 1):
        page = pages_by_index.get(n)
        if page is not None:
            w = page["width_pts"]
            h = page["height_pts"]
            lines.append(f"Page {n}: {w}x{h} pts")
        else:
            lines.append(f"Page {n}: dimensions unknown")
    return "\n".join(lines)


def build_user_prompt(start, end, dimensions_text):
    """Build the user prompt text for a batch."""
    count = end - start + 1
    layout_hint = ""
    if VLM_MODEL_PROFILE == "weak":
        layout_hint = (
            "\nFor each page you receive TWO images:\n"
            "1. The original page image (for reading text content)\n"
            "2. A layout analysis image with colored boxes showing detected regions:\n"
            "   - doc_title (red) = heading level 1\n"
            "   - paragraph_title (green) = heading level 2\n"
            "   - table (blue) = table element\n"
            "   - image (yellow) = image element\n"
            "   - header/footer (pink/gray) = page header/footer\n"
            "   - text = paragraph element\n"
            "Use the layout image to determine element types and reading order.\n"
        )
    return (
        f"Analyze these {count} document page images (pages {start}-{end}).\n"
        f"For each page, produce an XML <page> element following the schema exactly.\n"
        f"Wrap all pages in a <pages> root element.\n"
        f"{layout_hint}\n"
        f"Page dimensions:\n"
        f"{dimensions_text}"
    )


def build_image_content_items(workspace, start, end):
    """Build the list of image_url content items for the API message.

    Returns a list of dicts suitable for the 'content' array in the
    OpenAI-compatible messages format.

    For weak profiles, also includes layout visualization images from
    glm-ocr (colored bounding boxes with region labels) when available.
    """
    items = []
    for n in range(start, end + 1):
        # Raw page image
        image_path = os.path.join(workspace, "input-images", f"page-{n}.png")
        b64 = encode_image_to_base64(image_path)
        items.append({
            "type": "image_url",
            "image_url": {
                "url": f"data:image/png;base64,{b64}",
            },
        })

        # Layout visualization (weak profile only — helps model see structure)
        # Resized to max 768px wide to keep inference time reasonable
        if VLM_MODEL_PROFILE == "weak":
            page_idx = n - 1  # layout_vis uses 0-based indexing
            layout_path = os.path.join(
                workspace, "ocr-output", "input", "layout_vis",
                f"input_page{page_idx}.jpg"
            )
            if os.path.exists(layout_path):
                b64_layout = _encode_resized_image(layout_path, max_width=768)
                items.append({
                    "type": "image_url",
                    "image_url": {
                        "url": f"data:image/jpeg;base64,{b64_layout}",
                    },
                })
    return items


def build_messages(workspace, image_info, start, end):
    """Build the full messages array for the VLM API call."""
    dimensions_text = build_page_dimensions_text(image_info, start, end)
    user_text = build_user_prompt(start, end, dimensions_text)

    image_items = build_image_content_items(workspace, start, end)
    user_content = [{"type": "text", "text": user_text}] + image_items

    return [
        {"role": "system", "content": get_system_prompt()},
        {"role": "user", "content": user_content},
    ]


def call_vlm(messages):
    """Send request to VLM API. Returns the response dict.

    Retries once on failure after RETRY_DELAY seconds.
    """
    payload = {
        "model": VLM_MODEL,
        "messages": messages,
        "max_tokens": VLM_MAX_TOKENS,
        "temperature": VLM_TEMPERATURE,
        "stream": False,
    }
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {VLM_API_KEY}",
    }

    for attempt in range(2):
        try:
            resp = requests.post(
                VLM_ENDPOINT,
                json=payload,
                headers=headers,
                timeout=VLM_TIMEOUT,
            )
            resp.raise_for_status()
            return resp.json()
        except Exception as e:
            if attempt == 0:
                print(f"  VLM call failed: {e}. Retrying in {RETRY_DELAY}s...",
                      file=sys.stderr)
                time.sleep(RETRY_DELAY)
            else:
                raise


def extract_response_text(response):
    """Extract the text content from the VLM API response dict.

    Handles thinking models (e.g. qwen3.5) that put reasoning in
    'reasoning_content' and the actual answer in 'content'. If 'content'
    is empty but 'reasoning_content' has data, fall back to reasoning_content.
    """
    msg = response["choices"][0]["message"]
    content = msg.get("content", "") or ""
    if content.strip():
        return content
    # Fallback: thinking model may have put everything in reasoning_content
    reasoning = msg.get("reasoning_content", "") or ""
    if reasoning.strip():
        print("  Warning: 'content' was empty, using 'reasoning_content' as fallback", file=sys.stderr)
        return reasoning
    return content


def print_token_usage(response):
    """Print token usage stats if available in the response."""
    usage = response.get("usage")
    if usage:
        prompt_tokens = usage.get("prompt_tokens", "?")
        completion_tokens = usage.get("completion_tokens", "?")
        total_tokens = usage.get("total_tokens", "?")
        print(f"  Tokens: prompt={prompt_tokens}, "
              f"completion={completion_tokens}, total={total_tokens}")


def clean_xml_text(raw_text):
    """Strip markdown code fences and surrounding whitespace from VLM output."""
    text = raw_text.strip()
    # Remove ```xml ... ``` wrapping if present
    text = re.sub(r"^```(?:xml)?\s*\n?", "", text)
    text = re.sub(r"\n?```\s*$", "", text)
    return text.strip()


def clean_ocr_artifacts_in_xml(page_element):
    """Remove common OCR artifacts from text content within a parsed XML page.

    Modifies the element tree in place. Targets:
    - Markdown code fence strings embedded in text (```markdown ```)
    - Stray leading slashes in captions
    """
    for elem in page_element.iter():
        if elem.text:
            original = elem.text
            # Remove inline markdown code fences
            cleaned = re.sub(r"```(?:markdown|xml|json|html)?\s*```", "", original)
            cleaned = re.sub(r"```(?:markdown|xml|json|html)?", "", cleaned)
            cleaned = cleaned.strip()
            if cleaned != original:
                if not cleaned:
                    elem.text = None
                else:
                    elem.text = cleaned


def _repair_xml_text(text):
    """Attempt to repair common XML malformations from weak VLMs.

    Handles: unclosed tags, truncated output, stray text outside elements,
    HTML entities, and missing root wrappers. Returns repaired XML string.
    """
    # Strip any leading non-XML text (VLM explanations before XML)
    first_tag = re.search(r"<(?:pages?|heading|paragraph|table|image)\b", text)
    if first_tag and first_tag.start() > 0:
        text = text[first_tag.start():]

    # Strip any trailing non-XML text after last closing tag
    last_close = text.rfind(">")
    if last_close >= 0 and last_close < len(text) - 1:
        text = text[:last_close + 1]

    # Fix common HTML entity issues in VLM output
    # Replace unescaped & (but not &amp; &lt; &gt; &quot; &apos;)
    text = re.sub(r"&(?!amp;|lt;|gt;|quot;|apos;|#\d+;|#x[0-9a-fA-F]+;)", "&amp;", text)

    # Fix unclosed self-closing tags: <image ... > → <image ... />
    text = re.sub(r"(<image\b[^/]*?)(?<!/)>", r"\1/>", text)
    text = re.sub(r"(<horizontal-rule\b[^/]*?)(?<!/)>", r"\1/>", text)

    # If text is truncated mid-element, try to close open tags
    # Count unclosed <page> tags
    open_pages = len(re.findall(r"<page\b", text))
    close_pages = len(re.findall(r"</page>", text))
    if open_pages > close_pages:
        text += "</page>" * (open_pages - close_pages)

    # Ensure <pages> wrapper
    if not text.startswith("<pages"):
        if text.startswith("<page"):
            text = f"<pages>{text}</pages>"

    # Close <pages> if needed
    if text.startswith("<pages") and not text.rstrip().endswith("</pages>"):
        text += "</pages>"

    return text


def parse_vlm_response(raw_text):
    """Parse VLM XML response into a list of lxml page Elements.

    Handles both <pages>...</pages> wrapper and single <page>...</page>.
    For weak models, applies XML repair before parsing.
    Returns a list of etree.Element objects, one per page.
    """
    text = clean_xml_text(raw_text)

    # Try parsing as-is first (may have <pages> root or single <page>)
    # Wrap in <pages> if it doesn't already have it
    if not text.startswith("<pages"):
        # Check if it's a bare <page ...> element (possibly multiple)
        if text.startswith("<page"):
            text = f"<pages>{text}</pages>"
        else:
            # Try to find <pages> or <page> somewhere in the text
            pages_match = re.search(r"<pages[\s>].*</pages>", text, re.DOTALL)
            if pages_match:
                text = pages_match.group(0)
            else:
                page_match = re.search(r"<page[\s>].*</page>", text, re.DOTALL)
                if page_match:
                    text = f"<pages>{page_match.group(0)}</pages>"

    # Try strict parsing first
    try:
        root = etree.fromstring(text.encode("utf-8"))
    except etree.XMLSyntaxError:
        # Apply XML repair and retry
        print("  XML parse failed, attempting repair...", file=sys.stderr)
        repaired = _repair_xml_text(text)
        try:
            root = etree.fromstring(repaired.encode("utf-8"))
            print("  XML repair successful (strict parse)", file=sys.stderr)
        except etree.XMLSyntaxError:
            # Last resort: lxml recovery mode (tolerates malformed XML)
            try:
                parser = etree.XMLParser(recover=True)
                root = etree.fromstring(repaired.encode("utf-8"), parser=parser)
                print("  XML repair successful (recovery mode)", file=sys.stderr)
            except Exception as e:
                raise ValueError(f"XML parsing failed after repair: {e}")

    # root might be <pages> or <page>
    if root.tag == "pages":
        return list(root.findall("page"))
    elif root.tag == "page":
        return [root]
    else:
        # Unexpected root — look for page children anyway
        pages = root.findall(".//page")
        if pages:
            return pages
        raise ValueError(f"No <page> elements found in VLM response (root tag: {root.tag})")


def indent_xml(element):
    """Pretty-print an lxml element with 2-space indentation.

    Returns the indented XML as a unicode string.
    """
    etree.indent(element, space="  ")
    return etree.tostring(element, pretty_print=True, encoding="unicode")


def save_page_xml(page_element, workspace, page_number):
    """Save a single page element to $WORKSPACE/dsl-vlm/page-{N}.xml."""
    output_dir = os.path.join(workspace, "dsl-vlm")
    os.makedirs(output_dir, exist_ok=True)

    xml_str = indent_xml(page_element)
    output_path = os.path.join(output_dir, f"page-{page_number}.xml")
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(xml_str)
    return output_path


def save_raw_response(raw_text, workspace, batch_index):
    """Save raw VLM response for debugging when parsing fails."""
    output_dir = os.path.join(workspace, "dsl-vlm")
    os.makedirs(output_dir, exist_ok=True)
    path = os.path.join(output_dir, f"raw-batch-{batch_index}.txt")
    with open(path, "w", encoding="utf-8") as f:
        f.write(raw_text)
    return path


def process_batch(workspace, image_info, start, end, batch_index, total_batches):
    """Process a single batch: call VLM, parse response, save XMLs.

    Returns the number of pages successfully saved.
    """
    print(f"Batch {batch_index}/{total_batches}: sending pages {start}-{end}...")

    messages = build_messages(workspace, image_info, start, end)

    try:
        response = call_vlm(messages)
    except Exception as e:
        print(f"  ERROR: VLM call failed after retry: {e}", file=sys.stderr)
        print(f"  FIX: Check LMStudio is running at {VLM_ENDPOINT} with model {VLM_MODEL} loaded.", file=sys.stderr)
        print(f"  FIX: If timeout, increase VLM_TIMEOUT (current: {VLM_TIMEOUT}s).", file=sys.stderr)
        return 0

    print_token_usage(response)
    raw_text = extract_response_text(response)

    try:
        page_elements = parse_vlm_response(raw_text)
    except Exception as e:
        raw_path = save_raw_response(raw_text, workspace, batch_index)
        print(f"  WARNING: XML parsing failed: {e}", file=sys.stderr)
        print(f"  Raw response saved to {raw_path}", file=sys.stderr)
        print(f"  FIX: Inspect {raw_path} for malformed XML. If model output is truncated, reduce batch size or increase VLM_MAX_TOKENS.", file=sys.stderr)
        return 0

    saved_count = 0
    expected_pages = list(range(start, end + 1))

    for i, page_el in enumerate(page_elements):
        # Clean OCR artifacts from text within the page XML
        clean_ocr_artifacts_in_xml(page_el)

        # Determine page number from the element's 'number' attribute,
        # falling back to positional inference
        num_attr = page_el.get("number")
        if num_attr is not None:
            try:
                page_num = int(num_attr)
            except ValueError:
                page_num = expected_pages[i] if i < len(expected_pages) else start + i
        else:
            page_num = expected_pages[i] if i < len(expected_pages) else start + i

        path = save_page_xml(page_el, workspace, page_num)
        print(f"  Saved {path}")
        saved_count += 1

    return saved_count


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def main():
    global VLM_MODEL_PROFILE, VLM_MAX_TOKENS, VLM_TEMPERATURE

    parser = argparse.ArgumentParser(
        description="Send page images to VLM to generate XML DSL"
    )
    parser.add_argument("--workspace", required=True, help="Workspace directory path")
    parser.add_argument("--pages", required=True, type=int, help="Total number of pages")
    parser.add_argument("--batch-size", type=int, default=None,
                        help="Pages per batch (default: profile-dependent)")
    parser.add_argument("--model-profile", choices=["strong", "weak"], default=None,
                        help="Model profile (default: from VLM_MODEL_PROFILE env or 'strong')")
    args = parser.parse_args()

    # Apply model profile from CLI flag (takes precedence over env var)
    if args.model_profile:
        VLM_MODEL_PROFILE = args.model_profile

    # Apply profile defaults for settings not explicitly overridden
    if not os.environ.get("VLM_MAX_TOKENS"):
        VLM_MAX_TOKENS = _get_profile_default("max_tokens")
    if not os.environ.get("VLM_TEMPERATURE"):
        VLM_TEMPERATURE = _get_profile_default("temperature")

    workspace = args.workspace
    total_pages = args.pages
    batch_size = args.batch_size or _get_profile_default("batch_size")

    image_info = load_image_info(workspace)
    batches = compute_batches(total_pages, batch_size)

    print(f"Model profile: {VLM_MODEL_PROFILE}")
    print(f"Processing {total_pages} pages in {len(batches)} batch(es) "
          f"(batch size: {batch_size})")

    total_saved = 0
    for batch_idx, (start, end) in enumerate(batches, start=1):
        saved = process_batch(workspace, image_info, start, end, batch_idx, len(batches))
        total_saved += saved

    print(f"\nDone. Saved {total_saved}/{total_pages} page XML files.")
    if total_saved < total_pages:
        print("WARNING: Some pages were not saved. Check errors above.",
              file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
