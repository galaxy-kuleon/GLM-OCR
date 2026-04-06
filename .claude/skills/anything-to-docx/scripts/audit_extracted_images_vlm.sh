#!/usr/bin/env bash

set -euo pipefail

usage() {
  cat <<'EOF'
Usage:
  audit_extracted_images_vlm.sh --workspace WORKSPACE [--pdf INPUT.pdf] [--output DIR]
                                [--endpoint URL] [--model MODEL] [--api-key KEY]

What it does:
  1. Renders the source PDF to fresh 200dpi page PNGs with pdftocairo.
  2. Collects extracted images from:
     - glm-ocr crops:   ocr-output/input/imgs/cropped_page*_idx*.jpg
     - pdfimages:       pdf-images/img-*.png
  3. Sends each page image plus that page's extracted candidates to an
     OpenAI-compatible VLM endpoint via curl.
  4. Writes per-page JSON verdicts and a combined summary.

Defaults:
  endpoint: ${VLM_ENDPOINT:-https://openrouter.ai/api/v1/chat/completions}
  model:    ${VLM_MODEL:-z-ai/glm-5v-turbo}
  api key:  VLM_API_KEY / OPENROUTER_API_KEY / .env lookup

Example:
  ./.claude/skills/anything-to-docx/scripts/audit_extracted_images_vlm.sh \
    --workspace "/tmp/my-workspace" \
    --pdf "/path/to/input.pdf"
EOF
}

WORKSPACE=""
INPUT_PDF=""
OUTPUT_DIR=""
ENDPOINT="${VLM_ENDPOINT:-https://openrouter.ai/api/v1/chat/completions}"
MODEL="${VLM_MODEL:-z-ai/glm-5v-turbo}"
API_KEY="${VLM_API_KEY:-${OPENROUTER_API_KEY:-}}"

dotenv_get() {
  local file="$1"
  local key="$2"
  python3 - "$file" "$key" <<'PY'
import pathlib
import sys

path = pathlib.Path(sys.argv[1])
key = sys.argv[2]
if not path.exists():
    raise SystemExit(0)

for raw in path.read_text(encoding='utf-8').splitlines():
    line = raw.strip()
    if not line or line.startswith('#') or '=' not in line:
        continue
    k, v = line.split('=', 1)
    if k.strip() != key:
        continue
    print(v.strip().strip('"\''))
    break
PY
}

while [[ $# -gt 0 ]]; do
  case "$1" in
    --workspace)
      WORKSPACE="$2"
      shift 2
      ;;
    --pdf)
      INPUT_PDF="$2"
      shift 2
      ;;
    --output)
      OUTPUT_DIR="$2"
      shift 2
      ;;
    --endpoint)
      ENDPOINT="$2"
      shift 2
      ;;
    --model)
      MODEL="$2"
      shift 2
      ;;
    --api-key)
      API_KEY="$2"
      shift 2
      ;;
    -h|--help)
      usage
      exit 0
      ;;
    *)
      echo "Unknown argument: $1" >&2
      usage >&2
      exit 1
      ;;
  esac
done

if [[ -z "$WORKSPACE" ]]; then
  echo "ERROR: --workspace is required" >&2
  exit 1
fi

if [[ ! -d "$WORKSPACE" ]]; then
  echo "ERROR: workspace not found: $WORKSPACE" >&2
  exit 1
fi

if [[ -z "$INPUT_PDF" ]]; then
  if [[ -f "$WORKSPACE/input.pdf" ]]; then
    INPUT_PDF="$WORKSPACE/input.pdf"
  else
    echo "ERROR: --pdf not provided and $WORKSPACE/input.pdf not found" >&2
    exit 1
  fi
fi

if [[ ! -f "$INPUT_PDF" ]]; then
  echo "ERROR: input PDF not found: $INPUT_PDF" >&2
  exit 1
fi

if ! command -v pdftocairo >/dev/null 2>&1; then
  echo "ERROR: pdftocairo not found" >&2
  exit 1
fi

if ! command -v curl >/dev/null 2>&1; then
  echo "ERROR: curl not found" >&2
  exit 1
fi

if ! command -v python3 >/dev/null 2>&1; then
  echo "ERROR: python3 not found" >&2
  exit 1
fi

if [[ -z "$OUTPUT_DIR" ]]; then
  OUTPUT_DIR="$WORKSPACE/image-audit-vlm"
fi

if [[ -z "$API_KEY" ]]; then
  API_KEY="$(dotenv_get "$WORKSPACE/.env" "VLM_API_KEY")"
fi

if [[ -z "$API_KEY" ]]; then
  API_KEY="$(dotenv_get "$WORKSPACE/.env" "OPENROUTER_API_KEY")"
fi

if [[ -z "$API_KEY" ]]; then
  API_KEY="$(dotenv_get ".env" "VLM_API_KEY")"
fi

if [[ -z "$API_KEY" ]]; then
  API_KEY="$(dotenv_get ".env" "OPENROUTER_API_KEY")"
fi

if [[ -z "$API_KEY" ]]; then
  echo "ERROR: no OpenRouter API key found. Set --api-key, VLM_API_KEY, OPENROUTER_API_KEY, or .env." >&2
  exit 1
fi

RENDER_DIR="$OUTPUT_DIR/rendered-pages-200dpi"
RESP_DIR="$OUTPUT_DIR/page-reports"
mkdir -p "$RENDER_DIR" "$RESP_DIR"

rm -f "$RENDER_DIR"/page-*.png
pdftocairo -png -r 200 "$INPUT_PDF" "$RENDER_DIR/page"

python3 - "$WORKSPACE" "$RENDER_DIR" "$RESP_DIR" "$ENDPOINT" "$MODEL" "$API_KEY" <<'PY'
import base64
import json
import mimetypes
import os
import pathlib
import re
import subprocess
import sys

workspace = pathlib.Path(sys.argv[1])
render_dir = pathlib.Path(sys.argv[2])
resp_dir = pathlib.Path(sys.argv[3])
endpoint = sys.argv[4]
model = sys.argv[5]
api_key = sys.argv[6]


def natural_page_key(path: pathlib.Path):
    m = re.search(r"(\d+)", path.stem)
    return int(m.group(1)) if m else 0


def data_url(path: pathlib.Path) -> str:
    mime = mimetypes.guess_type(str(path))[0] or "application/octet-stream"
    raw = path.read_bytes()
    return f"data:{mime};base64," + base64.b64encode(raw).decode("ascii")


def ocr_page_num(path: pathlib.Path):
    m = re.search(r"cropped_page(\d+)_idx(\d+)\.jpg$", path.name)
    if not m:
        return None
    return int(m.group(1)) + 1


def pdfimages_page_num(path: pathlib.Path):
    m = re.search(r"img-(\d+)-\d+\.[A-Za-z0-9]+$", path.name)
    if not m:
        return None
    return int(m.group(1))


def build_candidates(page_num: int):
    candidates = []

    ocr_dir = workspace / "ocr-output" / "input" / "imgs"
    if ocr_dir.exists():
        for path in sorted(ocr_dir.glob("cropped_page*_idx*.jpg")):
            if ocr_page_num(path) == page_num:
                candidates.append({
                    "label": f"ocr:{path.name}",
                    "source_type": "glmocr_crop",
                    "path": str(path),
                    "basename": path.name,
                })

    pdf_dir = workspace / "pdf-images"
    if pdf_dir.exists():
        for path in sorted(pdf_dir.glob("img-*.*")):
            if pdfimages_page_num(path) == page_num:
                candidates.append({
                    "label": f"pdfimages:{path.name}",
                    "source_type": "pdfimages_asset",
                    "path": str(path),
                    "basename": path.name,
                })

    return candidates


def strip_fences(text: str) -> str:
    text = text.strip()
    if text.startswith("```"):
        parts = text.split("\n", 1)
        if len(parts) == 2:
            text = parts[1]
        if text.endswith("```"):
            text = text.rsplit("```", 1)[0]
    return text.strip()


SYSTEM_PROMPT = """You are auditing document image extraction quality.

You will receive:
1. A full rasterized PDF page at 200dpi.
2. Several extracted candidate images from the same page, from either glm-ocr crops or pdfimages.

Your job:
- Determine which candidate images actually correspond to visual content on the page.
- Judge whether each candidate is good, damaged, partial, duplicate header/footer noise, or unrelated.
- Identify important visual content on the page that is missing from all candidates.
- Pay special attention to logos, signatures, stamps, charts, icons, and table-embedded pictures.

Return JSON only with this exact shape:
{
  "page": 1,
  "overall_assessment": "short summary",
  "matches": [
    {
      "candidate_label": "ocr:cropped_page0_idx0.jpg",
      "source_type": "glmocr_crop|pdfimages_asset",
      "matched": true,
      "region_description": "what area on page this corresponds to",
      "quality": "good|damaged|partial|duplicate_header_footer|noise|unrelated",
      "keep": true,
      "issues": ["short issue list"]
    }
  ],
  "missing_regions": [
    {
      "description": "important image-like content present on page but not covered by any good candidate",
      "importance": "high|medium|low"
    }
  ],
  "page_verdict": {
    "good_extractions": 0,
    "bad_extractions": 0,
    "missing_important_regions": 0
  }
}

Be strict. If a logo is blurry or visibly degraded in a crop, mark it damaged. If a pdfimages asset is crisp and clearly matches the page graphic, mark it good. Distinguish repeating header/footer assets from actual content worth keeping in DOCX.
"""


def build_payload(page_num: int, page_png: pathlib.Path, candidates):
    content = [
        {
            "type": "text",
            "text": (
                f"Page audit {page_num}. The first image is the full page rasterized from the source PDF at 200dpi. "
                f"Each following image is an extracted candidate from the same page. "
                f"Classify what is correct, damaged, missing, good, and bad."
            ),
        },
        {"type": "image_url", "image_url": {"url": data_url(page_png)}},
    ]

    for candidate in candidates:
        content.append(
            {
                "type": "text",
                "text": (
                    f"Candidate {candidate['label']} from {candidate['source_type']}. "
                    f"Judge whether it correctly matches a visual region on the page and whether its extraction quality is good or bad."
                ),
            }
        )
        content.append(
            {"type": "image_url", "image_url": {"url": data_url(pathlib.Path(candidate['path']))}}
        )

    return {
        "model": model,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": content},
        ],
        "temperature": 0.1,
        "max_tokens": 8192,
        "response_format": {
            "type": "json_schema",
            "json_schema": {
                "name": "page_image_audit",
                "schema": {
                    "type": "object",
                    "properties": {
                        "page": {"type": "integer"},
                        "overall_assessment": {"type": "string"},
                        "matches": {
                            "type": "array",
                            "items": {
                                "type": "object",
                                "properties": {
                                    "candidate_label": {"type": "string"},
                                    "source_type": {"type": "string"},
                                    "matched": {"type": "boolean"},
                                    "region_description": {"type": "string"},
                                    "quality": {"type": "string"},
                                    "keep": {"type": "boolean"},
                                    "issues": {
                                        "type": "array",
                                        "items": {"type": "string"},
                                    },
                                },
                                "required": [
                                    "candidate_label",
                                    "source_type",
                                    "matched",
                                    "region_description",
                                    "quality",
                                    "keep",
                                    "issues",
                                ],
                                "additionalProperties": True,
                            },
                        },
                        "missing_regions": {
                            "type": "array",
                            "items": {
                                "type": "object",
                                "properties": {
                                    "description": {"type": "string"},
                                    "importance": {"type": "string"},
                                },
                                "required": ["description", "importance"],
                                "additionalProperties": True,
                            },
                        },
                        "page_verdict": {
                            "type": "object",
                            "properties": {
                                "good_extractions": {"type": "integer"},
                                "bad_extractions": {"type": "integer"},
                                "missing_important_regions": {"type": "integer"},
                            },
                            "required": [
                                "good_extractions",
                                "bad_extractions",
                                "missing_important_regions",
                            ],
                            "additionalProperties": True,
                        },
                    },
                    "required": [
                        "page",
                        "overall_assessment",
                        "matches",
                        "missing_regions",
                        "page_verdict",
                    ],
                    "additionalProperties": True,
                },
            },
        },
    }


def call_api(payload):
    proc = subprocess.run(
        [
            "curl",
            "-sS",
            "-X",
            "POST",
            endpoint,
            "-H",
            "Content-Type: application/json",
            "-H",
            f"Authorization: Bearer {api_key}",
            "--data-binary",
            "@-",
            "-w",
            "\nHTTP_STATUS:%{http_code}",
        ],
        input=json.dumps(payload),
        text=True,
        capture_output=True,
        timeout=1800,
    )
    if proc.returncode != 0:
        raise RuntimeError(f"curl failed: {proc.stderr.strip()}")

    body, _, status_text = proc.stdout.rpartition("\nHTTP_STATUS:")
    status = int(status_text.strip() or "0")
    if status >= 400:
        raise RuntimeError(f"HTTP {status}: {body.strip()}")

    data = json.loads(body)
    msg = data["choices"][0]["message"]
    content = msg.get("content") or msg.get("reasoning_content") or ""
    return json.loads(strip_fences(content))


page_pngs = sorted(render_dir.glob("page-*.png"), key=natural_page_key)
if not page_pngs:
    raise SystemExit(f"No rendered page PNGs found in {render_dir}")

summary = {
    "workspace": str(workspace),
    "render_dir": str(render_dir),
    "endpoint": endpoint,
    "model": model,
    "pages": [],
}

for page_png in page_pngs:
    page_num = natural_page_key(page_png)
    candidates = build_candidates(page_num)
    out_path = resp_dir / f"page-{page_num}.json"

    if not candidates:
        result = {
            "page": page_num,
            "overall_assessment": "No extracted candidates found for this page.",
            "matches": [],
            "missing_regions": [],
            "page_verdict": {
                "good_extractions": 0,
                "bad_extractions": 0,
                "missing_important_regions": 0,
            },
        }
    else:
        payload = build_payload(page_num, page_png, candidates)
        result = call_api(payload)
        if not isinstance(result, dict):
            raise RuntimeError(f"Unexpected response for page {page_num}: {type(result)!r}")

    result["page"] = page_num
    result["candidates_seen"] = [
        {
            "label": c["label"],
            "source_type": c["source_type"],
            "path": c["path"],
        }
        for c in candidates
    ]
    out_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")

    verdict = result.get("page_verdict") or {}
    summary["pages"].append(
        {
            "page": page_num,
            "report": str(out_path),
            "candidate_count": len(candidates),
            "good_extractions": verdict.get("good_extractions", 0),
            "bad_extractions": verdict.get("bad_extractions", 0),
            "missing_important_regions": verdict.get("missing_important_regions", 0),
            "overall_assessment": result.get("overall_assessment", ""),
        }
    )
    print(f"[audit] page {page_num}: {len(candidates)} candidates -> {out_path}")

summary_path = resp_dir.parent / "summary.json"
summary_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
print(f"[audit] summary: {summary_path}")
PY
