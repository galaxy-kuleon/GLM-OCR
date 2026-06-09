#!/usr/bin/env python3
"""Visual evaluator for PDF→DOCX battle outputs.

Renders source PDF and generated DOCX to PNG pages, then asks a local
OpenAI-compatible vision model to compare visual fidelity.
"""
from __future__ import annotations

import argparse
import base64
import json
import subprocess
import sys
import time
from pathlib import Path
from typing import Any

import requests

API_BASE_DEFAULT = "http://localhost:11234/v1"
API_KEY_DEFAULT = "change-me-local-key"


def run(cmd: list[str], cwd: Path | None = None, timeout: int = 180) -> None:
    subprocess.run(cmd, cwd=str(cwd) if cwd else None, check=True, timeout=timeout)


def render_pdf_to_pngs(pdf: Path, out_dir: Path, prefix: str = "page", dpi: int = 144, max_pages: int = 3) -> list[Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    # Render all pages, then cap list. pdftocairo naming: prefix-1.png...
    run(["pdftocairo", "-png", "-r", str(dpi), str(pdf), str(out_dir / prefix)], timeout=300)
    pages = sorted(out_dir.glob(f"{prefix}-*.png"))
    return pages[:max_pages]


def convert_docx_to_pdf(docx: Path, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(out_dir), str(docx)], timeout=300)
    pdf = out_dir / f"{docx.stem}.pdf"
    if not pdf.exists():
        candidates = sorted(out_dir.glob("*.pdf"))
        if candidates:
            return candidates[-1]
        raise FileNotFoundError(f"DOCX conversion produced no PDF in {out_dir}")
    return pdf


def image_data_url(path: Path) -> str:
    data = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:image/png;base64,{data}"


def models(api_base: str, api_key: str) -> list[str]:
    r = requests.get(f"{api_base}/models", headers={"Authorization": f"Bearer {api_key}"}, timeout=30)
    r.raise_for_status()
    return [m["id"] for m in r.json().get("data", [])]


def judge_pair(
    source_png: Path,
    output_png: Path,
    model: str,
    api_base: str,
    api_key: str,
    max_tokens: int = 4096,
    timeout: int = 180,
) -> dict[str, Any]:
    prompt = """You are a brutally honest document-conversion visual fidelity judge.
Compare SOURCE PDF page render vs OUTPUT DOCX page render.

Return ONLY valid JSON with this schema:
{
  "score": 0-10,
  "verdict": "pass|borderline|fail",
  "text_accuracy": 0-10,
  "layout_accuracy": 0-10,
  "typography_accuracy": 0-10,
  "table_accuracy": 0-10,
  "image_accuracy": 0-10,
  "major_defects": ["..."],
  "minor_defects": ["..."],
  "actionable_fixes": ["..."]
}

Scoring guidance:
- 9-10: nearly indistinguishable and editable
- 7-8: usable with minor layout/style defects
- 5-6: recognizable but commercial-quality failure
- 0-4: severe missing content/layout collapse
Be strict: missing tables/images, text overlap, wrong page geometry, or missing CJK text are major defects.
"""
    payload = {
        "model": model,
        "messages": [{
            "role": "user",
            "content": [
                {"type": "text", "text": prompt},
                {"type": "text", "text": "SOURCE PDF page render:"},
                {"type": "image_url", "image_url": {"url": image_data_url(source_png)}},
                {"type": "text", "text": "OUTPUT DOCX page render:"},
                {"type": "image_url", "image_url": {"url": image_data_url(output_png)}},
            ],
        }],
        "temperature": 0.0,
        "max_tokens": max_tokens,
    }
    started = time.time()
    r = requests.post(
        f"{api_base}/chat/completions",
        headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
        json=payload,
        timeout=timeout,
    )
    elapsed = time.time() - started
    out: dict[str, Any] = {"model": model, "elapsed_s": round(elapsed, 1), "status_code": r.status_code}
    try:
        r.raise_for_status()
        msg = r.json()["choices"][0]["message"]
        content = msg.get("content") or ""
        out["raw"] = content
        # Extract JSON if wrapped.
        import re
        m = re.search(r"\{.*\}", content, flags=re.S)
        if m:
            out["json"] = json.loads(m.group(0))
        else:
            out["error"] = "no_json"
    except Exception as e:
        out["error"] = str(e)
        out["raw_response"] = r.text[:1000]
    return out


def prepare_pair(source_pdf: Path, output_docx: Path, out_dir: Path, max_pages: int = 3) -> tuple[list[Path], list[Path], Path]:
    source_pages = render_pdf_to_pngs(source_pdf, out_dir / "source_png", "source", max_pages=max_pages)
    output_pdf = convert_docx_to_pdf(output_docx, out_dir / "docx_pdf")
    output_pages = render_pdf_to_pngs(output_pdf, out_dir / "output_png", "output", max_pages=max_pages)
    return source_pages, output_pages, output_pdf


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--source-pdf", type=Path, required=True)
    ap.add_argument("--output-docx", type=Path, required=True)
    ap.add_argument("--out-dir", type=Path, required=True)
    ap.add_argument("--api-base", default=API_BASE_DEFAULT)
    ap.add_argument("--api-key", default=API_KEY_DEFAULT)
    ap.add_argument("--model", action="append", help="Model to test; can repeat. Default: all /models")
    ap.add_argument("--max-pages", type=int, default=1)
    ap.add_argument("--max-tokens", type=int, default=4096)
    args = ap.parse_args()

    source_pages, output_pages, output_pdf = prepare_pair(args.source_pdf, args.output_docx, args.out_dir, args.max_pages)
    if not source_pages or not output_pages:
        raise SystemExit("No rendered pages")
    model_ids = args.model or models(args.api_base, args.api_key)
    results = []
    for model_id in model_ids:
        print(f"Judging with {model_id}...")
        res = judge_pair(source_pages[0], output_pages[0], model_id, args.api_base, args.api_key, args.max_tokens)
        print(json.dumps(res.get("json") or {"error": res.get("error"), "elapsed_s": res.get("elapsed_s")}, ensure_ascii=False)[:500])
        results.append(res)
    out = {"source_pages": [str(p) for p in source_pages], "output_pages": [str(p) for p in output_pages], "output_pdf": str(output_pdf), "results": results}
    args.out_dir.mkdir(parents=True, exist_ok=True)
    (args.out_dir / "visual_eval.json").write_text(json.dumps(out, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"Saved: {args.out_dir / 'visual_eval.json'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
