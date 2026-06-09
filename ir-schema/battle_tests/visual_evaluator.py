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
    prompt = """You are a brutally honest human reviewer evaluating PDF→DOCX conversion quality.
Compare SOURCE PDF page render vs OUTPUT DOCX page render from an end-user / lawyer / office-worker point of view.

Judge what a real human would notice when opening the generated DOCX:
- Does it look like the original page at first glance?
- Would the user trust it as a converted editable document?
- Are text, tables, images, text boxes, headers/footers, spacing, colors, and typography visually faithful?
- Are defects merely cosmetic, or would they require manual rework?
- Does the rendered DOCX suggest the underlying DOCX is editable without screenshot cheating?

Return ONLY valid JSON with this schema:
{
  "score": 0-10,
  "verdict": "pass|borderline|fail",
  "human_acceptability": "acceptable|needs_minor_touchup|needs_major_rework|unusable",
  "first_impression": "one sentence human reaction, as if a person just opened both files side-by-side",
  "human_delta_summary": "plain-language answer to: how different does the output feel from the original?",
  "would_user_call_this_useless": true|false,
  "uselessness_reason": "if true, explain why a normal user would feel the DOCX is 廢 / not worth using; if false, explain why it is still useful",
  "manual_rework_estimate": "none|minutes|hours|rebuild_from_scratch",
  "top_5_human_noticed_differences": ["the first things a human would notice, ordered by visibility"],
  "text_accuracy": 0-10,
  "layout_accuracy": 0-10,
  "typography_accuracy": 0-10,
  "table_accuracy": 0-10,
  "image_accuracy": 0-10,
  "editability_confidence": 0-10,
  "human_visible_defects": [
    {"severity": "critical|major|minor", "area": "text|layout|typography|table|image|textbox|header_footer|editability", "description": "specific visible defect", "human_impact": "why a user would care", "likely_fix_priority": 1}
  ],
  "major_defects": ["..."],
  "minor_defects": ["..."],
  "actionable_fixes": ["prioritized engineering fixes"],
  "would_accept_for_client_delivery": true|false,
  "client_delivery_reason": "short reason"
}

Scoring guidance from a human's point of view:
- 9-10: nearly indistinguishable, editable, client-deliverable; user feels impressed
- 7-8: clearly useful, only minor touch-up; user would keep editing this DOCX
- 5-6: recognizable but not commercial/client quality; user may feel disappointed but can salvage it
- 3-4: technically produced a DOCX, but a human would likely say it looks bad / needs major rework
- 0-2: severe missing content/layout collapse/unusable; a human would likely call the output 廢
Be strict: missing tables/images, text overlap, wrong page geometry, missing CJK text, displaced text boxes, wrong reading order, clipped text, or screenshot-cheating are critical/major defects. Do not let "DOCX exists" count as quality.
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


def write_human_review_report(out_dir: Path, source_pdf: Path, output_docx: Path, results: list[dict[str, Any]]) -> Path:
    """Write a human-POV markdown report from visual judge JSON outputs."""
    lines: list[str] = []
    lines.append("# Human-POV DOCX Fidelity Review")
    lines.append("")
    lines.append(f"Source PDF: `{source_pdf}`")
    lines.append(f"Output DOCX: `{output_docx}`")
    lines.append("")
    for res in results:
        model = res.get("model", "?")
        j = res.get("json") or {}
        lines.append(f"## Judge: `{model}`")
        lines.append("")
        if not j:
            lines.append(f"- Error: `{res.get('error')}`")
            lines.append("")
            continue
        lines.append(f"- Score: **{j.get('score', '?')}/10**")
        lines.append(f"- Verdict: **{j.get('verdict', '?')}**")
        lines.append(f"- Human acceptability: **{j.get('human_acceptability', '?')}**")
        lines.append(f"- Client-deliverable: **{j.get('would_accept_for_client_delivery', '?')}**")
        if j.get("first_impression"):
            lines.append(f"- First impression: {j['first_impression']}")
        if j.get("human_delta_summary"):
            lines.append(f"- Human delta: {j['human_delta_summary']}")
        if "would_user_call_this_useless" in j:
            lines.append(f"- Would a user call this 廢/useless? **{j.get('would_user_call_this_useless')}**")
        if j.get("uselessness_reason"):
            lines.append(f"- Uselessness reason: {j['uselessness_reason']}")
        if j.get("manual_rework_estimate"):
            lines.append(f"- Manual rework estimate: **{j['manual_rework_estimate']}**")
        if j.get("client_delivery_reason"):
            lines.append(f"- Client delivery reason: {j['client_delivery_reason']}")
        lines.append("")
        lines.append("### Category scores")
        for key in ["text_accuracy", "layout_accuracy", "typography_accuracy", "table_accuracy", "image_accuracy", "editability_confidence"]:
            if key in j:
                lines.append(f"- {key}: {j[key]}/10")
        lines.append("")
        visible_diffs = j.get("top_5_human_noticed_differences") or []
        if visible_diffs:
            lines.append("### Top human-noticed differences")
            for diff in visible_diffs:
                lines.append(f"- {diff}")
            lines.append("")
        visible = j.get("human_visible_defects") or []
        if visible:
            lines.append("### Human-visible defects")
            for d in visible:
                if isinstance(d, dict):
                    lines.append(f"- **{d.get('severity', '?')} / {d.get('area', '?')}**: {d.get('description', '')}")
                    if d.get("human_impact"):
                        lines.append(f"  - Human impact: {d['human_impact']}")
                else:
                    lines.append(f"- {d}")
            lines.append("")
        fixes = j.get("actionable_fixes") or []
        if fixes:
            lines.append("### Actionable fixes")
            for fix in fixes:
                lines.append(f"- {fix}")
            lines.append("")
    report = out_dir / "human_review.md"
    report.write_text("\n".join(lines), encoding="utf-8")
    return report


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
    report = write_human_review_report(args.out_dir, args.source_pdf, args.output_docx, results)
    print(f"Saved: {args.out_dir / 'visual_eval.json'}")
    print(f"Human review: {report}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
