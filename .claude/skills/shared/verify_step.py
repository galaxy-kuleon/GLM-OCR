#!/usr/bin/env python3
"""verify_step.py - Unified step verification for all skills.

Checks expected outputs after each pipeline step. Returns clear OK/FAIL
with prescriptive fix suggestions. Agent runs one command per step.

Usage:
    uv run .claude/skills/shared/verify_step.py --step B2 --workspace /path/to/workspace
    uv run .claude/skills/shared/verify_step.py --step A1 --workspace /path/to/workspace --docx-path /path/to/file.docx
"""

import argparse
import glob
import json
import os
import sys


def _ok(step: str, msg: str) -> None:
    print(f"{step} OK: {msg}")


def _fail(step: str, msg: str, fix: str) -> None:
    print(f"{step} FAIL: {msg}", file=sys.stderr)
    print(f"FIX: {fix}", file=sys.stderr)
    sys.exit(1)


def _file_exists(path: str) -> bool:
    return os.path.isfile(path)


def _count_glob(pattern: str) -> int:
    return len(glob.glob(pattern))


def verify_a1(ws: str, args: argparse.Namespace) -> None:
    docx = args.docx_path or ""
    if not docx:
        _fail("A1", "No --docx-path provided", "Pass --docx-path to verify_step.py")
    if _file_exists(docx):
        size = os.path.getsize(docx)
        _ok("A1", f"{docx} ({size} bytes)")
    else:
        _fail("A1", f"DOCX not found: {docx}",
              "Check soffice/pandoc conversion completed. Re-run the A1 command.")


def verify_a2a(ws: str, args: argparse.Namespace) -> None:
    path = os.path.join(ws, "texts.json")
    if not _file_exists(path):
        _fail("A2a", "texts.json missing",
              "Re-run extract_docx_texts.py. Check input DOCX is valid.")
    data = json.load(open(path, encoding="utf-8"))
    segs = data.get("total_segments", 0)
    if segs == 0:
        _fail("A2a", "texts.json has 0 segments",
              "Input DOCX may be empty or image-only. Check the file.")
    _ok("A2a", f"{segs} segments extracted")


def verify_a2c(ws: str, args: argparse.Namespace) -> None:
    path = os.path.join(ws, "translated-output.docx")
    if not _file_exists(path):
        _fail("A2c", "translated-output.docx missing",
              "Re-run apply_docx_translations.py. Check translations.json exists.")
    size = os.path.getsize(path)
    if size < 500:
        _fail("A2c", f"translated-output.docx too small ({size} bytes)",
              "Translation may have failed. Check translations.json completeness.")
    _ok("A2c", f"{size} bytes")


def verify_b1(ws: str, args: argparse.Namespace) -> None:
    page1 = os.path.join(ws, "input-images", "page-1.png")
    info = os.path.join(ws, "image-info.json")
    if not _file_exists(page1):
        _fail("B1", "input-images/page-1.png missing",
              "Check pdftocairo/resize_images completed. Check input file is valid.")
    if not _file_exists(info):
        _fail("B1", "image-info.json missing",
              "Re-run create_pdf_image_info.py or resize_images.py.")
    count = _count_glob(os.path.join(ws, "input-images", "page-*.png"))
    _ok("B1", f"{count} pages prepared")


def verify_b2(ws: str, args: argparse.Namespace) -> None:
    path = os.path.join(ws, "ocr-output", "input", "input.json")
    if not _file_exists(path):
        _fail("B2", "ocr-output/input/input.json missing",
              "Re-run glmocr. Check input images exist. Run with --log-level DEBUG.")
    data = json.load(open(path, encoding="utf-8"))
    if not isinstance(data, list) or len(data) == 0:
        _fail("B2", "OCR data is empty",
              "glmocr produced no results. Check input image quality.")
    regions = sum(len(p) for p in data)
    _ok("B2", f"{len(data)} pages, {regions} total regions")


def verify_b3(ws: str, args: argparse.Namespace) -> None:
    count = _count_glob(os.path.join(ws, "dsl-vlm", "page-*.xml"))
    if count == 0:
        _fail("B3", "No dsl-vlm/page-*.xml files",
              "VLM generation failed. Check VLM server is running. Re-run vlm_generate_dsl.py.")
    _ok("B3", f"{count} VLM XML files generated")


def verify_b4(ws: str, args: argparse.Namespace) -> None:
    count = _count_glob(os.path.join(ws, "dsl", "page-*.xml"))
    if count == 0:
        _fail("B4", "No dsl/page-*.xml files",
              "Merge failed. Check dsl-vlm/ files exist. Re-run merge step.")
    _ok("B4", f"{count} merged XML files")


def verify_b5(ws: str, args: argparse.Namespace) -> None:
    path = os.path.join(ws, "output.docx")
    if not _file_exists(path):
        _fail("B5", "output.docx missing",
              "Re-run dsl_to_docx.py. Check dsl/page-*.xml files are valid XML.")
    size = os.path.getsize(path)
    if size < 1000:
        _fail("B5", f"output.docx too small ({size} bytes)",
              "DOCX assembly may have failed. Check dsl/ XML validity.")
    _ok("B5", f"{size} bytes")


def verify_b7a(ws: str, args: argparse.Namespace) -> None:
    path = os.path.join(ws, "texts.json")
    if not _file_exists(path):
        _fail("B7a", "texts.json missing",
              "Re-run extract_dsl_texts.py. Check dsl/page-*.xml exist.")
    data = json.load(open(path, encoding="utf-8"))
    segs = data.get("total_segments", 0)
    if segs == 0:
        _fail("B7a", "texts.json has 0 segments",
              "No translatable text found. Check dsl/ XML files have text content.")
    _ok("B7a", f"{segs} segments extracted")


def verify_b7d(ws: str, args: argparse.Namespace) -> None:
    path = os.path.join(ws, "translated-output.docx")
    if not _file_exists(path):
        _fail("B7d", "translated-output.docx missing",
              "Re-run dsl_to_docx.py with --dsl-dir dsl-translated.")
    size = os.path.getsize(path)
    if size < 1000:
        _fail("B7d", f"translated-output.docx too small ({size} bytes)",
              "Check dsl-translated/ XML files. Re-run apply_dsl_translations.py then dsl_to_docx.py.")
    _ok("B7d", f"{size} bytes")


# Step registry: maps step name to verification function
_STEPS = {
    "A1": verify_a1,
    "A2a": verify_a2a,
    "A2c": verify_a2c,
    "B1": verify_b1,
    "B2": verify_b2,
    "B3": verify_b3,
    "B4": verify_b4,
    "B5": verify_b5,
    "B7a": verify_b7a,
    "B7d": verify_b7d,
}


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Verify pipeline step outputs. Run after each step to confirm success.",
        epilog="Available steps: " + ", ".join(sorted(_STEPS.keys())),
    )
    parser.add_argument("--step", required=True, choices=sorted(_STEPS.keys()),
                        help="Step to verify (e.g., B2, B5, A1)")
    parser.add_argument("--workspace", required=True,
                        help="Workspace directory path")
    parser.add_argument("--docx-path", default="",
                        help="DOCX file path (only needed for A1 verification)")
    args = parser.parse_args()

    ws = args.workspace
    if not os.path.isdir(ws):
        print(f"ERROR: workspace not found: {ws}", file=sys.stderr)
        print(f"FIX: Check WORKSPACE variable is set correctly.", file=sys.stderr)
        sys.exit(1)

    _STEPS[args.step](ws, args)


if __name__ == "__main__":
    main()
