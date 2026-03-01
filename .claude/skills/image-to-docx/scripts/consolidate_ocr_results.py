"""
consolidate_ocr_results.py
--------------------------
合併 per-page OCR 輸出，重命名裁圖索引，建立 pdf-to-docx 相容的
ocr-output/input/ 目錄。

Args:
  --workspace <absolute-path>
  --pages N

Input (per page N=1..PAGE_COUNT):
  $WORKSPACE/ocr-output-pages/page-{N}/input/input.json  → [[regions]]
  $WORKSPACE/ocr-output-pages/page-{N}/input/input.md
  $WORKSPACE/ocr-output-pages/page-{N}/input/imgs/cropped_page0_idx{M}.*

Output:
  $WORKSPACE/ocr-output/input/input.json   → [[page1],[page2],...]
  $WORKSPACE/ocr-output/input/input.md
  $WORKSPACE/ocr-output/input/imgs/cropped_page{N-1}_idx{M}.{ext}
"""

import argparse
import glob
import json
import re
import shutil
from pathlib import Path


def main():
    parser = argparse.ArgumentParser(description="Consolidate per-page OCR results")
    parser.add_argument("--workspace", required=True, help="Absolute path to workspace directory")
    parser.add_argument("--pages", type=int, required=True, help="Total number of pages")
    args = parser.parse_args()

    workspace = Path(args.workspace)
    page_count = args.pages

    out_dir = workspace / "ocr-output" / "input"
    imgs_out_dir = out_dir / "imgs"
    out_dir.mkdir(parents=True, exist_ok=True)
    imgs_out_dir.mkdir(parents=True, exist_ok=True)

    all_pages_regions = []
    all_pages_md = []

    for n in range(1, page_count + 1):
        page_dir = workspace / "ocr-output-pages" / f"page-{n}" / "input"

        # --- 讀取 OCR JSON ---
        json_path = page_dir / "input.json"
        if not json_path.exists():
            raise FileNotFoundError(f"OCR JSON not found: {json_path}")

        data = json.loads(json_path.read_text(encoding="utf-8"))
        # glmocr 單頁輸出格式：[[regions,...]]，取第一頁（index 0）
        if isinstance(data, list) and len(data) > 0:
            regions = data[0]
        else:
            raise ValueError(f"Unexpected JSON structure in {json_path}: {type(data)}")
        all_pages_regions.append(regions)
        print(f"[consolidate] Page {n}: {len(regions)} regions")

        # --- 讀取 Markdown ---
        md_path = page_dir / "input.md"
        if md_path.exists():
            all_pages_md.append(md_path.read_text(encoding="utf-8"))
        else:
            all_pages_md.append(f"<!-- Page {n}: no markdown -->\n")
            print(f"[consolidate] Warning: {md_path} not found, using placeholder")

        # --- 複製裁圖並重命名 ---
        imgs_src_dir = page_dir / "imgs"
        if imgs_src_dir.exists():
            # 匹配 cropped_page0_idx{M}.{ext}
            for src_file in sorted(imgs_src_dir.iterdir()):
                m = re.match(r"cropped_page\d+_idx(\d+)(.*)", src_file.name)
                if not m:
                    # 嘗試更寬鬆的匹配
                    m2 = re.match(r"cropped_.*?idx(\d+)(.*)", src_file.name)
                    if m2:
                        idx_str = m2.group(1)
                        suffix = m2.group(2)
                    else:
                        print(f"[consolidate] Warning: skipping unrecognized file {src_file.name}")
                        continue
                else:
                    idx_str = m.group(1)
                    suffix = m.group(2)

                new_name = f"cropped_page{n-1}_idx{idx_str}{suffix}"
                dst_file = imgs_out_dir / new_name
                shutil.copy2(src_file, dst_file)
                print(f"[consolidate] Copied: {src_file.name} → {new_name}")
        else:
            print(f"[consolidate] Warning: imgs dir not found: {imgs_src_dir}")

    # --- 寫入合併結果 ---
    combined_json = out_dir / "input.json"
    combined_json.write_text(
        json.dumps(all_pages_regions, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )
    print(f"[consolidate] Written: {combined_json} ({page_count} pages)")

    combined_md = out_dir / "input.md"
    combined_md.write_text(
        "\n\n---\n\n".join(all_pages_md),
        encoding="utf-8"
    )
    print(f"[consolidate] Written: {combined_md}")

    print(f"[consolidate] Done. Consolidated {page_count} pages into {out_dir}")


if __name__ == "__main__":
    main()
