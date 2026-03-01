"""
prepare_images.py
-----------------
將使用者提供的圖像排序、轉 PNG、測量尺寸，生成 workspace 初始結構。

Args:
  --images "img1.jpg,img2.jpg,..."   comma-separated 絕對路徑，已由 SKILL 排序
  --workspace <absolute-path>

Output: $WORKSPACE/image-info.json
{
  "page_count": N,
  "pages": [
    {"page":1,"source":"scan.jpg","ext":"jpg",
     "width_pts":595.3,"height_pts":841.9,"dpi":200,
     "ocr_input":"ocr-input/page-1/input.jpg"},
    ...
  ]
}
"""

import argparse
import json
import shutil
from pathlib import Path

from PIL import Image


def px_to_pts(px: float, dpi: float) -> float:
    return px * 72.0 / dpi


def main():
    parser = argparse.ArgumentParser(description="Prepare images for image-to-docx pipeline")
    parser.add_argument("--images", required=True, help="Comma-separated list of image paths")
    parser.add_argument("--workspace", required=True, help="Absolute path to workspace directory")
    args = parser.parse_args()

    workspace = Path(args.workspace)
    image_paths = [Path(p.strip()) for p in args.images.split(",") if p.strip()]

    if not image_paths:
        raise ValueError("No images provided")

    # 建立必要目錄
    png_dir = workspace / "input-pdf-rendered-pngs"
    png_dir.mkdir(parents=True, exist_ok=True)

    pages_info = []

    for idx, img_path in enumerate(image_paths, start=1):
        if not img_path.exists():
            raise FileNotFoundError(f"Image not found: {img_path}")

        ext = img_path.suffix.lstrip(".").lower()
        if ext == "jpeg":
            ext = "jpg"

        # 讀取圖像取得尺寸與 DPI
        with Image.open(img_path) as img:
            width_px, height_px = img.size
            dpi_info = img.info.get("dpi")
            if dpi_info:
                # dpi_info 可能是 tuple 或 IFDRational
                try:
                    dpi = float(dpi_info[0])
                except (TypeError, IndexError):
                    dpi = float(dpi_info)
                # 某些圖像 DPI 值為 0 或極小，fallback 到 200
                if dpi < 1:
                    dpi = 200.0
            else:
                dpi = 200.0

        width_pts = px_to_pts(width_px, dpi)
        height_pts = px_to_pts(height_px, dpi)

        # 儲存 PNG 到 input-pdf-rendered-pngs/page-{N}.png
        png_out = png_dir / f"page-{idx}.png"
        with Image.open(img_path) as img:
            # 確保 RGBA 轉成 RGB（PNG 儲存）
            if img.mode in ("RGBA", "LA", "P"):
                img = img.convert("RGB")
            img.save(png_out, format="PNG")

        # 複製原始圖像到 ocr-input/page-{N}/input.{ext}
        ocr_input_dir = workspace / "ocr-input" / f"page-{idx}"
        ocr_input_dir.mkdir(parents=True, exist_ok=True)
        ocr_input_path = ocr_input_dir / f"input.{ext}"
        shutil.copy2(img_path, ocr_input_path)

        ocr_input_rel = f"ocr-input/page-{idx}/input.{ext}"

        pages_info.append({
            "page": idx,
            "source": img_path.name,
            "ext": ext,
            "width_pts": round(width_pts, 2),
            "height_pts": round(height_pts, 2),
            "dpi": dpi,
            "ocr_input": ocr_input_rel,
        })

        print(f"[prepare] Page {idx}: {img_path.name} → {width_px}x{height_px}px @ {dpi}dpi → {width_pts:.1f}x{height_pts:.1f}pts")

    image_info = {
        "page_count": len(pages_info),
        "pages": pages_info,
    }

    info_path = workspace / "image-info.json"
    info_path.write_text(json.dumps(image_info, ensure_ascii=False, indent=2))
    print(f"[prepare] Written: {info_path}")
    print(f"[prepare] Total pages: {len(pages_info)}")


if __name__ == "__main__":
    main()
