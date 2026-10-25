from __future__ import annotations

import argparse
import math
from pathlib import Path

import pypdfium2 as pdfium
from PIL import Image, ImageDraw, ImageFont


def load_font(size: int):
    for path in (
        "/System/Library/Fonts/PingFang.ttc",
        "/System/Library/Fonts/STHeiti Light.ttc",
        "/System/Library/Fonts/Supplemental/Arial.ttf",
    ):
        try:
            return ImageFont.truetype(path, size=size)
        except OSError:
            pass
    return ImageFont.load_default()


def render(pdf_path: Path, output_dir: Path, scale: float, per_sheet: int) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    pages_dir = output_dir / "pages"
    pages_dir.mkdir(parents=True, exist_ok=True)
    pdf = pdfium.PdfDocument(str(pdf_path))
    rendered = []
    for index in range(len(pdf)):
        page = pdf[index]
        bitmap = page.render(scale=scale)
        image = bitmap.to_pil().convert("RGB")
        out = pages_dir / f"page_{index + 1:03d}.png"
        image.save(out, quality=92)
        rendered.append(image)

    cols = 3 if rendered and rendered[0].width > rendered[0].height else 4
    rows = math.ceil(per_sheet / cols)
    thumb_w = 380 if cols == 3 else 260
    ratio = rendered[0].height / rendered[0].width
    thumb_h = int(thumb_w * ratio)
    gutter = 28
    label_h = 34
    bg = (14, 31, 38)
    font = load_font(18)

    for start in range(0, len(rendered), per_sheet):
        chunk = rendered[start:start + per_sheet]
        sheet_w = cols * thumb_w + (cols + 1) * gutter
        sheet_h = rows * (thumb_h + label_h) + (rows + 1) * gutter
        sheet = Image.new("RGB", (sheet_w, sheet_h), bg)
        draw = ImageDraw.Draw(sheet)
        for j, image in enumerate(chunk):
            col = j % cols
            row = j // cols
            x = gutter + col * (thumb_w + gutter)
            y = gutter + row * (thumb_h + label_h + gutter)
            thumb = image.copy()
            thumb.thumbnail((thumb_w, thumb_h), Image.Resampling.LANCZOS)
            sheet.paste(thumb, (x, y))
            draw.text((x, y + thumb_h + 7), f"{start + j + 1:02d}", font=font, fill=(167, 187, 194))
        sheet.save(output_dir / f"contact_{start // per_sheet + 1:02d}.png", quality=92)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("pdf", type=Path)
    parser.add_argument("output_dir", type=Path)
    parser.add_argument("--scale", type=float, default=1.4)
    parser.add_argument("--per-sheet", type=int, default=12)
    args = parser.parse_args()
    render(args.pdf, args.output_dir, args.scale, args.per_sheet)

