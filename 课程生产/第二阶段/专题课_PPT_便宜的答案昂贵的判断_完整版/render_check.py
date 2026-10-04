from pathlib import Path
import fitz
from PIL import Image, ImageOps, ImageDraw

ROOT = Path(__file__).resolve().parent
RENDER = ROOT / "render"
PDF = RENDER / "专题课_便宜的答案昂贵的判断_完整PPT_v1.0.pdf"

doc = fitz.open(PDF)
paths = []
for i, page in enumerate(doc, 1):
    pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5), alpha=False)
    out = RENDER / f"slide{i:02d}.png"
    pix.save(out)
    paths.append(out)

thumb_w, thumb_h = 480, 270
cols, rows = 4, 3
for sheet_no, start in enumerate(range(0, len(paths), cols * rows), 1):
    batch = paths[start:start + cols * rows]
    canvas = Image.new("RGB", (cols * thumb_w, rows * (thumb_h + 26)), "#e5e7eb")
    draw = ImageDraw.Draw(canvas)
    for idx, image_path in enumerate(batch):
        img = Image.open(image_path).convert("RGB")
        img.thumbnail((thumb_w - 8, thumb_h - 8))
        x = (idx % cols) * thumb_w + (thumb_w - img.width) // 2
        y = (idx // cols) * (thumb_h + 26) + 4
        canvas.paste(img, (x, y))
        draw.text((x, y + thumb_h), f"{start + idx + 1:02d}", fill="#111827")
    canvas.save(RENDER / f"contact_{sheet_no:02d}.png")

print(f"rendered {len(paths)} slides")
