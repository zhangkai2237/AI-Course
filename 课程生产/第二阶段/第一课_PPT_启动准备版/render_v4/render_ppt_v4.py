from pathlib import Path
import sys
import fitz
from PIL import Image, ImageDraw

pdf = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).with_name("第一课_第二阶段项目启动与准备_详细版PPT_v4.0.pdf")
out = pdf.parent
doc = fitz.open(pdf)
thumbs = []
for i, page in enumerate(doc, 1):
    pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5), alpha=False)
    path = out / f"slide{i:02d}.png"
    pix.save(path)
    image = Image.open(path).convert("RGB")
    image.thumbnail((384, 216))
    canvas = Image.new("RGB", (400, 245), "white")
    canvas.paste(image, ((400-image.width)//2, 5))
    ImageDraw.Draw(canvas).text((12, 222), f"{i:02d}", fill="black")
    thumbs.append(canvas)

for start in range(0, len(thumbs), 12):
    group = thumbs[start:start+12]
    sheet = Image.new("RGB", (1600, 735), "#dddddd")
    for j, thumb in enumerate(group):
        x = (j % 4) * 400
        y = (j // 4) * 245
        sheet.paste(thumb, (x, y))
    sheet.save(out / f"contact_{start//12+1:02d}.png")
print(f"rendered {len(thumbs)} slides")
