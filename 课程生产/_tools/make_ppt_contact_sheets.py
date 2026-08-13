from pathlib import Path
import sys
from PIL import Image, ImageDraw

src = Path(sys.argv[1])
out_prefix = sys.argv[2]
files = sorted(src.glob('page-*.png'))
if not files:
    files = sorted(src.glob('*.png'))
cols, rows, tw, th = 4, 3, 320, 180

for k in range(0, len(files), cols * rows):
    group = files[k:k + cols * rows]
    sheet = Image.new('RGB', (cols * tw, rows * th), (235, 235, 235))
    draw = ImageDraw.Draw(sheet)
    for i, path in enumerate(group):
        im = Image.open(path).convert('RGB').resize((tw, th))
        x, y = (i % cols) * tw, (i // cols) * th
        sheet.paste(im, (x, y))
        draw.rectangle((x + 6, y + 6, x + 78, y + 28), fill=(255, 255, 255))
        draw.text((x + 10, y + 9), path.stem, fill=(195, 83, 49))
    sheet.save(f'{out_prefix}_{k // (cols * rows) + 1}.jpg', quality=90)
