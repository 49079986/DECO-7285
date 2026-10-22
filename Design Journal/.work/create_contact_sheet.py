from pathlib import Path
from PIL import Image, ImageDraw, ImageFont, ImageOps

src = Path(r"C:\Users\yuyan\Desktop\S4\DECO 7285\img")
out = Path(r"C:\Users\yuyan\Desktop\S4\DECO 7285\Design Journal\.work\contact_sheet.png")
thumb_w, thumb_h = 360, 240
label_h = 34
cols = 3
files = sorted(src.glob("*.jpeg"), key=lambda p: int(p.stem))
rows = (len(files) + cols - 1) // cols
canvas = Image.new("RGB", (cols * thumb_w, rows * (thumb_h + label_h)), "white")
draw = ImageDraw.Draw(canvas)
font = ImageFont.load_default()

for idx, path in enumerate(files):
    image = Image.open(path).convert("RGB")
    fitted = ImageOps.contain(image, (thumb_w - 20, thumb_h - 20))
    x0 = (idx % cols) * thumb_w
    y0 = (idx // cols) * (thumb_h + label_h)
    x = x0 + (thumb_w - fitted.width) // 2
    y = y0 + (thumb_h - fitted.height) // 2
    canvas.paste(fitted, (x, y))
    draw.rectangle((x0, y0, x0 + thumb_w - 1, y0 + thumb_h + label_h - 1), outline="#B8B8B8")
    draw.text((x0 + 10, y0 + thumb_h + 8), path.name, fill="black", font=font)

canvas.save(out)
print(out)
