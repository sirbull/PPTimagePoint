"""Lager testbilder og en testpresentasjon i tests/fixtures."""
import random
from pathlib import Path
from PIL import Image, ImageDraw, ImageFilter
from pptx import Presentation
from pptx.util import Emu, Inches

out = Path(__file__).parent / "fixtures"
out.mkdir(exist_ok=True)
random.seed(1)

# Groenn bakgrunn med stoey og svak gradient, roed smultring (hull i midten)
W, H = 1600, 1000
im = Image.new("RGB", (W, H))
px = im.load()
for y in range(H):
    for x in range(W):
        g = 175 + int(25 * x / W) + random.randint(-6, 6)
        px[x, y] = (40 + random.randint(-5, 5), g, 60 + random.randint(-5, 5))
d = ImageDraw.Draw(im)
d.ellipse([500, 200, 1100, 800], fill=(200, 40, 50))
d.ellipse([720, 420, 880, 580], fill=(40, 190, 60))  # hull med bakgrunnsfarge
d.rectangle([1200, 300, 1400, 700], fill=(30, 60, 200))
im = im.filter(ImageFilter.GaussianBlur(1.2))
im.save(out / "green.jpg", quality=92)

# Hvit bakgrunn produktbilde
im2 = Image.new("RGB", (900, 900), (252, 252, 252))
d2 = ImageDraw.Draw(im2)
d2.rounded_rectangle([200, 150, 700, 750], 60, fill=(80, 80, 90))
d2.ellipse([350, 300, 550, 500], fill=(240, 200, 30))
im2.save(out / "white.png")

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])
slide.shapes.add_picture(str(out / "white.png"), Inches(0.5), Inches(0.5), Inches(2), Inches(2))
pic = slide.shapes.add_picture(str(out / "green.jpg"), Inches(3), Inches(1), Inches(5), Inches(3.125))
pic.crop_left = 0.25
pic.crop_right = 0.25
pic.name = "Testbilde"
prs.save(out / "test.pptx")
print("picture id", pic.shape_id, "name", pic.name)
