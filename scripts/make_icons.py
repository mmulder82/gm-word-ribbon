"""Generate simple branded PNG icons for the Office Add-in ribbon buttons.
Brand color: #441D42 (Gemeente Gooise Meren), font family intent: Corbel
(falls back to a bundled sans-serif since Corbel is not installed on Linux).
"""
from PIL import Image, ImageDraw, ImageFont
import os

BRAND = (68, 29, 66)      # #441D42
WHITE = (255, 255, 255)

OUT_DIR = os.path.join(os.path.dirname(__file__), "..", "assets")
os.makedirs(OUT_DIR, exist_ok=True)

SIZES = [16, 32, 64, 80]

def find_font(size):
    candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
    ]
    for c in candidates:
        if os.path.exists(c):
            return ImageFont.truetype(c, size)
    return ImageFont.load_default()

for size in SIZES:
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    # rounded-square brand tile
    radius = max(2, size // 6)
    draw.rounded_rectangle([0, 0, size - 1, size - 1], radius=radius, fill=BRAND)
    text = "G"
    font = find_font(int(size * 0.62))
    bbox = draw.textbbox((0, 0), text, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    draw.text(((size - tw) / 2 - bbox[0], (size - th) / 2 - bbox[1]), text, font=font, fill=WHITE)
    img.save(os.path.join(OUT_DIR, f"icon-{size}.png"))

print("icons written to", os.path.abspath(OUT_DIR))
