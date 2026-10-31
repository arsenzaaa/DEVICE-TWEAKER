"""Frames the program screenshots for the README and the site: rounded corners, a thin edge
and a soft shadow on a transparent background, so they sit well on light and dark pages.

    python dev/branding/screens.py

Reads assets/screenshots/*.png (captured by dev/qa/Capture-Showcase.ps1) and writes
assets/screenshots/framed/<name>.png.
"""
from pathlib import Path

from PIL import Image, ImageChops, ImageDraw, ImageFilter

ROOT = Path(__file__).resolve().parents[2]
SOURCE = ROOT / "assets" / "screenshots"
OUT = SOURCE / "framed"

RADIUS = 14
PAD = 56
SHADOW_BLUR = 26
SHADOW_OFFSET = 18
SHADOW_ALPHA = 110
EDGE = (255, 255, 255, 34)


def rounded_mask(size, radius, scale=4):
    # Drawn at 4x and reduced, so the corners are antialiased.
    w, h = size
    big = Image.new("L", (w * scale, h * scale), 0)
    ImageDraw.Draw(big).rounded_rectangle((0, 0, w * scale - 1, h * scale - 1), radius * scale, fill=255)
    return big.resize(size, Image.LANCZOS)


def frame(shot):
    shot = shot.convert("RGBA")
    w, h = shot.size
    mask = rounded_mask((w, h), RADIUS)
    shot.putalpha(ImageChops.multiply(shot.getchannel("A"), mask))

    edge = Image.new("RGBA", (w * 4, h * 4), (0, 0, 0, 0))
    ImageDraw.Draw(edge).rounded_rectangle((2, 2, w * 4 - 3, h * 4 - 3), RADIUS * 4, outline=EDGE, width=4)
    shot.alpha_composite(edge.resize((w, h), Image.LANCZOS))

    canvas = Image.new("RGBA", (w + PAD * 2, h + PAD * 2), (0, 0, 0, 0))
    shadow = Image.new("L", canvas.size, 0)
    shadow.paste(mask.point(lambda v: v * SHADOW_ALPHA // 255), (PAD, PAD + SHADOW_OFFSET))
    shadow = shadow.filter(ImageFilter.GaussianBlur(SHADOW_BLUR))
    canvas.putalpha(shadow)
    canvas.alpha_composite(shot, (PAD, PAD))
    return canvas


def main():
    OUT.mkdir(exist_ok=True)
    for old in OUT.glob("*.png"):
        old.unlink()
    for path in sorted(SOURCE.glob("*.png")):
        framed = frame(Image.open(path))
        framed.save(OUT / path.name, optimize=True)
        print(path.name, framed.size)


if __name__ == "__main__":
    main()
