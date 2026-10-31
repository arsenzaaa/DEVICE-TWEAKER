"""Frames the program screenshots for the README: rounded corners, a thin edge
and a soft shadow on a transparent background, so they sit well on light and dark pages.

    python dev/branding/screens.py

Reads assets/screenshots/*.png (captured by dev/qa/Capture-Showcase.ps1) and writes
assets/screenshots/framed/<name>.png (full size, for high-density screens) and
assets/screenshots/readme/<name>.png, reduced in advance to exactly the 838 px README column of
github.com, so the browser shows it 1:1 instead of resampling it.
"""
from pathlib import Path

from PIL import Image, ImageChops, ImageDraw, ImageFilter

ROOT = Path(__file__).resolve().parents[2]
SOURCE = ROOT / "assets" / "screenshots"
OUT = SOURCE / "framed"
README_OUT = SOURCE / "readme"
COLUMN = 838

# radius, padding, shadow blur, shadow offset, shadow alpha
FULL = (14, 56, 26, 18, 110)
COMPACT = (8, 20, 11, 6, 100)
EDGE = (255, 255, 255, 34)


def rounded_mask(size, radius, scale=4):
    # Drawn at 4x and reduced, so the corners are antialiased.
    w, h = size
    big = Image.new("L", (w * scale, h * scale), 0)
    ImageDraw.Draw(big).rounded_rectangle((0, 0, w * scale - 1, h * scale - 1), radius * scale, fill=255)
    return big.resize(size, Image.LANCZOS)


def frame(shot, style=FULL):
    radius, pad, blur, offset, alpha = style
    shot = shot.convert("RGBA")
    w, h = shot.size
    mask = rounded_mask((w, h), radius)
    shot.putalpha(ImageChops.multiply(shot.getchannel("A"), mask))

    edge = Image.new("RGBA", (w * 4, h * 4), (0, 0, 0, 0))
    ImageDraw.Draw(edge).rounded_rectangle((2, 2, w * 4 - 3, h * 4 - 3), radius * 4, outline=EDGE, width=4)
    shot.alpha_composite(edge.resize((w, h), Image.LANCZOS))

    canvas = Image.new("RGBA", (w + pad * 2, h + pad * 2), (0, 0, 0, 0))
    shadow = Image.new("L", canvas.size, 0)
    shadow.paste(mask.point(lambda v: v * alpha // 255), (pad, pad + offset))
    shadow = shadow.filter(ImageFilter.GaussianBlur(blur))
    canvas.putalpha(shadow)
    canvas.alpha_composite(shot, (pad, pad))
    return canvas


def reduced(shot):
    """Lanczos down to the README column, then a light unsharp mask to keep the program text legible."""
    width = COLUMN - COMPACT[1] * 2
    height = round(shot.height * width / shot.width)
    small = shot.convert("RGB").resize((width, height), Image.LANCZOS)
    return small.filter(ImageFilter.UnsharpMask(radius=0.8, percent=60, threshold=0))



def main():
    for folder in (OUT, README_OUT):
        folder.mkdir(exist_ok=True)
        for old in folder.glob("*.png"):
            old.unlink()
    for path in sorted(SOURCE.glob("*.png")):
        shot = Image.open(path)
        framed = frame(shot)
        framed.save(OUT / path.name, optimize=True)
        small = frame(reduced(shot), COMPACT)
        small.save(README_OUT / path.name, optimize=True)
        print(path.name, framed.size, small.size)


if __name__ == "__main__":
    main()
