"""Draws the README wordmark in assets/readme: DEVICE TWEAKER in the bold, letter-spaced mono of the
program title, in a dark and a light version. Text is converted to outlines, so it looks the same on
every system, and the SVG is placed 1:1 in the README (width attribute = SVG width).

    pip install fonttools
    python dev/branding/readme.py --fonts <folder with JetBrainsMono-Bold.ttf>

Font: JetBrains Mono 2.304 (OFL), https://github.com/JetBrains/JetBrainsMono/releases/tag/v2.304
"""
import argparse
from pathlib import Path

from fontTools.pens.svgPathPen import SVGPathPen
from fontTools.pens.transformPen import TransformPen
from fontTools.ttLib import TTFont

ROOT = Path(__file__).resolve().parents[2]
WIDTH = 560
WORD = "DEVICE TWEAKER"
THEMES = {"dark": "#f0f0f0", "light": "#0a0a0c"}


def num(value):
    return f"{value:.2f}".rstrip("0").rstrip(".")


def wordmark(font, color):
    glyphs, cmap = font.getGlyphSet(), font.getBestCmap()
    upm, hmtx = font["head"].unitsPerEm, font["hmtx"]
    names = [cmap[ord(ch)] for ch in WORD]

    def width(size, tracking):
        return sum(hmtx[n][0] for n in names) * size / upm + tracking * (len(names) - 1)

    size = 100
    while width(size, round(size * 0.06)) > WIDTH - 8:
        size -= 1
    tracking = round(size * 0.06)
    cap = font["OS/2"].sCapHeight * size / upm
    height = round(cap) + 24
    x, y, scale = round((WIDTH - width(size, tracking)) / 2), round(12 + cap), size / upm
    pen = SVGPathPen(glyphs, ntos=num)
    for name in names:
        glyphs[name].draw(TransformPen(pen, (scale, 0, 0, -scale, x, y)))
        x += hmtx[name][0] * scale + tracking
    return (f'<svg xmlns="http://www.w3.org/2000/svg" width="{WIDTH}" height="{height}" viewBox="0 0 {WIDTH} {height}" '
            f'role="img" aria-label="{WORD}"><title>{WORD}</title><path d="{pen.getCommands()}" fill="{color}"/></svg>\n')


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--fonts", required=True, type=Path)
    parser.add_argument("--out", default=ROOT / "assets" / "readme", type=Path)
    args = parser.parse_args()
    found = sorted(args.fonts.rglob("JetBrainsMono-Bold.ttf"))
    if not found:
        raise SystemExit("Font not found: JetBrainsMono-Bold.ttf")
    font = TTFont(found[0])
    args.out.mkdir(parents=True, exist_ok=True)
    for theme, color in THEMES.items():
        (args.out / f"wordmark-{theme}.svg").write_text(wordmark(font, color), encoding="utf-8")
    print(f"wordmark: {WIDTH} px")


if __name__ == "__main__":
    main()
