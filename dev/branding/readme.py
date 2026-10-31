"""Draws the README wordmark in assets/readme: DEVICE TWEAKER in the bold, letter-spaced mono of the
program title, in a dark and a light version. It is the only drawn element of the README; everything
else there is plain GitHub text.

    python dev/branding/readme.py --fonts <folder with JetBrainsMono-*.ttf>

The SVG is placed 1:1 (width attribute = SVG width) and the baseline sits on a whole pixel,
so GitHub shows it without resampling.
"""
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from render import Canvas, Font  # noqa: E402

ROOT = Path(__file__).resolve().parents[2]
WIDTH = 560
WORD = "DEVICE TWEAKER"

# Text colours of dev/site/site.css.
THEMES = {"dark": "#f0f0f0", "light": "#0a0a0c"}


def wordmark(font, color):
    size = 100
    while font.width(WORD, size, round(size * 0.06)) > WIDTH - 8:
        size -= 1
    tracking = round(size * 0.06)
    cap = font.cap_height(size)
    c = Canvas(WIDTH, round(cap) + 24, WORD)
    c.text(font, WORD, size, round((WIDTH - font.width(WORD, size, tracking)) / 2), round(12 + cap), color, tracking)
    return c.svg()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--fonts", required=True, type=Path)
    parser.add_argument("--out", default=ROOT / "assets" / "readme", type=Path)
    args = parser.parse_args()
    found = sorted(args.fonts.rglob("JetBrainsMono-Bold.ttf"))
    if not found:
        raise SystemExit("Font not found: JetBrainsMono-Bold.ttf")
    font = Font(found[0])
    args.out.mkdir(parents=True, exist_ok=True)
    for old in args.out.glob("*.svg"):
        old.unlink()
    for theme, color in THEMES.items():
        (args.out / f"wordmark-{theme}.svg").write_text(wordmark(font, color), encoding="utf-8")
    print(f"wordmark: {WIDTH} px")


if __name__ == "__main__":
    main()
