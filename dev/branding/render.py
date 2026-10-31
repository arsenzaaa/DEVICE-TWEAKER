"""Renders the project mark and diagrams in assets/branding.

Text is converted to outlines, so the SVGs look the same on every system. Font (OFL):
JetBrains Mono 2.304 (https://github.com/JetBrains/JetBrainsMono/releases/tag/v2.304).

    python dev/branding/render.py --fonts <folder with JetBrainsMono-*.ttf>

Everything is black and white, like the program window.
"""
import argparse
import html
from pathlib import Path

from fontTools.pens.svgPathPen import SVGPathPen
from fontTools.pens.transformPen import TransformPen
from fontTools.ttLib import TTFont

# Black and white, like the program window.
THEMES = {
    "dark": {
        "text": "#ededed", "muted": "#8f8f8f", "line": "#333333", "fill": "#111111", "page": "#0a0a0a",
        "usb": "#ededed", "gpu": "#ededed", "net": "#ededed",
    },
    "light": {
        "text": "#111111", "muted": "#6b6b6b", "line": "#cfcfcf", "fill": "#f5f5f5", "page": "#ffffff",
        "usb": "#111111", "gpu": "#111111", "net": "#111111",
    },
}

# The mark is the same on light and dark pages: a dark tile like the program window.
MARK = {"tile": "#0a0a0a", "line": "#2a2a2a", "pin": "#5c5c5c", "usb": "#f2f2f2", "gpu": "#f2f2f2", "net": "#f2f2f2"}

CORE = {"ru": "ЯДРО", "en": "CORE"}

# Cores lit in the mark and their device: the same plan the auto-optimization diagram shows.
LIT = {(0, 2): "usb", (1, 3): "gpu", (2, 1): "gpu", (3, 0): "net"}


def num(value):
    return f"{value:.2f}".rstrip("0").rstrip(".")


class Font:
    def __init__(self, path):
        self.font = TTFont(path)
        self.glyphs = self.font.getGlyphSet()
        self.cmap = self.font.getBestCmap()
        self.upm = self.font["head"].unitsPerEm
        self.hmtx = self.font["hmtx"]

    def _glyph(self, ch):
        return self.cmap.get(ord(ch)) or self.cmap[ord("?")]

    def width(self, value, size, tracking=0.0):
        scale = size / self.upm
        return sum(self.hmtx[self._glyph(ch)][0] * scale for ch in value) + tracking * max(0, len(value) - 1)

    def cap_height(self, size):
        return self.font["OS/2"].sCapHeight * size / self.upm

    def path(self, value, size, x, y, tracking=0.0, anchor="start"):
        scale = size / self.upm
        width = self.width(value, size, tracking)
        cx = x - (width if anchor == "end" else width / 2 if anchor == "middle" else 0)
        pen = SVGPathPen(self.glyphs, ntos=num)
        for ch in value:
            name = self._glyph(ch)
            self.glyphs[name].draw(TransformPen(pen, (scale, 0, 0, -scale, cx, y)))
            cx += self.hmtx[name][0] * scale + tracking
        return pen.getCommands()


class Canvas:
    def __init__(self, width, height, title):
        self.width, self.height, self.title = width, height, title
        self.parts = []

    def text(self, font, value, size, x, y, fill, tracking=0.0, anchor="start"):
        self.parts.append(f'<path d="{font.path(value, size, x, y, tracking, anchor)}" fill="{fill}"/>')

    def rect(self, x, y, w, h, stroke=None, fill="none", rx=3, width=1.0):
        # Strokes sit on half pixels so 1px lines stay sharp.
        s = f' stroke="{stroke}" stroke-width="{num(width)}"' if stroke else ""
        self.parts.append(f'<rect x="{num(x)}" y="{num(y)}" width="{num(w)}" height="{num(h)}" rx="{num(rx)}" fill="{fill}"{s}/>')

    def line(self, d, stroke, width=1.0, cap="butt"):
        self.parts.append(f'<path d="{d}" fill="none" stroke="{stroke}" stroke-width="{num(width)}" stroke-linecap="{cap}" stroke-linejoin="round"/>')

    def svg(self):
        t = html.escape(self.title)
        return (f'<svg xmlns="http://www.w3.org/2000/svg" width="{self.width}" height="{self.height}" '
                f'viewBox="0 0 {self.width} {self.height}" role="img" aria-label="{t}">'
                f'<title>{t}</title>{"".join(self.parts)}</svg>\n')


def mark(th):
    """Project mark: a CPU package with 4 x 4 cores; the cores given to USB, GPU and NET are lit."""
    c = Canvas(128, 128, "DEVICE TWEAKER")
    c.rect(0.75, 0.75, 126.5, 126.5, th["line"], th["tile"], 28, 1.5)
    cell, gap, pad = 12, 4, 8
    size = pad * 2 + cell * 4 + gap * 3
    x = y = (128 - size) / 2
    pins = ""
    for i in range(4):
        p = x + pad + i * (cell + gap) + cell / 2
        pins += (f"M{num(x - 8)} {num(p)}H{num(x - 3)}M{num(x + size + 3)} {num(p)}H{num(x + size + 8)}"
                 f"M{num(p)} {num(y - 8)}V{num(y - 3)}M{num(p)} {num(y + size + 3)}V{num(y + size + 8)}")
    c.line(pins, th["pin"], 2, "round")
    c.rect(x + 1, y + 1, size - 2, size - 2, th["pin"], "none", 7, 2)
    for row in range(4):
        for col in range(4):
            cx, cy = x + pad + col * (cell + gap), y + pad + row * (cell + gap)
            key = LIT.get((row, col))
            if key:
                c.rect(cx, cy, cell, cell, None, th[key], 2.5)
            else:
                c.rect(cx + 0.75, cy + 0.75, cell - 1.5, cell - 1.5, th["pin"], "none", 2, 1.5)
    return c.svg()


def device(c, f, th, key, x, y, w, h, size):
    c.rect(x + 0.75, y + 0.75, w - 1.5, h - 1.5, th["line"], th["fill"], 3, 1.5)
    c.text(f["mono"], key.upper(), size, x + w / 2, y + h / 2 + f["mono"].cap_height(size) / 2, th[key], 0.6, "middle")


# Wide layout for the README column, compact one for phones: canvas width, GPU box width,
# joint x, row box x and width, label x, first CPU cell x, cell step and cell width.
ALLOCATION = {
    "wide": (560, 76, 118, 182, 378, 200, 330, 114, 102),
    "compact": (380, 64, 92, 126, 254, 140, 214, 78, 70),
}


def allocation(f, th, core_label, layout):
    """Auto-optimization: a discrete GPU gets two different physical cores, never two SMT siblings."""
    width, bw, joint, rx, rw, lx, cx, step, cw = ALLOCATION[layout]
    c = Canvas(width, 132, f"GPU: {core_label} 2, {core_label} 4")
    rows = [(26, 2, 4), (106, 4, 8)]
    device(c, f, th, "gpu", 0.5, 48, bw, 36, 15)
    c.line(f"M{bw + 0.5} 66H{joint}" + "".join(f"M{joint} 66V{y}H{rx - 2}" for y, _, _ in rows), th["gpu"], 1.5)
    for y, core, cpu in rows:
        c.parts.append(f'<path d="M{rx - 8} {y - 4.5}L{rx - 1.5} {y}L{rx - 8} {y + 4.5}" fill="none" stroke="{th["gpu"]}" '
                       f'stroke-width="1.5" stroke-linejoin="round" stroke-linecap="round"/>')
        c.rect(rx + 0.75, y - 19.25, rw - 1.5, 38.5, th["line"], th["fill"], 3, 1.5)
        c.text(f["mono"], f"{core_label} {core}", 14, lx, y + f["mono"].cap_height(14) / 2, th["text"], 1.0)
        for i, n in enumerate((cpu, cpu + 1)):
            x = cx + i * step
            label = f"CPU {n}"
            if i == 0:
                c.rect(x, y - 12, cw, 24, None, th["gpu"], 2)
                c.text(f["mono"], label, 13, x + cw / 2, y + f["mono"].cap_height(13) / 2, th["page"], 0.6, "middle")
            else:
                c.rect(x + 0.75, y - 11.25, cw - 1.5, 22.5, th["line"], "none", 2, 1.5)
                c.text(f["mono"], label, 13, x + cw / 2, y + f["mono"].cap_height(13) / 2, th["muted"], 0.6, "middle")
    return c.svg()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--fonts", required=True, type=Path)
    parser.add_argument("--out", default=Path("assets/branding"), type=Path)
    args = parser.parse_args()

    def load(name):
        found = sorted(args.fonts.rglob(name))
        if not found:
            raise SystemExit(f"Font not found: {name}")
        return Font(found[0])

    f = {"mono": load("JetBrainsMono-Bold.ttf")}
    out = args.out
    out.mkdir(parents=True, exist_ok=True)
    for old in out.glob("*.svg"):
        old.unlink()

    (out / "mark.svg").write_text(mark(MARK), encoding="utf-8")
    for theme, th in THEMES.items():
        for lang, core in CORE.items():
            (out / f"cpu-allocation-{lang}-{theme}.svg").write_text(allocation(f, th, core, "wide"), encoding="utf-8")
            (out / f"cpu-allocation-mobile-{lang}-{theme}.svg").write_text(allocation(f, th, core, "compact"), encoding="utf-8")
    print("\n".join(sorted(p.name for p in out.glob("*.svg"))))


if __name__ == "__main__":
    main()
