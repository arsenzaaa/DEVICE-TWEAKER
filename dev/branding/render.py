"""Renders the README artwork in assets/branding.

Text is converted to outlines, so the SVGs look the same on every system. Fonts (OFL):
Inter 4.1 (https://github.com/rsms/inter/releases/tag/v4.1) for the title and
JetBrains Mono 2.304 (https://github.com/JetBrains/JetBrainsMono/releases/tag/v2.304) for everything else.

    python dev/branding/render.py --fonts <folder with InterDisplay-*.ttf and JetBrainsMono-*.ttf>

Device colours are the category colours of the program (src/GUI/DeviceBadgePalette.cs);
the light theme uses darker shades of the same hues.
"""
import argparse
import html
from pathlib import Path

from fontTools.pens.svgPathPen import SVGPathPen
from fontTools.pens.transformPen import TransformPen
from fontTools.ttLib import TTFont

THEMES = {
    "dark": {
        "text": "#f0f2f4", "muted": "#9198a1", "line": "#3d444d", "fill": "#151b23", "page": "#0d1117",
        "usb": "#48d9d0", "gpu": "#76b900", "net": "#62aef5",
    },
    "light": {
        "text": "#1f2328", "muted": "#59636e", "line": "#c8d1da", "fill": "#f6f8fa", "page": "#ffffff",
        "usb": "#0a8a82", "gpu": "#4f7d00", "net": "#1f6fc4",
    },
}

TEXT = {
    "ru": {
        "buttons": {"download": "СКАЧАТЬ 0.0.4", "settings": "НАСТРОЙКИ", "changelog": "ИЗМЕНЕНИЯ"},
        "core": "ЯДРО",
    },
    "en": {
        "buttons": {"download": "DOWNLOAD 0.0.4", "settings": "SETTINGS", "changelog": "CHANGELOG"},
        "core": "CORE",
    },
}

PLATFORM = "WINDOWS 10 / 11 x64"
FEATURES = "MSI / AFFINITY / RSS / IMOD / ITR"


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


def chip(c, f, th, x, y, assigned):
    """CPU package: 4 x 4 cores with pins, cores filled in the colour of their device."""
    cell, gap, pad = 24, 8, 18
    size = pad * 2 + cell * 4 + gap * 3
    for i in range(4):
        p = y + pad + i * (cell + gap) + cell / 2
        q = x + pad + i * (cell + gap) + cell / 2
        c.line(f"M{num(x - 7)} {num(p)}H{num(x)}M{num(x + size)} {num(p)}H{num(x + size + 7)}"
               f"M{num(q)} {num(y - 7)}V{num(y)}M{num(q)} {num(y + size)}V{num(y + size + 7)}", th["line"], 1.5)
    c.rect(x + 0.75, y + 0.75, size - 1.5, size - 1.5, th["line"], th["fill"], 8, 1.5)
    for row in range(4):
        for col in range(4):
            cx, cy = x + pad + col * (cell + gap), y + pad + row * (cell + gap)
            color = assigned.get((row, col))
            if color:
                c.rect(cx, cy, cell, cell, None, th[color], 2)
            else:
                c.rect(cx + 0.75, cy + 0.75, cell - 1.5, cell - 1.5, th["line"], "none", 2, 1.5)
    return size, pad, cell, gap


def device(c, f, th, key, x, y, w, h, size):
    c.rect(x + 0.75, y + 0.75, w - 1.5, h - 1.5, th["line"], th["fill"], 3, 1.5)
    c.text(f["mono"], key.upper(), size, x + w / 2, y + h / 2 + f["mono"].cap_height(size) / 2, th[key], 0.6, "middle")


def diagram(c, f, th, right, top):
    """USB / GPU / NET wired to the cores they get: GPU on two different physical cores."""
    probe = Canvas(0, 0, "")
    size, pad, cell, gap = chip(probe, f, th, 0, 0, {})
    cx = right - size - 8
    row_y = [top + pad + i * (cell + gap) + cell / 2 for i in range(4)]
    chip(c, f, th, cx, top, {(0, 2): "usb", (1, 3): "gpu", (2, 1): "gpu", (3, 0): "net"})
    bw, bh = 76, 36
    bx = cx - 7 - 172
    joint = bx + bw + 30
    for key, rows in (("usb", [0]), ("gpu", [1, 2]), ("net", [3])):
        mid = sum(row_y[r] for r in rows) / len(rows)
        device(c, f, th, key, bx, mid - bh / 2, bw, bh, 15)
        if len(rows) == 1:
            c.line(f"M{num(bx + bw)} {num(mid)}H{num(cx - 7)}", th[key], 1.5)
        else:
            d = f"M{num(bx + bw)} {num(mid)}H{num(joint)}"
            for r in rows:
                d += f"M{num(joint)} {num(mid)}V{num(row_y[r])}H{num(cx - 7)}"
            c.line(d, th[key], 1.5)
        for r in rows:
            c.parts.append(f'<circle cx="{num(cx - 7)}" cy="{num(row_y[r])}" r="2.5" fill="{th[key]}"/>')


def hero(f, th):
    c = Canvas(1080, 200, "DEVICE TWEAKER")
    c.text(f["label"], PLATFORM, 17, 2, 40, th["muted"], 1.6)
    c.text(f["title"], "DEVICE TWEAKER", 76, 0, 126, th["text"], -1.9)
    c.text(f["label"], FEATURES, 19, 2, 170, th["muted"], 1.4)
    diagram(c, f, th, 1080, 22)
    return c.svg()


def hero_mobile(f, th):
    size = 80
    while f["title"].width("DEVICE TWEAKER", size, -1.6) > 596:
        size -= 1
    label = f["label"].cap_height(21)
    top = 20
    platform = top + label
    title = platform + 28 + f["title"].cap_height(size)
    features = title + 28 + label
    c = Canvas(600, int(features + top), "DEVICE TWEAKER")
    c.text(f["label"], PLATFORM, 21, 2, platform, th["muted"], 1.6)
    c.text(f["title"], "DEVICE TWEAKER", size, 0, title, th["text"], -1.6)
    c.text(f["label"], FEATURES, 21, 2, features, th["muted"], 1.2)
    return c.svg()


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


def button(f, th, label, primary):
    size = 14
    width = int(round(f["mono"].width(label, size, 1.0) + 48))
    c = Canvas(width, 40, label)
    if primary:
        c.rect(0, 0, width, 40, None, th["text"], 3)
        color = th["page"]
    else:
        c.rect(0.75, 0.75, width - 1.5, 38.5, th["line"], "none", 3, 1.5)
        color = th["text"]
    c.text(f["mono"], label, size, width / 2, 20 + f["mono"].cap_height(size) / 2, color, 1.0, "middle")
    return c.svg(), width


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

    f = {"title": load("InterDisplay-ExtraBold.ttf"), "mono": load("JetBrainsMono-Bold.ttf"),
         "label": load("JetBrainsMono-Medium.ttf")}
    out = args.out
    out.mkdir(parents=True, exist_ok=True)
    for old in out.glob("*.svg"):
        old.unlink()

    widths = {}
    for theme, th in THEMES.items():
        (out / f"hero-{theme}.svg").write_text(hero(f, th), encoding="utf-8")
        (out / f"hero-mobile-{theme}.svg").write_text(hero_mobile(f, th), encoding="utf-8")
        for lang, t in TEXT.items():
            (out / f"cpu-allocation-{lang}-{theme}.svg").write_text(allocation(f, th, t["core"], "wide"), encoding="utf-8")
            (out / f"cpu-allocation-mobile-{lang}-{theme}.svg").write_text(allocation(f, th, t["core"], "compact"), encoding="utf-8")
            for key, label in t["buttons"].items():
                svg, width = button(f, th, label, key == "download")
                (out / f"button-{key}-{lang}-{theme}.svg").write_text(svg, encoding="utf-8")
                widths[f"{key}-{lang}"] = width
    for name, width in sorted(widths.items()):
        print(f"{name}: {width}px")


if __name__ == "__main__":
    main()
