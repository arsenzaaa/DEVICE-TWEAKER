"""Renders the README artwork in assets/branding.

Text is converted to outlines, so the SVGs look the same on every system.
Fonts (OFL): JetBrains Mono and Inter, see FONT_SOURCES. Run from the repository root:

    python dev/branding/render.py --fonts <folder with the .ttf files>
"""
import argparse
import html
from pathlib import Path

from fontTools.pens.svgPathPen import SVGPathPen
from fontTools.pens.transformPen import TransformPen
from fontTools.ttLib import TTFont

FONT_SOURCES = {
    "JetBrainsMono-Medium.ttf": "https://github.com/JetBrains/JetBrainsMono/releases/download/v2.304/JetBrainsMono-2.304.zip",
    "JetBrainsMono-Bold.ttf": "https://github.com/JetBrains/JetBrainsMono/releases/download/v2.304/JetBrainsMono-2.304.zip",
    "Inter-Regular.ttf": "https://github.com/rsms/inter/releases/download/v4.0/Inter-4.0.zip",
    "Inter-Medium.ttf": "https://github.com/rsms/inter/releases/download/v4.0/Inter-4.0.zip",
    "Inter-SemiBold.ttf": "https://github.com/rsms/inter/releases/download/v4.0/Inter-4.0.zip",
    "InterDisplay-ExtraBold.ttf": "https://github.com/rsms/inter/releases/download/v4.0/Inter-4.0.zip",
}

# Colours of the program itself (GUI/DeviceBadgePalette.cs, GUI/MainForm.Theme.cs);
# the light theme uses darker shades of the same hues for contrast on white.
THEMES = {
    "dark": {
        "card": "#0b0e13", "card2": "#11161d", "border": "#262c35", "text": "#f0f2f4",
        "muted": "#8b949e", "faint": "#3a414b", "btn": "#f0f2f4", "btnText": "#0b0e13",
        "usb": "#48d9d0", "gpu": "#76b900", "net": "#62aef5", "audio": "#ff9b7d",
        "pcore": "#ffe678", "smt": "#ffb43c", "ecore": "#78c8ff",
    },
    "light": {
        "card": "#ffffff", "card2": "#f6f8fa", "border": "#d0d7de", "text": "#1f2328",
        "muted": "#59636e", "faint": "#c4cbd3", "btn": "#1f2328", "btnText": "#ffffff",
        "usb": "#0e8f87", "gpu": "#4a7f00", "net": "#1f6fc0", "audio": "#c94f2a",
        "pcore": "#a07800", "smt": "#b86200", "ecore": "#1f74b8",
    },
}

TEXT = {
    "ru": {
        "eyebrow": "WINDOWS 10 / 11 · x64",
        "tagline": ["Настройка прерываний устройств", "под топологию вашего процессора"],
        "download": "Скачать 0.0.4", "settings": "Описание настроек", "changelog": "Что нового",
        "steps": [
            ("Скан", ["Устройства, роли USB,", "топология CPU"]),
            ("План", ["Ядра под GPU, ввод,", "сеть и аудио"]),
            ("Бекап", ["Снимок реестра", "до любой записи"]),
            ("Запись", ["MSI, affinity, IMOD,", "NIC ITR, RSS"]),
            ("Проверка", ["Повторное чтение", "каждого значения"]),
            ("Автозапуск", ["Вход, выход из сна,", "подключение USB"]),
        ],
    },
    "en": {
        "eyebrow": "WINDOWS 10 / 11 · x64",
        "tagline": ["Device interrupt tuning", "matched to your CPU topology"],
        "download": "Download 0.0.4", "settings": "Settings guide", "changelog": "What's new",
        "steps": [
            ("Scan", ["Devices, USB roles,", "CPU topology"]),
            ("Plan", ["Cores for GPU, input,", "network and audio"]),
            ("Backup", ["Registry snapshot", "before any write"]),
            ("Write", ["MSI, affinity, IMOD,", "NIC ITR, RSS"]),
            ("Verify", ["Every value is", "read back"]),
            ("Startup", ["Sign-in, resume,", "USB attach"]),
        ],
    },
}


class Font:
    def __init__(self, path):
        self.font = TTFont(path)
        self.glyphs = self.font.getGlyphSet()
        self.cmap = self.font.getBestCmap()
        self.upm = self.font["head"].unitsPerEm
        self.hmtx = self.font["hmtx"]

    def _glyph(self, ch):
        return self.cmap.get(ord(ch)) or self.cmap[ord("?")]

    def width(self, text, size, tracking=0.0):
        scale = size / self.upm
        advances = [self.hmtx[self._glyph(ch)][0] * scale for ch in text]
        return sum(advances) + tracking * max(0, len(text) - 1)

    def path(self, text, size, x, y, tracking=0.0, anchor="start"):
        scale = size / self.upm
        width = self.width(text, size, tracking)
        cx = x - (width if anchor == "end" else width / 2 if anchor == "middle" else 0)
        pen = SVGPathPen(self.glyphs, ntos=lambda v: f"{v:.2f}".rstrip("0").rstrip("."))
        for ch in text:
            name = self._glyph(ch)
            self.glyphs[name].draw(TransformPen(pen, (scale, 0, 0, -scale, cx, y)))
            cx += self.hmtx[name][0] * scale + tracking
        return pen.getCommands()


def text(font, value, size, x, y, fill, tracking=0.0, anchor="start", opacity=None):
    d = font.path(value, size, x, y, tracking, anchor)
    extra = f' fill-opacity="{opacity}"' if opacity is not None else ""
    return f'<path d="{d}" fill="{fill}"{extra}/>'


def svg(width, height, title, body):
    return (f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" '
            f'viewBox="0 0 {width} {height}" role="img" aria-label="{html.escape(title)}">'
            f"<title>{html.escape(title)}</title>{body}</svg>\n")


def fit(font, value, size, max_width, tracking=0.0):
    while font.width(value, size, tracking) > max_width and size > 8:
        size -= 1
    return size


def core_strip(f, c, x, y, cell, gap, label_y, label_size=14):
    """16 logical CPUs; the cores the AUTO plan gave to each device are filled."""
    owners = {2: "usb", 6: "net", 12: "gpu", 14: "gpu"}
    parts = []
    for lp in range(16):
        cx = x + lp * (cell + gap)
        kind = "pcore" if lp % 2 == 0 else "smt"
        owner = owners.get(lp)
        if owner:
            parts.append(f'<rect x="{cx}" y="{y}" width="{cell}" height="{cell}" rx="4" fill="{c[owner]}"/>')
        else:
            parts.append(f'<rect x="{cx + 0.75}" y="{y + 0.75}" width="{cell - 1.5}" height="{cell - 1.5}" rx="3.5" '
                         f'fill="none" stroke="{c[kind]}" stroke-opacity="0.45" stroke-width="1.5"/>')
    end = x + 16 * cell + 15 * gap
    parts.append(text(f["mono"], "CPU 0", label_size, x, label_y, c["muted"], 0.5))
    parts.append(text(f["mono"], "CPU 15", label_size, end, label_y, c["muted"], 0.5, anchor="end"))
    return "".join(parts), end


def hero(f, c, t):
    w, h = 1200, 360
    p = [f'<rect x="0.5" y="0.5" width="{w - 1}" height="{h - 1}" rx="16" fill="{c["card"]}" stroke="{c["border"]}"/>']
    p.append(text(f["mono"], t["eyebrow"], 16, 56, 84, c["muted"], 2.2))
    size = fit(f["display"], "DEVICE TWEAKER", 74, 560, -1.2)
    p.append(text(f["display"], "DEVICE TWEAKER", size, 54, 160, c["text"], -1.2))
    p.append(text(f["medium"], t["tagline"][0], 22, 56, 210, c["muted"]))
    p.append(text(f["medium"], t["tagline"][1], 22, 56, 240, c["muted"]))
    cx = 56
    for chip in ["MSI", "AFFINITY", "IMOD", "NIC ITR", "RSS"]:
        cw = f["mono"].width(chip, 15, 1.2) + 26
        p.append(f'<rect x="{cx + 0.5}" y="280.5" width="{cw:.1f}" height="33" rx="7" fill="{c["card2"]}" stroke="{c["border"]}"/>')
        p.append(text(f["mono"], chip, 15, cx + cw / 2, 302.5, c["text"], 1.2, anchor="middle"))
        cx += cw + 8
    p.append(f'<path d="M660 48V312" stroke="{c["border"]}"/>')
    rows = [("USB", "usb", "CPU 2", "IMOD 0x0"), ("GPU", "gpu", "CPU 12 · 14", "MSI-X"), ("NET", "net", "CPU 6", "RSS ×1")]
    for i, (name, key, target, note) in enumerate(rows):
        y = 74 + i * 56
        p.append(f'<rect x="700.5" y="{y + 0.5}" width="84" height="34" rx="7" fill="{c[key]}" fill-opacity="0.1" stroke="{c[key]}" stroke-opacity="0.6"/>')
        p.append(text(f["monoBold"], name, 15, 742.5, y + 23, c[key], 1.5, anchor="middle"))
        p.append(f'<path d="M796 {y + 17.5}H884" stroke="{c["faint"]}" stroke-width="1.5"/>'
                 f'<path d="M880 {y + 13}L886 {y + 17.5}L880 {y + 22}" fill="none" stroke="{c["faint"]}" stroke-width="1.5"/>')
        p.append(text(f["monoMedium"], target, 19, 900, y + 24, c["text"]))
        p.append(text(f["mono"], note, 15, 1142, y + 23, c["muted"], 0.5, anchor="end"))
    strip, _ = core_strip(f, c, 700, 258, 22, 5.33, 312)
    p.append(strip)
    return svg(w, h, "DEVICE TWEAKER", "".join(p))


def hero_mobile(f, c, t):
    w, h = 640, 330
    p = [f'<rect x="0.5" y="0.5" width="{w - 1}" height="{h - 1}" rx="14" fill="{c["card"]}" stroke="{c["border"]}"/>']
    p.append(text(f["mono"], t["eyebrow"], 18, 40, 66, c["muted"], 2))
    size = fit(f["display"], "DEVICE TWEAKER", 70, 560, -1)
    p.append(text(f["display"], "DEVICE TWEAKER", size, 38, 136, c["text"], -1))
    p.append(text(f["medium"], t["tagline"][0], 25, 40, 186, c["muted"]))
    p.append(text(f["medium"], t["tagline"][1], 25, 40, 220, c["muted"]))
    strip, _ = core_strip(f, c, 40, 250, 29, 7.07, 312, 19)
    p.append(strip)
    return svg(w, h, "DEVICE TWEAKER", "".join(p))


ICON_DOWNLOAD = "M0 0V13 M-6 7L0 13L6 7 M-9 18H9"


def button(f, c, label, primary, icon=None):
    size = 17
    tw = f["semibold"].width(label, size)
    icon_w = 26 if icon else 0
    w, h = int(tw + icon_w + 56), 48
    if primary:
        frame = f'<rect x="0" y="0" width="{w}" height="{h}" rx="10" fill="{c["btn"]}"/>'
        color = c["btnText"]
    else:
        frame = f'<rect x="1" y="1" width="{w - 2}" height="{h - 2}" rx="9" fill="{c["card"]}" stroke="{c["border"]}" stroke-width="1.5"/>'
        color = c["text"]
    x = (w - tw - icon_w) / 2
    body = frame
    if icon:
        body += f'<path d="{icon}" transform="translate({x + 9:.1f} 14)" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>'
    body += text(f["semibold"], label, size, x + icon_w, 30.5, color)
    return svg(w, h, label, body)


def pipeline(f, c, t):
    w, h = 1200, 210
    p = [f'<rect x="0.5" y="0.5" width="{w - 1}" height="{h - 1}" rx="14" fill="{c["card"]}" stroke="{c["border"]}"/>']
    step_w = 170
    gap = (w - 60 - 6 * step_w) / 5
    for i, (title, lines) in enumerate(t["steps"]):
        x = 30 + i * (step_w + gap)
        p.append(text(f["mono"], f"0{i + 1}", 15, x, 52, c["muted"], 1.5))
        p.append(text(f["semibold"], title, 24, x, 92, c["text"]))
        p.append(text(f["regular"], lines[0], 16.5, x, 132, c["muted"]))
        p.append(text(f["regular"], lines[1], 16.5, x, 156, c["muted"]))
        if i < 5:
            ax = x + step_w + gap / 2
            p.append(f'<path d="M{ax - 9} 78L{ax} 85L{ax - 9} 92" fill="none" stroke="{c["faint"]}" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"/>')
    return svg(w, h, " → ".join(s[0] for s in t["steps"]), "".join(p))


def pipeline_mobile(f, c, t):
    w, h = 640, 470
    p = [f'<rect x="0.5" y="0.5" width="{w - 1}" height="{h - 1}" rx="14" fill="{c["card"]}" stroke="{c["border"]}"/>']
    for i, (title, lines) in enumerate(t["steps"]):
        col, row = i % 2, i // 2
        x, y = 40 + col * 300, 40 + row * 145
        p.append(text(f["mono"], f"0{i + 1}", 16, x, y + 18, c["muted"], 1.5))
        p.append(text(f["semibold"], title, 27, x, y + 58, c["text"]))
        p.append(text(f["regular"], lines[0], 18, x, y + 90, c["muted"]))
        p.append(text(f["regular"], lines[1], 18, x, y + 114, c["muted"]))
    return svg(w, h, " → ".join(s[0] for s in t["steps"]), "".join(p))


def social(f, c):
    """Repository social preview, 1280x640 (Settings → Social preview)."""
    w, h = 1280, 640
    p = [f'<rect width="{w}" height="{h}" fill="{c["card"]}"/>']
    p.append(text(f["mono"], "WINDOWS 10 / 11 · x64", 22, 96, 170, c["muted"], 3))
    size = fit(f["display"], "DEVICE TWEAKER", 128, 1088, -2)
    p.append(text(f["display"], "DEVICE TWEAKER", size, 92, 300, c["text"], -2))
    p.append(text(f["mono"], "MSI · AFFINITY · IMOD · NIC ITR · RSS", 28, 96, 372, c["muted"], 2))
    strip, _ = core_strip(f, c, 96, 452, 52, 16.1, 548, 24)
    p.append(strip)
    return svg(w, h, "DEVICE TWEAKER", "".join(p))


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--fonts", required=True, type=Path)
    parser.add_argument("--out", default=Path("assets/branding"), type=Path)
    args = parser.parse_args()

    def load(name):
        found = list(args.fonts.rglob(name))
        if not found:
            raise SystemExit(f"Font not found: {name} ({FONT_SOURCES[name]})")
        return Font(found[0])

    f = {
        "mono": load("JetBrainsMono-Medium.ttf"), "monoMedium": load("JetBrainsMono-Medium.ttf"),
        "monoBold": load("JetBrainsMono-Bold.ttf"), "regular": load("Inter-Regular.ttf"),
        "medium": load("Inter-Medium.ttf"), "semibold": load("Inter-SemiBold.ttf"),
        "display": load("InterDisplay-ExtraBold.ttf"),
    }
    out = args.out
    out.mkdir(parents=True, exist_ok=True)
    for old in out.glob("*.svg"):
        old.unlink()
    for theme, c in THEMES.items():
        for lang, t in TEXT.items():
            (out / f"hero-{lang}-{theme}.svg").write_text(hero(f, c, t), encoding="utf-8")
            (out / f"hero-mobile-{lang}-{theme}.svg").write_text(hero_mobile(f, c, t), encoding="utf-8")
            (out / f"pipeline-{lang}-{theme}.svg").write_text(pipeline(f, c, t), encoding="utf-8")
            (out / f"pipeline-mobile-{lang}-{theme}.svg").write_text(pipeline_mobile(f, c, t), encoding="utf-8")
            (out / f"button-download-{lang}-{theme}.svg").write_text(button(f, c, t["download"], True, ICON_DOWNLOAD), encoding="utf-8")
            (out / f"button-settings-{lang}-{theme}.svg").write_text(button(f, c, t["settings"], False), encoding="utf-8")
            (out / f"button-changelog-{lang}-{theme}.svg").write_text(button(f, c, t["changelog"], False), encoding="utf-8")
    (out / "social-preview.svg").write_text(social(f, THEMES["dark"]), encoding="utf-8")
    print("\n".join(sorted(p.name for p in out.glob("*.svg"))))


if __name__ == "__main__":
    main()
