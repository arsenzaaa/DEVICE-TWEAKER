"""Draws the GitHub README artwork in assets/readme: the wordmark, framed buttons and guide cards,
in the same style as the site (and the program window). GitHub renders no CSS, so these parts of the
site are pictures here; every picture has a dark and a light version.

    python dev/branding/readme.py --fonts <folder with JetBrainsMono-*.ttf>

Card text comes from the first-line comment of each chapter in docs/guide, so the cards follow the guide.
"""
import argparse
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from render import Canvas, Font, num  # noqa: E402

ROOT = Path(__file__).resolve().parents[2]

# The same colours as dev/site/site.css.
THEMES = {
    "dark": {"panel": "#0c0c0f", "frame": "#3c3c44", "line": "#24242a", "border": "#969696", "text": "#f0f0f0",
             "muted": "#c8c8c8", "dim": "#8c8c94", "faint": "#5e5e66", "invert_bg": "#f0f0f0", "invert_text": "#050507"},
    "light": {"panel": "#ffffff", "frame": "#c9c9cf", "line": "#e2e2e6", "border": "#5c5c64", "text": "#0a0a0c",
              "muted": "#34343a", "dim": "#66666e", "faint": "#9a9aa2", "invert_bg": "#0a0a0c", "invert_text": "#fafafa"},
}

TEXT = {
    "ru": {
        "buttons": [("download", "СКАЧАТЬ {ver}", True), ("guide", "ГАЙД", False), ("site", "САЙТ", False), ("changelog", "ИЗМЕНЕНИЯ", False)],
        "chapter": "Глава", "reading": "Чтение", "minutes": "мин", "open": "ОТКРЫТЬ",
    },
    "en": {
        "buttons": [("download", "DOWNLOAD {ver}", True), ("guide", "GUIDE", False), ("site", "SITE", False), ("changelog", "CHANGELOG", False)],
        "chapter": "Chapter", "reading": "Reading", "minutes": "min", "open": "OPEN",
    },
}


def wrap(font, text, size, width):
    lines, line = [], ""
    for word in text.split():
        probe = f"{line} {word}".strip()
        if font.width(probe, size) <= width:
            line = probe
        else:
            lines.append(line)
            line = word
    return lines + [line]


# README column on github.com is 838 px wide at every desktop size: the artwork is drawn at exactly
# that size and placed 1:1, with baselines and text starts on whole pixels, so nothing is resampled.
COLUMN = 838


def put(c, font, value, size, x, y, fill, tracking=0.0, anchor="start"):
    width = font.width(value, size, tracking)
    start = x - (width if anchor == "end" else width / 2 if anchor == "middle" else 0)
    c.text(font, value, size, round(start), round(y), fill, tracking)


def header(f, th):
    """Only the wordmark is a picture; the lines under it are real text in the README."""
    size = 100
    while f["bold"].width("DEVICE TWEAKER", size, size * 0.06) > COLUMN - 40:
        size -= 1
    cap = f["bold"].cap_height(size)
    height = round(cap) + 24
    c = Canvas(COLUMN, height, "DEVICE TWEAKER")
    put(c, f["bold"], "DEVICE TWEAKER", size, COLUMN / 2, 12 + cap, th["text"], round(size * 0.06), "middle")
    return c.svg()


def button(f, th, label, primary):
    size, height = 15, 48
    w = int(round(f["bold"].width(label, size, 1.2) + 56))
    c = Canvas(w, height, label)
    if primary:
        c.rect(0, 0, w, height, None, th["invert_bg"], 0)
        color = th["invert_text"]
    else:
        c.rect(0.75, 0.75, w - 1.5, height - 1.5, th["border"], "none", 0, 1.5)
        color = th["text"]
    put(c, f["bold"], label, size, w / 2, height / 2 + f["bold"].cap_height(size) / 2, color, 1.2, "middle")
    return c.svg(), w


def chapters(lang):
    out = []
    for p in sorted((ROOT / "docs" / "guide").glob("[0-9]*.md")):
        if p.name.endswith(".en.md") != (lang == "en"):
            continue
        text = p.read_text(encoding="utf-8-sig")
        meta = re.match(r"<!--\s*tag:\s*([^|]+)\|\s*summary:\s*(.+?)\s*-->", text, re.S)
        title = re.search(r"^# (.+)$", text, re.M).group(1).strip()
        words = len(re.findall(r"\w+", re.sub(r"```.*?```", "", text, flags=re.S)))
        out.append({"file": p.name, "number": p.name[:2], "tag": meta.group(1).strip(), "title": title,
                    "summary": meta.group(2).strip(), "minutes": max(2, round(words / 170))})
    return out


CARD_W, PAD = 413, 18
TITLE, SUMMARY, FOOT, OPEN = 15, 13, 12, 11
TITLE_LH, SUMMARY_LH = 22, 21


def card_layout(f, ch):
    head_x = PAD + round(f["bold"].width(f"[{ch['tag']}]", TITLE) + 10 + f["regular"].width(ch["number"], TITLE) + 10)
    title_lines = wrap(f["bold"], ch["title"], TITLE, CARD_W - PAD - head_x)
    summary_lines = wrap(f["regular"], ch["summary"], SUMMARY, CARD_W - PAD * 2)
    return head_x, title_lines, summary_lines


def card(f, th, t, ch, height, max_title):
    head_x, title_lines, summary_lines = card_layout(f, ch)
    c = Canvas(CARD_W, height, f"[{ch['tag']}] {ch['number']} {ch['title']}")
    c.rect(0.5, 0.5, CARD_W - 1, height - 1, th["frame"], th["panel"], 0, 1)
    y = 30
    put(c, f["bold"], f"[{ch['tag']}]", TITLE, PAD, y, th["text"])
    put(c, f["regular"], ch["number"], TITLE, PAD + f["bold"].width(f"[{ch['tag']}]", TITLE) + 10, y, th["faint"])
    for i, line in enumerate(title_lines):
        put(c, f["bold"], line, TITLE, head_x, y + i * TITLE_LH, th["text"])
    rule = 30 + (max_title - 1) * TITLE_LH + 16
    c.line(f"M{PAD} {rule + 0.5}H{CARD_W - PAD}", th["line"], 1)
    y = rule + 27
    for line in summary_lines:
        put(c, f["regular"], line, SUMMARY, PAD, y, th["muted"])
        y += SUMMARY_LH
    foot = height - 27
    x = PAD
    for key, value in ((t["chapter"], ch["number"]), (t["reading"], f"{ch['minutes']} {t['minutes']}")):
        put(c, f["regular"], f"{key}:", FOOT, x, foot + 4, th["dim"])
        x += f["regular"].width(f"{key}: ", FOOT)
        put(c, f["bold"], value, FOOT, x, foot + 4, th["text"])
        x += f["bold"].width(value, FOOT) + 20
    bw = round(f["bold"].width(t["open"], OPEN, 1.0) + 28)
    bx = CARD_W - PAD - bw
    c.rect(bx + 0.5, foot - 14.5, bw - 1, 29, th["border"], "none", 0, 1)
    put(c, f["bold"], t["open"], OPEN, bx + bw / 2, foot + f["bold"].cap_height(OPEN) / 2, th["text"], 1.0, "middle")
    return c.svg()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--fonts", required=True, type=Path)
    parser.add_argument("--out", default=ROOT / "assets" / "readme", type=Path)
    args = parser.parse_args()

    def load(name):
        found = sorted(args.fonts.rglob(name))
        if not found:
            raise SystemExit(f"Font not found: {name}")
        return Font(found[0])

    f = {"bold": load("JetBrainsMono-Bold.ttf"), "regular": load("JetBrainsMono-Regular.ttf")}
    ver = re.search(r"<Version>([^<]+)</Version>", (ROOT / "src" / "DeviceTweakerCS.csproj").read_text(encoding="utf-8")).group(1)
    out = args.out
    out.mkdir(parents=True, exist_ok=True)
    for old in out.glob("*.svg"):
        old.unlink()
    widths = {}
    for theme, th in THEMES.items():
        (out / f"wordmark-{theme}.svg").write_text(header(f, th), encoding="utf-8")
    for lang, t in TEXT.items():
        chs = chapters(lang)
        layouts = [card_layout(f, ch) for ch in chs]
        max_title = max(len(l[1]) for l in layouts)
        max_summary = max(len(l[2]) for l in layouts)
        height = int(30 + (max_title - 1) * TITLE_LH + 16 + 27 + (max_summary - 1) * SUMMARY_LH + 22 + 50)
        for theme, th in THEMES.items():
            for key, label, primary in t["buttons"]:
                svg, w = button(f, th, label.format(ver=ver), primary)
                (out / f"button-{key}-{lang}-{theme}.svg").write_text(svg, encoding="utf-8")
                widths[f"{key}-{lang}"] = w
            for ch in chs:
                (out / f"card-{ch['number']}-{lang}-{theme}.svg").write_text(card(f, th, t, ch, height, max_title), encoding="utf-8")
        widths[f"card-{lang}"] = height
    for key, value in sorted(widths.items()):
        print(f"{key}: {value}")


if __name__ == "__main__":
    main()
