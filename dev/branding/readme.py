"""Draws the GitHub README artwork in assets/readme: the header, framed buttons and guide cards,
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
        "lead": "Прерывания устройств Windows под топологию процессора.",
        "sub": "MSI Utility V3 + Interrupt Affinity Policy Tool (IAPT) + ReservedCpuSets + RSS + USB IMOD / NIC ITR в одном интерфейсе, с интеллектуальной авто-оптимизацией.",
        "status": [("Windows", "10 / 11 x64"), ("Версия", "{ver}"), (".NET", "8"), ("Лицензия", "GPL-3.0")],
        "buttons": [("download", "СКАЧАТЬ {ver}", True), ("guide", "ГАЙД", False), ("site", "САЙТ", False), ("changelog", "ИЗМЕНЕНИЯ", False)],
        "chapter": "Глава", "reading": "Чтение", "minutes": "мин", "open": "ОТКРЫТЬ",
    },
    "en": {
        "lead": "Device interrupts on Windows, tuned to your CPU topology.",
        "sub": "MSI Utility V3 + Interrupt Affinity Policy Tool (IAPT) + ReservedCpuSets + RSS + USB IMOD / NIC ITR in one interface, with intelligent auto-optimization.",
        "status": [("Windows", "10 / 11 x64"), ("Version", "{ver}"), (".NET", "8"), ("License", "GPL-3.0")],
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


def status_row(c, f, th, items, size, y, center):
    sep = size * 1.1
    parts = [(f"{k}", f" - {v}") for k, v in items]
    widths = [f["bold"].width(k, size) + f["regular"].width(v, size) for k, v in parts]
    bar = f["regular"].width("|", size)
    total = sum(widths) + (len(parts) - 1) * (bar + sep * 2)
    x = center - total / 2
    for i, (k, v) in enumerate(parts):
        c.text(f["bold"], k, size, x, y, th["text"])
        x += f["bold"].width(k, size)
        c.text(f["regular"], v, size, x, y, th["muted"])
        x += f["regular"].width(v, size)
        if i < len(parts) - 1:
            c.text(f["regular"], "|", size, x + sep, y, th["faint"])
            x += bar + sep * 2


def header(f, th, t, ver, width, mobile):
    s = {"eyebrow": 16, "word": 104, "lead": 26, "sub": 19, "status": 17} if not mobile else \
        {"eyebrow": 22, "word": 0, "lead": 30, "sub": 23, "status": 22}
    if mobile:
        size = 80
        while f["bold"].width("DEVICE TWEAKER", size, size * 0.06) > width - 40:
            size -= 1
        s["word"] = size
    parts = []
    y = 34
    eyebrow = "[ WINDOWS 10 / 11 x64 ]"
    parts.append(("eyebrow", y))
    y += 34 + f["bold"].cap_height(s["word"])
    parts.append(("word", y))
    lead_lines = wrap(f["bold"], t["lead"], s["lead"], width - 80)
    y += 52
    lead_y = []
    for _ in lead_lines:
        lead_y.append(y)
        y += s["lead"] * 1.5
    sub_lines = wrap(f["regular"], t["sub"], s["sub"], width - (160 if not mobile else 60))
    y += 8
    sub_y = []
    for _ in sub_lines:
        sub_y.append(y)
        y += s["sub"] * 1.7
    status = [(k, v.format(ver=ver)) for k, v in t["status"]]
    rows = [status] if not mobile else [status[:2], status[2:]]
    y += 22
    status_y = []
    for _ in rows:
        status_y.append(y)
        y += s["status"] * 1.9
    c = Canvas(width, int(y), "DEVICE TWEAKER")
    mid = width / 2
    c.text(f["bold"], eyebrow, s["eyebrow"], mid, 34, th["faint"], s["eyebrow"] * 0.16, "middle")
    c.text(f["bold"], "DEVICE TWEAKER", s["word"], mid, parts[1][1], th["text"], s["word"] * 0.06, "middle")
    for line, ly in zip(lead_lines, lead_y):
        c.text(f["bold"], line, s["lead"], mid, ly, th["text"], 0, "middle")
    for line, ly in zip(sub_lines, sub_y):
        c.text(f["regular"], line, s["sub"], mid, ly, th["dim"], 0, "middle")
    for row, ry in zip(rows, status_y):
        status_row(c, f, th, row, s["status"], ry, mid)
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
    c.text(f["bold"], label, size, w / 2, height / 2 + f["bold"].cap_height(size) / 2, color, 1.2, "middle")
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


CARD_W, PAD = 600, 26
TITLE, SUMMARY, FOOT = 19, 17, 15
TITLE_LH, SUMMARY_LH = 28, 28


def card_layout(f, ch):
    head_x = PAD + f["bold"].width(f"[{ch['tag']}]", TITLE) + 14 + f["regular"].width(ch["number"], TITLE) + 14
    title_lines = wrap(f["bold"], ch["title"], TITLE, CARD_W - PAD - head_x)
    summary_lines = wrap(f["regular"], ch["summary"], SUMMARY, CARD_W - PAD * 2)
    return head_x, title_lines, summary_lines


def card(f, th, t, ch, height, max_title):
    head_x, title_lines, summary_lines = card_layout(f, ch)
    c = Canvas(CARD_W, height, f"[{ch['tag']}] {ch['number']} {ch['title']}")
    c.rect(0.5, 0.5, CARD_W - 1, height - 1, th["frame"], th["panel"], 0, 1)
    y = 40
    c.text(f["bold"], f"[{ch['tag']}]", TITLE, PAD, y, th["text"])
    c.text(f["regular"], ch["number"], TITLE, PAD + f["bold"].width(f"[{ch['tag']}]", TITLE) + 14, y, th["faint"])
    for i, line in enumerate(title_lines):
        c.text(f["bold"], line, TITLE, head_x, y + i * TITLE_LH, th["text"])
    rule = 40 + (max_title - 1) * TITLE_LH + 22
    c.line(f"M{PAD} {num(rule + 0.5)}H{CARD_W - PAD}", th["line"], 1)
    y = rule + 36
    for line in summary_lines:
        c.text(f["regular"], line, SUMMARY, PAD, y, th["muted"])
        y += SUMMARY_LH
    foot = height - 34
    x = PAD
    for key, value in ((t["chapter"], ch["number"]), (t["reading"], f"{ch['minutes']} {t['minutes']}")):
        c.text(f["regular"], f"{key}:", FOOT, x, foot + 6, th["dim"])
        x += f["regular"].width(f"{key}: ", FOOT)
        c.text(f["bold"], value, FOOT, x, foot + 6, th["text"])
        x += f["bold"].width(value, FOOT) + 28
    bw = f["bold"].width(t["open"], 14, 1.0) + 40
    bx = CARD_W - PAD - bw
    c.rect(bx + 0.75, foot - 19.25, bw - 1.5, 38.5, th["border"], "none", 0, 1.5)
    c.text(f["bold"], t["open"], 14, bx + bw / 2, foot + f["bold"].cap_height(14) / 2, th["text"], 1.0, "middle")
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
    for lang, t in TEXT.items():
        chs = chapters(lang)
        layouts = [card_layout(f, ch) for ch in chs]
        max_title = max(len(l[1]) for l in layouts)
        max_summary = max(len(l[2]) for l in layouts)
        height = int(40 + (max_title - 1) * TITLE_LH + 22 + 36 + (max_summary - 1) * SUMMARY_LH + 34 + 66)
        for theme, th in THEMES.items():
            (out / f"header-{lang}-{theme}.svg").write_text(header(f, th, t, ver, 1200, False), encoding="utf-8")
            (out / f"header-mobile-{lang}-{theme}.svg").write_text(header(f, th, t, ver, 640, True), encoding="utf-8")
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
