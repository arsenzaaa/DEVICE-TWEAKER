"""Builds the project site (GitHub Pages) from the repository documents.

    pip install markdown==3.7
    python dev/site/build.py --out _site

Russian pages are at the root, English ones under /en/. The guide chapters come from docs/guide,
the settings reference, development notes and changelog from docs/ and CHANGELOG*.md, so the
site never drifts from the repository. The page chrome copies the program window: a status line,
framed buttons, [ FILTER ] chips, device cards and the CPU Affinity checklist.
"""
import argparse
import html
import posixpath
import re
import shutil
from pathlib import Path

import markdown
from markdown.extensions.toc import slugify_unicode

ROOT = Path(__file__).resolve().parents[2]
HERE = Path(__file__).resolve().parent
REPO = "https://github.com/arsenzaaa/DEVICE-TWEAKER"
SITE = "https://arsenzaaa.github.io/DEVICE-TWEAKER/"
TELEGRAM = "https://t.me/arsenzaa"
NBSP = "\u00a0"

DOCS = {
    ("settings", "ru"): "docs/SETTINGS.md",
    ("settings", "en"): "docs/SETTINGS.en.md",
    ("development", "ru"): "docs/DEVELOPMENT.md",
    ("development", "en"): "docs/DEVELOPMENT.en.md",
    ("changelog", "ru"): "CHANGELOG.md",
    ("changelog", "en"): "CHANGELOG.en.md",
}

UI = {
    "ru": {
        "other": "en",
        "nav": [("guide", "ГАЙД"), ("settings", "НАСТРОЙКИ"), ("changelog", "ИЗМЕНЕНИЯ")],
        "docs_nav": [("settings", "Описание настроек"), ("development", "Разработка"), ("changelog", "История изменений")],
        "download": "СКАЧАТЬ", "guide": "Гайд", "chapter": "Глава", "reading": "Чтение", "minutes": "мин",
        "open": "ОТКРЫТЬ", "prev": "НАЗАД", "next": "ДАЛЕЕ", "contents": "Оглавление", "on_page": "На странице",
        "reference": "Справочник", "source": "Исходный текст", "all": "ВСЕ", "search": "Поиск по гайду... (Ctrl+F)",
        "nothing": "Ничего не найдено.", "license": "Распространяется по лицензии GPL-3.0.",
        "status": [("Windows", "10 / 11 x64"), ("Версия", "{ver}"), (".NET", "8"), ("Лицензия", "GPL-3.0")],
        "footer": [("Репозиторий", "github.com/arsenzaaa/DEVICE-TWEAKER", REPO), ("Релизы", "github.com/arsenzaaa/DEVICE-TWEAKER/releases", f"{REPO}/releases"),
                   ("Ошибки", "issues", f"{REPO}/issues"), ("Новости", "t.me/arsenzaa", TELEGRAM)],
    },
    "en": {
        "other": "ru",
        "nav": [("guide", "GUIDE"), ("settings", "SETTINGS"), ("changelog", "CHANGELOG")],
        "docs_nav": [("settings", "Settings guide"), ("development", "Development"), ("changelog", "Changelog")],
        "download": "DOWNLOAD", "guide": "Guide", "chapter": "Chapter", "reading": "Reading", "minutes": "min",
        "open": "OPEN", "prev": "BACK", "next": "NEXT", "contents": "Contents", "on_page": "On this page",
        "reference": "Reference", "source": "Source", "all": "ALL", "search": "Search the guide... (Ctrl+F)",
        "nothing": "Nothing found.", "license": "Distributed under the GPL-3.0 license.",
        "status": [("Windows", "10 / 11 x64"), ("Version", "{ver}"), (".NET", "8"), ("License", "GPL-3.0")],
        "footer": [("Repository", "github.com/arsenzaaa/DEVICE-TWEAKER", REPO), ("Releases", "github.com/arsenzaaa/DEVICE-TWEAKER/releases", f"{REPO}/releases"),
                   ("Issues", "issues", f"{REPO}/issues"), ("News", "t.me/arsenzaa", TELEGRAM)],
    },
}

LANDING = {
    "ru": {
        "title": "DEVICE TWEAKER — прерывания устройств Windows под топологию процессора",
        "description": "MSI, Interrupt Affinity, ReservedCpuSets, RSS, USB IMOD и NIC ITR в одной программе для Windows 10 / 11 x64, с интеллектуальной авто-оптимизацией и подробным гайдом.",
        "eyebrow": "WINDOWS 10 / 11 x64",
        "lead": "Прерывания устройств Windows под топологию процессора.",
        "sub": "MSI Utility V3 + Interrupt Affinity Policy Tool (IAPT) + ReservedCpuSets + RSS + USB IMOD / NIC ITR в одном интерфейсе, с интеллектуальной авто-оптимизацией.",
        "shot_alt": "DEVICE TWEAKER: USB-контроллер на Intel Core i9-14900K",
        "guide_title": "ГАЙД",
        "guide_lead": "ISR / DPC выполняются выше любого потока и не знают, какие ядра заняты игрой. Драйвер на ядре с GAME / RENDER / RHI прерывает его -> поток ждёт или теряет свой L1 / L2 -> растёт время кадра. Гайд разбирает, как это устроено в Windows, по каждому типу устройства, и что делает программа.",
        "install_title": "УСТАНОВКА",
        "files": [("DEVICE.TWEAKER.SelfContained.exe", ".NET встроен"), ("DEVICE.TWEAKER.exe", "нужен .NET 8 Desktop Runtime x64")],
        "steps": [("1", "Запустите программу от имени администратора."),
                  ("2", "<b>АВТООПТИМИЗАЦИЯ</b> строит план по топологии и применяет его. Ручная настройка — <b>ПРИМЕНИТЬ</b>."),
                  ("3", "Перезагрузите компьютер. <b>ПРОВЕРИТЬ</b> в блоке IMOD или NIC ITR покажет фактические значения регистров."),
                  ("4", "Откат — <b>ВОССТАНОВЛЕНИЕ → ИСХОДНЫЙ</b>: реестр возвращается к состоянию до первой авто-оптимизации.")],
        "driver": "IMOD и NIC ITR пишет собственный драйвер DTIMOD.sys, только в память выбранного устройства. Его загружает KDU, поэтому антивирусы относят программу к HackTool, а HVCI и Vulnerable Driver Blocklist блокируют загрузку. Остальные настройки от драйвера не зависят.",
        "driver_link": "Подробно — глава 11",
        "shots_title": "СКРИНШОТЫ",
        "shots": [("showcase_gpu_full", "GPU", "Видеокарта на двух физических ядрах"),
                  ("showcase_nic_full", "NET", "NIC ITR и RSS, Intel I226-V"),
                  ("showcase_amd_imod_full", "USB", "USB IMOD, Ryzen 9 9950X3D")],
        "shots_note": "На скриншотах тестовые устройства и значения.",
    },
    "en": {
        "title": "DEVICE TWEAKER — device interrupts on Windows, tuned to your CPU topology",
        "description": "MSI, Interrupt Affinity, ReservedCpuSets, RSS, USB IMOD and NIC ITR in one program for Windows 10 / 11 x64, with intelligent auto-optimization and an in-depth guide.",
        "eyebrow": "WINDOWS 10 / 11 x64",
        "lead": "Device interrupts on Windows, tuned to your CPU topology.",
        "sub": "MSI Utility V3 + Interrupt Affinity Policy Tool (IAPT) + ReservedCpuSets + RSS + USB IMOD / NIC ITR in one interface, with intelligent auto-optimization.",
        "shot_alt": "DEVICE TWEAKER: a USB controller on Intel Core i9-14900K",
        "guide_title": "GUIDE",
        "guide_lead": "ISRs / DPCs run above any thread and do not know which cores the game is using. A driver on the core of GAME / RENDER / RHI interrupts it -> the thread waits or loses its L1 / L2 -> frame time grows. The guide explains how this works in Windows, device type by device type, and what the program does about it.",
        "install_title": "INSTALLATION",
        "files": [("DEVICE.TWEAKER.SelfContained.exe", ".NET included"), ("DEVICE.TWEAKER.exe", "requires the .NET 8 Desktop Runtime x64")],
        "steps": [("1", "Run the program as administrator."),
                  ("2", "<b>AUTO-OPTIMIZATION</b> builds a plan from the topology and applies it. For manual settings use <b>APPLY</b>."),
                  ("3", "Restart the PC. <b>CHECK</b> in the IMOD or NIC ITR block shows the actual register values."),
                  ("4", "Rollback — <b>RESTORE → ORIGINAL</b> returns the registry to its state before the first auto-optimization.")],
        "driver": "IMOD and NIC ITR are written by the custom DTIMOD.sys driver, only to the selected device's memory. KDU loads it, so antivirus software flags the program as HackTool, and HVCI and the Vulnerable Driver Blocklist block the loading. All other settings work without the driver.",
        "driver_link": "Details in chapter 11",
        "shots_title": "SCREENSHOTS",
        "shots": [("showcase_gpu_full", "GPU", "Graphics card on two physical cores"),
                  ("showcase_nic_full", "NET", "NIC ITR and RSS, Intel I226-V"),
                  ("showcase_amd_imod_full", "USB", "USB IMOD, Ryzen 9 9950X3D")],
        "shots_note": "Screenshots show test devices and values.",
    },
}


def image_height(path):
    # PNG IHDR: width and height are big-endian at bytes 16..24.
    data = path.read_bytes()[16:24]
    return int.from_bytes(data[4:8], "big")


def version():
    match = re.search(r"<Version>([^<]+)</Version>", (ROOT / "src" / "DeviceTweakerCS.csproj").read_text(encoding="utf-8"))
    return match.group(1)


def page_url(slug, lang):
    """Site-relative directory of a page."""
    prefix = "en/" if lang == "en" else ""
    return prefix + (slug + "/" if slug else "")


def rel(from_url, to_url):
    """Relative link from one page directory to another site path ('' is the site root)."""
    target = posixpath.relpath(to_url.rstrip("/") or ".", from_url.rstrip("/") or ".")
    if target == ".":
        return "./"
    return target + "/" if to_url == "" or to_url.endswith("/") else target


def typography(markup):
    """No line break after one-letter Russian words and before a dash, outside tags and code."""
    parts = re.split(r"(<code>.*?</code>|<pre>.*?</pre>|<[^>]+>)", markup, flags=re.S)
    for i in range(0, len(parts), 2):
        parts[i] = re.sub(r"(?<![\w-])([иИвВсСкКуУаАоОяЯ]) ", lambda m: m.group(1) + NBSP, parts[i]).replace(" — ", NBSP + "— ")
    return "".join(parts)


# Guide chapters

class Chapter:
    def __init__(self, path, lang):
        self.path = path
        self.lang = lang
        text = path.read_text(encoding="utf-8-sig")
        meta = re.match(r"<!--\s*tag:\s*([^|]+)\|\s*summary:\s*(.+?)\s*-->", text, re.S)
        self.tag = meta.group(1).strip()
        self.summary = meta.group(2).strip()
        self.number = path.name[:2]
        self.slug = re.sub(r"(\.en)?\.md$", "", path.name)
        self.title = re.search(r"^# (.+)$", text, re.M).group(1).strip()
        self.text = text
        words = len(re.findall(r"\w+", re.sub(r"```.*?```", "", text, flags=re.S)))
        self.minutes = max(2, round(words / 170))

    @property
    def source(self):
        return self.path.relative_to(ROOT).as_posix()

    def url(self):
        return page_url(f"guide/{self.slug}", self.lang)


def chapters(lang):
    pattern = "[0-9]*.en.md" if lang == "en" else "[0-9]*.md"
    found = [Chapter(p, lang) for p in sorted((ROOT / "docs" / "guide").glob(pattern))]
    return [c for c in found if (lang == "en") == c.path.name.endswith(".en.md")]


def linked_pages():
    """Repository files with a page of their own; links to anything else go to GitHub."""
    pages = {path: page_url(slug, lang) for (slug, lang), path in DOCS.items()}
    pages["README.md"] = page_url("", "ru")
    pages["README.en.md"] = page_url("", "en")
    for lang in UI:
        pages[f"docs/guide/README{'.en' if lang == 'en' else ''}.md"] = page_url("guide", lang)
        for c in chapters(lang):
            pages[c.source] = c.url()
    return pages


def rewrite_links(body, source, here, pages):
    source_dir = posixpath.dirname(source)

    def fix(match):
        attr, url = match.group(1), match.group(2)
        if re.match(r"^[a-z]+:|^#|^//", url):
            return match.group(0)
        path, _, anchor = url.partition("#")
        target = posixpath.normpath(posixpath.join(source_dir, path)) if path else source
        new = rel(here, pages[target]) if target in pages else f"{REPO}/blob/main/{target}"
        return f'{attr}="{new}{"#" + anchor if anchor else ""}"'

    return re.sub(r'(href|src)="([^"]+)"', fix, body)


def render_markdown(text, source, here, pages):
    text = re.sub(r"^<!--.*?-->\s*", "", text, flags=re.S)
    title = re.search(r"^# (.+)$", text, re.M).group(1).strip()
    text = re.sub(r"^# .+\n", "", text, count=1, flags=re.M)
    # The navigation line under the title is replaced by the site navigation.
    text = re.sub(r"^\s*(\[[^\]]+\]\([^)]+\)(\s*·\s*)?)+\s*\n", "", text, count=1)
    # GitHub accepts list continuations indented by two spaces, Python-Markdown wants four.
    text = re.sub(r"^  (?=\S)", "    ", text, flags=re.M)
    md = markdown.Markdown(extensions=["tables", "fenced_code", "toc", "attr_list", "sane_lists"],
                           extension_configs={"toc": {"slugify": slugify_unicode, "toc_depth": "2-3"}})
    body = rewrite_links(md.convert(text), source, here, pages)
    body = re.sub(r'<blockquote>\s*<p>\[!(NOTE|TIP|IMPORTANT|WARNING|CAUTION)\]\s*',
                  lambda m: f'<blockquote class="note"><p class="note-title">{m.group(1)}</p><p>', body)
    return title, body, md.toc_tokens


# Page chrome

def head(lang, title, description, here, slug):
    root = rel(here, "")
    other = UI[lang]["other"]
    return f"""<!doctype html>
<html lang="{lang}">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{html.escape(title)}</title>
<meta name="description" content="{html.escape(description)}">
<meta name="color-scheme" content="dark light">
<link rel="canonical" href="{SITE}{page_url(slug, lang)}">
<link rel="alternate" hreflang="{lang}" href="{SITE}{page_url(slug, lang)}">
<link rel="alternate" hreflang="{other}" href="{SITE}{page_url(slug, other)}">
<meta property="og:title" content="{html.escape(title)}">
<meta property="og:description" content="{html.escape(description)}">
<meta property="og:type" content="website">
<link rel="icon" href="{root}assets/mark.svg" type="image/svg+xml">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;500;700;800&display=swap">
<link rel="stylesheet" href="{root}assets/site.css">
</head>
<body>
"""


def topbar(lang, here, slug, section):
    ui = UI[lang]
    root = rel(here, "")
    nav = "".join(
        f'<a class="chip{" on" if s == section else ""}" href="{rel(here, page_url(s, lang))}">[ {label} ]</a>'
        for s, label in ui["nav"])
    langs = "".join(
        f'<a class="lang{" on" if code == lang else ""}" href="{rel(here, page_url(slug, code))}" hreflang="{code}">{code.upper()}</a>'
        for code in ("en", "ru"))
    return f"""<header class="topbar">
<a class="brand" href="{rel(here, page_url('', lang))}"><img src="{root}assets/mark.svg" width="26" height="26" alt=""><span>DEVICE TWEAKER</span></a>
<nav class="nav">{nav}<a class="chip" href="{REPO}">[ GITHUB ]</a></nav>
<div class="langs">{langs}</div>
</header>
"""


def status_line(lang, ver):
    items = [(k, v.format(ver=ver)) for k, v in UI[lang]["status"]]
    return '<p class="status">' + '<span class="sep">|</span>'.join(f"<span><b>{k}</b> - {v}</span>" for k, v in items) + "</p>"


def footer(lang, here, ver):
    ui = UI[lang]
    rows = "".join(f'<p><span class="key">{k}:</span> <a href="{url}">{v}</a></p>' for k, v, url in ui["footer"])
    return f"""<footer class="footer">
<div class="footer-info">{rows}<p><span class="key">DEVICE TWEAKER:</span> {ver} · © arsenzaaa · {ui['license']}</p></div>
</footer>
<script src="{rel(here, '')}assets/site.js"></script>
</body>
</html>
"""


# Pages

def card(c, here, ui):
    return f"""<article class="card" data-tag="{html.escape(c.tag)}" data-text="{html.escape((c.title + ' ' + c.summary + ' ' + c.text).lower())}">
<a class="card-head" href="{rel(here, c.url())}"><span class="tag">[{c.tag}]</span><span class="num">{c.number}</span><span class="title">{c.title}</span></a>
<p class="card-text">{html.escape(c.summary)}</p>
<div class="card-foot"><span class="kv"><span class="key">{ui['chapter']}:</span> {c.number}</span><span class="kv"><span class="key">{ui['reading']}:</span> {c.minutes} {ui['minutes']}</span>
<a class="button small" href="{rel(here, c.url())}">{ui['open']}</a></div>
</article>"""


def guide_index(lang, here, ui):
    chs = chapters(lang)
    tags = []
    for c in chs:
        if c.tag not in tags:
            tags.append(c.tag)
    chips = f'<button class="chip on" data-filter="*">[ {ui["all"]} ]</button>' + "".join(
        f'<button class="chip" data-filter="{html.escape(t)}">[ {html.escape(t)} ]</button>' for t in tags)
    cards = "".join(card(c, here, ui) for c in chs)
    return f"""<div class="filterbar"><input class="search" type="search" placeholder="{ui['search']}" aria-label="{ui['search']}">
<div class="chips">{chips}</div></div>
<div class="cards">{cards}</div>
<p class="empty" hidden>{ui['nothing']}</p>"""


def landing(lang, ver):
    t, ui = LANDING[lang], UI[lang]
    here = page_url("", lang)
    root = rel(here, "")
    suffix = "en" if lang == "en" else "ru"
    shot = f"{root}assets/screenshots/showcase_intel_14900k_full_{suffix}.png"
    shot_h = image_height(ROOT / "assets" / "screenshots" / "readme" / f"showcase_intel_14900k_full_{suffix}.png")
    files = "".join(
        f'<p><span class="key">{html.escape(desc)}:</span> <a href="{REPO}/releases/latest/download/{name}">{name}</a></p>'
        for name, desc in t["files"])
    steps = "".join(f'<li><span class="step">{n}</span><span>{text}</span></li>' for n, text in t["steps"])
    ch11 = next(c for c in chapters(lang) if c.number == "11")
    shots = "".join(
        f'<figure class="shot-card"><figcaption><span class="tag">[{tag}]</span> {cap}</figcaption>'
        f'<a href="{root}assets/screenshots/raw/{name}_{suffix}.png"><img src="{root}assets/screenshots/thumbs/{name}_{suffix}.png" srcset="{root}assets/screenshots/thumbs/{name}_{suffix}.png 366w, {root}assets/screenshots/raw/{name}_{suffix}.png 1360w" sizes="(max-width: 980px) 100vw, 366px" alt="{html.escape(cap)}" width="366" height="275" loading="lazy"></a></figure>'
        for name, tag, cap in t["shots"])
    body = f"""<main>
<section class="hero">
<p class="eyebrow">[ {t['eyebrow']} ]</p>
<h1 class="wordmark">DEVICE TWEAKER</h1>
<p class="lead">{t['lead']}</p>
<p class="sub">{t['sub']}</p>
<div class="actions">
<a class="button primary" href="{REPO}/releases/latest">{ui['download']} {ver}</a>
<a class="button" href="{rel(here, page_url('guide', lang))}">{ui['guide'].upper()}</a>
</div>
{status_line(lang, ver)}
</section>
<section class="band"><a class="shot" href="{root}assets/screenshots/raw/showcase_intel_14900k_full_{suffix}.png"><img src="{root}assets/screenshots/readme/showcase_intel_14900k_full_{suffix}.png" srcset="{root}assets/screenshots/readme/showcase_intel_14900k_full_{suffix}.png 1x, {shot} 2x" alt="{html.escape(t['shot_alt'])}" width="838" height="{shot_h}"></a></section>
<section class="section" id="guide">
<h2 class="label">[ {t['guide_title']} ]</h2>
<p class="section-lead">{t['guide_lead']}</p>
{guide_index(lang, here, ui)}
</section>
<section class="section" id="install">
<h2 class="label">[ {t['install_title']} ]</h2>
<div class="panel">
<div class="panel-head"><span class="tag">[EXE]</span> <b>DEVICE TWEAKER {ver}</b></div>
<div class="info">{files}</div>
<ol class="steps">{steps}</ol>
<div class="actions left"><a class="button primary" href="{REPO}/releases/latest">{ui['download']} {ver}</a><a class="button" href="{rel(here, ch11.url())}">{t['driver_link']}</a></div>
</div>
<div class="note"><p class="note-title">DTIMOD.sys / KDU</p><p>{t['driver']}</p></div>
</section>
<section class="section" id="screenshots">
<h2 class="label">[ {t['shots_title']} ]</h2>
<div class="shots">{shots}</div>
<p class="faint">{t['shots_note']}</p>
</section>
</main>
"""
    if lang == "ru":
        body = typography(body)
    return head(lang, t["title"], t["description"], here, "") + topbar(lang, here, "", "") + body + footer(lang, here, ver)


def guide_page(lang, ver):
    ui = UI[lang]
    here = page_url("guide", lang)
    t = LANDING[lang]
    body = f"""<main class="wide">
<section class="section first">
<p class="crumbs"><a href="{rel(here, page_url('', lang))}">DEVICE TWEAKER</a> / {ui['guide']}</p>
<h1 class="page-title">{ui['guide']}</h1>
<p class="section-lead">{t['guide_lead']}</p>
{guide_index(lang, here, ui)}
</section>
</main>
"""
    if lang == "ru":
        body = typography(body)
    return head(lang, f"{ui['guide']} — DEVICE TWEAKER", t["guide_lead"], here, "guide") + topbar(lang, here, "guide", "guide") + body + footer(lang, here, ver)


def toc(tokens):
    if not tokens:
        return ""
    items = []
    for token in tokens:
        children = "".join(f'<li><a href="#{c["id"]}">{c["name"]}</a></li>' for c in token["children"])
        items.append(f'<li><a href="#{token["id"]}">{token["name"]}</a>{"<ul>" + children + "</ul>" if children else ""}</li>')
    return "<ul>" + "".join(items) + "</ul>"


def checklist(entries, here, current):
    """The left list, drawn like the CPU Affinity checklist of the program."""
    rows = "".join(
        f'<li><a href="{rel(here, url)}"{" aria-current=\"page\"" if url == current else ""}><span class="box"></span>{label}</a></li>'
        for url, label in entries)
    return f'<ul class="checklist">{rows}</ul>'


def doc_layout(lang, ver, here, slug, section, title, body, tokens, source, crumbs, entries, list_title, prevnext="", tag=""):
    ui = UI[lang]
    head_tag = f'<span class="tag">[{tag}]</span> ' if tag else ""
    content = f"""<main class="docs">
<aside class="sidebar"><div class="side">
<p class="side-label">{list_title}</p>
{checklist(entries, here, here)}
<p class="side-label">{ui['on_page']}</p>
<div class="side-toc">{toc(tokens)}</div>
</div></aside>
<article class="doc">
<p class="crumbs">{crumbs}</p>
<h1>{head_tag}{html.escape(title)}</h1>
{body}
{prevnext}
<p class="source"><a href="{REPO}/blob/main/{source}">{ui['source']}: {source}</a></p>
</article>
</main>
"""
    if lang == "ru":
        content = typography(content)
    description = re.sub(r"<[^>]+>", "", (re.search(r"<p>(.*?)</p>", body, re.S) or re.search(r"(.*)", title)).group(1))[:200]
    return head(lang, f"{title} — DEVICE TWEAKER", description, here, slug) + topbar(lang, here, slug, section) + content + footer(lang, here, ver)


def chapter_page(c, lang, ver, pages):
    ui = UI[lang]
    here = c.url()
    chs = chapters(lang)
    _, body, tokens = render_markdown(c.text, c.source, here, pages)
    i = [x.slug for x in chs].index(c.slug)
    prev_link = f'<a class="button" href="{rel(here, chs[i - 1].url())}">← {ui["prev"]}: {chs[i - 1].number}</a>' if i > 0 else "<span></span>"
    next_link = f'<a class="button" href="{rel(here, chs[i + 1].url())}">{ui["next"]}: {chs[i + 1].number} →</a>' if i < len(chs) - 1 else "<span></span>"
    entries = [(x.url(), f'<span class="n">{x.number}</span> {x.title}') for x in chs]
    crumbs = f'<a href="{rel(here, page_url("", lang))}">DEVICE TWEAKER</a> / <a href="{rel(here, page_url("guide", lang))}">{ui["guide"]}</a> / {c.number}'
    return doc_layout(lang, ver, here, f"guide/{c.slug}", "guide", c.title, body, tokens, c.source, crumbs, entries,
                      ui["contents"], f'<nav class="prevnext">{prev_link}{next_link}</nav>', c.tag)


def reference_page(slug, lang, ver, pages):
    ui = UI[lang]
    here = page_url(slug, lang)
    source = DOCS[(slug, lang)]
    title, body, tokens = render_markdown((ROOT / source).read_text(encoding="utf-8-sig"), source, here, pages)
    entries = [(page_url(s, lang), label) for s, label in ui["docs_nav"]]
    crumbs = f'<a href="{rel(here, page_url("", lang))}">DEVICE TWEAKER</a> / {ui["reference"]}'
    section = "changelog" if slug == "changelog" else "settings"
    return doc_layout(lang, ver, here, slug, section, title, body, tokens, source, crumbs, entries, ui["reference"])


def write(out, url, text):
    target = out / url / "index.html"
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(text, encoding="utf-8")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--out", default=Path("_site"), type=Path)
    args = parser.parse_args()
    out = args.out.resolve()
    if out.exists():
        shutil.rmtree(out)
    assets = out / "assets"
    (assets / "screenshots" / "raw").mkdir(parents=True)
    for name in ("site.css", "site.js"):
        shutil.copy(HERE / name, assets / name)
    for svg in (ROOT / "assets" / "branding").glob("*.svg"):
        shutil.copy(svg, assets / svg.name)
    for png in (ROOT / "assets" / "screenshots" / "framed").glob("*.png"):
        shutil.copy(png, assets / "screenshots" / png.name)
    for png in (ROOT / "assets" / "screenshots").glob("*.png"):
        shutil.copy(png, assets / "screenshots" / "raw" / png.name)
    for folder in ("readme", "thumbs"):
        (assets / "screenshots" / folder).mkdir()
        for png in (ROOT / "assets" / "screenshots" / folder).glob("*.png"):
            shutil.copy(png, assets / "screenshots" / folder / png.name)
    (out / ".nojekyll").write_text("", encoding="utf-8")

    ver = version()
    pages = linked_pages()
    for lang in UI:
        write(out, page_url("", lang), landing(lang, ver))
        write(out, page_url("guide", lang), guide_page(lang, ver))
        for c in chapters(lang):
            write(out, c.url(), chapter_page(c, lang, ver, pages))
        for slug in ("settings", "development", "changelog"):
            write(out, page_url(slug, lang), reference_page(slug, lang, ver, pages))
    print(f"Site built in {out}")


if __name__ == "__main__":
    main()
