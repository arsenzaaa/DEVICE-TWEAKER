"""Builds the project site (GitHub Pages) from the repository documents.

    pip install markdown==3.7
    python dev/site/build.py --out _site

Russian pages are at the root, English ones under /en/. The landing page text lives below in
LANDING; the settings guide, development notes and changelog are rendered from docs/ and
CHANGELOG*.md, so the site never drifts from the repository.
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

# Documents rendered as pages: (slug, language) -> source file.
PAGES = {
    ("settings", "ru"): "docs/SETTINGS.md",
    ("settings", "en"): "docs/SETTINGS.en.md",
    ("development", "ru"): "docs/DEVELOPMENT.md",
    ("development", "en"): "docs/DEVELOPMENT.en.md",
    ("changelog", "ru"): "CHANGELOG.md",
    ("changelog", "en"): "CHANGELOG.en.md",
}
# Repository files that have a page of their own; everything else links to GitHub.
LINKED = {path: key for key, path in PAGES.items()}
LINKED["README.md"] = ("", "ru")
LINKED["README.en.md"] = ("", "en")

UI = {
    "ru": {
        "lang": "ru", "other": "en", "other_label": "EN",
        "nav": [("", "Обзор"), ("settings", "Настройки"), ("development", "Разработка"), ("changelog", "Изменения")],
        "download": "Скачать", "docs": "Документация", "on_page": "На странице", "source": "Исходный текст страницы",
        "footer": "Распространяется по лицензии GPL-3.0.",
    },
    "en": {
        "lang": "en", "other": "ru", "other_label": "RU",
        "nav": [("", "Overview"), ("settings", "Settings"), ("development", "Development"), ("changelog", "Changelog")],
        "download": "Download", "docs": "Documentation", "on_page": "On this page", "source": "Page source",
        "footer": "Distributed under the GPL-3.0 license.",
    },
}

TAG_COLORS = ("cpu", "irq", "usb", "net", "sys", "dev", "io", "bak")

LANDING = {
    "ru": {
        "title": "DEVICE TWEAKER — прерывания устройств Windows под топологию процессора",
        "description": "MSI, Interrupt Affinity, ReservedCpuSets, RSS, USB IMOD и NIC ITR в одной программе для Windows 10 / 11 x64, с интеллектуальной авто-оптимизацией.",
        "eyebrow": "WINDOWS 10 / 11 x64",
        "lead": "Прерывания устройств Windows под топологию процессора.",
        "sub": "MSI Utility V3 + Interrupt Affinity Policy Tool (IAPT) + ReservedCpuSets + RSS + USB IMOD / NIC ITR в одном интерфейсе, с интеллектуальной авто-оптимизацией.",
        "meta": "GPL-3.0 · .NET 8 · RU / EN",
        "shot_alt": "DEVICE TWEAKER: USB-контроллер на Intel Core i9-14900K",
        "features_title": "ВОЗМОЖНОСТИ",
        "features": [
            ("cpu", "CPU", "Авто-оптимизация", "План по топологии процессора и ролям устройств: P- / E-cores, SMT / HT, CPPC, CCD / CCX. Если топология определена ненадёжно, запись не начинается."),
            ("irq", "IRQ", "Прерывания", "MSI Mode, MSI Limit, IRQ Priority, IRQ Policy и Interrupt Affinity. Режим из реестра, текущий IRQ Mode и возможности устройства показаны раздельно."),
            ("usb", "USB", "USB IMOD", "Интервал модерации каждого interrupter'а xHCI с шагом 250 нс: по ролям устройств, на контроллер или по индексам."),
            ("net", "NET", "NIC ITR и RSS", "Регистры модерации Intel и Realtek, включая Killer. RSS: базовый CPU и очереди, по одному процессору на физическое ядро."),
            ("sys", "SYS", "Автозапуск без EXE", "Профиль IMOD / NIC ITR применяет задача Планировщика при входе, после сна и при подключении USB. Файлы сверяются по SHA-256."),
            ("dev", "DEV", "Определение устройств", "GPU, USB, сеть, аудио и накопители. Для USB-контроллеров видно, где мышь, клавиатура, геймпад, микрофон, аудио и камера."),
            ("io", "I/O", "Ввод и питание", "Polling Rate с замером Raw Input, RawMouseThrottleDuration, энергосбережение USB и сети, ReservedCpuSets."),
            ("bak", "BAK", "Бекапы и логи", "Проверенный бекап перед каждым применением и возврат к исходному состоянию. Лог: что применилось, что пропущено и почему."),
        ],
        "auto_title": "АВТО-ОПТИМИЗАЦИЯ",
        "auto": [
            "Дискретная GPU получает два разных физических ядра: SMT-соседи одного ядра за два не считаются. Контроллер с мышью и клавиатурой, сеть и аудио распределяются по своим правилам, на процессорах с несколькими CCD выбирается целевой CCD.",
            "Накопителям включается MSI и высокий приоритет, affinity остаётся за Windows. Встроенной GPU affinity возвращается Windows, Wi-Fi не трогается. Перед первой авто-оптимизацией сохраняется исходное состояние.",
        ],
        "auto_alt": "GPU: CPU 4 на ядре 2 и CPU 8 на ядре 4, SMT-соседи CPU 5 и CPU 9 свободны",
        "writes_title": "ЧТО МЕНЯЕТ ПРОГРАММА",
        "writes_head": ("Настройка", "Куда записывается"),
        "writes": [
            ("MSI Mode, MSI Limit", "<code>MSISupported</code>, <code>MessageNumberLimit</code> в <code>Interrupt Management\\MessageSignaledInterruptProperties</code>"),
            ("IRQ Priority", "<code>DevicePriority</code> в <code>Interrupt Management\\Affinity Policy</code>"),
            ("IRQ Policy, Interrupt Affinity", "<code>DevicePolicy</code>, <code>AssignmentSetOverride</code> там же"),
            ("RSS", "<code>*RssBaseProcNumber</code>, <code>*NumRssQueues</code>, <code>*MaxRssProcessors</code>, <code>*RssMaxProcNumber</code> в ключе класса адаптера"),
            ("ReservedCpuSets", "<code>ReservedCpuSets</code> в <code>Session Manager\\kernel</code>"),
            ("Mouse Throttle", "<code>RawMouseThrottleDuration</code> в <code>HKCU\\Control Panel\\Mouse</code>"),
            ("Энергосбережение", "<code>PnPCapabilities</code>, <code>SelectiveSuspendEnabled</code>, <code>EnhancedPowerManagementEnabled</code>, USB selective suspend в плане питания"),
            ("USB IMOD", "регистр IMOD каждого interrupter'а xHCI, шаг 250 нс"),
            ("NIC ITR", "EITR / ITR у Intel, IntrMit у Realtek"),
        ],
        "writes_note": "Настройки реестра вступают в силу после перезагрузки. Регистры IMOD и NIC ITR меняются сразу и после записи читаются повторно.",
        "install_title": "УСТАНОВКА",
        "files": [
            ("DEVICE.TWEAKER.SelfContained.exe", "нет, .NET встроен"),
            ("DEVICE.TWEAKER.exe", '<a href="https://dotnet.microsoft.com/download/dotnet/8.0">.NET 8 Desktop Runtime x64</a>'),
        ],
        "files_head": ("Файл", "Требования"),
        "steps": [
            ("Запуск", "Запустите программу от имени администратора."),
            ("Применение", "<b>АВТООПТИМИЗАЦИЯ</b> строит план и применяет его. Ручная настройка — <b>ПРИМЕНИТЬ</b>."),
            ("Проверка", "Перезагрузите компьютер. <b>ПРОВЕРИТЬ</b> в блоке IMOD или NIC ITR покажет фактические значения регистров."),
        ],
        "driver": "Регистры IMOD и NIC ITR пишет собственный драйвер DTIMOD.sys, и только в память, которую Windows выделила выбранному устройству. Публичной подписи у драйвера нет, его загружает KDU, поэтому антивирусы могут относить программу к HackTool. HVCI и Microsoft Vulnerable Driver Blocklist блокируют такую загрузку, FACEIT и Vanguard могут с ней конфликтовать: состояние показывает <b>ПРОВЕРКА СИСТЕМЫ</b>. Остальные настройки от драйвера не зависят.",
        "driver_label": "Важно",
        "verify": "SHA-256 обоих EXE указаны в <code>SHA256SUMS.txt</code> на странице релиза, исходников — в <code>SOURCE-SHA256SUMS.txt</code>.",
        "rollback": "Откат: <b>ВОССТАНОВЛЕНИЕ → ИСХОДНЫЙ</b> возвращает реестр к состоянию до первой авто-оптимизации и удаляет задачу автозапуска. Бекапы хранятся в папке <code>Backups</code> рядом с EXE.",
        "gallery_title": "СКРИНШОТЫ",
        "gallery": [
            ("showcase_gpu_full", "GPU на двух физических ядрах"),
            ("showcase_nic_full", "NIC ITR и RSS, Intel I226-V"),
            ("showcase_amd_imod_full", "USB IMOD, Ryzen 9 9950X3D"),
        ],
        "gallery_note": "На скриншотах тестовые устройства и значения.",
        "more": "Подробно о каждой настройке",
    },
    "en": {
        "title": "DEVICE TWEAKER — device interrupts on Windows, tuned to your CPU topology",
        "description": "MSI, Interrupt Affinity, ReservedCpuSets, RSS, USB IMOD and NIC ITR in one program for Windows 10 / 11 x64, with intelligent auto-optimization.",
        "eyebrow": "WINDOWS 10 / 11 x64",
        "lead": "Device interrupts on Windows, tuned to your CPU topology.",
        "sub": "MSI Utility V3 + Interrupt Affinity Policy Tool (IAPT) + ReservedCpuSets + RSS + USB IMOD / NIC ITR in one interface, with intelligent auto-optimization.",
        "meta": "GPL-3.0 · .NET 8 · RU / EN",
        "shot_alt": "DEVICE TWEAKER: a USB controller on Intel Core i9-14900K",
        "features_title": "FEATURES",
        "features": [
            ("cpu", "CPU", "Auto-optimization", "A plan built from the CPU topology and device roles: P- / E-cores, SMT / HT, CPPC, CCD / CCX. If the topology cannot be detected reliably, nothing is written."),
            ("irq", "IRQ", "Interrupts", "MSI Mode, MSI Limit, IRQ Priority, IRQ Policy and Interrupt Affinity. The registry setting, the current IRQ Mode and the device's capabilities are shown separately."),
            ("usb", "USB", "USB IMOD", "The moderation interval of each xHCI interrupter in 250 ns steps: by device role, per controller or by index."),
            ("net", "NET", "NIC ITR and RSS", "Moderation registers of Intel and Realtek adapters, including Killer. RSS: base CPU and queues, one processor per physical core."),
            ("sys", "SYS", "Startup without the EXE", "An IMOD / NIC ITR profile is applied by a Task Scheduler task at sign-in, after sleep and when USB devices connect. Files are checked by SHA-256."),
            ("dev", "DEV", "Device detection", "GPUs, USB, network, audio and storage. USB controllers show where the mouse, keyboard, gamepad, microphone, audio and camera are."),
            ("io", "I/O", "Input and power", "Polling Rate with a Raw Input measurement, RawMouseThrottleDuration, USB and network power saving, ReservedCpuSets."),
            ("bak", "BAK", "Backups and logs", "A verified backup before every apply and a return to the original state. The log shows what was applied, what was skipped and why."),
        ],
        "auto_title": "AUTO-OPTIMIZATION",
        "auto": [
            "A discrete GPU gets two different physical cores: two SMT siblings of one core do not count as two. The controller with the mouse and keyboard, the network and audio have their own placement rules; on CPUs with several CCDs a target CCD is chosen.",
            "Storage gets MSI and high priority while affinity stays with Windows. Integrated GPU affinity goes back to Windows, Wi-Fi is left untouched. The original state is saved before the first auto-optimization.",
        ],
        "auto_alt": "GPU: CPU 4 on core 2 and CPU 8 on core 4, SMT siblings CPU 5 and CPU 9 stay free",
        "writes_title": "WHAT THE PROGRAM CHANGES",
        "writes_head": ("Setting", "Where it is written"),
        "writes": [
            ("MSI Mode, MSI Limit", "<code>MSISupported</code>, <code>MessageNumberLimit</code> in <code>Interrupt Management\\MessageSignaledInterruptProperties</code>"),
            ("IRQ Priority", "<code>DevicePriority</code> in <code>Interrupt Management\\Affinity Policy</code>"),
            ("IRQ Policy, Interrupt Affinity", "<code>DevicePolicy</code>, <code>AssignmentSetOverride</code> in the same key"),
            ("RSS", "<code>*RssBaseProcNumber</code>, <code>*NumRssQueues</code>, <code>*MaxRssProcessors</code>, <code>*RssMaxProcNumber</code> in the adapter class key"),
            ("ReservedCpuSets", "<code>ReservedCpuSets</code> in <code>Session Manager\\kernel</code>"),
            ("Mouse Throttle", "<code>RawMouseThrottleDuration</code> in <code>HKCU\\Control Panel\\Mouse</code>"),
            ("Power saving", "<code>PnPCapabilities</code>, <code>SelectiveSuspendEnabled</code>, <code>EnhancedPowerManagementEnabled</code>, USB selective suspend in the power plan"),
            ("USB IMOD", "the IMOD register of each xHCI interrupter, 250 ns steps"),
            ("NIC ITR", "EITR / ITR on Intel, IntrMit on Realtek"),
        ],
        "writes_note": "Registry settings take effect after a restart. IMOD and NIC ITR registers change immediately and are read back after writing.",
        "install_title": "INSTALLATION",
        "files": [
            ("DEVICE.TWEAKER.SelfContained.exe", "none, .NET included"),
            ("DEVICE.TWEAKER.exe", '<a href="https://dotnet.microsoft.com/download/dotnet/8.0">.NET 8 Desktop Runtime x64</a>'),
        ],
        "files_head": ("File", "Requirements"),
        "steps": [
            ("Run", "Run the program as administrator."),
            ("Apply", "<b>AUTO-OPTIMIZATION</b> builds a plan and applies it. For manual settings use <b>APPLY</b>."),
            ("Check", "Restart the PC. <b>CHECK</b> in the IMOD or NIC ITR block shows the actual register values."),
        ],
        "driver": "IMOD and NIC ITR registers are written by the custom DTIMOD.sys driver, and only to memory that Windows assigned to the selected device. The driver has no public signature and is loaded by KDU, so antivirus software may flag the program as HackTool. HVCI and the Microsoft Vulnerable Driver Blocklist block this kind of loading, and FACEIT and Vanguard may conflict with it: <b>SYSTEM CHECK</b> shows their state. All other settings work without the driver.",
        "driver_label": "Important",
        "verify": "SHA-256 hashes of both EXEs are in <code>SHA256SUMS.txt</code> on the release page, and of the sources in <code>SOURCE-SHA256SUMS.txt</code>.",
        "rollback": "Rollback: <b>RESTORE → ORIGINAL</b> returns the registry to its state before the first auto-optimization and removes the startup task. Backups are kept in the <code>Backups</code> folder next to the EXE.",
        "gallery_title": "SCREENSHOTS",
        "gallery": [
            ("showcase_gpu_full", "GPU on two physical cores"),
            ("showcase_nic_full", "NIC ITR and RSS, Intel I226-V"),
            ("showcase_amd_imod_full", "USB IMOD, Ryzen 9 9950X3D"),
        ],
        "gallery_note": "Screenshots show test devices and values.",
        "more": "Every setting in detail",
    },
}

GITHUB_ICON = ('<svg viewBox="0 0 16 16" width="18" height="18" aria-hidden="true"><path fill="currentColor" d="M8 0C3.58 0 0 3.58 0 8c0 3.54 2.29 6.53 5.47 7.59.4.07.55-.17.55-.38 0-.19-.01-.82-.01-1.49-2.01.37-2.53-.49-2.69-.94-.09-.23-.48-.94-.82-1.13-.28-.15-.68-.52-.01-.53.63-.01 1.08.58 1.23.82.72 1.21 1.87.87 2.33.66.07-.52.28-.87.51-1.07-1.78-.2-3.64-.89-3.64-3.95 0-.87.31-1.59.82-2.15-.08-.2-.36-1.02.08-2.12 0 0 .67-.21 2.2.82.64-.18 1.32-.27 2-.27.68 0 1.36.09 2 .27 1.53-1.04 2.2-.82 2.2-.82.44 1.1.16 1.92.08 2.12.51.56.82 1.27.82 2.15 0 3.07-1.87 3.75-3.65 3.95.29.25.54.73.54 1.48 0 1.07-.01 1.93-.01 2.2 0 .21.15.46.55.38A8.013 8.013 0 0016 8c0-4.42-3.58-8-8-8z"/></svg>')


def version():
    match = re.search(r"<Version>([^<]+)</Version>", (ROOT / "src" / "DeviceTweakerCS.csproj").read_text(encoding="utf-8"))
    return match.group(1)


def page_url(slug, lang):
    """Site-relative path of a page, ending with a slash."""
    prefix = "en/" if lang == "en" else ""
    return prefix + (slug + "/" if slug else "")


def rel(from_url, to_url):
    """Relative link from one page directory to another site path ('' is the site root)."""
    target = posixpath.relpath(to_url.rstrip("/") or ".", from_url.rstrip("/") or ".")
    if target == ".":
        return "./"
    return target + "/" if to_url == "" or to_url.endswith("/") else target


def rewrite_links(body, source, here):
    """Points links between documents at their pages; other repository files go to GitHub."""
    source_dir = posixpath.dirname(source)

    def fix(match):
        attr, url = match.group(1), match.group(2)
        if re.match(r"^[a-z]+:|^#|^//", url):
            return match.group(0)
        path, _, anchor = url.partition("#")
        target = posixpath.normpath(posixpath.join(source_dir, path)) if path else source
        if target in LINKED:
            slug, lang = LINKED[target]
            new = rel(here, page_url(slug, lang))
        else:
            new = f"{REPO}/blob/main/{target}"
        return f'{attr}="{new}{"#" + anchor if anchor else ""}"'

    return re.sub(r'(href|src)="([^"]+)"', fix, body)


ALERT = re.compile(r'<blockquote>\s*<p>\[!(NOTE|TIP|IMPORTANT|WARNING|CAUTION)\]\s*', re.S)


def alerts(body):
    def open_(match):
        kind = match.group(1).lower()
        return f'<blockquote class="alert alert-{kind}"><p class="alert-title">{match.group(1).title()}</p><p>'
    return ALERT.sub(open_, body)


def render_document(source, here):
    text = (ROOT / source).read_text(encoding="utf-8-sig")
    title = re.search(r"^# (.+)$", text, re.M).group(1).strip()
    text = re.sub(r"^# .+\n", "", text, count=1, flags=re.M)
    # The navigation line under the title is replaced by the site navigation.
    text = re.sub(r"^\s*(\[[^\]]+\]\([^)]+\)(\s*·\s*)?)+\s*\n", "", text, count=1)
    # GitHub accepts list continuations indented by two spaces, Python-Markdown wants four.
    text = re.sub(r"^  (?=\S)", "    ", text, flags=re.M)
    md = markdown.Markdown(extensions=["tables", "fenced_code", "toc", "attr_list", "sane_lists"],
                           extension_configs={"toc": {"slugify": slugify_unicode, "toc_depth": "2-3"}})
    body = md.convert(text)
    body = alerts(rewrite_links(body, source, here))
    return title, body, md.toc_tokens


def head(lang, title, description, here, slug):
    root = rel(here, "")
    alt_lang = UI[lang]["other"]
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
<link rel="alternate" hreflang="{alt_lang}" href="{SITE}{page_url(slug, alt_lang)}">
<meta property="og:title" content="{html.escape(title)}">
<meta property="og:description" content="{html.escape(description)}">
<meta property="og:type" content="website">
<link rel="icon" href="{root}assets/mark.svg" type="image/svg+xml">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;700&display=swap">
<link rel="stylesheet" href="{root}assets/site.css">
</head>
"""


def topbar(lang, here, slug):
    ui = UI[lang]
    root = rel(here, "")
    links = "".join(
        f'<a href="{rel(here, page_url(s, lang))}"{" aria-current=\"page\"" if s == slug else ""}>{label}</a>'
        for s, label in ui["nav"])
    other = rel(here, page_url(slug, ui["other"]))
    return f"""<header class="topbar">
<div class="topbar-inner">
<a class="brand" href="{rel(here, page_url('', lang))}"><img src="{root}assets/mark.svg" width="26" height="26" alt=""><span>DEVICE TWEAKER</span></a>
<nav class="nav">{links}</nav>
<div class="topbar-end">
<a class="lang" href="{other}" hreflang="{ui['other']}">{ui['other_label']}</a>
<a class="icon" href="{REPO}" aria-label="GitHub">{GITHUB_ICON}</a>
</div>
</div>
</header>
"""


def footer(lang, here):
    ui = UI[lang]
    links = " · ".join(f'<a href="{rel(here, page_url(s, lang))}">{label}</a>' for s, label in ui["nav"])
    return f"""<footer class="footer">
<div class="footer-inner">
<div class="footer-brand"><img src="{rel(here, '')}assets/mark.svg" width="20" height="20" alt=""><span>DEVICE TWEAKER</span></div>
<p>{links} · <a href="{REPO}">GitHub</a> · <a href="https://t.me/arsenzaa">Telegram</a></p>
<p class="muted">© arsenzaaa. {ui['footer']}</p>
</div>
</footer>
</body>
</html>
"""


def picture(here, base, alt, width):
    root = rel(here, "")
    return (f'<picture><source media="(prefers-color-scheme: dark)" srcset="{root}assets/{base}-dark.svg">'
            f'<img src="{root}assets/{base}-light.svg" alt="{html.escape(alt)}" width="{width}"></picture>')


def section(label, inner, ident):
    return f'<section class="section" id="{ident}"><h2 class="label">[ {label} ]</h2>{inner}</section>'


def landing(lang, ver):
    t, ui = LANDING[lang], UI[lang]
    here = page_url("", lang)
    root = rel(here, "")
    suffix = "en" if lang == "en" else "ru"
    shot = f"{root}assets/screenshots/showcase_intel_14900k_full_{suffix}.png"
    cards = "".join(
        f'<article class="card"><span class="tag tag-{key}">[{tag}]</span><h3>{title}</h3><p>{text}</p></article>'
        for key, tag, title, text in t["features"])
    writes = "".join(f"<tr><td>{a}</td><td>{b}</td></tr>" for a, b in t["writes"])
    files = "".join(
        f'<tr><td><a href="{REPO}/releases/latest/download/{name}"><code>{name}</code></a></td><td>{req}</td></tr>'
        for name, req in t["files"])
    steps = "".join(f'<li><span class="step-title">{a}</span><span>{b}</span></li>' for a, b in t["steps"])
    gallery = "".join(
        f'<figure><a href="{root}assets/screenshots/{name}_{suffix}.png"><img src="{root}assets/screenshots/{name}_{suffix}.png" alt="{html.escape(cap)}" loading="lazy"></a><figcaption>{cap}</figcaption></figure>'
        for name, cap in t["gallery"])
    settings = rel(here, page_url("settings", lang))
    body = f"""<main>
<section class="hero">
<p class="eyebrow">[ {t['eyebrow']} ]</p>
<h1 class="wordmark">DEVICE TWEAKER</h1>
<p class="lead">{t['lead']}</p>
<p class="sub">{t['sub']}</p>
<div class="actions">
<a class="button primary" href="{REPO}/releases/latest">{ui['download'].upper()} {ver}</a>
<a class="button" href="{settings}">{ui['docs'].upper()}</a>
</div>
<p class="meta">v{ver} · {t['meta']}</p>
<a class="shot" href="{shot}"><img src="{shot}" alt="{html.escape(t['shot_alt'])}" width="1472" height="1412"></a>
</section>
{section(t['features_title'], f'<div class="cards">{cards}</div>', 'features')}
{section(t['auto_title'], '<div class="split"><div>' + ''.join(f'<p>{p}</p>' for p in t['auto']) + f'<p><a class="more" href="{settings}#{"интеллектуальная-авто-оптимизация-и-восстановление" if lang == "ru" else "intelligent-auto-optimization-and-restore"}">{t["more"]} →</a></p></div><div class="figure">{picture(here, "cpu-allocation-" + suffix, t["auto_alt"], 560)}</div></div>', 'auto')}
{section(t['writes_title'], f'<div class="table-wrap"><table><thead><tr><th>{t["writes_head"][0]}</th><th>{t["writes_head"][1]}</th></tr></thead><tbody>{writes}</tbody></table></div><p class="note">{t["writes_note"]} <a href="{settings}">{t["more"]} →</a></p>', 'writes')}
{section(t['install_title'], f'<ol class="steps">{steps}</ol><div class="table-wrap"><table><thead><tr><th>{t["files_head"][0]}</th><th>{t["files_head"][1]}</th></tr></thead><tbody>{files}</tbody></table></div><blockquote class="alert alert-important"><p class="alert-title">{t["driver_label"]}</p><p>{t["driver"]}</p></blockquote><p>{t["verify"]}</p><pre><code>(Get-FileHash .\\DEVICE.TWEAKER.SelfContained.exe -Algorithm SHA256).Hash\nSelect-String -Path .\\SHA256SUMS.txt -Pattern \'DEVICE.TWEAKER.SelfContained.exe\'</code></pre><p>{t["rollback"]}</p>', 'install')}
{section(t['gallery_title'], f'<div class="gallery">{gallery}</div><p class="note">{t["gallery_note"]}</p>', 'screenshots')}
</main>
"""
    return head(lang, t["title"], t["description"], here, "") + '<body class="landing">\n' + topbar(lang, here, "") + body + footer(lang, here)


def toc(tokens):
    if not tokens:
        return ""
    items = []
    for token in tokens:
        children = "".join(f'<li><a href="#{c["id"]}">{c["name"]}</a></li>' for c in token["children"])
        items.append(f'<li><a href="#{token["id"]}">{token["name"]}</a>{"<ul>" + children + "</ul>" if children else ""}</li>')
    return "<ul>" + "".join(items) + "</ul>"


def document(slug, lang, source):
    ui = UI[lang]
    here = page_url(slug, lang)
    title, body, tokens = render_document(source, here)
    pages = "".join(
        f'<li><a href="{rel(here, page_url(s, lang))}"{" aria-current=\"page\"" if s == slug else ""}>{label}</a></li>'
        for s, label in ui["nav"])
    description = re.sub(r"<[^>]+>", "", re.search(r"<p>(.*?)</p>", body, re.S).group(1))[:200]
    content = f"""<div class="docs">
<aside class="sidebar">
<nav><p class="side-label">{ui['docs']}</p><ul class="side-pages">{pages}</ul>
<p class="side-label">{ui['on_page']}</p><div class="side-toc">{toc(tokens)}</div></nav>
</aside>
<main class="doc">
<h1>{html.escape(title)}</h1>
{body}
<p class="source"><a href="{REPO}/blob/main/{source}">{ui['source']}: {source}</a></p>
</main>
</div>
"""
    return head(lang, f"{title} — DEVICE TWEAKER", description, here, slug) + "<body>\n" + topbar(lang, here, slug) + content + footer(lang, here)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--out", default=Path("_site"), type=Path)
    args = parser.parse_args()
    out = args.out.resolve()
    if out.exists():
        shutil.rmtree(out)
    assets = out / "assets"
    (assets / "screenshots").mkdir(parents=True)
    shutil.copy(HERE / "site.css", assets / "site.css")
    for svg in (ROOT / "assets" / "branding").glob("*.svg"):
        shutil.copy(svg, assets / svg.name)
    for png in (ROOT / "assets" / "screenshots" / "framed").glob("*.png"):
        shutil.copy(png, assets / "screenshots" / png.name)
    (out / ".nojekyll").write_text("", encoding="utf-8")

    ver = version()
    for lang in UI:
        target = out / page_url("", lang) / "index.html"
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(landing(lang, ver), encoding="utf-8")
    for (slug, lang), source in PAGES.items():
        target = out / page_url(slug, lang) / "index.html"
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(document(slug, lang, source), encoding="utf-8")
    print(f"Site built in {out}")


if __name__ == "__main__":
    main()
