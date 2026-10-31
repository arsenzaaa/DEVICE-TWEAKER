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
TELEGRAM = "https://t.me/arsenzaa"
NBSP = "\u00a0"

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
        "other": "en", "other_label": "EN",
        "nav": [("settings", "Документация"), ("changelog", "Изменения")],
        "docs_nav": [("", "Обзор"), ("settings", "Настройки"), ("development", "Разработка"), ("changelog", "Изменения")],
        "download": "Скачать", "docs": "Документация", "on_page": "На странице", "source": "Исходный текст страницы",
        "license": "Распространяется по лицензии GPL-3.0.",
    },
    "en": {
        "other": "ru", "other_label": "RU",
        "nav": [("settings", "Docs"), ("changelog", "Changelog")],
        "docs_nav": [("", "Overview"), ("settings", "Settings"), ("development", "Development"), ("changelog", "Changelog")],
        "download": "Download", "docs": "Documentation", "on_page": "On this page", "source": "Page source",
        "license": "Distributed under the GPL-3.0 license.",
    },
}

INTERRUPT_KEYS = {
    "ru": """HKLM\\SYSTEM\\CurrentControlSet\\Enum\\PCI\\<устройство>\\Device Parameters\\Interrupt Management
  MessageSignaledInterruptProperties  MSISupported           1 = MSI / MSI-X
                                      MessageNumberLimit     удаляется при Unlimited
  Affinity Policy                     DevicePriority         3 = High
                                      DevicePolicy           4 = SpecifiedProcessors
                                      AssignmentSetOverride  маска CPU""",
    "en": """HKLM\\SYSTEM\\CurrentControlSet\\Enum\\PCI\\<device>\\Device Parameters\\Interrupt Management
  MessageSignaledInterruptProperties  MSISupported           1 = MSI / MSI-X
                                      MessageNumberLimit     removed for Unlimited
  Affinity Policy                     DevicePriority         3 = High
                                      DevicePolicy           4 = SpecifiedProcessors
                                      AssignmentSetOverride  CPU mask""",
}
RSS_KEYS = {
    lang: f"""HKLM\\SYSTEM\\CurrentControlSet\\Control\\Class\\{{4d36e972-e325-11ce-bfc1-08002be10318}}\\<{name}>
  *RssBaseProcNumber  *NumRssQueues  *MaxRssProcessors  *RssMaxProcNumber"""
    for lang, name in (("ru", "адаптер"), ("en", "adapter"))
}
VERIFY = """(Get-FileHash .\\DEVICE.TWEAKER.SelfContained.exe -Algorithm SHA256).Hash
Select-String -Path .\\SHA256SUMS.txt -Pattern 'DEVICE.TWEAKER.SelfContained.exe'"""

# Landing page. Text uses **bold** and `code`; each section is (id, title, blocks), where a block is
# a paragraph string, ("items", [(term, text), ...]) or ("code", text).
LANDING = {
    "ru": {
        "title": "DEVICE TWEAKER — прерывания устройств Windows под топологию процессора",
        "description": "MSI, Interrupt Affinity, ReservedCpuSets, RSS, USB IMOD и NIC ITR в одной программе для Windows 10 / 11 x64, с интеллектуальной авто-оптимизацией.",
        "eyebrow": "WINDOWS 10 / 11 x64",
        "lead": "Прерывания устройств Windows под топологию процессора.",
        "sub": "MSI Utility V3 + Interrupt Affinity Policy Tool (IAPT) + ReservedCpuSets + RSS + USB IMOD / NIC ITR в одном интерфейсе, с интеллектуальной авто-оптимизацией.",
        "meta": "GPL-3.0 · .NET 8 · RU / EN",
        "shot_alt": "DEVICE TWEAKER: USB-контроллер на Intel Core i9-14900K",
        "sections": [
            ("why", "Зачем", [
                "ISR / DPC выполняются выше любого потока и не знают, какие ядра заняты игрой. Драйвер на ядре с важным потоком (GAME / RENDER / RHI) прерывает его -> поток ждёт или уходит на другое ядро и теряет свой L1 / L2 -> растёт время кадра, появляются статтеры. Два драйвера на одном ядре задерживают ISR / DPC друг друга, на одном SMT-SET делят кэш и EXECUTION-UNITS.",
                "DEVICE TWEAKER разносит прерывания по физическим ядрам: драйвера не делят ядра и SMT-SET, ядро 0 по возможности остаётся свободным.",
                ("items", [
                    ("MSI / Affinity", "MSI Mode, MSI Limit, IRQ Priority, IRQ Policy и Interrupt Affinity для каждого устройства"),
                    ("ReservedCpuSets", "ядра, на которые Windows не ставит обычные потоки"),
                    ("RSS", "базовый CPU и очереди проводных адаптеров NDIS"),
                    ("USB IMOD", "модерация прерываний interrupter'ов xHCI"),
                    ("NIC ITR", "модерация прерываний сетевых контроллеров Intel и Realtek"),
                ]),
            ]),
            ("auto", "Авто-оптимизация", [
                "**АВТООПТИМИЗАЦИЯ** читает топологию (P-Core / E-Core, SMT / HT, CPPC, CCD / CCX), находит роли устройств и строит план:",
                ("items", [
                    ("GPU", "-> два разных физических ядра. На одном ядре ISR видеокарты прерывает её же DPC -> спайки задержки. SMT-соседи одного ядра за два ядра не считаются."),
                    ("USB / сеть / аудио", "-> свои ядра по правилам совместного размещения. Контроллер с мышью, клавиатурой и геймпадом определяется автоматически."),
                    ("CCD", "-> на процессорах с несколькими CCD выбирается целевой CCD."),
                    ("Накопители", "-> MSI и высокий приоритет, affinity остаётся за Windows: на части NVMe `SpecifiedProcessors` не даёт Windows загрузиться."),
                    ("iGPU / Wi-Fi", "-> affinity встроенной GPU возвращается Windows, Wi-Fi не трогается."),
                ]),
                "Если топология определена ненадёжно, запись не начинается. Перед первым запуском сохраняется исходное состояние, его возвращает **ВОССТАНОВЛЕНИЕ → ИСХОДНЫЙ**.",
                ("figure", "cpu-allocation", "GPU: CPU 4 на ядре 2 и CPU 8 на ядре 4, SMT-соседи CPU 5 и CPU 9 свободны"),
            ]),
            ("interrupts", "Прерывания", [
                "**MSI MODE** — прерывания сообщениями вместо линии. Без MSI на трейсе xHCI выполнение DPC прерывается двумя ISR, с MSI не прерывается. **MSI LIMIT** — `Unlimited` удаляет ограничение числа сообщений.",
                "**IRQ PRIORITY** влияет на приоритет аппаратного прерывания. Это не приоритет игры и не приоритет DPC. **INTERRUPT AFFINITY** задаёт логические процессоры для IRQ устройства; Policy `SpecCPU` использует отмеченную маску, остальные политики оставляют выбор Windows.",
                ("code", INTERRUPT_KEYS["ru"]),
                "**IRQ MODE** показывает, что устройство использует сейчас: Line / MSI / MSI-X. **HW** — что оно поддерживает. **СЧИТАТЬ IRQ** уточняет режим по битам `MSI Enable` / `MSI-X Enable` в конфигурации PCI через DTIMOD.sys.",
            ]),
            ("imod", "USB IMOD", [
                "IMOD — интервал модерации прерываний у **interrupter'а** контроллера xHCI, а не у отдельной мыши. Шаг 250 нс: `0xC8` = 50 мкс (Windows по умолчанию), `0xFA0` = 1 мс, `0x0` = модерация выключена.",
                "Interrupter выдаётся устройству при подключении и после перезагрузки может смениться -> программа каждый раз заново читает карту «устройство → interrupter» из контекстов xHCI в памяти.",
                ("items", [
                    ("Устройства", "значения по ролям. Профиль AUTO: мышь `0x0`, аудио `0xFA0`, клавиатура и геймпад `0xC8`."),
                    ("XHCI", "одно значение на весь контроллер."),
                    ("Interrupters", "значения по индексам."),
                ]),
                "После записи регистры читаются повторно. **ПРОВЕРИТЬ** показывает фактические значения и время чтения.",
            ]),
            ("network", "Сеть", [
                "У адаптеров NDIS прерывания распределяет RSS, поэтому вместо маски настраиваются базовый CPU и очереди. Процессоры берутся по одному на физическое ядро, SMT-соседи пропускаются, число очередей ограничено значениями драйвера и векторами MSI-X.",
                ("code", RSS_KEYS["ru"]),
                "**NIC ITR** — регистры модерации самого контроллера: EITR у Intel I210 / I211, I225 / I226, I350, 82576 / 82580, ITR у I219, IntrMit у Realtek RTL8111 / 8168, RTL8125 / 8126, включая совместимые Killer. **СОХРАНИТЬ** выключает собственную модерацию драйвера (`*InterruptModeration`, у Intel также `ITR`) -> иначе драйвер перезаписывает регистр.",
            ]),
            ("startup", "Автозапуск", [
                "Регистры IMOD и NIC ITR держат значение до перезагрузки, сна или сброса контроллера. Сохранённый профиль применяет задача Планировщика `\\DEVICE TWEAKER\\IMOD` от SYSTEM: при входе в Windows, после выхода из сна и при подключении USB / HID-устройства.",
                "Задача запускает открытый `ApplyIMOD.ps1` из `%ProgramData%\\DEVICE TWEAKER\\Startup` и работает без EXE. Перед запуском сверяются SHA-256 runtime, драйвера и загрузчика, изменять папку могут только администраторы и SYSTEM.",
            ]),
            ("more", "Остальное", [
                ("items", [
                    ("ReservedCpuSets", "Windows не планирует на зарезервированные ядра обычные потоки. Драйвер с маской на такое ядро продолжает работать там, а потоки от него защищены."),
                    ("Mouse Throttle", "`RawMouseThrottleDuration`: частота Raw Input для фоновых окон Windows 11, 1000 / значение Гц."),
                    ("Polling Rate", "мышь и клавиатура по данным USB и замер частоты событий Raw Input."),
                    ("Энергосбережение", "USB-контроллеры, корневые USB-хабы, проводные адаптеры и USB selective suspend; значения перечитываются через WMI."),
                    ("Бекапы и логи", "проверенный бекап перед каждым применением, лог: что применилось, что пропущено и почему."),
                ]),
            ]),
            ("install", "Установка", [
                ("items", [
                    ("SelfContained.exe", ".NET встроен."),
                    ("DEVICE.TWEAKER.exe", "нужен .NET 8 Desktop Runtime x64."),
                ]),
                "Запустите от имени администратора -> **АВТООПТИМИЗАЦИЯ** или ручная настройка и **ПРИМЕНИТЬ** -> перезагрузка. **ПРОВЕРИТЬ** в блоке IMOD или NIC ITR покажет фактические значения регистров.",
                ("note", "DTIMOD.sys и KDU", "Регистры IMOD и NIC ITR недоступны из пользовательского режима, их пишет собственный драйвер DTIMOD.sys, и только в память, которую Windows выделила выбранному устройству. Публичной подписи у драйвера нет, его загружает KDU -> антивирусы относят такие загрузчики к HackTool. HVCI и Microsoft Vulnerable Driver Blocklist блокируют загрузку, FACEIT и Vanguard могут с ней конфликтовать: состояние показывает **ПРОВЕРКА СИСТЕМЫ**. Остальные настройки от драйвера не зависят."),
                "SHA-256 файлов — в `SHA256SUMS.txt` на странице релиза, исходников — в `SOURCE-SHA256SUMS.txt`:",
                ("code", VERIFY),
                "Откат: **ВОССТАНОВЛЕНИЕ → ИСХОДНЫЙ** возвращает реестр к состоянию до первой авто-оптимизации и удаляет задачу автозапуска.",
            ]),
        ],
        "gallery": [
            ("showcase_gpu_full", "Fig 1. GPU на двух физических ядрах"),
            ("showcase_nic_full", "Fig 2. NIC ITR и RSS, Intel I226-V"),
            ("showcase_amd_imod_full", "Fig 3. USB IMOD, Ryzen 9 9950X3D"),
        ],
        "gallery_title": "Скриншоты",
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
        "sections": [
            ("why", "Why", [
                "ISR / DPC run above any thread and do not know which cores the game is using. A driver on the core of an important thread (GAME / RENDER / RHI) interrupts it -> the thread waits or moves to another core and loses its L1 / L2 -> frame time grows and stutters appear. Two drivers on one core delay each other's ISR / DPC; on one SMT-SET they share cache and EXECUTION-UNITS.",
                "DEVICE TWEAKER spreads interrupts across physical cores: drivers do not share cores or SMT-SETs, and core 0 stays free when possible.",
                ("items", [
                    ("MSI / Affinity", "MSI Mode, MSI Limit, IRQ Priority, IRQ Policy and Interrupt Affinity for each device"),
                    ("ReservedCpuSets", "cores where Windows does not place ordinary threads"),
                    ("RSS", "base CPU and queues of wired NDIS adapters"),
                    ("USB IMOD", "interrupt moderation of xHCI interrupters"),
                    ("NIC ITR", "interrupt moderation of Intel and Realtek network controllers"),
                ]),
            ]),
            ("auto", "Auto-optimization", [
                "**AUTO-OPTIMIZATION** reads the topology (P-Core / E-Core, SMT / HT, CPPC, CCD / CCX), finds the device roles and builds a plan:",
                ("items", [
                    ("GPU", "-> two different physical cores. On one core the GPU's ISR interrupts its own DPC -> latency spikes. Two SMT siblings of one core do not count as two cores."),
                    ("USB / network / audio", "-> their own cores, by placement rules. The controller with the mouse, keyboard and gamepad is found automatically."),
                    ("CCD", "-> on CPUs with several CCDs a target CCD is chosen."),
                    ("Storage", "-> MSI and high priority, affinity stays with Windows: on some NVMe drives `SpecifiedProcessors` stops Windows from booting."),
                    ("iGPU / Wi-Fi", "-> integrated GPU affinity goes back to Windows, Wi-Fi is left untouched."),
                ]),
                "If the topology cannot be detected reliably, nothing is written. The original state is saved before the first run, and **RESTORE → ORIGINAL** brings it back.",
                ("figure", "cpu-allocation", "GPU: CPU 4 on core 2 and CPU 8 on core 4, SMT siblings CPU 5 and CPU 9 stay free"),
            ]),
            ("interrupts", "Interrupts", [
                "**MSI MODE** — message-signaled interrupts instead of a line. Without MSI, a trace of an xHCI controller shows its DPC interrupted by two ISRs; with MSI it is not interrupted. **MSI LIMIT** — `Unlimited` removes the message limit.",
                "**IRQ PRIORITY** affects the priority of the hardware interrupt. It is not the game's priority and not the DPC's priority. **INTERRUPT AFFINITY** sets the logical processors for the device's IRQ; the `SpecCPU` policy uses the checked mask, the other policies leave the choice to Windows.",
                ("code", INTERRUPT_KEYS["en"]),
                "**IRQ MODE** shows what the device uses now: Line / MSI / MSI-X. **HW** shows what it supports. **READ IRQ MODE** confirms the mode from the `MSI Enable` / `MSI-X Enable` bits in PCI configuration space through DTIMOD.sys.",
            ]),
            ("imod", "USB IMOD", [
                "IMOD is the interrupt moderation interval of an xHCI **interrupter**, not of a single mouse. 250 ns steps: `0xC8` = 50 µs (Windows default), `0xFA0` = 1 ms, `0x0` = moderation off.",
                "A device gets its interrupter when it connects, and it may change after a restart -> the program reads the device → interrupter map from the xHCI contexts in memory every time.",
                ("items", [
                    ("Devices", "values by role. The AUTO profile: mouse `0x0`, audio `0xFA0`, keyboard and gamepad `0xC8`."),
                    ("XHCI", "one value for the whole controller."),
                    ("Interrupters", "values by index."),
                ]),
                "Registers are read back after every write. **CHECK** shows the actual values and the time they were read.",
            ]),
            ("network", "Network", [
                "On NDIS adapters RSS distributes the interrupts, so the base CPU and queues are set instead of a mask. Processors are taken one per physical core, SMT siblings are skipped, and the queue count is limited by the driver's values and the MSI-X vectors.",
                ("code", RSS_KEYS["en"]),
                "**NIC ITR** — the controller's own moderation registers: EITR on Intel I210 / I211, I225 / I226, I350, 82576 / 82580, ITR on I219, IntrMit on Realtek RTL8111 / 8168, RTL8125 / 8126, including compatible Killer adapters. **SAVE** turns off the driver's own moderation (`*InterruptModeration`, on Intel also `ITR`) -> otherwise the driver overwrites the register.",
            ]),
            ("startup", "Startup", [
                "IMOD and NIC ITR registers keep their value until a restart, sleep or controller reset. A saved profile is applied by the Task Scheduler task `\\DEVICE TWEAKER\\IMOD`, running as SYSTEM: at sign-in, after resume from sleep and when a USB / HID device connects.",
                "The task runs the readable `ApplyIMOD.ps1` from `%ProgramData%\\DEVICE TWEAKER\\Startup` and works without the EXE. SHA-256 of the runtime, driver and loader is checked before each run; only administrators and SYSTEM can modify the folder.",
            ]),
            ("more", "More", [
                ("items", [
                    ("ReservedCpuSets", "Windows does not schedule ordinary threads on reserved cores. A driver with a mask on such a core keeps running there, and threads are protected from it."),
                    ("Mouse Throttle", "`RawMouseThrottleDuration`: the Raw Input rate for background windows on Windows 11, 1000 / value Hz."),
                    ("Polling Rate", "mouse and keyboard from USB data, plus a Raw Input event-rate measurement."),
                    ("Power saving", "USB controllers, root USB hubs, wired adapters and USB selective suspend; values are read back through WMI."),
                    ("Backups and logs", "a verified backup before every apply; the log shows what was applied, what was skipped and why."),
                ]),
            ]),
            ("install", "Installation", [
                ("items", [
                    ("SelfContained.exe", ".NET included."),
                    ("DEVICE.TWEAKER.exe", "requires the .NET 8 Desktop Runtime x64."),
                ]),
                "Run as administrator -> **AUTO-OPTIMIZATION**, or manual settings and **APPLY** -> restart. **CHECK** in the IMOD or NIC ITR block shows the actual register values.",
                ("note", "DTIMOD.sys and KDU", "IMOD and NIC ITR registers are not reachable from user mode; they are written by the custom DTIMOD.sys driver, and only to memory that Windows assigned to the selected device. The driver has no public signature and is loaded by KDU -> antivirus software flags such loaders as HackTool. HVCI and the Microsoft Vulnerable Driver Blocklist block the loading, and FACEIT and Vanguard may conflict with it: **SYSTEM CHECK** shows their state. All other settings work without the driver."),
                "SHA-256 hashes of the files are in `SHA256SUMS.txt` on the release page, of the sources in `SOURCE-SHA256SUMS.txt`:",
                ("code", VERIFY),
                "Rollback: **RESTORE → ORIGINAL** returns the registry to its state before the first auto-optimization and removes the startup task.",
            ]),
        ],
        "gallery": [
            ("showcase_gpu_full", "Fig 1. GPU on two physical cores"),
            ("showcase_nic_full", "Fig 2. NIC ITR and RSS, Intel I226-V"),
            ("showcase_amd_imod_full", "Fig 3. USB IMOD, Ryzen 9 9950X3D"),
        ],
        "gallery_title": "Screenshots",
        "gallery_note": "Screenshots show test devices and values.",
        "more": "Every setting in detail",
    },
}


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


def inline(text):
    """**bold** and `code` of the landing text, HTML-escaped."""
    parts = re.split(r"(`[^`]+`)", text)
    out = []
    for part in parts:
        if part.startswith("`"):
            out.append(f"<code>{html.escape(part[1:-1])}</code>")
        else:
            out.append(re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", html.escape(part, quote=False)))
    return "".join(out)


def typography(markup):
    """No line break after one-letter Russian words and before a dash, outside tags and code."""
    parts = re.split(r"(<code>.*?</code>|<pre>.*?</pre>|<[^>]+>)", markup, flags=re.S)
    for i in range(0, len(parts), 2):
        parts[i] = re.sub(r"(?<![\w-])([иИвВсСкКуУаАоОяЯ]) ", lambda m: m.group(1) + NBSP, parts[i]).replace(" — ", NBSP + "— ")
    return "".join(parts)


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
    return ALERT.sub(lambda m: f'<blockquote class="note"><p class="note-title">{m.group(1).title()}</p><p>', body)


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
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;500;700;800&display=swap">
<link rel="stylesheet" href="{root}assets/site.css">
</head>
"""


def topbar(lang, here, slug, ver):
    ui = UI[lang]
    root = rel(here, "")
    links = "".join(
        f'<a href="{rel(here, page_url(s, lang))}"{" aria-current=\"page\"" if s == slug else ""}>{label}</a>'
        for s, label in ui["nav"])
    other = rel(here, page_url(slug, ui["other"]))
    return f"""<body>
<div class="frame">
<header class="topbar">
<a class="brand" href="{rel(here, page_url('', lang))}"><img src="{root}assets/mark.svg" width="24" height="24" alt=""><span>DEVICE TWEAKER</span></a>
<nav class="nav">{links}<a href="{REPO}">GitHub</a><a href="{other}" hreflang="{ui['other']}">{ui['other_label']}</a></nav>
<a class="button primary small" href="{REPO}/releases/latest">{ui['download']}</a>
</header>
"""


def footer(lang, here, ver):
    ui = UI[lang]
    cells = [
        (REPO, "GitHub"),
        (rel(here, page_url("settings", lang)), ui["docs"]),
        (rel(here, page_url("changelog", lang)), ui["docs_nav"][3][1]),
        (TELEGRAM, "Telegram"),
        (f"{REPO}/releases/latest", f"{ui['download']} <span>[{ver}]</span>"),
    ]
    grid = "".join(f'<a href="{href}">{label}</a>' for href, label in cells)
    return f"""<footer class="footer">
<nav class="footer-grid">{grid}</nav>
</footer>
</div>
<p class="colophon">© arsenzaaa · {ui['license']}</p>
</body>
</html>
"""


def picture(here, base, alt, width):
    root = rel(here, "")
    return (f'<picture><source media="(max-width: 600px) and (prefers-color-scheme: dark)" srcset="{root}assets/{base.replace("allocation-", "allocation-mobile-")}-dark.svg">'
            f'<source media="(max-width: 600px)" srcset="{root}assets/{base.replace("allocation-", "allocation-mobile-")}-light.svg">'
            f'<source media="(prefers-color-scheme: dark)" srcset="{root}assets/{base}-dark.svg">'
            f'<img src="{root}assets/{base}-light.svg" alt="{html.escape(alt)}" width="{width}"></picture>')


def block(item, here, suffix):
    if isinstance(item, str):
        return f"<p>{inline(item)}</p>"
    kind = item[0]
    if kind == "items":
        rows = "".join(f'<li><span class="mark">[*]</span><span><b>{inline(term)}</b> {inline(text)}</span></li>' for term, text in item[1])
        return f'<ul class="points">{rows}</ul>'
    if kind == "code":
        return f"<pre><code>{html.escape(item[1])}</code></pre>"
    if kind == "note":
        return f'<div class="note"><p class="note-title">{inline(item[1])}</p><p>{inline(item[2])}</p></div>'
    if kind == "figure":
        return f'<div class="figure">{picture(here, f"{item[1]}-{suffix}", item[2], 560)}</div>'
    raise ValueError(kind)


def landing(lang, ver):
    t, ui = LANDING[lang], UI[lang]
    here = page_url("", lang)
    root = rel(here, "")
    suffix = "en" if lang == "en" else "ru"
    shot = f"{root}assets/screenshots/showcase_intel_14900k_full_{suffix}.png"
    settings = rel(here, page_url("settings", lang))
    sections = "".join(
        f'<section class="section" id="{ident}"><h2>{title}</h2>{"".join(block(b, here, suffix) for b in blocks)}</section>'
        for ident, title, blocks in t["sections"])
    gallery = "".join(
        f'<figure><a href="{root}assets/screenshots/{name}_{suffix}.png"><img src="{root}assets/screenshots/{name}_{suffix}.png" alt="{html.escape(cap)}" loading="lazy"></a><figcaption>{cap.split(". ", 1)[0]}. <b>{cap.split(". ", 1)[1]}</b></figcaption></figure>'
        for name, cap in t["gallery"])
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
</section>
<section class="band"><a class="shot" href="{shot}"><img src="{shot}" alt="{html.escape(t['shot_alt'])}" width="1472" height="1412"></a></section>
{sections}
<section class="section" id="screenshots"><h2>{t['gallery_title']}</h2><div class="gallery">{gallery}</div><p class="faint">{t['gallery_note']}</p>
<p><a class="button" href="{settings}">{t['more']} →</a></p></section>
</main>
"""
    if lang == "ru":
        body = typography(body)
    return (head(lang, t["title"], t["description"], here, "") + topbar(lang, here, "", ver) + body
            + footer(lang, here, ver))


def toc(tokens):
    if not tokens:
        return ""
    items = []
    for token in tokens:
        children = "".join(f'<li><a href="#{c["id"]}">{c["name"]}</a></li>' for c in token["children"])
        items.append(f'<li><a href="#{token["id"]}">{token["name"]}</a>{"<ul>" + children + "</ul>" if children else ""}</li>')
    return "<ul>" + "".join(items) + "</ul>"


def document(slug, lang, source, ver):
    ui = UI[lang]
    here = page_url(slug, lang)
    title, body, tokens = render_document(source, here)
    pages = "".join(
        f'<li><a href="{rel(here, page_url(s, lang))}"{" aria-current=\"page\"" if s == slug else ""}>{label}</a></li>'
        for s, label in ui["docs_nav"])
    description = re.sub(r"<[^>]+>", "", re.search(r"<p>(.*?)</p>", body, re.S).group(1))[:200]
    content = f"""<div class="docs">
<aside class="sidebar">
<nav><ul class="side-pages">{pages}</ul>
<p class="side-label">{ui['on_page']}</p><div class="side-toc">{toc(tokens)}</div></nav>
</aside>
<main class="doc">
<h1>{html.escape(title)}</h1>
{body}
<p class="source"><a href="{REPO}/blob/main/{source}">{ui['source']}: {source}</a></p>
</main>
</div>
"""
    return (head(lang, f"{title} — DEVICE TWEAKER", description, here, slug) + topbar(lang, here, slug, ver)
            + content + footer(lang, here, ver))


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
        target.write_text(document(slug, lang, source, ver), encoding="utf-8")
    print(f"Site built in {out}")


if __name__ == "__main__":
    main()
