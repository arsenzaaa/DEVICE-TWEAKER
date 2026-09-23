<div align="center">

<img src="./DEVICE%20TWEAKER/assets/banner.svg" alt="DEVICE TWEAKER Banner" width="900">

<br><br>

[![Release](https://img.shields.io/github/v/release/arsenzaaa/DEVICE-TWEAKER?include_prereleases&style=for-the-badge&color=007acc&label=Release)](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases)
[![Unit Tests](https://img.shields.io/badge/Unit%20Tests-21%20Passed-2ea44f?style=for-the-badge&logo=dotnet)](https://github.com/arsenzaaa/DEVICE-TWEAKER)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%20%7C%2011%20x64-0078d4?style=for-the-badge&logo=windows)](https://github.com/arsenzaaa/DEVICE-TWEAKER)
[![License](https://img.shields.io/badge/License-GPLv3-2ea44f?style=for-the-badge)](LICENSE)
[![Telegram](https://img.shields.io/badge/Telegram-@arsenzaa-2CA5E0?style=for-the-badge&logo=telegram&logoColor=white)](https://t.me/arsenzaa)

<p align="center">
  <b>Русский</b> • <a href="./README.en.md">English</a>
</p>

---

<br>

<img src="./DEVICE%20TWEAKER/assets/screenshots/main_interface_ru.png" alt="DEVICE TWEAKER Главный интерфейс" width="900">

</div>

<br>

**DEVICE TWEAKER** — низкоуровневая утилита для Windows 10 и 11, предназначенная для тонкого управления прерываниями устройств, конвейером обработки ввода и системными задержками ядра Windows NT.

Программа объединяет перевод устройств в режим прерываний по сообщению (**MSI / MSI-X**), топологическую привязку прерываний к ядрам процессора (**CPU Affinity**), системную изоляцию через **`ReservedCpuSets`**, разделение очередей сетевого стека (**RSS**, **NIC ITR**) и прямое программирование физических MMIO-регистров модерации контроллеров USB (**xHCI IMOD** с дискретностью 250 нс через собственный драйвер ядра `DTIMOD.sys`).

---

## Архитектура задержек и прерываний в Windows

В архитектуре Windows NT обработка аппаратных событий построена на строгой приоритетной иерархии уровней прерываний (**IRQL**):

```
[ Аппаратное устройство: USB / GPU / NIC ]
                    │
                    ▼  (Аппаратный IRQ)
      [ Обработчик прерывания ISR ]  ──►  Выполняется на уровне DIRQL (блокирует ядро)
                    │
                    ▼  (Запрос отложенной процедуры)
      [ Очередь DPC (WDF01000.sys) ]  ──►  Выполняется на уровне DISPATCH_LEVEL
                    │
                    ▼  (Сигнал потоку подсистемы ввода)
      [ Поток ввода подсистемы ОС ]
         ├─ Windows 10: CSRSS (Win32k Raw Input Thread)
         └─ Windows 11: DWM   (Kernel Sensor Thread / Master Input Thread)
                    │
                    ▼  (Передача сырого ввода через буфер или IPI)
      [ Игровой процесс / Приложение ]
         ├─ GameThread / App Input Thread
         └─ RenderThread / RHIThread
```

### 1. Вытеснение основного потока рендера (Execution Preemption)
Обработчики прерываний (**ISR**) и очереди отложенных вызовов процедур (**DPC**) выполняются на уровнях `DIRQL` и `DISPATCH_LEVEL`, которые аппаратно выше любого потока пользовательского режима (`PASSIVE_LEVEL`). 

Когда прерывания высокочастотной мыши (1000–8000 Гц), сетевой карты или видеокарты назначаются на то же физическое ядро, где исполняется критический поток игры (`RenderThread` или `GameThread`), ядро принудительно прерывает выполнение игры для обработки очередей драйверов. Это приводит к разрыву конвейера кадров, микростаттерам и тяжелой вариативности времени кадра (**Frame Time jitter**).

### 2. Заторы очередей на CPU 0
По умолчанию планировщик Windows направляет на нулевое логическое ядро прерывания системного таймера, дисковых контроллеров NVMe/SATA, системных шин и большинства Plug-and-Play устройств. Если оставить видеокарту и периферию на `CPU 0`, их обработчики ISR/DPC встают в общую системную очередь, порождая задержки диспетчеризации.

### 3. Аппаратная модерация USB (xHCI IMOD)
Спецификация Intel xHCI предусматривает в контроллерах USB аппаратную модерацию прерываний (**Interrupt Moderation**, регистры `IMODI` / `IMODC`). По умолчанию в Windows для каждого интерраптера включен интервал ~50 мкс (200 единиц с шагом 250 нс). Контроллер намеренно задерживает отправку прерывания процессору, накапливая пакеты в аппаратном буфере.

Для мышей с частотой опроса 1000–8000 Гц это создает постоянный джиттер передачи отсчетов. **DEVICE TWEAKER** через драйвер `DTIMOD.sys` выполняет прямое чтение и запись физических MMIO-регистров контроллера, сбрасывая интервал `IMODI` в `0 мкс` (полное отключение задержки). При этом для контроллеров с аудиоустройствами интервал IMOD настраивается индивидуально, чтобы снизить DPC-нагрузку без деградации звукового потока.

### 4. Топологические штрафы (CCD0 3D V-Cache vs E-Cores)
- **AMD Ryzen Dual-CCD (X3D):** Чиплет CCD0 оснащен быстрым массивом 3D V-Cache с минимальной латентностью доступа, тогда как чиплет CCD1 работает со стандартным кэшем. Обмен данными между чиплетами через шину Infinity Fabric накладывает штраф в 60–80 нс и приводит к сбросу кэш-линий. Направление прерываний критических устройств на CCD1 при запущенной на CCD0 игре вынуждает систему непрерывно синхронизировать память через межчиплетную шину.
- **Intel Hybrid Architecture (12–14th Gen):** Энергоэффективные E-ядра не имеют Hyper-Threading, обладают урезанным конвейером и более высокой латентностью пробуждения из C-стейтов. Обработка прерываний реального времени на E-ядрах гарантирует просадку скорости диспетчеризации.

### 5. Сетевой стек: RSS, NIC ITR и NetAdapterCx
Сетевые адаптеры при интенсивном сетевом трафике генерируют непрерывный поток прерываний. Без явного разделения очередей **Receive Side Scaling (RSS)** сетевые прерывания обрабатываются в общем стеке `WDF01000.sys`, вытесняя ввод мыши. Отключение аппаратной модерации сетевой карты (`NIC ITR = 0 / Off` для Intel и Realtek) и деактивация **Energy Efficient Ethernet (EEE)** исключают засыпание PHY-контроллера и устраняют буферизацию входящих пакетов.

### 6. Системная маска ReservedCpuSets
Параметр реестра `ReservedCpuSets` в `HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\kernel` задает битовую маску ядер, на которые планировщик Windows NT перестает назначать непривязанные системные потоки и фоновые службы ОС. Закрепляя прерывания критических устройств за зарезервированными ядрами, мы изолируем их обработку от стороннего системного шума.

---

## Возможности

- **Раздельный мониторинг по шинам и классам:** Видеокарты, USB-контроллеры (с определением физической привязки CHIP 0 / CHIP 1 по PCI ID), сетевые адаптеры, накопители и аудиоустройства разделены по независимым блокам.
- **MSI / MSI-X и IRQ Priority:** Перевод устройств из устаревшего Line-Based режима в режим прерываний по сообщению (MSI), снятие ограничения количества сообщений (`MessageNumberLimit`) и повышение приоритета обработки (`IRQ Priority = High`).
- **Аппаратная карта распределения ядер (Affinity Engine):**
  - Разделение физических ядер и Hyper-Threading / SMT пар.
  - Учет чиплетов AMD (CCD0 с 3D V-Cache / CCD1) и кластеров CCX.
  - Учет ядер P-Core / E-Core на гетерогенных процессорах Intel.
  - Чтение аппаратных рейтингов ядер через прямое чтение событий ядра Windows ETW (`EventLogReader`, Event ID 53) за <60 мс без задержек WMI.
- **Низкоуровневый драйвер ядра `DTIMOD.sys`:**
  - Прямое чтение и запись MMIO-регистров xHCI IMOD с аппаратным шагом 250 нс.
  - Безопасная валидация границ PCI BAR и xHCI operational registers.
  - Настройка `0 мкс` (Zero Moderation) для контроллеров мыши.
- **Оптимизация сетевого стека:**
  - Настройка очередей Receive Side Scaling (`*NumRssQueues`, `*RssBaseProcNumber`).
  - Прямое отключение аппаратной модерации `ITR = 0 / Off` для контроллеров Intel (I210, I211, I225-V/LM/K/I, I226, I350) и Realtek.
  - Отключение энергосбережения Energy Efficient Ethernet (EEE).
- **1-Click Авто-оптимизация под обнаруженную топологию:**
  - Полный увод видеокарты и игровой периферии с перегруженного `CPU 0`.
  - Привязка видеокарты к паре смежных физических P-ядер.
  - Разведение прерываний мыши и сетевой карты по разным ядрам во избежание пересечения в очередях DPC.
  - Фиксация задержко-чувствительных устройств на чиплете CCD0 для процессоров AMD Ryzen X3D.
- **Управление энергосбережением шины:** Быстрое отключение Selective Suspend и Device Power Saving для USB-контроллеров и проводных адаптеров NDIS.
- **Защита заводского состояния и откат:**
  - Автоматическое создание бекапа перед применением любых настроек.
  - Неперезаписываемый защищенный снимок исходной системы (**ORIGINAL STATE**).
  - Проверка валидности форматов и автоматический rollback при сбоях во время записи.
- **Мгновенная смена языка:** Полная локализация интерфейса (RU / EN) с переключением на лету без перезапуска.

---

<details>
<summary><b>📸 Визуальный обзор и примеры топологий (нажмите, чтобы развернуть)</b></summary>
<br>

### 1. Распределение топологии: AMD Ryzen Dual-CCD (CCD0 с 3D V-Cache vs CCD1)
<p align="center">
  <img src="./DEVICE%20TWEAKER/assets/screenshots/showcase_amd_9950x3d_dual_ccd_ru.png" alt="AMD Dual-CCD Topology" width="850">
</p>

### 2. Распределение топологии: Intel Hybrid Architecture (P-Core / E-Core)
<p align="center">
  <img src="./DEVICE%20TWEAKER/assets/screenshots/showcase_intel_14900k_hybrid_ru.png" alt="Intel Hybrid Topology" width="850">
</p>

### 3. Таблица прямого управления регистрами USB xHCI IMOD (шаг 250 нс)
<p align="center">
  <img src="./DEVICE%20TWEAKER/assets/screenshots/showcase_amd_imod_table_en.png" alt="xHCI IMOD Table" width="850">
</p>

### 4. Безопасность: Автобекапы и восстановление заводского состояния (ORIGINAL STATE)
<p align="center">
  <img src="./DEVICE%20TWEAKER/assets/screenshots/restore_dialog_ru.png" alt="Бекапы и восстановление" width="850">
</p>

### 5. Панель симуляции топологий оборудования (Test Admin Panel)
<p align="center">
  <img src="./DEVICE%20TWEAKER/assets/screenshots/test_admin_panel.png" alt="Test Admin Panel" width="850">
</p>

</details>

---

## Архитектура безопасности и работа с ядром

- **Безопасный опрос оборудования:** Кнопка **ОБНОВИТЬ (REFRESH)** опрашивает устройства через SetupAPI и системный реестр без загрузки драйвера ядра.
- **Изоляция драйвера ядра `DTIMOD.sys`:** Драйвер поднимается через библиотеку KDU только при явном нажатии **ПРОВЕРИТЬ (CHECK)** или сохранении значений IMOD. Загруженный драйвер безопасно удерживается в адресном пространстве ядра до перезагрузки Windows во избежание сбоев ядра при динамической выгрузке.
- **Применение параметров:** Параметры прерываний PCI/MSI и сетевых очерей NDIS читаются операционной системой во время инициализации устройств, поэтому после применения настроек требуется перезагрузка ПК.
- **Журналирование:** Подробные диагностические логи формируются при каждом запуске в каталоге `logs` рядом с исполняемым файлом.

---

## Варианты сборки и системные требования

- **Операционная система:** Windows 10 или Windows 11 (64-bit, редакции Home, Pro, Enterprise, LTSC).
- **Варианты поставки:**
  - `DEVICE.TWEAKER.exe` (~4 МБ) — стандартная версия, требует установленный [.NET 8 Desktop Runtime](https://dotnet.microsoft.com/download/dotnet/8.0).
  - `DEVICE.TWEAKER.NET.FRAMEWORK.exe` (~150 МБ) — автономная версия (Self-Contained со встроенным рантаймом .NET 8, работает на любой чистой системе без доустановки библиотек).

---

## Сборка из исходного кода

Для сборки необходим [.NET 8.0 SDK](https://dotnet.microsoft.com/download/dotnet/8.0).

Сборка обеих версий приложения с генерацией контрольных сумм SHA-256 выполняется одной командой из корня репозитория:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File ".\DEVICE TWEAKER\build.ps1" -Flavor both -Configuration Release -SkipImodDriverBuild
```

Скомпилированные бинарники и хеш-файлы формируются в каталоге `DEVICE TWEAKER\bin\ReleasePackages\<версия>\`.

---

## Автоматизированное тестирование

Проект содержит набор из **21 модульного теста (xUnit)**, валидирующих парсинг топологии процессора, разбор событий CPPC из журнала ETW, битовые маски affinity, формат маски `ReservedCpuSets` и физические границы IMOD:

```powershell
dotnet test ".\DEVICE TWEAKER\tests\DeviceTweaker.Tests\DeviceTweaker.Tests.csproj" -c Release -p:BuildImodDriver=false
```

---

## Структура проекта

```
DEVICE-TWEAKER/
├── .gitignore
├── LICENSE
├── README.md                      # Документация (Русский)
├── README.en.md                   # Документация (English)
└── DEVICE TWEAKER/                # Исходный код и ресурсы приложения
    ├── Core/                      # Службы, резервное копирование и логика отката
    ├── Devices/                   # Опрос шин PCI, USB-контроллеров и сетевых карт NDIS
    ├── GUI/                       # Пользовательский интерфейс WinForms и диалоговые окна
    ├── IMOD/                      # Драйвер прямого доступа DTIMOD.sys и библиотека KDU
    ├── Interop/                   # Низкоуровневые P/Invoke обертки к Win32 API
    ├── Localization/              # Локализация интерфейса (RU / EN)
    ├── Models/                    # Модели данных устройств, топологий и отчетов
    ├── Tweaks/                    # Алгоритмы авто-оптимизации под топологию CPU
    ├── Scripts/                   # Скрипты автоматизации и валидации
    ├── tests/                     # Модульные тесты DeviceTweaker.Tests (xUnit)
    ├── assets/                    # Графические ресурсы, иконки и скриншоты
    ├── docs/                      # Техническая документация и архив релизов
    ├── DeviceTweakerCS.csproj     # Файл проекта .NET 8
    └── build.ps1                  # Скрипт автоматизированной сборки
```

---

## Автор и контакты

- **Автор:** [@arsenzaa](https://t.me/arsenzaa)
- **GitHub:** [arsenzaaa](https://github.com/arsenzaaa)
- **Репозиторий:** [DEVICE-TWEAKER](https://github.com/arsenzaaa/DEVICE-TWEAKER)

## Лицензия

Проект распространяется под лицензией [GNU General Public License v3.0](LICENSE).  
Компоненты KDU в каталоге `IMOD/KDU` распространяются под лицензией MIT.
