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

**DEVICE TWEAKER** — низкоуровневая утилита для Windows 10 и 11, объединяющая настройку режимов прерываний (MSI / MSI-X), привязку к ядрам с учетом аппаратной топологии процессора (CPU Affinity), системную изоляцию (ReservedCpuSets), оптимизацию сетевого стека (RSS, NIC ITR) и прямое программирование регистров USB xHCI IMOD (с шагом 250 нс через собственный драйвер) в едином интерфейсе с алгоритмом автоматического подбора.

---

## Зачем нужна оптимизация прерываний

- **Вытеснение игрового потока (ISR / DPC Latency):**  
  Аппаратные прерывания (ISR) и очереди отложенных вызовов процедур (DPC) выполняются на уровне ядра с наивысшим приоритетом (`DISPATCH_LEVEL`). Если прерывания сетевого адаптера или высокочастотной мыши (1000–8000 Гц) обрабатываются на том же ядре, где выполняется главный рендер-поток игры, он принудительно вытесняется. Это приводит к микрофризам, тайминг-вариативности и нестабильному времени кадра (Frame Time jitter).
- **Разгрузка CPU 0:**  
  По умолчанию Windows направляет на нулевое логическое ядро системный таймер, дисковый ввод-вывод и прерывания большинства системных шин. Увод игровой периферии и видеокарты с CPU 0 устраняет заторы в системных очередях DPC.
- **Аппаратная модерация USB (xHCI IMOD):**  
  В xHCI-контроллерах по умолчанию включена модерация прерываний (~50 мкс), которая накапливает пакеты данных в буфере и отправляет их пачками. Установка значения `0 мкс` через низкоуровневый драйвер ядра `DTIMOD.sys` полностью отключает искусственную задержку — пакеты опроса мыши поступают процессору мгновенно.
- **Топологическая адресация (CCD & Hybrid Cores):**  
  Исключение медленных E-ядер на процессорах Intel Hybrid и изоляция прерываний критических устройств на чиплете с 3D V-Cache (CCD0) на процессорах AMD Ryzen исключают паразитные задержки межъядерной синхронизации через Infinity Fabric.

---

## Возможности

- **Группировка и мониторинг устройств:** Видеокарты, USB-контроллеры, сетевые адаптеры, накопители и аудиоустройства разделены по отдельным блокам с отображением их текущих параметров, очередей и статуса.
- **MSI / MSI-X и IRQ Priority:** Включение режима Message Signaled Interrupts, снятие лимита векторов (`MessageNumberLimit`) и выставление наивысшего приоритета обслуживания (`IRQ Priority = High`).
- **Распределение CPU Affinity:** Ручная и автоматическая привязка устройств к физическим ядрам и потокам с наглядной визуализацией:
  - Разделение P-Core и E-Core (Intel 12–14th Gen).
  - Пары Hyper-Threading / SMT.
  - Чиплеты AMD (CCD0 с 3D V-Cache / CCD1).
  - Кластеры CCX и аппаратный рейтинг ядер (CPPC).
- **Прямое управление USB xHCI IMOD через `DTIMOD.sys`:** Чтение и прямая запись физических регистров модерации прерываний (MMIO) контроллеров USB с дискретностью 250 нс. Возможность выставить `0 мкс` для мыши или индивидуальные интервалы для каждого интерраптера.
- **Оптимизация сетевого стека (RSS и NIC ITR):** Полное отключение модерации прерываний сетевой карты (`ITR = 0 / Off` для адаптеров Intel и Realtek), настройка очередей Receive Side Scaling (RSS) и отключение энергосбережения Energy Efficient Ethernet (EEE).
- **Системная изоляция ядер (`ReservedCpuSets`):** Настройка системной маски изоляции на уровне ядра Windows. На выделенные ядра перестают назначаться фоновые системные потоки и службы ОС.
- **1-Click Авто-оптимизация под архитектуру CPU:**
  - Полная разгрузка `CPU 0` от периферии и видеокарты.
  - Закрепление GPU за выделенной парой смежных физических P-ядер.
  - Разведение высокочастотной мыши и сетевой карты по разным ядрам во избежание конкуренции DPC.
  - Фиксация критических устройств на скоростном чиплете CCD0 для процессоров AMD Ryzen X3D.
- **Управление энергосбережением шины:** Быстрое отключение Power Saving и Selective Suspend для USB-контроллеров и проводных адаптеров Ethernet.
- **Безопасность и заводской снимок:** Автоматический бекап перед любыми изменениями, защита заводского снимка Windows (`ORIGINAL STATE`) от перезаписи и возможность полного отката.
- **Локализация:** Поддержка русского и английского языков с переключением на лету без перезапуска приложения.

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

## Архитектура и безопасность

- **Опрос без нагрузки на ядро:** Кнопка **ОБНОВИТЬ (REFRESH)** запрашивает параметры оборудования исключительно через штатные SetupAPI и реестр Windows, не загружая драйвер в ядро.
- **Изоляция драйвера ядра `DTIMOD.sys`:** Драйвер задействуется только при явном нажатии **ПРОВЕРИТЬ (CHECK)** или сохранении значений IMOD. Загруженный через KDU драйвер безопасно удерживается в памяти до перезагрузки Windows во избежание сбоев ядра при горячей выгрузке.
- **Вступление изменений в силу:** Для окончательного применения настроек прерываний и очередей реестра требуется перезагрузка ПК.
- **Журналирование:** Полные диагностические логи формируются при каждом запуске в каталоге `logs` рядом с исполняемым файлом.

---

## Требования к системе

- **ОС:** Windows 10 или Windows 11 (64-bit, все редакции).
- **Варианты сборки:**
  - `DEVICE.TWEAKER.exe` (~4 МБ) — стандартная версия, требует установленный [.NET 8 Desktop Runtime](https://dotnet.microsoft.com/download/dotnet/8.0).
  - `DEVICE.TWEAKER.NET.FRAMEWORK.exe` (~150 МБ) — автономная версия (Self-Contained со встроенным рантаймом .NET 8, работает на любой чистой системе без доустановки библиотек).

---

## Запуск и использование

1. Загрузите актуальный релиз со страницы [GitHub Releases](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases).
2. Запустите приложение от имени **Администратора**.
3. Нажмите **АВТООПТИМИЗАЦИЯ** для автоматического подбора конфигурации под ваш процессор или задайте распределение вручную.
4. Нажмите **ПРИМЕНИТЬ**.
5. Перезагрузите компьютер.

Полная история версий доступна в [CHANGELOG.md](./DEVICE%20TWEAKER/CHANGELOG.md).

---

## Сборка из исходного кода

Для компиляции проекта требуется [.NET 8.0 SDK](https://dotnet.microsoft.com/download/dotnet/8.0).

Сборка обеих версий приложения с генерацией контрольных сумм SHA-256 выполняется одной командой из корня репозитория:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File ".\DEVICE TWEAKER\build.ps1" -Flavor both -Configuration Release -SkipImodDriverBuild
```

Готовые бинарники и хеш-суммы формируются в каталоге `DEVICE TWEAKER\bin\ReleasePackages\<версия>\`.

---

## Автоматизированное тестирование

Репозиторий включает набор из **21 модульного теста (xUnit)**, проверяющих парсинг топологии процессоров, битовые маски affinity, форматирование `ReservedCpuSets` и корректность таблиц IMOD:

```powershell
dotnet test ".\DEVICE TWEAKER\tests\DeviceTweaker.Tests\DeviceTweaker.Tests.csproj" -c Release -p:BuildImodDriver=false
```

---

## Структура проекта

```
DEVICE-TWEAKER/
├── .gitignore
├── LICENSE
├── README.md                      # Основное описание (Русский)
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
    ├── docs/                      # Дополнительная документация и история релизов
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
