<div align="right"><a href="./README.en.md">English</a></div>

<p align="center"><samp>[ WINDOWS 10 / 11 x64 ]</samp></p>

<p align="center"><picture><source media="(prefers-color-scheme: dark)" srcset="./assets/readme/wordmark-dark.svg"><img src="./assets/readme/wordmark-light.svg" alt="DEVICE TWEAKER" width="838"></picture></p>

<p align="center"><b>Прерывания устройств Windows под топологию процессора.</b></p>

<p align="center">MSI Utility V3 + Interrupt Affinity Policy Tool (IAPT) + ReservedCpuSets + RSS + USB IMOD / NIC ITR в одном интерфейсе, с интеллектуальной авто-оптимизацией.</p>

<p align="center"><samp><b>Windows</b> - 10 / 11 x64 &nbsp;|&nbsp; <b>Версия</b> - 0.0.4 &nbsp;|&nbsp; <b>.NET</b> - 8 &nbsp;|&nbsp; <b>Лицензия</b> - GPL-3.0</samp></p>

<p align="center">
<a href="https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest"><picture><source media="(prefers-color-scheme: dark)" srcset="./assets/readme/button-download-ru-dark.svg"><img src="./assets/readme/button-download-ru-light.svg" alt="Скачать 0.0.4" width="187" height="48"></picture></a>
<a href="https://arsenzaaa.github.io/DEVICE-TWEAKER/guide/"><picture><source media="(prefers-color-scheme: dark)" srcset="./assets/readme/button-guide-ru-dark.svg"><img src="./assets/readme/button-guide-ru-light.svg" alt="Гайд" width="96" height="48"></picture></a>
<a href="https://arsenzaaa.github.io/DEVICE-TWEAKER/"><picture><source media="(prefers-color-scheme: dark)" srcset="./assets/readme/button-site-ru-dark.svg"><img src="./assets/readme/button-site-ru-light.svg" alt="Сайт" width="96" height="48"></picture></a>
<a href="./CHANGELOG.md"><picture><source media="(prefers-color-scheme: dark)" srcset="./assets/readme/button-changelog-ru-dark.svg"><img src="./assets/readme/button-changelog-ru-light.svg" alt="История изменений" width="147" height="48"></picture></a>
</p>

<p align="center"><a href="./assets/screenshots/showcase_intel_14900k_full_ru.png"><picture><source srcset="./assets/screenshots/readme/showcase_intel_14900k_full_ru.png 1x, ./assets/screenshots/framed/showcase_intel_14900k_full_ru.png 2x"><img src="./assets/screenshots/readme/showcase_intel_14900k_full_ru.png" alt="DEVICE TWEAKER: USB-контроллер на Intel Core i9-14900K, IMOD и привязка прерываний" width="838"></picture></a></p>

## <samp>[ ЗАЧЕМ ]</samp>

ISR / DPC выполняются выше любого потока и не знают, какие ядра заняты игрой. Драйвер на ядре с важным потоком (GAME / RENDER / RHI) прерывает его -> поток ждёт или уходит на другое ядро и теряет свой L1 / L2 -> растёт время кадра, появляются статтеры. Два драйвера на одном ядре задерживают ISR / DPC друг друга, на одном SMT-SET делят кэш и EXECUTION-UNITS.

DEVICE TWEAKER разносит прерывания по физическим ядрам: драйвера не делят ядра и SMT-SET, ядро 0 по возможности остаётся свободным. **АВТООПТИМИЗАЦИЯ** строит план по топологии (P-Core / E-Core, SMT / HT, CPPC, CCD / CCX) и ролям устройств. Вручную настраивается каждое устройство: MSI, affinity, приоритет, ReservedCpuSets, RSS, USB IMOD, NIC ITR, Mouse Throttle, энергосбережение.

## <samp>[ ГАЙД ]</samp>

Как это устроено в Windows и что делает программа, по каждому типу устройства. Те же главы с поиском — на [сайте](https://arsenzaaa.github.io/DEVICE-TWEAKER/guide/).

<p align="center">
<a href="./docs/guide/01-interrupts.md"><picture><source media="(prefers-color-scheme: dark)" srcset="./assets/readme/card-01-ru-dark.svg"><img src="./assets/readme/card-01-ru-light.svg" alt="[IRQ] 01 Прерывания: ISR, DPC и ядра: Как устроены ISR и DPC, почему они важнее любого потока игры и что происходит, когда драйвера стоят на одном ядре с GAME / RENDER / RHI." width="413"></picture></a>
<a href="./docs/guide/02-msi.md"><picture><source media="(prefers-color-scheme: dark)" srcset="./assets/readme/card-02-ru-dark.svg"><img src="./assets/readme/card-02-ru-light.svg" alt="[MSI] 02 MSI и MSI-X: Линия против сообщения: почему без MSI MODE появляются лишние и более длинные ISR, зачем удалять MessageNumberLimit и как программа читает реальный режим устройства." width="413"></picture></a>
<a href="./docs/guide/03-affinity.md"><picture><source media="(prefers-color-scheme: dark)" srcset="./assets/readme/card-03-ru-dark.svg"><img src="./assets/readme/card-03-ru-light.svg" alt="[CPU] 03 Affinity, политики и приоритет прерываний: Политики Interrupt Affinity, маска AssignmentSetOverride, что на самом деле делает IRQ Priority, Interrupt Steering и альтернативный метод «все ядра, кроме 0»." width="413"></picture></a>
<a href="./docs/guide/04-gpu.md"><picture><source media="(prefers-color-scheme: dark)" srcset="./assets/readme/card-04-ru-dark.svg"><img src="./assets/readme/card-04-ru-light.svg" alt="[GPU] 04 Видеокарта: Почему видеокарте нужны два разных физических ядра, чем драйвер NVIDIA отличается от AMD и что делать, если ISR / DPC всё равно собираются на одном ядре." width="413"></picture></a>
<a href="./docs/guide/05-usb-imod.md"><picture><source media="(prefers-color-scheme: dark)" srcset="./assets/readme/card-05-ru-dark.svg"><img src="./assets/readme/card-05-ru-light.svg" alt="[USB] 05 USB-контроллеры и IMOD: Как устроен контроллер xHCI: interrupter, Event Ring и регистр IMOD. Почему interrupter у мыши меняется при каждом подключении, какой IMOD ставить звуку и как разложить устройства по контроллерам." width="413"></picture></a>
<a href="./docs/guide/06-network.md"><picture><source media="(prefers-color-scheme: dark)" srcset="./assets/readme/card-06-ru-dark.svg"><img src="./assets/readme/card-06-ru-light.svg" alt="[NET] 06 Сеть: RSS, NDIS и NIC ITR: NDIS и NetAdapterCx, почему маска не работает на NDIS и вместо неё нужен RSS, сколько очередей реально использует игра, поток NDIS и MMCSS, NIC ITR и модерация драйвера." width="413"></picture></a>
<a href="./docs/guide/07-cpu-topology.md"><picture><source media="(prefers-color-scheme: dark)" srcset="./assets/readme/card-07-ru-dark.svg"><img src="./assets/readme/card-07-ru-light.svg" alt="[CPU] 07 Топология процессора и ReservedCpuSets: SMT-SET, P- и E-cores, CCD и CPPC-рейтинг: как читать список CPU в программе. ReservedCpuSets, формат маски и две стратегии для процессоров с двумя CCD или E-cores." width="413"></picture></a>
<a href="./docs/guide/08-input.md"><picture><source media="(prefers-color-scheme: dark)" srcset="./assets/readme/card-08-ru-dark.svg"><img src="./assets/readme/card-08-ru-light.svg" alt="[I/O] 08 Ввод: Polling Rate, Raw Input и Mouse Throttle: Путь ввода мыши, клавиатуры и контроллера в Windows 10 и 11, ограничение фоновых окон и RawMouseThrottleDuration, Raw Input Buffer и выделенный поток ввода, что на самом деле измеряет замер Polling Rate." width="413"></picture></a>
<a href="./docs/guide/09-power.md"><picture><source media="(prefers-color-scheme: dark)" srcset="./assets/readme/card-09-ru-dark.svg"><img src="./assets/readme/card-09-ru-light.svg" alt="[PWR] 09 Энергосбережение устройств: USB selective suspend, разрешение Windows отключать устройство и параметры Device Parameters: что из этого реально работает, где хранится и как программа проверяет запись." width="413"></picture></a>
<a href="./docs/guide/10-verification.md"><picture><source media="(prefers-color-scheme: dark)" srcset="./assets/readme/card-10-ru-dark.svg"><img src="./assets/readme/card-10-ru-light.svg" alt="[LOG] 10 Проверка: трейсы и логи: Как проверить результат не по FPS одного прогона, а по факту: трейс xperf / WPR, где смотреть ISR и DPC в WPA и MXA, что писать в отчёт о железе и как читать лог программы." width="413"></picture></a>
<a href="./docs/guide/11-driver.md"><picture><source media="(prefers-color-scheme: dark)" srcset="./assets/readme/card-11-ru-dark.svg"><img src="./assets/readme/card-11-ru-light.svg" alt="[SYS] 11 DTIMOD.sys, KDU и автозапуск: Зачем программе собственный драйвер, куда он может писать, почему его загружает KDU и что об этом думают антивирусы, HVCI и античиты. Как устроен автозапуск IMOD / NIC ITR без EXE." width="413"></picture></a>
</p>

## <samp>[ УСТАНОВКА ]</samp>

- [`DEVICE.TWEAKER.SelfContained.exe`](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest/download/DEVICE.TWEAKER.SelfContained.exe) — .NET встроен.
- [`DEVICE.TWEAKER.exe`](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest/download/DEVICE.TWEAKER.exe) — нужен [.NET 8 Desktop Runtime x64](https://dotnet.microsoft.com/download/dotnet/8.0).

Запустите от имени администратора -> **АВТООПТИМИЗАЦИЯ** или ручная настройка и **ПРИМЕНИТЬ** -> перезагрузка. **ПРОВЕРИТЬ** в блоке IMOD или NIC ITR покажет фактические значения регистров.

> **DTIMOD.sys и KDU.** IMOD и NIC ITR пишет собственный драйвер DTIMOD.sys, и только в память, которую Windows выделила выбранному устройству. Публичной подписи у драйвера нет, его загружает [KDU](./src/IMOD/KDU/README.md) -> антивирусы относят такие загрузчики к HackTool, HVCI и Microsoft Vulnerable Driver Blocklist блокируют загрузку, FACEIT и Vanguard могут с ней конфликтовать. Остальные настройки от драйвера не зависят. Подробно — [глава 11](./docs/guide/11-driver.md).

SHA-256 файлов — в `SHA256SUMS.txt` на [странице релиза](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest), исходников — в `SOURCE-SHA256SUMS.txt`:

```powershell
(Get-FileHash .\DEVICE.TWEAKER.SelfContained.exe -Algorithm SHA256).Hash
Select-String -Path .\SHA256SUMS.txt -Pattern 'DEVICE.TWEAKER.SelfContained.exe'
```

## <samp>[ ОТКАТ ]</samp>

Перед каждым применением программа создаёт проверенный бекап в папке `Backups` рядом с EXE. **ВОССТАНОВЛЕНИЕ → ИСХОДНЫЙ** возвращает реестр к состоянию до первой авто-оптимизации и удаляет задачу автозапуска. В том же меню — любой из бекапов и **СБРОС WINDOWS**.

## <samp>[ ОБРАТНАЯ СВЯЗЬ ]</samp>

Ошибки — в [issues](https://github.com/arsenzaaa/DEVICE-TWEAKER/issues/new?template=bug_report.yml) с логом из папки `logs`, версией Windows и конфигурацией ПК. Чтобы помочь с авто-оптимизацией, отправьте [отчёт о железе](https://github.com/arsenzaaa/DEVICE-TWEAKER/issues/new?template=hardware_report.yml) с тремя логами: до применения, сразу после и после перезагрузки. Обновления — в [Telegram](https://t.me/arsenzaa).

<p align="center"><sub><a href="https://arsenzaaa.github.io/DEVICE-TWEAKER/">Сайт</a> · <a href="./docs/SETTINGS.md">Описание настроек</a> · <a href="./docs/DEVELOPMENT.md">Разработка и сборка</a> · <a href="./docs/THIRD_PARTY_NOTICES.md">Сторонние компоненты</a> · <a href="./.github/SECURITY.md">Безопасность</a> · <a href="./LICENSE">GPL-3.0</a></sub></p>
