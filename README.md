<p align="center">
  <a href="https://arsenzaaa.github.io/DEVICE-TWEAKER/">
    <picture>
      <source media="(prefers-color-scheme: dark)" srcset="./assets/readme/wordmark-dark.svg">
      <img src="./assets/readme/wordmark-light.svg" alt="DEVICE TWEAKER" width="560">
    </picture>
  </a>
</p>

<p align="center">Прерывания устройств Windows 10 / 11 x64 под топологию процессора.</p>

<p align="center">
  <a href="https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest"><b>Скачать 0.0.4</b></a> ·
  <a href="https://arsenzaaa.github.io/DEVICE-TWEAKER/guide/">Гайд</a> ·
  <a href="https://arsenzaaa.github.io/DEVICE-TWEAKER/">Сайт</a> ·
  <a href="./CHANGELOG.md">Изменения</a> ·
  <a href="https://t.me/arsenzaa">Telegram</a> ·
  <a href="./README.en.md">English</a>
</p>

<p align="center"><a href="./assets/screenshots/showcase_intel_14900k_full_ru.png"><picture><source srcset="./assets/screenshots/readme/showcase_intel_14900k_full_ru.png 1x, ./assets/screenshots/framed/showcase_intel_14900k_full_ru.png 2x"><img src="./assets/screenshots/readme/showcase_intel_14900k_full_ru.png" alt="DEVICE TWEAKER: USB-контроллер на Intel Core i9-14900K, IMOD и привязка прерываний" width="838"></picture></a></p>

---

## <samp>[ О ПРОГРАММЕ ]</samp>

ISR / DPC выполняются выше любого потока и не знают, какие ядра заняты игрой. Драйвер на ядре с важным потоком (GAME / RENDER / RHI) прерывает его -> поток ждёт или уходит на другое ядро и теряет свой L1 / L2 -> растёт время кадра. Два драйвера на одном ядре задерживают ISR / DPC друг друга, на одном SMT-SET делят кэш и EXECUTION-UNITS.

DEVICE TWEAKER разносит прерывания по физическим ядрам: MSI Utility V3 + Interrupt Affinity Policy Tool (IAPT) + ReservedCpuSets + RSS + USB IMOD / NIC ITR в одном интерфейсе, с интеллектуальной авто-оптимизацией.

## <samp>[ ВОЗМОЖНОСТИ ]</samp>

- **Авто-оптимизация.** План по топологии (P / E-cores, SMT / HT, CPPC, CCD / CCX) и ролям устройств: GPU на двух разных физических ядрах, драйвера не делят ядра и SMT-SET, ядро 0 по возможности свободно. [Глава 01](./docs/guide/01-interrupts.md)
- **MSI и MSI-X.** MSI Mode, MSI Limit, текущий IRQ Mode и возможности устройства раздельно; **СЧИТАТЬ IRQ** читает реальный режим из конфигурации PCI. [Глава 02](./docs/guide/02-msi.md)
- **Interrupt Affinity и приоритет.** Политика, маска и IRQ Priority для каждого устройства, список CPU с физическими ядрами, CCD и CPPC-рейтингом. [Глава 03](./docs/guide/03-affinity.md)
- **USB IMOD.** Модерация каждого interrupter'а xHCI по ролям устройств; карта «устройство -> interrupter» читается из памяти контроллера при каждом применении. [Глава 05](./docs/guide/05-usb-imod.md)
- **RSS и NIC ITR.** Базовый CPU и очереди адаптеров NDIS по одному на физическое ядро, регистры модерации Intel и Realtek. [Глава 06](./docs/guide/06-network.md)
- **ReservedCpuSets, Mouse Throttle, Polling Rate, энергосбережение.** [Главы 07–09](./docs/guide/README.md)
- **Автозапуск без EXE.** Профиль IMOD / NIC ITR применяет задача Планировщика при входе, после сна и при подключении USB; открытый скрипт, файлы сверяются по SHA-256. [Глава 11](./docs/guide/11-driver.md)
- **Бекапы и логи.** Проверенный бекап перед каждым применением, **ВОССТАНОВЛЕНИЕ -> ИСХОДНЫЙ** возвращает состояние до первой авто-оптимизации; в логе — что применилось, что пропущено и почему.

## <samp>[ УСТАНОВКА ]</samp>

Скачайте один из файлов [последнего релиза](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest):

- `DEVICE.TWEAKER.SelfContained.exe` — .NET встроен;
- `DEVICE.TWEAKER.exe` — нужен [.NET 8 Desktop Runtime x64](https://dotnet.microsoft.com/download/dotnet/8.0).

Запустите от имени администратора -> **АВТООПТИМИЗАЦИЯ** или ручная настройка и **ПРИМЕНИТЬ** -> перезагрузка. **ПРОВЕРИТЬ** в блоке IMOD или NIC ITR покажет фактические значения регистров.

> [!IMPORTANT]
> IMOD и NIC ITR пишет собственный драйвер DTIMOD.sys, и только в память, которую Windows выделила выбранному устройству. Публичной подписи у драйвера нет, его загружает [KDU](./src/IMOD/KDU/README.md), поэтому антивирусы относят программу к HackTool, а HVCI и Microsoft Vulnerable Driver Blocklist блокируют загрузку. Остальные настройки от драйвера не зависят. Подробно — [глава 11](./docs/guide/11-driver.md).

SHA-256 файлов — в `SHA256SUMS.txt` на странице релиза:

```powershell
(Get-FileHash .\DEVICE.TWEAKER.SelfContained.exe -Algorithm SHA256).Hash
```

## <samp>[ ГАЙД ]</samp>

Как это устроено в Windows и что делает программа. Те же главы с поиском — на [сайте](https://arsenzaaa.github.io/DEVICE-TWEAKER/guide/).

1. [Прерывания: ISR, DPC и ядра](./docs/guide/01-interrupts.md)
2. [MSI и MSI-X](./docs/guide/02-msi.md)
3. [Affinity, политики и приоритет прерываний](./docs/guide/03-affinity.md)
4. [Видеокарта](./docs/guide/04-gpu.md)
5. [USB-контроллеры и IMOD](./docs/guide/05-usb-imod.md)
6. [Сеть: RSS, NDIS и NIC ITR](./docs/guide/06-network.md)
7. [Топология процессора и ReservedCpuSets](./docs/guide/07-cpu-topology.md)
8. [Ввод: Polling Rate, Raw Input и Mouse Throttle](./docs/guide/08-input.md)
9. [Энергосбережение устройств](./docs/guide/09-power.md)
10. [Проверка: трейсы и логи](./docs/guide/10-verification.md)
11. [DTIMOD.sys, KDU и автозапуск](./docs/guide/11-driver.md)

Справочник по каждому элементу интерфейса — [описание настроек](./docs/SETTINGS.md).

## <samp>[ ОБРАТНАЯ СВЯЗЬ ]</samp>

Ошибки — в [issues](https://github.com/arsenzaaa/DEVICE-TWEAKER/issues/new?template=bug_report.yml) с логом из папки `logs` рядом с EXE. Результаты на своём железе — в [отчёт о железе](https://github.com/arsenzaaa/DEVICE-TWEAKER/issues/new?template=hardware_report.yml): три лога, до применения, сразу после и после перезагрузки. Обновления — в [Telegram](https://t.me/arsenzaa).

Сборка из исходников — [docs/DEVELOPMENT.md](./docs/DEVELOPMENT.md). Лицензия — [GPL-3.0](./LICENSE).
