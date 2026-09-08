# DEVICE TWEAKER

[![.NET](https://img.shields.io/badge/.NET-8.0--windows-512BD4?logo=dotnet&logoColor=white)](https://dotnet.microsoft.com/download/dotnet/8.0)
[![Platform](https://img.shields.io/badge/platform-Windows%2010%2F11%20x64-0078D6?logo=windows&logoColor=white)](#требования)
[![License](https://img.shields.io/badge/license-GPLv3-blue.svg)](LICENSE)
[![Telegram](https://img.shields.io/badge/Telegram-arsenzaa-2CA5E0?logo=telegram&logoColor=white)](https://t.me/arsenzaa)
[![Latest release](https://img.shields.io/badge/release-v0.0.4--alpha.2-orange)](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.4-alpha.2)

![DEVICE TWEAKER](./DEVICE%20TWEAKER/assets/DEVICE%20TWEAKER-wordmark.svg)

**DEVICE TWEAKER** — утилита для Windows 10/11, объединяющая MSI Utility, Interrupt Affinity Policy Tool, ReservedCpuSets, RSS, NIC ITR, USB IMOD и авто-оптимизацию в одном интерфейсе.

Текущая версия **v0.0.4-alpha.2** (pre-release): [Releases](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.4-alpha.2).

Исходный код находится в папке [`DEVICE TWEAKER`](./DEVICE%20TWEAKER).

## Возможности

- Блоки устройств: USB-контроллеры, GPU, накопители, аудио и сетевые адаптеры.
- MSI Mode / MSI Limit, IRQ Priority и политика распределения прерываний (Interrupt Affinity Policy).
- Ручное назначение affinity по логическим процессорам.
- `ReservedCpuSets` с отображением текущего значения.
- RSS и ITR на поддерживаемых сетевых адаптерах.
- IMOD через `DTIMOD.sys` на поддерживаемых USB-контроллерах.
- Power Saving для USB-контроллеров и проводных NIC.
- Классификация USB CHIP 0 / CHIP 1 и отображение Selective Suspend.
- `AUTO-OPTIMIZATION` с учетом P/E-Core, SMT/Hyper-Threading, CPPC, CCD и CCX.
- Резервные копии, восстановление и полный сброс (`RESET`).
- Интерфейс EN/RU без перезапуска программы.

## Требования

- Windows 10 или Windows 11 x64.
- Права администратора.
- Обычная сборка (`DEVICE.TWEAKER.exe`) — нужен .NET 8 Desktop Runtime или новее.
- `DEVICE.TWEAKER.NET.FRAMEWORK.exe` — self-contained, отдельная установка .NET 8 не требуется.

## Важно

Программа меняет параметры прерываний и значения в `HKLM`; при работе с IMOD загружает драйвер `DTIMOD.sys`.

- Перед изменениями создавайте резервную копию.
- Не задавайте случайные значения IMOD/ITR без понимания их смысла.
- После серьезных изменений может потребоваться перезагрузка.
- Логи пишутся автоматически в папку `logs` рядом с EXE.
- `REFRESH` не загружает `DTIMOD.sys`. Для загрузки драйвера и чтения IMOD / NIC ITR используйте `CHECK`.
- Драйвер, поднятый через KDU, остается в памяти до перезагрузки. Принудительная выгрузка не выполняется из‑за риска BSOD.

## Запуск

1. Скачайте EXE из [GitHub Releases](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases).
2. При наличии `SHA256SUMS.txt` (в локальном пакете сборки он есть; на GitHub сейчас обычно только два EXE) сверьте контрольные суммы.
3. Запустите от имени администратора.
4. Проверьте устройства и параметры, настройте вручную или через `AUTO-OPTIMIZATION`.
5. Перезагрузите ПК, если программа сообщит о необходимости перезагрузки.

## Сборка

Из корня репозитория:

```powershell
cd "DEVICE TWEAKER"
powershell -NoProfile -ExecutionPolicy Bypass -File .\build.ps1 -Flavor both -Configuration Release -SkipImodDriverBuild
```

Пакет появится в `DEVICE TWEAKER\bin\ReleasePackages\<version>\`.

Подробнее: [`DEVICE TWEAKER/docs/BUILD.md`](./DEVICE%20TWEAKER/docs/BUILD.md). История изменений: [`DEVICE TWEAKER/CHANGELOG.md`](./DEVICE%20TWEAKER/CHANGELOG.md).

## Структура проекта

Каталог [`DEVICE TWEAKER`](./DEVICE%20TWEAKER):

- `Affinity` — CPU Affinity, RSS и ReservedCpuSets.
- `Core` — основная логика, бекапы, Raw Input и IMOD.
- `Devices` — обнаружение устройств, USB и NDIS топология.
- `GUI` — интерфейс WinForms.
- `Localization` — строки EN/RU.
- `Tweaks` — авто-оптимизация и сброс.
- `IMOD` — `DTIMOD.sys`, KDU и загрузчики.
- `assets` — иконка, wordmark и manifest.
- `Scripts` — сертификат драйвера, `Smoke-SafeGui.ps1` и вспомогательные тесты.
- `docs` — сборка и релиз.
- `Models` / `Interop` — модели данных и P/Invoke.

## Разработчик

Telegram — [@arsenzaa](https://t.me/arsenzaa)

## Лицензия

Проект — [GNU GPLv3](LICENSE).

KDU в `DEVICE TWEAKER/IMOD/KDU` — MIT.
