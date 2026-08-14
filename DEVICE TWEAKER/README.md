# DEVICE TWEAKER

[![.NET](https://img.shields.io/badge/.NET-8.0--windows-512BD4?logo=dotnet&logoColor=white)](https://dotnet.microsoft.com/download/dotnet/8.0)
[![Platform](https://img.shields.io/badge/platform-Windows%2010%2F11%20x64-0078D6?logo=windows&logoColor=white)](#требования)
[![License](https://img.shields.io/badge/license-GPLv3-blue.svg)](LICENSE)
[![Telegram](https://img.shields.io/badge/Telegram-arsenzaa-2CA5E0?logo=telegram&logoColor=white)](https://t.me/arsenzaa)

**DEVICE TWEAKER** — утилита для Windows 10 и Windows 11, объединяющая MSI Utility, Interrupt Affinity Policy Tool, ReservedCpuSets, RSS, NIC ITR и управление USB IMOD в одном интерфейсе с авто-оптимизацией.

Текущая версия **v0.0.4-alpha.2**.

## Возможности

- Отображение USB-контроллеров, видеокарт, накопителей, аудиоустройств и сетевых адаптеров отдельными блоками.
- Настройка MSI Mode, MSI Limit, IRQ Priority и политики распределения прерываний.
- Ручное распределение устройств по логическим процессорам.
- Настройка `ReservedCpuSets` с отображением текущего значения.
- Настройка RSS и ITR у поддерживаемых сетевых адаптеров.
- Настройка IMOD у поддерживаемых USB-контроллеров через `DTIMOD.sys`.
- Авто-оптимизация с учетом P-Core, E-Core, SMT/Hyper-Threading, CPPC, CCD и CCX.
- Создание резервных копий, восстановление выбранного бекапа и полный сброс изменений.

## Требования

- Windows 10 или Windows 11 x64.
- Права администратора.
- Для обычной версии требуется .NET 8 Desktop Runtime или новее.
- Версия `DEVICE.TWEAKER.NET.FRAMEWORK.exe` уже содержит необходимый .NET 8 Runtime. Название сохранено для совместимости с предыдущими релизами. Это не старый .NET Framework.

## Важно

Программа изменяет параметры прерываний, значения в `HKLM` и при использовании IMOD загружает драйвер `DTIMOD.sys`.

- Перед применением настроек создавайте резервную копию.
- Не используйте случайные IMOD/ITR-значения, если не понимаете их назначение.
- После серьезных изменений может потребоваться перезагрузка Windows.
- Подробное логирование включается автоматически при запуске. Для каждого запуска создается отдельный `DeviceTweaker`-лог в папке `logs` рядом с EXE. Работа автозапускного IMOD-скрипта записывается в `ApplyIMOD_дата.log` в той же папке.
- `REFRESH` не загружает `DTIMOD.sys`. Для загрузки драйвера и чтения текущих значений IMOD или NIC ITR используется кнопка `CHECK`.
- Загруженный через KDU драйвер `DTIMOD.sys` остается в памяти до перезагрузки Windows. Программа не выполняет его принудительную выгрузку из-за риска BSOD на отдельных системах.

## Запуск

1. Скачайте один из EXE-файлов из [GitHub Releases](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases).
2. При наличии `SHA256SUMS.txt` сверьте контрольные суммы.
3. Запустите программу от имени администратора.
4. Проверьте найденные устройства и предлагаемые параметры.
5. Настройте устройства вручную или используйте `AUTO-OPTIMIZATION`.
6. Перезагрузите компьютер, если программа сообщит о необходимости перезагрузки.

## Сборка

Команда для сборки обеих версий с использованием готового `DTIMOD.sys`.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\build.ps1 -Flavor both -Configuration Release -SkipImodDriverBuild
```

После сборки готовый пакет появляется в `bin\ReleasePackages\v0.0.4-alpha.2\`.

Подробная инструкция находится в [docs/BUILD.md](docs/BUILD.md). История изменений — в [CHANGELOG.md](CHANGELOG.md).

## Структура проекта

- `Affinity` - CPU Affinity, RSS и ReservedCpuSets.
- `Core` - основная логика, резервные копии, Raw Input и IMOD.
- `Devices` - поиск устройств и определение USB и NDIS топологии.
- `GUI` - интерфейс WinForms.
- `Tweaks` - авто-оптимизация и сброс настроек.
- `IMOD` - драйвер `DTIMOD.sys`, KDU и необходимые загрузчики.
- `assets` - иконка, wordmark и manifest приложения.
- `Scripts` - сертификат драйвера и безопасный GUI smoke (`Smoke-SafeGui.ps1`).
- `docs` - инструкции по сборке и релизу.
- `Models` / `Interop` - модели данных и P/Invoke.

## Разработчик

Telegram - [@arsenzaa](https://t.me/arsenzaa)

## Лицензия

Проект распространяется по лицензии [GNU GPLv3](LICENSE).

Исходный код KDU находится в `IMOD/KDU` и распространяется по лицензии MIT.
