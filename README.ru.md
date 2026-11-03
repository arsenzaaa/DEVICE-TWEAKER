<h1 align="center">
  <picture>
    <source media="(prefers-color-scheme: dark)" srcset="./assets/branding/wordmark-dark.svg">
    <img src="./assets/branding/wordmark-light.svg" width="320" height="97" alt="DEVICE TWEAKER">
  </picture>
</h1>

<p align="center">Настройка прерываний устройств в Windows.</p>

<p align="center"><a href="https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest/download/DEVICE.TWEAKER.SelfContained.exe"><strong>Скачать</strong></a> · <a href="./docs/README.ru.md">Документация</a> · <a href="./CHANGELOG.ru.md">История изменений</a> · <a href="./docs/SCREENSHOTS.ru.md">Скриншоты</a></p>

<p align="center"><a href="./README.md">English</a> · <strong>Русский</strong></p>

## Возможности

- **Прерывания устройств:** MSI Mode и Limit, IRQ Priority, CPU affinity, чтение текущего режима IRQ и возможностей PCI. AUTO подбирает настройки по топологии процессора.
- **USB:** xHCI IMOD для контроллера, роли устройства или отдельного interrupter, чтение регистров и профили автозапуска.
- **Сеть:** ядра и очереди RSS, NIC ITR для поддерживаемых адаптеров Intel и Realtek.

Также доступны [ReservedCpuSets](./docs/guide/07-cpu-topology.ru.md), [замер частоты событий мыши и Mouse Throttle](./docs/guide/08-input.ru.md), [энергосбережение устройств](./docs/guide/09-power.ru.md), резервное копирование и восстановление настроек.

## Запуск

1. Скачайте [сборку со встроенным .NET](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest/download/DEVICE.TWEAKER.SelfContained.exe) и запустите EXE от имени администратора. Нужна Windows 10/11 x64; установка не требуется.
2. Выберите устройство, измените нужные значения и нажмите **ПРИМЕНИТЬ**. Для автоматического подбора используйте **АВТООПТИМИЗАЦИЮ**. Перед записью сохраняются исходные настройки.
3. После изменения MSI или Interrupt Affinity перезагрузите Windows. Для IMOD/NIC ITR кнопка **ПРОВЕРИТЬ** считывает регистры. Порядок проверки результата описан в [руководстве](./docs/guide/10-verification.ru.md).

Сохранённую копию можно выбрать через **ВОССТАНОВЛЕНИЕ**. Результаты операций находятся в `logs` рядом с EXE.

[Сборка без встроенного .NET](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/download/v0.0.4/DEVICE.TWEAKER.exe) требует [.NET 10 Desktop Runtime x64](https://dotnet.microsoft.com/download/dotnet/10.0). Файлы и контрольные суммы доступны на [странице выпуска](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest).

Для IMOD, NIC ITR и чтения конфигурации PCI используется DTIMOD.sys с загрузкой через KDU. Поддержка зависит от оборудования, драйвера устройства и настроек безопасности Windows. Условия загрузки описаны в [руководстве по драйверу](./docs/guide/11-driver.ru.md).

## Документация

- [Руководство](./docs/guide/README.ru.md) - прерывания Windows и настройка устройств.
- [Справочник настроек](./docs/SETTINGS.ru.md) - поля, кнопки, применение и восстановление.
- [Сборка и разработка](./docs/DEVELOPMENT.ru.md) - исходники, зависимости и проверки.
- [Участие в проекте](./.github/CONTRIBUTING.ru.md) - исправления, предложения и результаты на оборудовании.

## Поддержка

[Сообщить об ошибке](https://github.com/arsenzaaa/DEVICE-TWEAKER/issues/new?template=bug_report.yml) · [Отчёт об оборудовании](https://github.com/arsenzaaa/DEVICE-TWEAKER/issues/new?template=hardware_report.yml) · [Telegram](https://t.me/arsenzaa) · [Boosty](https://boosty.to/arsenza)

В сообщении укажите версию программы, сборку Windows, устройство и шаги воспроизведения. Прикладывая лог, удалите личные пути и идентификаторы.

---

[GPL-3.0](./LICENSE) · [Сторонние компоненты](./docs/THIRD_PARTY_NOTICES.ru.md) · [Безопасность](./.github/SECURITY.ru.md)
