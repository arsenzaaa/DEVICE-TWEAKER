# DEVICE TWEAKER

[English](./README.md) · **Русский**

Настройка аппаратных прерываний и устройств Windows в одном окне — MSI/MSI-X, привязка прерываний к ядрам, ReservedCpuSets, USB IMOD, RSS и NIC ITR с автооптимизацией под топологию CPU, проверкой фактических значений и безопасными бэкапами.

**[Скачать для Windows](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest/download/DEVICE.TWEAKER.SelfContained.exe)** · [Руководство](./docs/guide/README.ru.md) · [Скриншоты](./docs/SCREENSHOTS.ru.md) · [История изменений](./CHANGELOG.ru.md)

## Возможности

- **Прерывания и привязка к ядрам:** настройка MSI/MSI-X, MSI Limit, IRQ Priority, масок CPU Affinity и ReservedCpuSets с чтением фактического режима IRQ для PCI-устройств.
- **Автооптимизация под топологию CPU:** автоматическое распределение прерываний GPU, ввода и сети по оптимальным физическим ядрам (P/E-ядра, CCD/CCX, SMT) — или ручная настройка любого устройства через **ПРИМЕНИТЬ**.
- **USB IMOD и сеть:** настройка xHCI IMOD для всего контроллера, по ролям устройств или по индексам interrupter с обратным чтением; выбор очередей и базового CPU для RSS, а также чтение, применение и сохранение NIC ITR на поддерживаемых адаптерах Intel и Realtek.
- **Ввод, питание и безопасный откат:** измерение частоты опроса мыши в реальном времени, управление Mouse Throttle и энергосбережением, восстановление настроек из автоматических бэкапов.

## Начало работы

Нужны **Windows 10/11 x64** и права администратора. Основной EXE включает .NET. Компактная сборка требует .NET 10 Desktop Runtime x64 и доступна на [странице выпуска](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest).

1. Запустите EXE от имени администратора и выберите устройство в списке.
2. Измените настройки и нажмите **ПРИМЕНИТЬ** или выберите **АВТООПТИМИЗАЦИЯ** для автоматической настройки.
3. Перезагрузите Windows после изменения MSI, привязки прерываний или ReservedCpuSets. Проверьте результат по [руководству](./docs/guide/10-verification.ru.md).

Программа использует DTIMOD для USB IMOD, NIC ITR и чтения конфигурации PCI. Условия загрузки и возможные конфликты с античитами описаны в [руководстве по драйверу](./docs/guide/11-driver.ru.md).

## Документация

- [Справочник настроек](./docs/SETTINGS.ru.md)
- [Сборка из исходников и QA](./docs/DEVELOPMENT.ru.md)

## Обратная связь и поддержка

[Сообщить об ошибке](https://github.com/arsenzaaa/DEVICE-TWEAKER/issues/new?template=bug_report.yml) · [Отчёт об оборудовании](https://github.com/arsenzaaa/DEVICE-TWEAKER/issues/new?template=hardware_report.yml) · [Новости в Telegram](https://t.me/arsenzaa) · [Поддержать на Boosty](https://boosty.to/arsenza)

[Участие в проекте](./.github/CONTRIBUTING.ru.md) · [GPL-3.0](./LICENSE) · [Сторонние компоненты](./docs/THIRD_PARTY_NOTICES.ru.md) · [Безопасность](./.github/SECURITY.ru.md)
