# DEVICE TWEAKER

[English](./README.md) · **Русский**

Утилита для настройки MSI, привязки прерываний, ReservedCpuSets и USB IMOD в Windows. Она объединяет параметры, знакомые по MSI Utility V3 и Interrupt Affinity Policy Tool, и добавляет автооптимизацию с учётом топологии процессора.

**[Скачать для Windows](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest/download/DEVICE.TWEAKER.SelfContained.exe)** · [Руководство](./docs/guide/README.ru.md) · [Скриншоты](./docs/SCREENSHOTS.ru.md) · [История изменений](./CHANGELOG.ru.md)

## Возможности

- Настраивайте MSI Mode, MSI Limit, IRQ Priority и привязку прерываний. Считывайте текущий режим IRQ и настраивайте ReservedCpuSets в отдельной панели.
- Подбирайте ядра для видеокарт, устройств ввода и сетевых адаптеров с помощью **АВТООПТИМИЗАЦИИ**. Меняйте настройки отдельных устройств через **ПРИМЕНИТЬ**.
- Настраивайте USB IMOD для всего контроллера, по ролям устройств или по номерам interrupter. Выбирайте CPU и число очередей RSS. Считывайте, применяйте или сохраняйте NIC ITR на поддерживаемых адаптерах Intel и Realtek.
- Измеряйте частоту событий мыши, настраивайте Mouse Throttle и энергосбережение, восстанавливайте настройки из бэкапов.

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
