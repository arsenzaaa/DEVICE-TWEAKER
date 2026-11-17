# DEVICE TWEAKER

[English](./README.md) · **Русский**

**DEVICE TWEAKER** - утилита для Windows, объединяющая возможности MSI Utility V3 и Interrupt Affinity Policy Tool, настройку ReservedCpuSets и управление USB IMOD.

**[Скачать для Windows](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest/download/DEVICE.TWEAKER.SelfContained.exe)** · [Руководство](./docs/guide/README.ru.md) · [Скриншоты](./docs/SCREENSHOTS.ru.md) · [История изменений](./CHANGELOG.ru.md)

## Возможности

### Прерывания и привязка к ядрам

Для каждого устройства можно задать MSI Mode, MSI Limit, IRQ Priority и Interrupt Affinity. Текущий режим IRQ отображается рядом с сохранённой конфигурацией. Отдельно доступна настройка ReservedCpuSets.

### Автооптимизация с учётом топологии CPU

Автоматический подбор ядер для видеокарты, устройств ввода и сетевых адаптеров учитывает физические ядра и SMT-потоки, P/E-ядра, CCD и доступные оценки CPPC. Настройки каждого устройства можно скорректировать вручную.

### Настройка USB и сети

USB IMOD настраивается для контроллера целиком, по ролям устройств или по индексам interrupter. Для сети доступны выбор процессоров и числа очередей RSS, а на поддерживаемых адаптерах Intel и Realtek - управление NIC ITR. Значения IMOD и NIC ITR можно прочитать из регистров и сохранить в профиль автозапуска.

### Резервные копии и диагностика

Исходные настройки устройств и последующие бэкапы доступны через ВОССТАНОВЛЕНИЕ. Логи содержат сведения о сборке, устройстве, запрошенных значениях и выполненных этапах. Дополнительные инструменты включают замер частоты событий мыши, Mouse Throttle и управление энергосбережением устройств.

## Начало работы

Нужны **Windows 10/11 x64** и права администратора.

**[Сборка со встроенным .NET](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest/download/DEVICE.TWEAKER.SelfContained.exe)** готова к запуску без установки. [Сборка без встроенного .NET](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest/download/DEVICE.TWEAKER.exe) занимает меньше места, но требует установленного [.NET 10 Desktop Runtime x64](https://dotnet.microsoft.com/download/dotnet/10.0).

1. Запустите EXE от имени администратора и выберите устройство в списке.
2. Измените нужные параметры и нажмите **ПРИМЕНИТЬ** либо используйте **АВТООПТИМИЗАЦИЮ** для автоматического подбора.
3. После изменения MSI, Interrupt Affinity или ReservedCpuSets перезагрузите Windows. Порядок проверки описан в [руководстве](./docs/guide/10-verification.ru.md).

Сохранённые копии доступны через **ВОССТАНОВЛЕНИЕ**, а логи находятся в папке `logs` рядом с EXE. Все файлы выпуска и контрольные суммы размещены на [странице выпусков](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest).

USB IMOD, NIC ITR и аппаратное чтение PCI используют DTIMOD.sys через KDU. Остальные настройки работают независимо от этого драйвера. DTIMOD может конфликтовать с античитами; условия загрузки и работа автозапуска описаны в [главе о драйвере](./docs/guide/11-driver.ru.md).

## Документация

- [Руководство](./docs/guide/README.ru.md) - работа прерываний, топология процессора и настройка устройств.
- [Справочник настроек](./docs/SETTINGS.ru.md) - элементы интерфейса, применение изменений и восстановление.
- [Разработка](./docs/DEVELOPMENT.ru.md) - сборка из исходников и целевые проверки с помощью QA-инструментов.

## Обратная связь и поддержка

Ошибки и предложения можно оставить в [Issues](https://github.com/arsenzaaa/DEVICE-TWEAKER/issues). Приложите лог с нужной операцией, предварительно убрав личные пути и имена. Результатами проверки на своём оборудовании можно поделиться через [форму отчёта](https://github.com/arsenzaaa/DEVICE-TWEAKER/issues/new?template=hardware_report.yml). Новости проекта публикуются в [Telegram-канале](https://t.me/arsenzaa).

Поддержать развитие DEVICE TWEAKER можно на [Boosty](https://boosty.to/arsenza).

[Участие в проекте](./.github/CONTRIBUTING.ru.md) · [GPL-3.0](./LICENSE) · [Сторонние компоненты](./docs/THIRD_PARTY_NOTICES.ru.md) · [Безопасность](./.github/SECURITY.ru.md)
