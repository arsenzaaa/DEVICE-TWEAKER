# Документация

[DEVICE TWEAKER](../README.ru.md) / Документация · [English](./README.md)

Настройка устройств, принципы работы прерываний Windows и разработка DEVICE TWEAKER.

## Начало работы

- [Скриншоты](./SCREENSHOTS.ru.md) - настройки GPU, USB и сети.
- [Справочник настроек](./SETTINGS.ru.md) - поля, кнопки, применение и восстановление.
- [Руководство](./guide/README.ru.md) - 11 глав от ISR/DPC и MSI до IMOD и RSS.
- [Логи и проверка](./guide/10-verification.ru.md) - сохранённые настройки, регистры и поведение системы.

## По устройствам

- **[Видеокарта](./guide/04-gpu.ru.md)** - физические ядра, SMT и политика привязки.
- **[USB](./guide/05-usb-imod.ru.md)** - xHCI-контроллеры, interrupter и IMOD для мыши, клавиатуры и аудио.
- **[Сеть](./guide/06-network.ru.md)** - RSS, NDIS/NetAdapterCx и NIC ITR.
- **[Процессор](./guide/07-cpu-topology.ru.md)** - P/E-ядра, CPPC, CCD/CCX и ReservedCpuSets.
- **[Ввод](./guide/08-input.ru.md)** - Polling Rate, Raw Input и Mouse Throttle.
- **[Энергосбережение](./guide/09-power.ru.md)** - USB selective suspend и отключение устройств.

## Драйвер и автозапуск

[DTIMOD.sys, KDU и PowerShell-профиль](./guide/11-driver.ru.md): какие функции требуют драйвера, как ограничена запись в память и как профиль работает без основного EXE.

## Для разработчиков

[Сборка и устройство проекта](./DEVELOPMENT.ru.md) · [Локальный QA](../dev/qa/README.ru.md) · [Стенд GUI](../dev/qa/WORKBENCH.ru.md) · [Подготовка релиза](./RELEASE.ru.md) · [Участие в проекте](../.github/CONTRIBUTING.ru.md)

---

[История изменений](../CHANGELOG.ru.md) · [Выпуски](./releases/README.ru.md) · [Сторонние компоненты](./THIRD_PARTY_NOTICES.ru.md) · [Политика безопасности](../.github/SECURITY.ru.md)
