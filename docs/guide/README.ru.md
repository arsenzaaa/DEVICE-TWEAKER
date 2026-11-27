# Руководство

[Документация](../README.ru.md) / Руководство · [English](./README.md)

Как устроены аппаратные прерывания в Windows и как DEVICE TWEAKER настраивает GPU, USB-контроллеры, сетевые адаптеры и процессор.

## Основы

**01 · [Прерывания: ISR, DPC и ядра](./01-interrupts.ru.md)**

Уровни IRQL (`CR8`), почему приоритет потока не спасает от DPC, очереди по CPU и конкуренция на SMT.

**02 · [MSI и MSI-X](./02-msi.ru.md)**

Разделяемые линии INTx (`KiChainedDispatch`) против MSI/MSI-X, `MessageNumberLimit` и чтение реального режима PCI.

**03 · [Affinity и приоритет](./03-affinity.ru.md)**

`IRQ_DEVICE_POLICY`, `AssignmentSetOverride`, Interrupt Steering и на что на самом деле влияет `DevicePriority` (DIRQL).

## Устройства и процессор

**04 · [Видеокарта](./04-gpu.ru.md)**

Почему AUTO выделяет дискретной GPU пару соседних физических ядер и как проверять влияние на время кадра.

**05 · [USB и IMOD](./05-usb-imod.ru.md)**

Event Ring и interrupter'ы xHCI, `0xC8` (`usbxhci.sys`) против `0xFA0` (спецификация xHCI), карта ролей и компромиссы по частоте опроса.

**06 · [Сеть, RSS и NIC ITR](./06-network.ru.md)**

NDIS против NetAdapterCx, очереди RSS и выбор физических ядер, регистры NIC ITR и компромиссы под нагрузкой.

**07 · [Топология CPU](./07-cpu-topology.ru.md)**

Планировщик SMT против прерываний, P/E-ядра, рейтинги CPPC, домены CCD/L3 и `ReservedCpuSets`.

**08 · [Ввод и Polling Rate](./08-input.ru.md)**

Путь от USB-опроса до Raw Input (`win32k`), `GetRawInputBuffer` и `RawMouseThrottleDuration` в Windows 11.

**09 · [Энергосбережение](./09-power.ru.md)**

USB selective suspend, `PnPCapabilities` / WMI `MSPower_DeviceEnable`, проверка после сна и восстановление AC/DC.

## Проверка и реализация

**10 · [Проверка результата](./10-verification.ru.md)**

Логи, обратное чтение MMIO, контроль VBS/HVCI и анализ трасс ETW в Xperf, WPA и MXA.

**11 · [Драйвер и автозапуск](./11-driver.ru.md)**

Ограничения доступа DTIMOD, KDU и античиты, задача автозапуска и ограничение жизненного цикла PnP/питания.

---

[Справочник настроек](../SETTINGS.ru.md) · [Сборка и разработка](../DEVELOPMENT.ru.md)
