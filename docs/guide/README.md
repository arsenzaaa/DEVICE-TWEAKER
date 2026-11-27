# Guide

[Documentation](../README.md) / Guide · [Русский](./README.ru.md)

How Windows handles hardware interrupts and how DEVICE TWEAKER configures GPUs, USB controllers, network adapters and CPUs.

## Foundations

**01 · [Interrupts: ISR, DPC and cores](./01-interrupts.md)**

IRQL levels (`CR8`), why thread priorities cannot preempt a DPC, per-CPU queues and SMT contention.

**02 · [MSI and MSI-X](./02-msi.md)**

Shared INTx lines (`KiChainedDispatch`) vs posted MSI/MSI-X writes, `MessageNumberLimit` and live PCI mode readback.

**03 · [Affinity and priority](./03-affinity.md)**

`IRQ_DEVICE_POLICY`, `AssignmentSetOverride`, Interrupt Steering and what `DevicePriority` (DIRQL) actually changes.

## Devices and processor

**04 · [Graphics card](./04-gpu.md)**

Why AUTO assigns two neighboring physical cores to a discrete GPU and how to verify frame-time impact.

**05 · [USB and IMOD](./05-usb-imod.md)**

xHCI Event Rings and interrupters, `0xC8` (`usbxhci.sys`) vs `0xFA0` (xHCI spec), role mapping and polling-rate trade-offs.

**06 · [Networking, RSS and NIC ITR](./06-network.md)**

NDIS vs NetAdapterCx, RSS queues and physical-core selection, hardware ITR registers and throughput trade-offs.

**07 · [CPU topology](./07-cpu-topology.md)**

SMT scheduling vs interrupt placement, P/E cores, CPPC ranks, CCD/L3 domains and `ReservedCpuSets`.

**08 · [Input and Polling Rate](./08-input.md)**

USB polling to Raw Input (`win32k`), `GetRawInputBuffer` and Windows 11 `RawMouseThrottleDuration`.

**09 · [Power saving](./09-power.md)**

USB selective suspend, `PnPCapabilities` / WMI `MSPower_DeviceEnable`, sleep verification and AC/DC restore.

## Verification and implementation

**10 · [Verifying results](./10-verification.md)**

Logs, MMIO readback, controlled VBS/HVCI comparisons and ETW trace analysis with Xperf, WPA and MXA.

**11 · [Driver and startup](./11-driver.md)**

DTIMOD scoping, KDU/anti-cheat boundaries, scheduled startup profiles and the PnP/power lifecycle limitation.

---

[Settings reference](../SETTINGS.md) · [Build and develop](../DEVELOPMENT.md)
