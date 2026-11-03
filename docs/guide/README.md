# Guide

[Documentation](../README.md) / Guide · [Русский](./README.ru.md)

From interrupt handling to device settings. Read in order or jump to the hardware you are configuring.

## Foundations

**01 · [Interrupts: ISR, DPC and cores](./01-interrupts.md)**

The path from a device to an application thread. Execution levels, interrupts and competition for a physical core.

**02 · [MSI and MSI-X](./02-msi.md)**

Lines and messages, interrupt vectors, MSI Limit and reading the current device mode.

**03 · [Affinity and priority](./03-affinity.md)**

Windows policies, CPU masks, IRQ Priority and the rules used by the program.

## Devices and processor

**04 · [Graphics card](./04-gpu.md)**

Selecting a pair of physical cores, SMT siblings and GPU interrupt handling.

**05 · [USB and IMOD](./05-usb-imod.md)**

xHCI controllers, Event Rings and interrupters. Role-based values and rereading the device map.

**06 · [Networking, RSS and NIC ITR](./06-network.md)**

NDIS and NetAdapterCx, RSS queues, moderation registers and the network driver's own settings.

**07 · [CPU topology](./07-cpu-topology.md)**

Physical cores, P/E, CPPC, CCD/CCX and ReservedCpuSets. How to read the processor map.

**08 · [Input and Polling Rate](./08-input.md)**

USB polling, Raw Input events, background input processing and Mouse Throttle.

**09 · [Power saving](./09-power.md)**

USB selective suspend, permission to turn off a device and verification of written settings.

## Verification and implementation

**10 · [Verifying results](./10-verification.md)**

Logs, readback, external traces and evidence for understanding behaviour on particular hardware.

**11 · [Driver and startup](./11-driver.md)**

DTIMOD.sys, KDU, write boundaries and the independent PowerShell profile.

---

[Settings reference](../SETTINGS.md) · [Build and develop](../DEVELOPMENT.md)
