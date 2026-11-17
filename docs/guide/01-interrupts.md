# Interrupts: ISR, DPC and cores

[Documentation](../README.md) / [Guide](./README.md) / 01 · [Русский](./01-interrupts.ru.md)

## From device to application

An interrupt notifies a processor of a device event. A driver runs a short ISR (interrupt service routine) and defers part of the work to a DPC (deferred procedure call); the application receives data later. The processing path depends on the driver, input API and Windows version.

```text
device event → ISR → deferred driver processing → application
```

A driver can use several DPCs and worker threads. See [Introduction to DPC objects](https://learn.microsoft.com/en-us/windows-hardware/drivers/kernel/introduction-to-dpc-objects).

## IRQL and thread priority

IRQL controls which interrupts are permitted on the current processor; it differs from Windows thread priority. A conventional ISR runs at the device's DIRQL and an ordinary DPC at DISPATCH_LEVEL. Interrupts at the current or a lower level are masked; a higher level can interrupt processing. See [Managing hardware priorities](https://learn.microsoft.com/en-us/windows-hardware/drivers/kernel/managing-hardware-priorities).

## Which CPU performs the work

Interrupt Affinity controls device interrupt assignment. DPCs and driver threads have their own processor targets. A DPC defaults to the CPU on which it was queued; the driver can select a different target. These rules are described in [Organization of DPC queues](https://learn.microsoft.com/en-us/windows-hardware/drivers/kernel/organization-of-dpc-queues).

## Observing interrupt processing

DEVICE TWEAKER displays the saved policy and current interrupt information. An ETW trace records processor placement and execution time for ISR/DPC handlers. The [verification chapter](./10-verification.md) covers trace capture and before/after comparisons.

---

[Guide](./README.md) · [Next: MSI and MSI-X](./02-msi.md)
