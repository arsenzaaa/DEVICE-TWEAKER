# Interrupts: ISR, DPC and cores

[Documentation](../README.md) / [Guide](./README.md) / 01 · [Русский](./01-interrupts.ru.md)

## From device to application

An interrupt signals a device event. The ISR (interrupt service routine) handles the immediate work and can queue a DPC (deferred procedure call) for work that can wait. A driver may also use worker threads.

```text
device event → ISR → deferred driver processing → application
```

## IRQL and thread priority

IRQL is the interrupt priority level on a logical processor. At a raised IRQL, Windows masks interrupts at the current level and below on that processor. A device ISR normally runs at its device IRQL and an ordinary DPC at `DISPATCH_LEVEL`. Thread priority determines which ready thread the scheduler chooses. See *Windows Internals*, 7th ed., Part 1, pp. 488–492, and [Managing hardware priorities](https://learn.microsoft.com/en-us/windows-hardware/drivers/kernel/managing-hardware-priorities).

## Which CPU performs the work

Interrupt Affinity controls device interrupt assignment. By default, an ISR that queues a DPC places it on the current CPU. The driver can target a different CPU. Worker threads are scheduled separately. See [Organization of DPC queues](https://learn.microsoft.com/en-us/windows-hardware/drivers/kernel/organization-of-dpc-queues).

## Observing interrupt processing

Record a trace under load with the original settings. The [verification chapter](./10-verification.md) shows how to find the driver module and compare ISR and DPC placement before and after an affinity change.

---

[Guide](./README.md) · [Next: MSI and MSI-X](./02-msi.md)
