# Interrupts: ISR, DPC and cores

[Documentation](../README.md) / [Guide](./README.md) / 01 · [Русский](./01-interrupts.ru.md)

## From device to application

An interrupt notifies a processor of a device event. A typical driver handles it briefly in an ISR and defers some work to a DPC; the application receives data later. The processing path depends on the driver, input API and Windows version.

```text
device event → ISR → deferred driver processing → application
```

A driver can use several DPCs and worker threads. See [Introduction to DPC objects](https://learn.microsoft.com/en-us/windows-hardware/drivers/kernel/introduction-to-dpc-objects).

## IRQL and thread priority

IRQL controls which interrupts are permitted on the current processor; it differs from Windows thread priority. A conventional ISR runs at the device's DIRQL and an ordinary DPC at DISPATCH_LEVEL. Interrupts at the current or a lower level are masked; a higher level can interrupt processing. See [Managing hardware priorities](https://learn.microsoft.com/en-us/windows-hardware/drivers/kernel/managing-hardware-priorities).

[Threaded DPCs](https://learn.microsoft.com/en-us/windows-hardware/drivers/kernel/introduction-to-threaded-dpcs) can run at PASSIVE_LEVEL. Use event details in the trace to identify the execution context.

## Which CPU performs the work

Interrupt Affinity controls device interrupt assignment, not the affinity of every driver thread. A DPC has a target processor: by default, the CPU on which it was queued, although the driver can choose another. See [Organization of DPC queues](https://learn.microsoft.com/en-us/windows-hardware/drivers/kernel/organization-of-dpc-queues).

Moving an ISR and moving all device processing are different outcomes. Actual work placement, handler duration and competition with the application matter.

## What to verify

DEVICE TWEAKER shows the saved policy and available current interrupt information. ISR/DPC placement requires a separate ETW trace. A displayed mask, one FPS measurement or one maximum DPC duration does not establish better input latency.

Record the original conditions and compare the same workload before and after a change. See [chapter 10](./10-verification.md).

---

[Guide](./README.md) · [Next: MSI and MSI-X](./02-msi.md)
