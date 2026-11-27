# Interrupts: ISR, DPC and cores

[Documentation](../README.md) / [Guide](./README.md) / 01 · [Русский](./01-interrupts.ru.md)

## From device to application

When hardware needs attention, it raises an interrupt on a target logical processor. The kernel dispatches the driver's **ISR** (Interrupt Service Routine), which does only the time-critical work (acknowledging the hardware and reading status registers) and typically queues a **DPC** (Deferred Procedure Call) to finish processing later. Longer or pageable work is handed off to system worker threads or the receiving application's thread.

```text
device event → ISR (DIRQL) → DPC (DISPATCH_LEVEL) → worker / application thread (PASSIVE_LEVEL)
```

## IRQL and why thread priority cannot preempt a DPC

Every logical processor in Windows maintains an **IRQL** (Interrupt Request Level), stored in `CR8` (the local APIC Task Priority Register) on x64. Raising IRQL masks all interrupts at the current level and below on that CPU:

- `0` (`PASSIVE_LEVEL`): normal user-mode and kernel thread execution.
- `1` (`APC_LEVEL`): asynchronous procedure calls; thread context switches are still allowed.
- `2` (`DISPATCH_LEVEL`): the thread scheduler itself and ordinary DPCs.
- `3–11` (and `12` for high-priority interrupts): **DIRQL** (Device IRQL), where hardware ISRs run.
- `13–15`: clock interrupt, inter-processor interrupts (IPI), power and high/profile interrupts.

Thread priorities (`0–31`) only decide which ready thread runs while `IRQL < 2`. As soon as a CPU enters an ISR or a DPC (`IRQL >= 2`), the scheduler cannot switch threads on that core until the routine completes and IRQL drops back below `DISPATCH_LEVEL`. Both ISRs and DPCs execute in the context of whichever thread happened to be running on that core. Even a priority-31 real-time thread is preempted immediately.

This is why raising a game's process priority does nothing against driver stalls: long ISR or DPC bursts on the same physical core take CPU time straight out of the application. A minority of drivers use **Threaded DPCs** (which execute on a priority-31 kernel thread at `PASSIVE_LEVEL` and can be preempted by interrupts or other priority-31 threads), but most GPU, USB, audio and NDIS network drivers still rely on standard DPCs.

See *Windows Internals*, 7th ed., Part 1, pp. 488–492, Part 2, ch. 8, and [Managing hardware priorities](https://learn.microsoft.com/en-us/windows-hardware/drivers/kernel/managing-hardware-priorities).

## Which CPU executes the work

Interrupt Affinity controls which logical processors may receive a device's hardware interrupts. When an ISR calls `KeInsertQueueDpc`, Windows places the DPC in the **current CPU's** DPC queue by default, so moving the interrupt usually moves its DPC with it. A driver can still target another CPU's DPC queue explicitly (or delegate work to system threads scheduled independently), and neither ISRs nor DPCs follow the thread scheduler's SMT-avoidance rules. They fire on their assigned logical processor regardless of what is running on its SMT sibling.

See [Organization of DPC queues](https://learn.microsoft.com/en-us/windows-hardware/drivers/kernel/organization-of-dpc-queues).

## Observing interrupt processing

Before moving devices around, capture a baseline ETW trace under your actual workload. As a general reference from Microsoft's multimedia guidelines, well-behaved drivers aim for ISR durations under ~25 µs and DPC durations under ~100 µs; long spikes on a core running active game or audio threads are the primary target. The [verification chapter](./10-verification.md) explains how to identify the exact driver module and compare ISR/DPC placement before and after changing affinity.

---

[Guide](./README.md) · [Next: MSI and MSI-X](./02-msi.md)
