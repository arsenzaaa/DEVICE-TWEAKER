# Graphics card

[Documentation](../README.md) / [Guide](./README.md) / 04 · [Русский](./04-gpu.ru.md)

## AUTO policy for GPUs

For a discrete GPU, AUTO selects a pair of physical cores on the target CCD, uses SpecCPU and High, enables MSI and removes the custom MSI Limit. An SMT sibling is not counted as a second physical core; CPU 0 is avoided where possible. AUTO returns integrated-GPU affinity to Windows.

Work distribution across the selected cores depends on the GPU driver. Compare AUTO with the original settings under the same workload.

## Manual configuration

The GPU card exposes CPU Affinity, MSI Mode/Limit, IRQ Priority and Policy. Reboot after applying interrupt policy and read IRQ Mode again.

Compare the original behavior with AUTO first. For subsequent experiments, change one parameter and keep the game scene, FPS limiter, power mode and graphics settings consistent. Compare repeated runs rather than the best individual result.

## What to inspect in a trace

Inspect actual graphics-stack ISR/DPC and worker-thread activity. Interrupt Affinity does not set the affinity of every GPU-driver thread; the driver can select a DPC's target CPU separately. See [DPC queue organization](https://learn.microsoft.com/en-us/windows-hardware/drivers/kernel/organization-of-dpc-queues).

Record module names and execution times together with the driver version, GPU, HAGS state and workload.

## HAGS and BIOS

The application does not manage HAGS, graphics-driver threads or firmware. Comparing HAGS is a separate experiment with its own measurement conditions.

Judge the outcome by stability and reproducible measurements on your system. Use **RESTORE** if behavior gets worse.

---

[Previous: Affinity and priority](./03-affinity.md) · [Guide](./README.md) · [Next: USB and IMOD](./05-usb-imod.md)
