# Graphics card

[Documentation](../README.md) / [Guide](./README.md) / 04 · [Русский](./04-gpu.ru.md)

## AUTO policy for GPUs

For a discrete GPU, AUTO selects a pair of physical cores on the target CCD, uses SpecCPU and High, enables MSI and removes the custom MSI Limit. An SMT sibling is not counted as a second physical core; CPU 0 is avoided where possible. AUTO returns integrated-GPU affinity to Windows.

Work distribution across the selected cores depends on the GPU driver. Compare AUTO with the original settings under the same workload.

## Manual configuration

The GPU card exposes CPU Affinity, MSI Mode/Limit, IRQ Priority and Policy. Reboot after applying interrupt policy and read IRQ Mode again.

Compare the original behavior with AUTO first. For subsequent experiments, change one parameter and keep the game scene, FPS limiter, power mode and graphics settings consistent. Compare repeated runs rather than the best individual result.

## Checking the result

Compare frame times and graphics-driver ISR/DPC activity using the method in [chapter 10](./10-verification.md). Record the GPU model, driver version, HAGS state and workload with each result. Saved settings are available through **RESTORE**.

---

[Previous: Affinity and priority](./03-affinity.md) · [Guide](./README.md) · [Next: USB and IMOD](./05-usb-imod.md)
