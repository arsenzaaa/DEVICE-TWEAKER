# Graphics card

[Documentation](../README.md) / [Guide](./README.md) / 04 · [Русский](./04-gpu.ru.md)

## AUTO policy for GPUs

For a discrete GPU, AUTO looks for a pair of neighboring physical cores in the preferred topology, excluding E-cores. It avoids the core that contains CPU 0 when possible. AUTO configures **High** IRQ Priority, enables MSI and clears the custom MSI Limit. It sets **SpecCPU** when it finds a pair and otherwise skips the affinity assignment. Integrated graphics use Windows interrupt placement.

The selected cores are eligible to receive the GPU's interrupts. Which cores handle the interrupts and DPCs depends on the driver and the allocated vectors.

## Manual configuration

The GPU card exposes CPU Affinity, MSI Mode/Limit, IRQ Priority and Policy. Restart Windows after applying changes. If you changed MSI, click **READ IRQ MODE** and check the **IRQ Mode** field.

## Checking the result

Compare frame times and graphics-driver ISR/DPC activity with the original settings. Keep the game scene, graphics settings and FPS limit fixed. Use the procedure in [verifying changes](./10-verification.md).

---

[Previous: Affinity and priority](./03-affinity.md) · [Guide](./README.md) · [Next: USB and IMOD](./05-usb-imod.md)
