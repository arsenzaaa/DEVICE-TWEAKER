# MSI and MSI-X

[Documentation](../README.md) / [Guide](./README.md) / 02 · [Русский](./02-msi.ru.md)

## Line-based INTx vs message-signaled interrupts

Legacy PCI devices signal interrupts over level-triggered **INTx** lines routed through the IOAPIC. Because lines are often shared across multiple devices (`IRQ sharing`), a single trigger forces Windows to walk a chain of registered ISRs (`KiChainedDispatch`) until a driver claims the interrupt, which is a well-known source of elevated interrupt latency and stability issues (*Windows Internals*, 7th ed., Part 2, ch. 8). Drivers on level-triggered lines also frequently have to perform an extra MMIO read to flush posted DMA writes and deassert the line.

**MSI** replaces the shared pin with a PCIe memory write directed straight to the local APIC. Each device gets dedicated vectors (shown as negative IRQ numbers in Device Manager), and PCIe transaction ordering guarantees that prior DMA data is visible in memory by the time the interrupt arrives.

**MSI-X** extends this model up to **2048 vectors**, with an independent message address and data entry per vector stored in the device's MSI-X table (in BAR space). That allows a driver to route separate queues or interrupters (such as network RSS receive queues or xHCI interrupters) directly to different CPUs or NUMA nodes. Availability depends on both the hardware capability and the driver. See [Introduction to message-signaled interrupts](https://learn.microsoft.com/en-us/windows-hardware/drivers/kernel/introduction-to-message-signaled-interrupts).

## Configured setting vs live IRQ Mode

**MSI Mode** controls the device's `MSISupported` registry value under `Device Parameters\Interrupt Management\MessageSignaledInterruptProperties`. Clicking **APPLY** writes the registry setting; Windows reads it when the device stack starts (after a reboot).

**IRQ Mode** reports what the device is actually using right now. Click **READ IRQ MODE** in the toolbar to query every listed PCI device, including cards hidden by filters. That command combines Windows allocated interrupt resources with PCI configuration space (`MSI Enable` / `MSI-X Enable` bits read through DTIMOD), which distinguishes MSI from MSI-X even when only one IRQ resource is allocated. Background reads on startup or **REFRESH** never load DTIMOD on their own; they rely on Windows resources and any already loaded driver.

A card showing **MSI Mode: Enabled** alongside **IRQ Mode: Line** means MSI is requested in the registry, yet the device is currently running on INTx. If that mismatch remains after rebooting, the driver or platform did not adopt MSI.

The **HW** line lists the device's PCI capabilities and hardware message ceiling, while **IRQ** shows the interrupt-resource count reported by WMI. A device can support dozens of vectors in hardware while its driver only requests one. When Windows resources confirm message-signaled interrupts without PCI config readback, the card displays **MSI/MSI-X** until `READ IRQ MODE` resolves the exact bit.

## MSI Limit

**MSI Limit** writes `MessageNumberLimit`, capping how many message interrupts Windows allocates to the device. Selecting **Unlimited** removes the registry value so Windows and the driver negotiate the vector count normally. The maximum shown in **HW** is a hardware ceiling, not a target. Artificially raising or capping the limit rarely helps unless troubleshooting a specific driver quirk.

In the Windows registry, valid limits are powers of two (**1, 2, 4, 8 or 16**) for MSI and up to **2048** for MSI-X. See [Enabling message-signaled interrupts in the registry](https://learn.microsoft.com/en-us/windows-hardware/drivers/kernel/enabling-message-signaled-interrupts-in-the-registry). DEVICE TWEAKER validates manual input against the detected capabilities of the device.

## Applying, safeguards and rollback

After changing MSI Mode or MSI Limit, click **APPLY**, reboot, and run **READ IRQ MODE** to confirm the active state under the same workload.

If a device's PCI capabilities explicitly report message-signaled interrupts without line-based INTx support, DEVICE TWEAKER refuses to disable MSI so the device cannot be left without a working interrupt path.

**AUTO-OPTIMIZATION** enables MSI and clears custom limits on the devices it tunes, preserves any existing limit on display audio, and skips Wi-Fi adapters. **RESET WINDOWS DEFAULT** intentionally leaves `MSISupported` and `MessageNumberLimit` untouched; use **RESTORE** with a backup to revert them.

---

[Previous: Interrupts: ISR, DPC and cores](./01-interrupts.md) · [Guide](./README.md) · [Next: Affinity and priority](./03-affinity.md)
