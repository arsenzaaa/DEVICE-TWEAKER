# MSI and MSI-X

[Documentation](../README.md) / [Guide](./README.md) / 02 · [Русский](./02-msi.ru.md)

## Line and message

A PCI device can signal an event through an interrupt line or a message. MSI uses a memory write instead of an interrupt pin; MSI-X gives the driver more control over its messages. The device and driver determine which modes are available. See [message-signaled interrupts](https://learn.microsoft.com/en-us/windows-hardware/drivers/kernel/introduction-to-message-signaled-interrupts).

## Saved setting and current mode

**MSI Mode** selects the device setting. **APPLY** writes `MSISupported` to its registry key.

**IRQ Mode** shows the mode currently reported for the device. Click **READ IRQ MODE** at the top to read all listed PCI devices, including cards hidden by a filter. This command can load DTIMOD. Automatic background reads use Windows resources or an already loaded driver.

For example, **MSI Mode: Enabled** with **IRQ Mode: Line** means that MSI is configured and the device is currently using a line interrupt. If both readings persist after a reboot, check the device and driver.

The **HW** part of the card lists supported modes and the maximum number of messages. The nearby **IRQ** number counts interrupt resources reported by WMI. A device can have one allocated interrupt resource while supporting multiple messages. When the data confirms message-based interrupts, the card shows **MSI/MSI-X** until PCI configuration can distinguish the two.

## MSI Limit

**MSI Limit** caps the requested message allocation. **Unlimited** removes the custom limit and lets Windows and the driver choose it. The hardware maximum shown in **HW** describes capability, not a recommended limit.

Windows accepts 1, 2, 4, 8 or 16 for MSI, and up to 2048 for MSI-X. See [enabling MSI in the registry](https://learn.microsoft.com/en-us/windows-hardware/drivers/kernel/enabling-message-signaled-interrupts-in-the-registry). The program validates the entered value against the information available for the device.

## Apply and restore

Select the desired mode and limit, click **APPLY**, then restart Windows. Read **IRQ Mode** again and check normal use of the device under the same workload. Keep the original backup so you can return to it through **RESTORE**.

If PCI capabilities explicitly report message interrupts without line-based INTx support, the program refuses to disable MSI and leaves that setting unchanged.

AUTO enables MSI and removes the custom limit on devices it configures. It leaves the limit intact for display audio and skips Wi-Fi. **RESET WINDOWS DEFAULT** preserves MSI Mode and Limit. Restore a backup to return those two settings to their previous values.

---

[Previous: Interrupts: ISR, DPC and cores](./01-interrupts.md) · [Guide](./README.md) · [Next: Affinity and priority](./03-affinity.md)
