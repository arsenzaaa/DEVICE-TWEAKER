# MSI and MSI-X

[Documentation](../README.md) / [Guide](./README.md) / 02 · [Русский](./02-msi.ru.md)

## Line and message

A line-based interrupt uses an interrupt line; MSI is delivered by writing to a special memory address. MSI-X provides more flexible message management. Availability depends on the device and its driver. See [Introduction to message-signaled interrupts](https://learn.microsoft.com/en-us/windows-hardware/drivers/kernel/introduction-to-message-signaled-interrupts).

MSI does not automatically guarantee lower latency or uninterrupted DPC execution. DEVICE TWEAKER changes the setting; compatibility and results must be checked on the actual device after a reboot.

## Saved setting and current mode

| Field | Meaning |
| :--- | :--- |
| MSI Mode | The saved `MSISupported` registry setting. |
| IRQ Mode | Available current-mode evidence: Line, MSI, MSI-X or combined MSI/MSI-X. |
| HW | The PCI device's capabilities and available message maximum. |
| IRQ Count | Allocated-resource information, not interrupt frequency. |

**READ IRQ MODE** queries all PCI devices in the list and loads DTIMOD if needed. It reads Windows interrupt resources and PCI configuration. Background reads at startup or refresh use Windows resources or an already loaded driver. If the mode cannot be determined, the result includes the reason; the registry setting remains a separate field.

## MSI Limit

`MessageNumberLimit` caps the requested number of messages. **Unlimited** removes this override and leaves the decision to Windows and the driver; it does not promise the maximum number of vectors.

Windows documents 1, 2, 4, 8 and 16 for multi-message MSI, and up to 2048 for MSI-X. See [Enabling MSI in the registry](https://learn.microsoft.com/en-us/windows-hardware/drivers/kernel/enabling-message-signaled-interrupts-in-the-registry). The application also checks a known hardware maximum and rejects incompatible input.

```text
HKLM\SYSTEM\CurrentControlSet\Enum\PCI\<device>\<instance>\Device Parameters\Interrupt Management\MessageSignaledInterruptProperties
  MSISupported        REG_DWORD
  MessageNumberLimit  REG_DWORD
```

## Configuration sequence

1. Record the original MSI Mode, IRQ Mode and HW information.
2. Apply the chosen setting. The application creates a recovery backup before writing.
3. Reboot, read IRQ mode again and exercise the device under load.
4. Restore the saved values if device behavior gets worse.

AUTO selects MSI and removes the custom limit for the device categories it configures. **RESET WINDOWS DEFAULT** deliberately preserves MSI Mode and Limit; restoring those values requires a suitable backup.

---

[Previous: Interrupts: ISR, DPC and cores](./01-interrupts.md) · [Guide](./README.md) · [Next: Affinity and priority](./03-affinity.md)
