# MSI and MSI-X

## Line and message

The old way is a **line-based** interrupt: the device raises a signal on an IOAPIC line. Several devices may share a line (IRQ-SHARING), and then every interrupt makes Windows call the ISR of every driver on that line in turn (`KiChainedDispatch`). Windows Internals names shared interrupts as a cause of high interrupt latency and instability.

**MSI** (Message Signaled Interrupts) is an interrupt by a memory write: the device writes a message to an address, and it arrives at the LAPIC of the target processor. There are no shared lines, and the data arrives with the interrupt.

- **MSI** — up to 32 messages, one shared address.
- **MSI-X** — up to 2048 messages, each with its own address and CPU binding. MSI-X is what lets a network card spread RSS queues over different cores.

## What MSI MODE does in practice

A test on a USB xHCI Controller with one vector:

```text
MSI MODE off   the DPC (WDF01000) is interrupted by two ISRs (WDF01000), and the ISRs themselves are longer
MSI MODE on    the DPC (WDF01000) runs uninterrupted, the ISRs are shorter
```

Without MSI you get extra ISRs, and each one is longer. That is a widespread and significant problem; IRQ-SHARING is rarer, and MSI fixes it too.

**Conclusion:** if you are still not using MSI MODE, you are making a mistake. The exceptions are devices and drivers that do not work with MSI: on some network adapters enabling MSI led to a BSOD or lost connectivity. Always restart after enabling it and check the device.

## Message limit

`MessageNumberLimit` caps the number of messages. **Do not set a limit**: without one the device uses the number of vectors it needs (the maximum). In DEVICE TWEAKER that is **MSI Limit: Unlimited**, which deletes the value from the registry. The real vector count is visible in `msinfo32` → Hardware Resources → IRQs: count your device's IRQs (MSI ones are negative).

## Where it is stored

```text
HKLM\SYSTEM\CurrentControlSet\Enum\PCI\<device>\<instance>\Device Parameters\Interrupt Management\MessageSignaledInterruptProperties
  MSISupported         REG_DWORD   1 = MSI / MSI-X, 0 = line
  MessageNumberLimit   REG_DWORD   removed for Unlimited
```

The setting takes effect after a restart. A driver may ignore it: NVIDIA, for example, enables MSI selectively by architecture.

## Registry setting and real mode are different things

`MSISupported` is only a request. What the device actually uses DEVICE TWEAKER shows separately:

- **MSI Mode** — the stored setting (`MSISupported`).
- **IRQ Mode** — the current mode: Line / MSI / MSI-X, taken from the resources Windows assigned or through an already loaded DTIMOD.sys. If the resources cannot tell MSI from MSI-X, it shows `MSI/MSI-X`.
- **HW** — what the device itself supports, without the parent controller's capabilities.
- **READ IRQ MODE** at the top reads the `MSI Enable` / `MSI-X Enable` bits in PCI configuration space for every listed device through DTIMOD.sys, loading the driver if needed. No settings are changed.

MSI Mode and IRQ Mode may differ until a restart or when the device driver decides so. The IRQ count from WMI does not determine the mode by itself.

## What DEVICE TWEAKER does

- **AUTO-OPTIMIZATION** enables MSI and removes the limit on the devices it configures, storage included.
- MSI Mode and MSI Limit can be set manually on the device card; a general reset keeps MSI Mode and MSI Limit.

**Conclusion:** MSI MODE on, limit empty, real mode confirmed with **READ IRQ MODE** after a restart.

---

[← Chapter 01](./01-interrupts.en.md) · [Contents](./README.en.md) · [Chapter 03 →](./03-affinity.en.md)
