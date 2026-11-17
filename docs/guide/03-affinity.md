# Affinity, policies and interrupt priority

[Documentation](../README.md) / [Guide](./README.md) / 03 · [Русский](./03-affinity.ru.md)

## Choosing CPUs

Interrupt Affinity selects the logical CPUs allowed to service a device's interrupts. Moving that work away from a busy game core can reduce competition for CPU time. Start with the device and workload you want to tune, then compare the result with the original configuration.

Use the topology labels in the CPU list: two SMT siblings share a physical core. On a hybrid processor, distinguish P-cores and E-cores; on a multi-CCD processor, check which CCD contains the selected cores. [Chapter 7](./07-cpu-topology.md) explains the labels.

For manual selection, tick the required CPUs and use **SpecCPU**. Click **APPLY** to save the policy and mask, then restart Windows. AUTO can choose cores using the detected topology; [GPU](./04-gpu.md) and [network](./06-network.md) settings have their own allocation rules.

## Policies

**SpecCPU** uses the CPUs selected in the card. **MachineDefault** returns processor selection to Windows. Clearing the CPU selection on an ordinary device removes its affinity override when applied; network cards have separate RSS/IRQ controls.

The other policies ask Windows to choose processors according to a rule:

- **AllClose** uses processors close to the device in the NUMA topology. On a non-NUMA system it behaves like **All**.
- **Single** requests one close processor; on a non-NUMA system Windows can choose any processor.
- **All** allows assignment across the machine's processors.
- **SpreadMessages** requests distribution of different interrupt messages across processors where possible. It uses the messages the device and driver allocate.

See [IRQ_DEVICE_POLICY](https://learn.microsoft.com/en-us/windows-hardware/drivers/ddi/wdm/ne-wdm-_irq_device_policy) for Windows policy definitions.

## Affinity mask

The program calculates the mask from the checked CPUs. Bit N represents logical CPU N in processor group 0, up to 64 logical processors. For example, CPUs 8 and 9 produce `0x300`.

The saved values are under the device's registry key:

```text
Device Parameters\Interrupt Management\Affinity Policy
  DevicePolicy           REG_DWORD
  AssignmentSetOverride  REG_BINARY
  DevicePriority         REG_DWORD
```

`AssignmentSetOverride` stores that example as `00 03 00 00 00 00 00 00` in little-endian order. See [Interrupt affinity](https://learn.microsoft.com/en-us/windows-hardware/drivers/kernel/interrupt-affinity-and-priority) for the registry format.

## IRQ Priority

**Undefined**, **Low**, **Normal** and **High** express a preference for interrupt servicing. Windows considers it when assigning IRQLs for the platform and device resources. AUTO uses **High** for devices it configures. This setting does not change the game's thread priority or a DPC's queue priority. See [IRQ_PRIORITY](https://learn.microsoft.com/en-us/windows-hardware/drivers/ddi/wdm/ne-wdm-_irq_priority).

## Check the result

After rebooting, compare the card's saved policy and mask with the log. Use an ETW trace to see where the driver's ISR, DPCs and worker threads actually run: a driver can target them separately. [Chapter 10](./10-verification.md) describes this check.

The mask defines the CPUs available for interrupts; the device driver determines how work is distributed within that set.

AUTO leaves storage-device affinity to Windows. For manual storage changes, keep a recovery backup and check boot, I/O and stability.

---

[Previous: MSI and MSI-X](./02-msi.md) · [Guide](./README.md) · [Next: Graphics card](./04-gpu.md)
