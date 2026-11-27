# Affinity, policies and interrupt priority

[Documentation](../README.md) / [Guide](./README.md) / 03 · [Русский](./03-affinity.ru.md)

## Why interrupt placement matters

By default, Windows concentrates a large share of system timers (the clock owner on CPU 0), interrupts and DPCs on the first cores, while **Interrupt Steering** can dynamically move steerable interrupts away from parked cores (*Windows Internals*, 7th ed., Part 2, ch. 8). Moving heavy device interrupts (GPU, xHCI, audio, wired NIC) onto dedicated physical cores reduces direct IRQL preemption of latency-sensitive application threads.

When choosing cores manually, read the topology labels carefully:

- **SMT siblings** share execution units and L1/L2 caches on the same physical core. Pinning a heavy device to the SMT sibling of a busy game thread still creates contention.
- **Hybrid CPUs** distinguish P-cores from E-cores.
- **Multi-CCD processors** separate L3 cache domains across chiplets; keep related cores within the intended CCD.

See [Chapter 7](./07-cpu-topology.md) for reading the processor map. For manual placement, select the target CPUs, choose **SpecCPU**, click **APPLY**, and reboot. AUTO applies topology-aware rules; [GPU](./04-gpu.md) and [network](./06-network.md) devices follow dedicated policies.

## Policies (`IRQ_DEVICE_POLICY`)

- **SpecCPU** (`IrqPolicySpecifiedProcessors`, value `4`): pins interrupts strictly to the logical CPUs checked in the card.
- **MachineDefault** (`IrqPolicyMachineDefault`, value `0`): hands placement back to Windows. On standard devices, selecting **MachineDefault** (or clearing all CPU checkboxes) and clicking **APPLY** removes the manual override. On NDIS adapters, use the **RSS / IRQ / BOTH** selector instead.
- **AllClose** (`IrqPolicyAllCloseProcessors`): targets processors on the device's NUMA node (equivalent to **All** on single-node desktops).
- **Single** (`IrqPolicyOneCloseProcessor`): requests a single processor close to the device; on non-NUMA systems Windows may pick any CPU.
- **All** (`IrqPolicyAllProcessorsInMachine`): permits assignment across all processors in the machine.
- **SpreadMessages** (`IrqPolicySpreadMessagesAcrossAllProcessors`): asks Windows to spread multiple MSI/MSI-X messages across distinct processors, using the vectors allocated to the driver.

Selecting any Windows-managed policy clears manual CPU checkmarks in the card. See *Windows Internals*, 7th ed., Part 2, pp. 52–54 (Table 8-5), and [IRQ_DEVICE_POLICY](https://learn.microsoft.com/en-us/windows-hardware/drivers/ddi/wdm/ne-wdm-_irq_device_policy).

## Affinity mask and registry format

DEVICE TWEAKER computes the 64-bit bitmask from the checked logical CPUs in processor group 0 (bit `N` = logical CPU `N`). For example, CPUs 8 and 9 yield `0x300`.

Settings are stored under the device's hardware key:

```text
Device Parameters\Interrupt Management\Affinity Policy
  DevicePolicy           REG_DWORD
  AssignmentSetOverride  REG_BINARY
  DevicePriority         REG_DWORD
```

In little-endian `REG_BINARY`, `0x300` is stored as `00 03 00 00 00 00 00 00` (when reading existing registry data, DEVICE TWEAKER also accepts shorter 1–7 byte little-endian masks written by other tools). See [Interrupt affinity and priority](https://learn.microsoft.com/en-us/windows-hardware/drivers/kernel/interrupt-affinity-and-priority).

## IRQ Priority

`DevicePriority` (**Undefined**, **Low**, **Normal**, **High**) is a hint used by the PnP interrupt arbiter when assigning IDT vectors and their corresponding **DIRQL** on x64 (*Windows Internals*, 7th ed., Part 2, ch. 8, Table 8-6):

| Setting | Target DIRQL (x64) |
|---|---|
| `Low` (`IrqPriorityLow`) | 3–4 |
| `Normal` / `Undefined` | 5–11 (assigned by vector) |
| `High` (`IrqPriorityHigh`) | 12 |

Windows is not a real-time OS: `DevicePriority` affects **only the hardware IRQL (ISR preemption order among devices)**. It does **not** change DPC priority (standard DPCs always run at `DISPATCH_LEVEL`, IRQL 2) and does not boost user-mode threads.

AUTO sets **High** on devices it actively places, leaves integrated graphics and excluded devices at **Undefined**, and preserves the existing priority on display audio. See [IRQ_PRIORITY](https://learn.microsoft.com/en-us/windows-hardware/drivers/ddi/wdm/ne-wdm-_irq_priority).

## Storage controllers and verification

For NVMe/SATA storage controllers, Windows and the storage miniport manage multi-queue CPU distribution directly, so storage cards expose MSI and IRQ Priority rather than manual CPU masks.

After rebooting, verify the applied policy in the log and inspect an ETW trace under load. Compare where the driver's ISRs and DPCs actually execute against your application's busiest threads (for instance, checking whether moving a noisy device off CPU 0 eliminates overlap with the main game thread). See [Chapter 10](./10-verification.md).

---

[Previous: MSI and MSI-X](./02-msi.md) · [Guide](./README.md) · [Next: Graphics card](./04-gpu.md)
