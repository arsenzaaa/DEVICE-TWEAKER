# Affinity, policies and interrupt priority

[Documentation](../README.md) / [Guide](./README.md) / 03 · [Русский](./03-affinity.ru.md)

## What Interrupt Affinity changes

A device policy controls the processor assignment of its interrupts. **SpecCPU** uses the selected mask; it does not pin every DPC and driver thread to those CPUs. Microsoft recommends the default policy where appropriate. See [Interrupt affinity](https://learn.microsoft.com/en-us/windows-hardware/drivers/kernel/interrupt-affinity-and-priority).

| Application | Code | CPU selection |
| :--- | :--- | :--- |
| MachineDefault | 0 | Windows default |
| AllClose | 1 | Close to device |
| Single | 2 | One close CPU |
| All | 3 | All available |
| SpecCPU | 4 | Selected mask |
| SpreadMessages | 5 | Different CPUs for messages |

See [IRQ_DEVICE_POLICY](https://learn.microsoft.com/en-us/windows-hardware/drivers/ddi/wdm/ne-wdm-_irq_device_policy). SpreadMessages spreads existing messages when possible; it does not create extra vectors.

## Mask and persistence

```text
...\Device Parameters\
  Interrupt Management\Affinity Policy
  DevicePolicy           REG_DWORD
  AssignmentSetOverride  REG_BINARY
  DevicePriority         REG_DWORD
```

In DEVICE TWEAKER, mask bit N represents logical CPU N in supported processor group 0. CPUs 8 and 9 give `0x300`, stored as `00 03 00 00 00 00 00 00`. The application calculates the mask from the selected CPUs. Clearing the selection for an ordinary device restores MachineDefault and removes the override. Network devices have separate [RSS/IRQ modes](./06-network.md).

Saved policy and actual processing placement can differ. Compare readback and an ETW trace after rebooting.

## IRQ Priority

Undefined, Low, Normal and High express interrupt servicing preference. Windows considers it when assigning IRQLs for the platform and device resources. This is neither application thread priority nor DPC queue priority. See [IRQ_PRIORITY](https://learn.microsoft.com/en-us/windows-hardware/drivers/ddi/wdm/ne-wdm-_irq_priority).

AUTO uses High for the devices it configures. Evaluate the effect under your workload.

## Unexpected results

Check the device instance, policy, mask, reboot and current interrupt mode. Then distinguish ISR, DPC and worker-thread activity in a trace: the driver can distribute them differently. A single busy CPU under a two-CPU mask does not by itself establish a fault.

Preserve settings, versions and the trace for further diagnosis. Firmware management is outside DEVICE TWEAKER; examine interrupt behavior through Windows resources and the driver.

AUTO leaves storage-device affinity to Windows. Manual changes should be recoverable experiments with boot, I/O and stability checks; another device's profile is not a universal setting.

---

[Previous: MSI and MSI-X](./02-msi.md) · [Guide](./README.md) · [Next: Graphics card](./04-gpu.md)
