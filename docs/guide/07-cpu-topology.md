# CPU topology and ReservedCpuSets

[Documentation](../README.md) / [Guide](./README.md) / 07 · [Русский](./07-cpu-topology.ru.md)

## Reading the CPU list

Cards show logical CPUs, physical cores, SMT siblings and available CCD/CCX and CPPC information. Homogeneous CPUs use Core and Core/SMT labels; hybrid topologies use P-Core and E-Core labels. For example, `CPU 6 (Core, R131, #4)` and `CPU 7 (Core/SMT, R131, #4)` belong to one physical core.

CPPC ratings describe the firmware's relative core-performance ranking. CCD/CCX labels identify chiplets and core groups; cache information helps distinguish their roles on multi-CCD processors.

## How AUTO uses topology

AUTO selects a target CCD and physical cores using project rules, SMT relationships and available ratings. A discrete GPU receives a physical-core pair; integrated-GPU affinity returns to Windows. Other roles receive their own assignments. Unreliable detection of required topology blocks the relevant automatic write.

AUTO changes device settings while leaving SMT, E-cores and game-thread affinity unchanged. Core assignments depend on the detected topology; compare the resulting performance under your workload.

## ReservedCpuSets

The **Reserved CPU Sets** panel selects CPUs that Windows should try to leave aside during ordinary thread scheduling. Checking a CPU in this panel reserves it; checking a CPU in a device card assigns that device's interrupt affinity. These are separate settings.

The reservation mask is stored here:

```text
HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Kernel
  ReservedCpuSets  REG_BINARY
```

Choose the CPUs in the panel and click **APPLY**. The program creates a backup before writing the mask; restart Windows to check its effect. If writing fails, the edits remain available for another attempt. The panel is disabled when CPU Set IDs cannot be determined reliably.

This system-wide mask does not assign the selected CPUs to a particular game. Check actual thread scheduling and ISR/DPC activity after rebooting, and keep the original mask for restoration. The [CPU Sets API documentation](https://learn.microsoft.com/en-us/windows/win32/procthread/cpu-sets) describes per-process assignments and CPU Set IDs; those APIs are distinct from this registry setting.

## Checking the selection

Use core and CCD/CCX labels to check which physical cores are selected. The [verification chapter](./10-verification.md) describes how to inspect thread scheduling and interrupt processing after rebooting.

---

[Previous: Networking, RSS and NIC ITR](./06-network.md) · [Guide](./README.md) · [Next: Input and Polling Rate](./08-input.md)
