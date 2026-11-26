# CPU topology and ReservedCpuSets

[Documentation](../README.md) / [Guide](./README.md) / 07 · [Русский](./07-cpu-topology.ru.md)

## Reading the CPU list

Use the CPU list to see which logical processors share a physical core. Where available, the labels include CCD, CCX and CPPC information. On homogeneous CPUs, look for Core and Core/SMT. On hybrid CPUs, P-Core and E-Core distinguish core types. For example, `CPU 6 (Core, R131, #4)` and `CPU 7 (Core/SMT, R131, #4)` belong to one physical core.

CPPC ratings rank cores by the firmware's estimate of their performance. CCD and CCX labels identify chiplets and core groups. Cache information helps distinguish CCDs on multi-chiplet processors.

## How AUTO uses topology

AUTO selects a target CCD and physical cores using its topology rules, SMT relationships and available ratings. It assigns two physical cores to a discrete GPU. Windows chooses the cores for integrated graphics. Other roles receive their own assignments. Unreliable detection of required topology blocks the relevant automatic write.

AUTO changes device settings while leaving SMT, E-cores and game-thread affinity unchanged. Core assignments depend on the detected topology; compare the resulting performance under your workload.

## ReservedCpuSets

Use **Reserved CPU Sets** to reserve logical CPUs at the system level. Windows can still use them when required, for example when a thread's affinity restricts it to those CPUs. The CPU checkboxes in a device card configure its interrupt affinity.

The reservation mask is stored here:

```text
HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Kernel
  ReservedCpuSets  REG_BINARY
```

Choose the CPUs and click **APPLY**. Clearing all selections and applying removes the registry value. The program creates a backup before writing. If writing fails, the edits remain pending. The panel is available when CPU Set IDs can be determined reliably.

**AUTO-OPTIMIZATION** clears the reservation when this panel is available. To keep a custom reservation, use **APPLY** to configure device settings manually.

For the CPU Sets and thread-affinity model, see *Windows Internals*, 7th ed., Part 1, pp. 278–283, and the [CPU Sets API documentation](https://learn.microsoft.com/en-us/windows/win32/procthread/cpu-sets).

## Checking the selection

Use the core and CCD/CCX labels to confirm the selection. After rebooting, check where threads and interrupt handlers ran using the [trace procedure](./10-verification.md#inspecting-a-trace).

---

[Previous: Networking, RSS and NIC ITR](./06-network.md) · [Guide](./README.md) · [Next: Input and Polling Rate](./08-input.md)
