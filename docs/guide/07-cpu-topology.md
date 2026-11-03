# CPU topology and ReservedCpuSets

[Documentation](../README.md) / [Guide](./README.md) / 07 · [Русский](./07-cpu-topology.ru.md)

## Reading the CPU list

Cards show logical CPUs, physical cores, SMT siblings and available P/E-core, CCD/CCX and CPPC information. For example, `CPU 6 (P-Core, R131, #4)` and `CPU 7 (P-Core/HT, R131, #4)` belong to one physical core.

A CPPC rating is relative firmware guidance, not an application benchmark. CCD/CCX and cache information describe topology; moving work between them has no single fixed cost across processors and workloads.

## How AUTO uses topology

AUTO selects a target CCD and physical cores using project rules, SMT relationships and available ratings. A GPU receives a physical-core pair; other roles receive their own assignments. Unreliable detection of required topology blocks the relevant automatic write.

AUTO changes device settings while leaving SMT, E-cores and game-thread affinity unchanged. Core assignments depend on the detected topology; compare the resulting performance under your workload.

## ReservedCpuSets

The separate panel manages this value:

```text
HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Kernel
  ReservedCpuSets  REG_BINARY
```

Edits remain pending until **APPLY**, after a backup. A failed write preserves the pending state. The panel is unavailable without reliable CPU Set IDs.

Do not equate this value with per-process CPU Set assignment APIs. The [CPU Sets documentation](https://learn.microsoft.com/en-us/windows/win32/procthread/cpu-sets) describes IDs, affinity constraints and reservation policies; it does not establish isolation of all ISR/DPC activity through this registry setting.

Reservation does not prove that system events disappeared from a CPU or that a game automatically received the intended cores. Check scheduling and interrupts on your Windows build, preserve the original mask and verify post-reboot behavior.

## Practical verification

Confirm how selected CPUs map to physical cores, then compare ISR/DPC and application-thread placement. Two SMT siblings are not two independent physical cores. Measure contention under your own workload.

---

[Previous: Networking, RSS and NIC ITR](./06-network.md) · [Guide](./README.md) · [Next: Input and Polling Rate](./08-input.md)
