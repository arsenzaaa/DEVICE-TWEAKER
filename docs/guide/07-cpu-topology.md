# CPU topology and ReservedCpuSets

[Documentation](../README.md) / [Guide](./README.md) / 07 · [Русский](./07-cpu-topology.ru.md)

## Reading the processor map

DEVICE TWEAKER queries Windows topology (`GetSystemCpuSetInformation` and processor-relationship data) to map logical processors onto physical cores, cache domains and firmware performance ranks:

- **Homogeneous CPUs** label the two logical processors of one physical core as `Core` and `Core/SMT` (for example, `CPU 6 (Core, R131, #4)` and `CPU 7 (Core/SMT, R131, #4)` share the same physical core).
- **Hybrid CPUs** distinguish `P-Core` and `E-Core` via `EfficiencyClass` and SMT pairing.
- **CPPC ranks** reflect the firmware's estimate of each core's preferred boost capability.
- **CCD / CCX and L3 cache domains** (`LastLevelCacheIndex`) separate chiplets on multi-CCD processors, where crossing between CCDs incurs inter-chiplet latency.

Understanding SMT matters differently for threads versus interrupts. The Windows thread scheduler already prioritizes an idle physical core over a busy SMT sibling (*Windows Internals*, 7th ed., Part 1, ch. 4), so normal threads naturally spread across physical cores while idle cores exist. Hardware interrupts and DPCs do **not** follow those SMT-avoidance rules: an interrupt fires on its assigned logical CPU regardless of what is running on its sibling thread.

## How AUTO uses topology

**AUTO-OPTIMIZATION** chooses a target CCD and allocates physical cores using SMT relationships, core classes and CPPC ranks: two neighboring physical cores for a discrete GPU, Windows defaults for integrated graphics, and separate placement rules for input, wired networking and audio. If required topology metadata cannot be read reliably, AUTO skips the affected automatic affinity writes rather than guessing.

AUTO modifies only device settings. It never disables SMT, parks/unparks E-cores or alters game-thread affinity.

## ReservedCpuSets

**Reserved CPU Sets** applies a system-wide core reservation via the Windows **CPU Sets** mechanism. Unlike traditional hard affinity masks (`SetProcessAffinityMask`), CPU Sets act as soft scheduling boundaries that **also constrain kernel and system threads**: reserved logical processors are marked `Allocated` in `SYSTEM_CPU_SET_INFORMATION`, kept clear of general background threads, and kept unparked. Windows will still run a thread on a reserved CPU if that thread's explicit affinity or CPU Set targets it.

The reservation bitmask is stored in the registry as little-endian `REG_BINARY`, padded to a multiple of 8 bytes (one 64-bit `KAFFINITY` mask per processor group, as required by `ntoskrnl`):

```text
HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Kernel
  ReservedCpuSets  REG_BINARY
```

Select the logical CPUs to reserve and click **APPLY** (clearing all selections and applying deletes the registry value). DEVICE TWEAKER enables the panel only when CPU Set IDs can be mapped reliably, backs up the value before writing, and leaves edits pending if a write fails.

> **Avoid mixing reserved and unreserved CPUs in a single process affinity mask.** As noted in Windows scheduling behavior and *Windows Internals* (7th ed., Part 1, pp. 278–283), forcing a hard affinity mask that spans both reserved and unreserved processors can cause erratic thread placement. Keep ordinary processes on the unreserved pool and target reserved cores explicitly only for dedicated workloads.

**AUTO-OPTIMIZATION** clears `ReservedCpuSets` whenever the panel is available. If you want to keep a custom reservation alongside device tuning, apply device changes manually via **APPLY**.

See *Windows Internals*, 7th ed., Part 1, pp. 278–283, and the [CPU Sets documentation](https://learn.microsoft.com/en-us/windows/win32/procthread/cpu-sets).

## Verification

Confirm that your selected logical processors match the intended physical cores and CCD/CCX labels. After rebooting, verify actual thread and ISR/DPC placement in an ETW trace ([Chapter 10](./10-verification.md#inspecting-a-trace)).

---

[Previous: Networking, RSS and NIC ITR](./06-network.md) · [Guide](./README.md) · [Next: Input and Polling Rate](./08-input.md)
