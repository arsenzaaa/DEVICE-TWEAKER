# Graphics card

## Which drivers run

When you move a graphics card to a core, all of its interrupts move with it:

```text
NVIDIA   NVLDDMKM (ISR + DPC) and DXGKRNL (ISR + DPC)
AMD      DXGKRNL only (ISR + DPC)
```

AMD produces more DXGKRNL interrupts, but the total CPU load is lower than with NVIDIA, whose driver needs the ISRs / DPCs of two drivers. The NVIDIA driver also has NVLDDMKM system threads at priority 29; AMD has no such threads.

## Why two cores

On a single core, the graphics driver's ISR interrupts its own DPC:

```text
core N   DPC NVLDDMKM ... [ ISR NVLDDMKM ] ... DPC continues
```

The DPC cannot move to another core and waits for the ISR → latency spikes at random moments and broken parallelism. The "one core for cache" gain is measured in microseconds and is not worth those spikes. So the card goes on **two different physical cores**: two logical processors of one SMT-SET share cache and EXECUTION-UNITS and do not count as two cores.

Tests on an unstable benchmark easily "prove" that one core is better: under PBO a CPU has good and bad cores, and that difference gets mistaken for one core versus two. Check with a trace, not with the FPS of a single run.

## Rules

1. Two different physical cores free of other drivers.
2. Not core 0: do not put the card on "core 0 + core X".
3. If a trace shows the ISRs / DPCs still go to one core, set **Interrupt Redirection Mode Selection = Round Robin** in BIOS (a hidden option, via SCEWIN / GRUB). See chapter [03](./03-affinity.en.md).

## Integrated graphics

An integrated GPU does not need dedicated cores: AUTO-OPTIMIZATION returns its affinity to Windows and does not spend scarce physical cores on it.

## Around DEVICE TWEAKER

The program only configures interrupts. A few related things worth knowing:

- Keep **HAGS** on with NVIDIA: without it the NVLDDMKM system thread at priority 29 is constantly active (suspending it freezes the system), and DXGMMS2 is raised to 29. On AMD HAGS is always on.
- DXGMMS threads start rarely and run briefly; they can stay where they are, since free cores matter more for the game's threads.

## What DEVICE TWEAKER does

- **AUTO-OPTIMIZATION** gives a discrete GPU a pair of adjacent physical cores on the target CCD, never counting an SMT sibling as the second core and avoiding core 0 when possible. `SpecCPU` policy, High priority, MSI on, limit removed.
- The card's block shows IRQ Mode (usually MSI-X), HW and the mask; cores can be checked manually in the CPU Affinity list.

**Conclusion:** a discrete graphics card gets two physical cores with no other drivers and not core 0, verified with a trace, with Round Robin in BIOS if needed.

---

[← Chapter 03](./03-affinity.en.md) · [Contents](./README.en.md) · [Chapter 05 →](./05-usb-imod.en.md)
