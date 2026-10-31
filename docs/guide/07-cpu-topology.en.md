# CPU topology and ReservedCpuSets

## Reading the CPU list

Every device card in DEVICE TWEAKER lists all logical processors with their place in the topology:

```text
CPU 6  (P-Core, R131, #4)          P-core, CPPC rank 131, physical core #4
CPU 7  (P-Core/HT, R131, #4)       the second thread of the same core (SMT / HT)
CPU 16 (E-Core, R70, #9)           E-core
CPU 8  (P-Core, CCD0, R136, #5)    on AMD with two CCDs the CCD is shown
```

- **SMT-SET** — the two logical processors of one physical core. They share L1 / L2 and EXECUTION-UNITS: load on one slows the other. Drivers and important threads must not share an SMT-SET.
- **CPPC rank** (`R`) — the firmware's rating of a core: higher is faster. `Pref` marks preferred cores.
- **CCD / CCX** — AMD CPUs with two CCDs have no shared L3 between CCDs, and core-to-core latency between them is much higher. A thread must not hop between CCDs; related threads belong on one CCD.
- **E-cores** on Intel are slower and have higher latency when exchanging data with P-cores.

The status line at the top of the window shows Hyper-Threading / SMT, Hybrid CPU, CPPC and Dual-CCD.

## Hyper-Threading / SMT

When one logical processor is loaded, its sibling slows down. If a game has enough physical cores, SMT can lower FPS: important threads share SMT-SETs with background work, and SMT rules do not apply to ISRs / DPCs at all. But if the game is heavily threaded (GAME / RENDER / RHI plus many BACKGROUND / FOREGROUND WORKERS) and FPS drops without SMT, keep it on. DEVICE TWEAKER works in both modes: with SMT it simply never counts siblings as two cores.

## ReservedCpuSets

`ReservedCpuSets` reserves logical processors: Windows does not use them for ordinary thread scheduling, and driver ISRs / DPCs without an explicit mask do not land there. On Windows 11 clock interrupts and system events remain on reserved cores (they take under 1 ns and cost practically nothing).

```text
HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\kernel
  ReservedCpuSets   REG_BINARY   CPU mask, least significant byte first
```

Example: 12 logical processors, reserving CPUs 1–11 → `1111 1111 1110` = `0xFFE` → in REG_BINARY the least significant byte first: `FE 0F`.

Know whom the reservation protects. If you reserve a core where a masked driver runs, the driver keeps running there: **threads are protected from that driver, not the other way round**. Reservation is not hard isolation and not a substitute for device Interrupt Affinity. Never give a process an affinity that spans both reserved and unreserved CPUs: the behaviour is unpredictable.

## Two CCDs and E-cores

This is where reservation pays off. Two approaches:

```text
Option 1   reserve the P-cores / CCD0, put drivers on the E-cores / CCD1,
           move the game and important system threads manually onto reserved cores without drivers
Option 2   reserve the E-cores / CCD1 and move all background work there manually,
           drivers and the game go on the P-cores / CCD0 the usual way
```

Option 1 leaves storage interrupts and DPCs such as KBDHID on the slow cores, and the game's threads must be placed after every restart. Option 2 requires moving all background work by hand after every restart. On an X3D keep the game on the CCD with 3D V-Cache; drivers do not need much cache.

## What DEVICE TWEAKER does

- Detects P / E-cores, SMT / HT, CCD / CCX and CPPC rank, and shows which physical core each thread belongs to.
- **AUTO-OPTIMIZATION** chooses a target CCD and places drivers on physical cores by rank. If the topology cannot be detected reliably, nothing is written.
- ReservedCpuSets is set in its own block; the mask is built from the checked processors.

**Conclusion:** count physical cores, not logical ones. Drivers do not share SMT-SETs, threads do not hop between CCDs, and reservation is a tool for two CCDs and E-cores, not a way to "isolate a driver".

---

[← Chapter 06](./06-network.en.md) · [Contents](./README.en.md) · [Chapter 08 →](./08-input.en.md)
