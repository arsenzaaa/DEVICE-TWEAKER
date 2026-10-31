# Interrupts: ISR, DPC and cores

## From the device to the game

A device reports new data with an interrupt (IRQ). The chain is always the same:

```text
IRQ → driver ISR → driver DPC → thread that processes the data → application thread
```

- **ISR** (Interrupt Service Routine) runs at the device IRQL (3–12) and is kept as short as possible: take the data from the hardware and queue a DPC.
- **DPC** (Deferred Procedure Call) runs at DISPATCH_LEVEL (IRQL 2). The driver does its main work here and signals the threads that data is ready.
- **Thread** — an ordinary Windows thread (priority 0–31) continues the work. Waking it on another core needs an IPI.

Input as an example (WDF01000 handles all input devices):

```text
Mouse, Windows 10       WDF01000 ISR → WDF01000 DPC → CSRSS (Win32k Raw Input Thread)   → game input thread
Mouse, Windows 11       WDF01000 ISR → WDF01000 DPC → DWM KERNEL SENSOR THREAD          → game input thread
Keyboard, 10 / 11       WDF01000 ISR → WDF01000 DPC → CSRSS (winsrvext.dll)             → game input thread
Controller, Windows 11  WDF01000 ISR → WDF01000 DPC → DWM MASTER INPUT THREAD           → game input thread
```

The more threads a DPC signals, the longer the DPC itself takes (WDF01000 especially), and the longer the threads that receive the signal run.

## Why ISRs and DPCs outrank any thread

1. The clock interrupt, ISRs and DPCs (in that order) run **above any thread**, even at priority 31.
2. **ISRs and DPCs do not know which cores are busy with threads.** On 4 cores:

   ```text
   core 0 - thread at priority 6
   core 1 - thread at priority 6
   core 2 - thread at priority 6
   core 3 - idle
   ```

   An incoming ISR or DPC may well interrupt a thread on cores 0–2 instead of taking the idle core 3. Scheduler rules do not apply to ISRs and DPCs.
3. **An interrupted thread waits on the same core.** It does not move to a free core; it waits until the ISR + DPC + context switches are done (one switch costs roughly 3–5 µs).
4. **A DPC interrupted by an ISR waits on its core too.** Any ISR can interrupt any other ISR and any DPC regardless of IRQ priority.
5. **SMT rules do not apply to ISRs and DPCs.** Two drivers on one SMT-SET (two logical processors of one physical core) share cache and EXECUTION-UNITS and slow each other down.

## What it means for a game

A game has several important threads: GAME (main), RENDER, and RHI on Unreal Engine. If a driver sits on the core where such a thread runs:

- the thread is interrupted → it runs longer → frame time grows;
- or the thread cannot start on its IDEAL PROCESSOR (a soft mask) and moves to another core → it loses its L1 / L2 and pollutes another core's cache → again it runs longer;
- with many drivers on one core, a stutter is almost guaranteed whenever an important thread lands there.

A telling test: 7800X3D + RTX 5070, CPU load 30–36 % (no CPU-BOUND), GPU 60–70 %, background fully tuned, HT on. With drivers spread across cores and the game's main thread on a driver-free core (its SMT-SET free as well), runs are nearly identical. With all drivers on core 0 and the main thread's IDEAL PROCESSOR there too, results are random and on average **40–60 FPS lower** — without any CPU-BOUND. Under CPU-BOUND and with background load the effect only grows.

DEVICE TWEAKER does not always change FPS: not when the game is GPU-BOUND, not when its important threads happen to avoid driver cores, not when the CPU has plenty of cache. The goal is to remove the chance of such collisions instead of relying on luck.

## ISR and DPC on the same core

A DPC usually starts at the end of its ISR or right after it. Running the DPC on another core needs an IPI (costing about the core-to-core latency) and refetching data from the ISR core's cache. Splitting one driver's ISR and DPC across cores gains nothing. ISRs and DPCs also benefit from cache when they always run on one core with little other work in between.

## What DEVICE TWEAKER does

- Shows each device's IRQ Mode, MSI mode, priority and mask, and for every logical processor its physical core, P / E-core, CCD and CPPC rank.
- **AUTO-OPTIMIZATION** spreads drivers across physical cores so they share neither cores nor SMT-SETs, and avoids core 0 when possible.
- Saves a verified backup before writing, and logs what was applied, what was skipped and why.

**Conclusion:** drivers belong on their own physical cores, apart from each other and from the game's important threads. The next chapters show how to get there for each device type. Verifying the result with a trace is covered in chapter [10](./10-verification.en.md).

---

[Contents](./README.en.md) · [Chapter 02 →](./02-msi.en.md)
