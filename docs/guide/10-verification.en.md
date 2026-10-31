<!-- tag: LOG | summary: How to check the result by facts rather than the FPS of one run: an xperf / WPR trace, where to look at ISRs and DPCs in WPA and MXA, what to send in a hardware report and how to read the program log. -->

# Verification: traces and logs

## Why not FPS

The FPS of a single run is a poor judge. It depends on the bound (CPU / RAM / GPU), on where the game's important threads happened to land, on good and bad cores under PBO. A before / after test on an unstable benchmark can show anything. Check what the program actually changes: **which cores ISRs and DPCs run on and how long they take**.

## Recording a trace

You need the Windows Performance Toolkit from the [Windows ADK](https://learn.microsoft.com/en-us/windows-hardware/get-started/adk-install). In an elevated command prompt:

```bat
xperf -on PROC_THREAD+LOADER+INTERRUPT+DPC
rem 10 seconds of normal load: the game, mouse movement, audio
xperf -stop -d C:\kernel.etl
xperf -i C:\kernel.etl -o C:\report.txt -a dpcisr
```

`report.txt` is a text summary: for every driver the ISR / DPC count, the duration distribution and **the distribution by CPU**. It shows at once whether a driver moved to the intended cores.

The same can be recorded with WPR (`wpr -start CPU -start GPU`, then `wpr -stop trace.etl`) and opened in WPA or Media eXperience Analyzer (MXA).

## What to look at

- **The CPU each driver's ISRs / DPCs run on**: NVLDDMKM and DXGKRNL on the graphics card's two cores, WDF01000 on the USB controller cores, NDIS on the network core.
- **ISR and DPC of one driver on one core.** A DPC on another core means IPIs.
- **Interrupted DPCs.** Without MSI, or with two drivers on one core, you can see a DPC interrupted by foreign ISRs.
- **Duration.** WDF01000 DPCs are usually 8–16 µs; growth after a change is a reason to investigate.
- In MXA: `CPU > ISRs and DPCs > Drivers`, drag a driver onto the panel and look per core.

If a mask is set but the ISRs / DPCs still stay on one core, it is not a program error but the interrupt redirection mode: chapter [03](./03-affinity.en.md), Round Robin in BIOS.

## The program log

DEVICE TWEAKER writes a detailed log to the `logs` folder next to the EXE, one file per run:

- what was found: devices, USB roles, CPU topology;
- the AUTO-OPTIMIZATION plan: which cores were chosen and why, what was skipped and for what reason;
- what was applied, what was not, and the read-back result: IMOD / NIC ITR registers, power-saving WMI, active RSS;
- the source and time of each IRQ Mode read.

An unhandled error also writes a crash file and `last-crash.txt` there.

## Hardware report

Auto-optimization has been verified on specific systems, and every new configuration makes it more accurate. To help, send a [hardware report](https://github.com/arsenzaaa/DEVICE-TWEAKER/issues/new?template=hardware_report.yml) with three logs:

```text
1. before auto-optimization (or after a settings reset)
2. right after applying it
3. after a restart and opening the program again
```

Attach before / after traces if you recorded them: they are the most valuable part.

**Conclusion:** the result is verified with a trace — where ISRs / DPCs run and how long they take — not with a single benchmark run.
