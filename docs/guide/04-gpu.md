# Graphics card

[Documentation](../README.md) / [Guide](./README.md) / 04 · [Русский](./04-gpu.ru.md)

## Why AUTO assigns two physical cores to a discrete GPU

Graphics stacks (`dxgkrnl.sys` together with `nvlddmkm.sys` or `amdkmdag.sys`) generate frequent, bursty interrupts and DPCs for VSync, DMA fence completions and GPU scheduling. Pinning a discrete GPU to a **single** logical CPU frequently degrades frame-time consistency: in WPA/MXA traces, incoming GPU ISRs preempt the driver's own in-flight DPC on that core and stall subsequent work.

To avoid self-contention while keeping GPU interrupts localized within the same L3 cache domain, **AUTO-OPTIMIZATION** picks a pair of **neighboring physical cores** in the preferred topology:

- E-cores are excluded on hybrid Intel processors.
- The physical core hosting CPU 0 is avoided whenever another suitable pair exists, because CPU 0 carries the heaviest baseline share of clock timers and unsteered system DPCs.
- AUTO enables MSI, clears any custom `MessageNumberLimit`, sets **High** IRQ Priority (`DIRQL 12` hint), and writes **SpecCPU** for the selected pair. If no valid pair is available, the affinity write is skipped rather than forcing a single-core pin.
- Integrated GPUs are left on Windows default interrupt routing.

Keep in mind that `AssignmentSetOverride` defines which CPUs are *eligible* to receive the GPU's interrupts. How ISR and DPC execution actually spreads across those cores depends on the allocated MSI/MSI-X vectors and how the vendor driver schedules its DPCs and system threads.

## Manual configuration

The GPU card exposes **CPU Affinity**, **MSI Mode / MSI Limit**, **IRQ Priority** and **Policy**. After clicking **APPLY**, reboot Windows and use **READ IRQ MODE** to confirm that the driver actually came up in MSI or MSI-X mode.

## Verification

Compare frame times (including 1% / 0.1% lows and frame-time variance) alongside `dxgkrnl` / vendor-driver ISR and DPC activity in an ETW trace. Keep the scene, graphics preset, resolution and FPS cap identical across runs; see [Chapter 10](./10-verification.md).

---

[Previous: Affinity and priority](./03-affinity.md) · [Guide](./README.md) · [Next: USB and IMOD](./05-usb-imod.md)
