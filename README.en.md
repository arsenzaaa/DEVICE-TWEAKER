<div align="center">

# DEVICE TWEAKER

Low-level Windows utility for hardware interrupt (MSI / MSI-X) configuration, topology-aware CPU queue steering (Interrupt Affinity), and physical timer moderation (xHCI IMOD / NIC ITR)

[Русский](./README.md) • **English**

</div>

<p align="center">
  <img src="./DEVICE%20TWEAKER/assets/screenshots/main_interface_en.png" alt="DEVICE TWEAKER Main Interface" width="100%">
</p>

## Overview

**DEVICE TWEAKER** is an advanced systems engineering tool for 64-bit Windows 10 and Windows 11. It provides deterministic control over hardware interrupt routing (MSI / MSI-X), steers DPC/ISR execution queues to dedicated physical cores (Interrupt Affinity) with full awareness of asymmetric CPU topologies (Intel Core Hybrid, AMD Multi-CCD), reprograms physical timer moderation registers (xHCI IMOD / NIC ITR), and isolates system resources from unpinned OS threads.

For direct physical MMIO register programming, the utility includes a custom signed Ring 0 kernel driver (`DTIMOD.sys`), fully compatible with Windows Core Isolation (HVCI / Memory Integrity) and kernel-level anti-cheat engines.

## Features

### Hardware Interrupts (MSI / MSI-X)

DEVICE TWEAKER migrates peripheral PCIe controllers from legacy line-based signaling (INTx pin assertion) to high-throughput Message Signaled Interrupts (MSI / MSI-X). In vector mode, interrupts are posted directly into the target CPU core's Local APIC address space (`0xFEE00000`) via PCIe Memory Write cycles:

- **Vector Mode Activation (`MSISupported = 1`):** moves USB host controllers, dedicated GPUs, NVMe drives, and NICs into MSI mode, eliminating shared IRQ line arbitration and latency spikes.
- **Vector Limit Unclamp (`MessageNumberLimit = 0` / Unlocked):** removes driver registry clamp values, enabling `pci.sys` to allocate the full hardware MSI-X vector allocation (up to 2048 vectors) requested by the peripheral controller.
- **HAL Interrupt Priority (`DevicePriority`):** raises dispatch queue priority for latency-sensitive devices (high-polling mouse, GPU render loop) inside the Windows HAL scheduler.

### Queue Steering across CPU Cores (Interrupt Affinity)

Standard Windows load balancing distributes DPC and ISR queues across logical processors without considering hybrid cores or multi-die topologies. DEVICE TWEAKER configures `DevicePolicy = 4` (`IrqPolicySpecifiedProcessors`) and the 64-bit `AssignmentSetOverride` affinity bitmask to precisely route interrupts:

#### Intel Core Hybrid Architecture (P-Cores vs E-Cores)

On Intel Core 12th-14th Gen and Core Ultra processors (Alder Lake, Raptor Lake, Arrow Lake), cores are partitioned into Performance Cores (P-Cores) and Efficient Cores (E-Cores, grouped into 4-core clusters sharing an L2 cache).

- **Hardware Behavior:** high-rate mouse packets (1000-8000 Hz), GPU render fences, or network queues scheduled onto slower E-cores trigger cross-ring synchronization penalties and frametime jitter.
- **Implementation:** DEVICE TWEAKER detects hybrid topology (`Hybrid CPU - On`), segregates core clusters, and provides a one-click **`[ P-CORES ]`** button to exclude E-cores and Hyper-Threading sibling threads from interrupt handling.

<p align="center">
  <img src="./DEVICE%20TWEAKER/assets/screenshots/showcase_intel_14900k_hybrid_en.png" alt="Intel Core i9-14900K Hybrid Topology" width="100%">
</p>

#### AMD Ryzen Multi-CCD Architecture (Dual-CCD / 3D V-Cache)

On multi-die AMD Ryzen processors (7900X, 7950X, 7950X3D, 9900X, 9950X, 9950X3D), cores reside across two physical dies (CCD0 and CCD1) linked over Infinity Fabric.

- **Hardware Behavior:** intra-die L3 cache access takes ~20 ns, whereas inter-die requests crossing Infinity Fabric incur a 70-90 ns roundtrip latency penalty. On X3D processors, the 96 MB 3D V-Cache is physically located on CCD0 only.
- **Implementation:** DEVICE TWEAKER flags multi-die topology (`Dual-CCD - True`), visually highlights secondary CCD1 cores, and anchors critical devices to the primary die with the **`[ CCD0 ]`** button, keeping execution caches warm inside local L3 / 3D V-Cache.

<p align="center">
  <img src="./DEVICE%20TWEAKER/assets/screenshots/showcase_gpu_affinity_en.png" alt="AMD Ryzen Dual-CCD GPU Affinity" width="100%">
</p>

#### Silicon Core Binning (ACPI CPPC2)

The application queries hardware silicon quality ratings natively via the Windows ETW event tracing infrastructure (provider `Microsoft-Windows-Kernel-Processor-Power`, Event ID 55, payload `MaximumPerformancePercent`). The query completes in under 60 ms and allows mapping mouse interrupt routines to the highest-binned core without third-party utilities.

### Hardware USB xHCI Timer Moderation (DTIMOD.sys)

Per Intel xHCI Specification Revision 1.2 (§5.5.2), each hardware interrupter implements a 32-bit `IMOD` register. The lower 16 bits (`IMODI`) specify interrupt moderation delay in 250-nanosecond hardware ticks. By default, the Windows `USBXHCI.SYS` driver imposes delays of 50 µs (`0xC8`) or 1 ms (`0xFA0`) to batch transfers.

- **Zero Moderation (`IMOD = 0` / 0x0):** the xHCI controller emits an interrupt onto the PCIe bus immediately upon completing an Event TRB, eliminating software buffering delay for mouse input.
- **Ring 0 Driver (`DTIMOD.sys`):** provides direct physical MMIO register access via `MmMapIoSpace`.
- **Controller Partitioning (CHIP 0 / CHIP 1):** the `UsbChipPath` engine analyzes PCIe bus topology to distinguish direct CPU-attached lanes (CHIP 0) from chipset hubs (CHIP 1 / PCH). This allows applying zero delay (`0x0`) to mouse input on direct CPU ports while maintaining 1 ms (`0xFA0`) buffers for chipset audio devices to prevent audio dropouts.

### Network Stack: NIC ITR Registers & RSS Queue Scaling

- **Direct Physical Timer Programming:** using `DTIMOD.sys`, the utility reads and updates physical hardware registers:
  - Intel PCIe (EITR / ITR): I225, I226, I210, I211, I350, I219, 82580, 82576, and Killer E3100 (MMIO offsets `0x1680` and `0x00C4` with 1-2 µs granularity).
  - Realtek PCIe (IntrMit / IntrMitV2): RTL8111, RTL8168, RTL8125, RTL8126, and Killer E2500/E2600 (registers `0x00E2` and `0x0A00` with per-queue mask `0x7F7F7F7F`).
- **Receive Side Scaling (RSS):** configures `*NumRssQueues` and `*RssBaseProcNumber` in the adapter driver key. Restricting queues to 2 and offsetting the base processing core from CPU 0 prevents network DPCs from preempting mouse input and graphics pipelines.
- **PCIe Bus Power Management:** disables device D3 sleep states (`PnPCapabilities` bit `0x08`, synchronized via WMI `MSPower_DeviceEnable`) and disables USB Selective Suspend.

### System Core Isolation (ReservedCpuSets)

Configures `[HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\kernel] ReservedCpuSets`. A 64-bit processor mask strips selected logical processors from the generic Windows thread scheduler pool (`KiSelectNextThread`). Background system services and general tasks will not run on isolated cores, reserving execution units and L1/L2 caches for dedicated hardware interrupts and pinned processes.

### Background Input Throttling (RawMouseThrottleDuration)

Configures `HKCU\Control Panel\Mouse\RawMouseThrottleDuration` (introduced in Windows 11 Build 22621.1928 / KB5028185). Allows setting raw input polling throttling (1 to 20 ms) for unfocused background windows, eliminating unnecessary CPU utilization from high-polling mice (1000-8000 Hz) in background apps (Discord, OBS, browser) while leaving active game input unaffected.

### System Checkpoints and Recovery

- **Immutable Baseline Snapshot:** on initial launch, the utility generates a complete baseline configuration snapshot (`DeviceTweakerBackup_ORIGINAL.json`) with overwrite protection.
- **Automatic Checkpoints:** an atomic checkpoint is taken before any modifications are committed (APPLY, reset, or IMOD/ITR writes). In the event of a write failure, parameters roll back automatically.
- **Single-Click Recovery:** restores settings to the most recent working checkpoint or returns the operating system to factory Windows defaults.

<p align="center">
  <img src="./DEVICE%20TWEAKER/assets/screenshots/restore_dialog_ru.png" alt="Restore and Checkpoint Management" width="700">
</p>

## Requirements & Binaries

### System Requirements
- **Operating System:** Windows 10 (Build 19041 or higher) / Windows 11 (64-bit).
- **Architecture:** x86-64.
- **Privileges:** Administrator access (required for HKLM registry access, Local APIC routing, and Ring 0 driver communication).

### Executable Flavors
- **`DEVICE.TWEAKER.exe`** (~4 MB): compact binary requiring [.NET 8 Desktop Runtime](https://dotnet.microsoft.com/download/dotnet/8.0).
- **`DEVICE.TWEAKER.SelfContained.exe`** (~150 MB): standalone binary with bundled .NET 8 runtime, ready to run on clean systems without installing prerequisites.

## Building from Source

Building requires [.NET 8.0 SDK](https://dotnet.microsoft.com/download/dotnet/8.0).

To build both application variants and generate the SHA-256 verification manifest:

```powershell
powershell -ExecutionPolicy Bypass -File ".\DEVICE TWEAKER\publish-variants.ps1" -Flavor both
```

Compiled binaries and the `SHA256SUMS.txt` manifest are placed in:
```
.\DEVICE TWEAKER\bin\ReleasePackages\v0.0.4-alpha.2\
```

## License and Community

- **Author:** [@arsenza](https://t.me/arsenzaa)
- **Channel & Feedback:** [Telegram @arsenzaa](https://t.me/arsenzaa)
- **License:** [GNU General Public License v3.0](./LICENSE)
