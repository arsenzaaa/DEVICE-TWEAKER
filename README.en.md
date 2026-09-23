<div align="center">

<img src="./DEVICE%20TWEAKER/assets/banner.svg" alt="DEVICE TWEAKER" width="900">

<br><br>

[![Release](https://img.shields.io/github/v/release/arsenzaaa/DEVICE-TWEAKER?include_prereleases&style=for-the-badge&color=007acc&label=Release)](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases)
[![Unit Tests](https://img.shields.io/badge/Unit%20Tests-21%20Passed-2ea44f?style=for-the-badge&logo=dotnet)](https://github.com/arsenzaaa/DEVICE-TWEAKER)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%20%7C%2011%20x64-0078d4?style=for-the-badge&logo=windows)](https://github.com/arsenzaaa/DEVICE-TWEAKER)
[![License](https://img.shields.io/badge/License-GPLv3-2ea44f?style=for-the-badge)](LICENSE)
[![Telegram](https://img.shields.io/badge/Telegram-@arsenzaa-2CA5E0?style=for-the-badge&logo=telegram&logoColor=white)](https://t.me/arsenzaa)

<p align="center">
  <a href="./README.md">Русский</a> • <b>English</b>
</p>

---

<br>

<img src="./DEVICE%20TWEAKER/assets/screenshots/main_interface_en.png" alt="DEVICE TWEAKER Main Interface" width="900">

</div>

<br>

**DEVICE TWEAKER** is a Windows 10 and 11 utility for configuring device interrupts (MSI), core affinity routing (CPU Affinity), direct USB interrupt moderation (xHCI IMOD), and network stack tuning.

It combines into a single interface everything that previously required scattered tools and manual registry edits: switching devices to MSI / MSI-X mode, removing vector limits, pinning interrupts to specific cores according to CPU architecture, reading and writing USB IMOD registers directly via our custom kernel driver `DTIMOD.sys`, configuring network queues (RSS / NIC ITR), and reserving CPU cores via `ReservedCpuSets`.

---

## Why Configure Interrupts

In Windows, device hardware signals follow a specific processing chain:
1. A device (mouse, keyboard, network card, GPU) fires a hardware interrupt (**IRQ**).
2. The CPU pauses current work and executes a fast Interrupt Service Routine (**ISR**).
3. The driver places the primary processing into a Deferred Procedure Call (**DPC**) queue, executing at the highest priority (`DISPATCH_LEVEL`).
4. Data is delivered to OS input subsystem threads (in Windows 10 — `CSRSS`, in Windows 11 — `DWM`), which then forward it to the game (`GameThread`, `RenderThread`).

### Why Micro-Stutters Happen Without Proper Configuration:

- **Game Thread Preemption:**  
  DPC queues execute with higher priority than games. When interrupts from a high-polling mouse (1000–8000 Hz) or network adapter are handled on the same core where the game renders frames, the game thread is forcibly preempted. This causes micro-stutters and uneven frame pacing.
- **CPU 0 Overload:**  
  By default, Windows routes the system timer, disk operations, and most device interrupts to core 0. Leaving graphics or input controllers on CPU 0 creates congestion in the shared system DPC queue.
- **Hardware USB Delay (xHCI IMOD):**  
  USB host controllers enable interrupt moderation (~50 μs) by default. The controller intentionally delays interrupts and batches packets instead of delivering them immediately. For high-rate gaming mice, this introduces latency and jitter. `DTIMOD.sys` writes directly to physical controller registers to set moderation to `0 μs` (zero delay).
- **Cross-CCD Latency on AMD (Infinity Fabric):**  
  On AMD Ryzen dual-CCD processors (such as 7950X3D, 9950X3D), the game runs on CCD0 with the fast 3D V-Cache. If GPU or peripheral interrupts execute on CCD1, data is constantly routed through the Infinity Fabric interconnect, adding 60–80 ns of round-trip latency.
- **Slow E-Cores on Intel:**  
  On Intel 12–14th Gen processors, latency-critical interrupts should not be assigned to efficiency cores (E-Cores) due to their narrower pipeline and slower wake-up times.
- **Mouse and Network Collisions:**  
  When the network card and mouse share the same core, heavy network traffic saturates the `WDF01000.sys` DPC queue, causing mouse packets to wait in line. Network traffic should be moved to separate cores via RSS queues, with network moderation disabled (`NIC ITR = 0`).
- **Why `ReservedCpuSets` Matters:**  
  This is a Windows registry setting (`HKLM\...\Session Manager\kernel`) that instructs the OS scheduler not to schedule background services and unpinned threads on selected cores. We reserve cores to shield them from background system noise.

---

## Features

- **Categorized Device Layout:** GPUs, USB controllers (with CHIP 0 / CHIP 1 chipset path detection via PCI ID), network adapters, storage, and audio grouped into clear categories.
- **MSI / MSI-X & IRQ Priority:** Switch devices to Message Signaled Interrupts, remove message limits (`MessageNumberLimit`), and set high priority (`IRQ Priority = High`).
- **CPU Affinity Routing:**
  - Physical core and Hyper-Threading / SMT pair mapping.
  - AMD Ryzen chiplet separation (CCD0 with 3D V-Cache vs CCD1) and CCX clusters.
  - Intel Performance (P-Core) and Efficiency (E-Core) core separation.
  - Core priority (CPPC) detection via Windows ETW kernel events (Event ID 53) in under 60 ms.
- **Direct USB IMOD Control via `DTIMOD.sys`:**
  - Direct access to xHCI controller registers with 250 ns precision.
  - Zero moderation (`0 μs`) for mouse controllers.
  - Per-interrupter tuning (e.g. increase moderation for USB audio to reduce CPU overhead).
- **Network Adapter Optimization:**
  - Receive Side Scaling queue tuning (`*NumRssQueues`, `*RssBaseProcNumber`).
  - Disabling interrupt moderation (`NIC ITR = 0 / Off`) on Intel (I210, I211, I225, I226, I350) and Realtek NICs.
  - Disabling Energy Efficient Ethernet (EEE).
- **1-Click Topology Auto-Optimization:**
  - Moves GPU and mouse interrupts off congested `CPU 0`.
  - Pins the GPU to an adjacent pair of physical P-Cores.
  - Separates mouse and network interrupts to dedicated cores to prevent DPC queue collisions.
  - Pins latency-sensitive devices to CCD0 on AMD Ryzen X3D processors.
- **Power Management:** Quickly disable Selective Suspend and Power Saving for USB controllers and network adapters.
- **Backups & Safety:**
  - Automatic backup creation before applying any changes.
  - Write-protected factory snapshot (**ORIGINAL STATE**) that is never overwritten.
  - Parameter validation before writing and automatic rollback on failure.
- **Bilingual Interface:** Real-time on-the-fly switching between English and Russian without restarting.

---

<details>
<summary><b>📸 Screenshots & Core Affinity Layouts (Click to expand)</b></summary>
<br>

### 1. Core Allocation: AMD Ryzen Dual-CCD (CCD0 with 3D V-Cache vs CCD1)
<p align="center">
  <img src="./DEVICE%20TWEAKER/assets/screenshots/showcase_amd_9950x3d_dual_ccd_ru.png" alt="AMD Dual-CCD Topology" width="850">
</p>

### 2. Core Allocation: Intel Hybrid (P-Core / E-Core)
<p align="center">
  <img src="./DEVICE%20TWEAKER/assets/screenshots/showcase_intel_14900k_hybrid_ru.png" alt="Intel Hybrid Topology" width="850">
</p>

### 3. USB xHCI IMOD Register Table (250 ns quantum)
<p align="center">
  <img src="./DEVICE%20TWEAKER/assets/screenshots/showcase_amd_imod_table_en.png" alt="xHCI IMOD Table" width="850">
</p>

### 4. Backups and Factory Snapshot Restore (ORIGINAL STATE)
<p align="center">
  <img src="./DEVICE%20TWEAKER/assets/screenshots/restore_dialog_ru.png" alt="Backups and Restore" width="850">
</p>

</details>

---

## Driver & Safety Architecture

- **Driver-Free Polling:** The **REFRESH** button queries devices using standard Windows SetupAPI and registry hives without loading the kernel driver.
- **Kernel Driver `DTIMOD.sys`:** Loaded via KDU only when clicking **CHECK** or writing IMOD values. The driver remains resident until reboot to prevent kernel instability from dynamic driver unloading.
- **Reboot:** Interrupt parameters are read by Windows during device initialization, so a system reboot is required after clicking **APPLY**.
- **Logging:** Structured logs are written to the `logs` folder alongside the executable on every launch.

---

## Build Variants & Requirements

- **OS:** Windows 10 or Windows 11 (64-bit, all editions).
- **Available Builds:**
  - `DEVICE.TWEAKER.exe` (~4 MB) — standard version, requires [.NET 8 Desktop Runtime](https://dotnet.microsoft.com/download/dotnet/8.0).
  - `DEVICE.TWEAKER.NET.FRAMEWORK.exe` (~150 MB) — standalone self-contained build with embedded .NET 8 runtime (runs out of the box on clean Windows installations).

---

## Building from Source

Compilation requires the [.NET 8.0 SDK](https://dotnet.microsoft.com/download/dotnet/8.0).

Build both application variants and generate SHA-256 checksums with a single command from the repository root:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File ".\DEVICE TWEAKER\build.ps1" -Flavor both -Configuration Release -SkipImodDriverBuild
```

Compiled binaries and checksum files will be in `DEVICE TWEAKER\bin\ReleasePackages\<version>\`.

---

## Automated Test Suite

The repository includes a suite of **21 unit tests (xUnit)** validating processor topology parsing, ETW CPPC event deserialization, affinity bitmasks, `ReservedCpuSets` formatting, and IMOD register bounds:

```powershell
dotnet test ".\DEVICE TWEAKER\tests\DeviceTweaker.Tests\DeviceTweaker.Tests.csproj" -c Release -p:BuildImodDriver=false
```

---

## Repository Layout

```
DEVICE-TWEAKER/
├── .gitignore
├── LICENSE
├── README.md                      # Documentation (Russian)
├── README.en.md                   # Documentation (English)
└── DEVICE TWEAKER/                # Application source code & assets
    ├── Core/                      # Services, backup and restore engine
    ├── Devices/                   # PCI, USB controller, and NDIS discovery
    ├── GUI/                       # WinForms interface and dialog forms
    ├── IMOD/                      # DTIMOD.sys physical MMIO driver and KDU
    ├── Interop/                   # P/Invoke wrappers for Win32 API
    ├── Localization/              # Interface language catalogs (RU / EN)
    ├── Models/                    # Data models, topologies, and reports
    ├── Tweaks/                    # Auto-tuning heuristics
    ├── Scripts/                   # Automation and validation scripts
    ├── tests/                     # DeviceTweaker.Tests xUnit test suite
    ├── assets/                    # Graphical assets, icons, and screenshots
    ├── docs/                      # Documentation & release archives
    ├── DeviceTweakerCS.csproj     # .NET 8 project definition
    └── build.ps1                  # Automated build script
```

---

## Author & Community

- **Author:** [@arsenzaa](https://t.me/arsenzaa)
- **GitHub:** [arsenzaaa](https://github.com/arsenzaaa)
- **Repository:** [DEVICE-TWEAKER](https://github.com/arsenzaaa/DEVICE-TWEAKER)

## License

This project is licensed under the [GNU General Public License v3.0](LICENSE).  
KDU components in `IMOD/KDU` are distributed under the MIT license.
