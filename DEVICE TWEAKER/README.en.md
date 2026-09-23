<div align="center">

<img src="./DEVICE%20TWEAKER/assets/banner.svg" alt="DEVICE TWEAKER Banner" width="900">

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

**DEVICE TWEAKER** is a low-level Windows 10 and 11 utility combining Message Signaled Interrupts (MSI / MSI-X) management, CPU interrupt affinity routing with hardware topology awareness, kernel core isolation (ReservedCpuSets), network stack tuning (RSS, NIC ITR), and direct physical register manipulation of USB xHCI IMOD (at a 250 ns quantum via a custom kernel driver) in a single unified interface with 1-click automatic optimization.

---

## Why Interrupt Optimization Matters

- **Game Thread Preemption (ISR / DPC Latency):**  
  Interrupt Service Routines (ISR) and Deferred Procedure Calls (DPC) execute at kernel level with highest CPU priority (`DISPATCH_LEVEL`). If interrupts from high-polling mice (1000–8000 Hz) or network cards execute on the same physical core running the game's primary render thread, the render thread is forcibly preempted. This causes micro-stutters, irregular frame times, and 0.1% low FPS drops.
- **Relieving CPU 0:**  
  By default, Windows routes the system timer, disk I/O, and general device interrupts to logical core 0. Steering high-load peripherals and GPU interrupts away from CPU 0 prevents DPC queue congestion.
- **Hardware USB Moderation (xHCI IMOD):**  
  xHCI USB controllers enable interrupt moderation (~50 μs) by default, batching incoming packets before signaling the CPU. Setting this register to `0 μs` via the custom `DTIMOD.sys` kernel driver disables moderation entirely, delivering mouse movement packets to the CPU with zero hardware delay.
- **Topological Routing (CCD & Hybrid Cores):**  
  Bypassing Intel efficiency cores (E-Cores) and pinning latency-critical device interrupts to the 3D V-Cache chiplet (CCD0) on AMD Ryzen processors avoids cross-CCX penalties and Infinity Fabric interconnect latency.

---

## Key Features

- **Categorized Device Monitoring:** GPUs, USB controllers, network adapters, storage controllers, and audio endpoints organized into discrete blocks with real-time parameter inspection.
- **MSI / MSI-X & IRQ Priority:** Enable Message Signaled Interrupts, remove vector limits (`MessageNumberLimit`), and enforce `IRQ Priority = High`.
- **CPU Affinity Engine:** Manual and automatic routing with hardware topology awareness:
  - Intel Hybrid Architecture (P-Core / E-Core separation).
  - Hyper-Threading / SMT logical pairs.
  - AMD Ryzen Chiplets (CCD0 with 3D V-Cache vs CCD1).
  - CCX clusters and Windows CPPC processor core energy/performance ratings.
- **Direct USB xHCI IMOD Control via `DTIMOD.sys`:** Read and write physical interrupt moderation registers (MMIO) with 250 ns precision. Configure `0 μs` for zero-delay mouse operation or set individual per-interrupter timings.
- **Network Stack Optimization (RSS & NIC ITR):** Complete disabling of network interrupt throttling (`ITR = 0 / Off` for supported Intel and Realtek NICs), Receive Side Scaling (RSS) queue tuning, and Energy Efficient Ethernet (EEE) disabling.
- **Kernel-Level Core Isolation (`ReservedCpuSets`):** Configure the Windows scheduler isolation mask so background OS tasks and worker threads avoid dedicated gaming cores.
- **1-Click Auto-Optimization Engine:**
  - Offloads `CPU 0` completely from peripherals and graphics controllers.
  - Pins the GPU to a dedicated adjacent pair of physical performance cores.
  - Separates high-frequency mouse and NIC interrupts to prevent DPC queue collisions.
  - Restricts device interrupts to CCD0 on AMD Ryzen X3D processors.
- **Power Management Control:** Fast toggling of selective suspend and bus power-saving policies for USB controllers and Ethernet NICs.
- **State Protection & Safe Rollback:** Automatic snapshot creation before any changes, write-protected Windows factory snapshot (`ORIGINAL STATE`), and comprehensive one-click restore.
- **Bilingual Interface:** Real-time on-the-fly switching between English and Russian without application restart.

---

<details>
<summary><b>📸 Visual Showcase & Topology Previews (Click to expand)</b></summary>
<br>

### 1. AMD Ryzen Dual-CCD Topology (CCD0 with 3D V-Cache vs CCD1)
<p align="center">
  <img src="./DEVICE%20TWEAKER/assets/screenshots/showcase_amd_dual_ccd_en.png" alt="AMD Dual-CCD Topology" width="850">
</p>

### 2. Intel Hybrid Architecture Topology (P-Cores & E-Cores)
<p align="center">
  <img src="./DEVICE%20TWEAKER/assets/screenshots/showcase_intel_hybrid_ru.png" alt="Intel Hybrid Topology" width="850">
</p>

### 3. USB xHCI IMOD Register Table (250 ns quantum)
<p align="center">
  <img src="./DEVICE%20TWEAKER/assets/screenshots/showcase_amd_imod_table_en.png" alt="xHCI IMOD Table" width="850">
</p>

### 4. Safety: Automated Backups & Factory Snapshot (ORIGINAL STATE)
<p align="center">
  <img src="./DEVICE%20TWEAKER/assets/screenshots/backup_choice_dialog.png" alt="Backup Dialog" width="850">
</p>

### 5. Topology Simulation Engine (Test Admin Panel)
<p align="center">
  <img src="./DEVICE%20TWEAKER/assets/screenshots/test_admin_panel.png" alt="Test Admin Panel" width="850">
</p>

</details>

---

## Safety & Driver Architecture

- **Driver-Free Polling:** The **REFRESH** button queries hardware through Windows SetupAPI and registry keys without loading any kernel driver.
- **Isolated Kernel Driver (`DTIMOD.sys`):** Loaded only upon explicit **CHECK** or when writing IMOD registers. Loaded via KDU, the driver remains resident until system reboot to prevent kernel instability from dynamic driver unloading.
- **System Reboot:** A system reboot is required for registry interrupt changes and hardware parameters to take effect.
- **Diagnostic Logging:** Structured logs are automatically generated on every launch inside the `logs` directory alongside the executable.

---

## System Requirements

- **Operating System:** Windows 10 or Windows 11 (64-bit, all editions).
- **Available Builds:**
  - `DEVICE.TWEAKER.exe` (~4 MB) — standard build, requires [.NET 8 Desktop Runtime](https://dotnet.microsoft.com/download/dotnet/8.0).
  - `DEVICE.TWEAKER.NET.FRAMEWORK.exe` (~150 MB) — standalone self-contained build with bundled .NET 8 runtime (runs out of the box on clean Windows installations).

---

## Quick Start

1. Download the latest release from [GitHub Releases](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases).
2. Run the application as **Administrator**.
3. Click **AUTO-OPTIMIZATION** for automated topology-aware matching, or configure affinities manually.
4. Click **APPLY**.
5. Restart your computer.

For full version history, see [CHANGELOG.md](./DEVICE%20TWEAKER/CHANGELOG.md).

---

## Building from Source

Compilation requires the [.NET 8.0 SDK](https://dotnet.microsoft.com/download/dotnet/8.0).

Build both application variants and generate SHA-256 checksums with a single command from the repository root:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File ".\DEVICE TWEAKER\build.ps1" -Flavor both -Configuration Release -SkipImodDriverBuild
```

Compiled binaries and checksums are placed into `DEVICE TWEAKER\bin\ReleasePackages\<version>\`.

---

## Automated Test Suite

The repository includes a suite of **21 unit tests (xUnit)** validating processor topology parsing, affinity bitmasks, `ReservedCpuSets` formatting, and IMOD hardware tables:

```powershell
dotnet test ".\DEVICE TWEAKER\tests\DeviceTweaker.Tests\DeviceTweaker.Tests.csproj" -c Release -p:BuildImodDriver=false
```

---

## Repository Layout

```
DEVICE-TWEAKER/
├── .gitignore
├── LICENSE
├── README.md                      # Primary Russian documentation
├── README.en.md                   # English documentation
└── DEVICE TWEAKER/                # Application source code & assets
    ├── Core/                      # Services, backup and restore engine
    ├── Devices/                   # PCI, USB controller, and NDIS discovery
    ├── GUI/                       # WinForms interface and dialog forms
    ├── IMOD/                      # DTIMOD.sys physical MMIO driver and KDU
    ├── Interop/                   # Low-level Win32 P/Invoke declarations
    ├── Localization/              # Interface language catalogs (RU / EN)
    ├── Models/                    # Data models, topologies, and reports
    ├── Tweaks/                    # Topology-aware auto-tuning heuristics
    ├── Scripts/                   # Automation and validation scripts
    ├── tests/                     # DeviceTweaker.Tests xUnit test suite
    ├── assets/                    # Graphical assets, icons, and screenshots
    ├── docs/                      # Technical documentation & release archives
    ├── DeviceTweakerCS.csproj     # .NET 8 project definition
    └── build.ps1                  # Multi-target automated build script
```

---

## Author & Community

- **Author:** [@arsenzaa](https://t.me/arsenzaa)
- **GitHub:** [arsenzaaa](https://github.com/arsenzaaa)
- **Repository:** [DEVICE-TWEAKER](https://github.com/arsenzaaa/DEVICE-TWEAKER)

## License

This project is licensed under the [GNU General Public License v3.0](LICENSE).  
KDU components in `IMOD/KDU` are distributed under the MIT license.
