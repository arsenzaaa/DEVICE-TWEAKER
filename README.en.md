<div align="center">

# DEVICE TWEAKER

[![Telegram](https://img.shields.io/badge/Telegram-arsenzaa-2CA5E0?style=for-the-badge&logo=telegram&logoColor=white)](https://t.me/arsenzaa)
[![Release](https://img.shields.io/github/v/release/arsenzaaa/DEVICE-TWEAKER?include_prereleases&style=for-the-badge&color=007acc&label=Release)](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases)
[![License](https://img.shields.io/badge/License-GPLv3-2ea44f?style=for-the-badge)](LICENSE)

<br>

<img src="./assets/DEVICE%20TWEAKER-wordmark.svg" alt="DEVICE TWEAKER" width="460">

<br>

[Русский](./README.md) | **English**

</div>

---

<div align="center">
  <img src="./assets/screenshots/main_interface_en.png" alt="DEVICE TWEAKER GUI" width="850">
</div>

---

**DEVICE TWEAKER** is a Windows 10 and 11 utility combining MSI Mode configuration, CPU Affinity interrupt steering, ReservedCpuSets core isolation, network stack optimization (RSS, NIC ITR), and direct USB xHCI IMOD register tuning in a unified interface with smart CPU topology auto-optimization.

---

## Why Interrupt Tuning Matters

- **Interrupt Preemption Over Game Threads:** Hardware Interrupt Service Routines (ISR) and Deferred Procedure Calls (DPC) execute with highest Windows kernel priority. When network cards or high-polling mice (1000–8000 Hz) process interrupts on the same core executing your game, the game thread is preempted, causing micro-stutters and uneven frametimes.
- **Offloading CPU 0:** By default, Windows routes the system timer, disk queues, and device interrupts to logical core 0, bottlenecking DPC queues.
- **Disabling USB Moderation Delay (xHCI IMOD):** xHCI controllers default to ~50 µs interrupt moderation, buffering mouse packets in batches. Setting IMOD to `0 µs` via the custom `DTIMOD.sys` driver disables this buffering entirely — mouse interrupts reach the CPU instantly.
- **CPU Silicon Topology Awareness:** Prevents device interrupts from landing on slow Intel E-cores and locks game traffic to the 3D V-Cache chiplet (CCD0) on AMD Ryzen processors, avoiding high-latency Infinity Fabric data transfers.

---

## Features

- **Device Categorization & Monitoring:** graphics cards, USB controllers, network adapters, NVMe storage, and audio devices organized in separate categories with their active interrupt settings.
- **MSI / MSI-X & IRQ Priority:** enable Message Signaled Interrupts (MSI), remove message vector caps (`MessageNumberLimit`), and set high interrupt priority (`IRQ Priority = High`).
- **CPU Affinity Steering:** manual and automatic device-to-core binding with a visual map: P-Core / E-Core separation, Hyper-Threading / SMT pairs, AMD chiplets (CCD0 / CCD1), CCX blocks, and CPPC core performance rankings.
- **Direct USB xHCI IMOD Control via `DTIMOD.sys`:** read and write physical MMIO interrupt moderation registers with 250 ns granularity. Set `0 µs` (unmoderated) for mouse ports or tune custom intervals per interrupter.
- **Network Stack Optimization (RSS & NIC ITR):** disable NIC interrupt moderation (`ITR = 0 / Off` on Intel and Realtek), configure Receive Side Scaling (RSS) queues, and disable Energy Efficient Ethernet (EEE).
- **Core Isolation via `ReservedCpuSets`:** set the kernel-level `ReservedCpuSets` bitmask. Reserved cores are exempt from general Windows background threads and services.
- **1-Click CPU Topology Auto-Optimization:**
  - Offloads `CPU 0` from peripheral and device interrupts.
  - Binds the GPU to a dedicated pair of physical P-cores.
  - Isolates mouse and network queues onto separate cores to prevent DPC contention.
  - Pins interrupts to the 3D V-Cache chiplet (CCD0) on AMD Ryzen X3D processors.
- **Power Management:** easily disable power-saving states and Selective Suspend on USB controllers and network adapters.
- **USB Controller Classification:** CHIP 0 / CHIP 1 detection by PCI ID and refined filtering for NVIDIA USB controllers.
- **Safety & Backups:** automatic snapshot before applying any change, protected original baseline (ORIGINAL STATE), and one-click rollback.
- **Localization:** native Russian and English interface with instant on-the-fly switching.

---

## Important

- The program modifies registry settings (`HKLM`) and uses the low-level `DTIMOD.sys` driver when working with IMOD.
- The **REFRESH** button queries devices through standard SetupAPI and registry without loading kernel drivers.
- The `DTIMOD.sys` driver is only loaded when clicking **CHECK** or applying IMOD settings. Loaded via KDU, it remains in memory until reboot to avoid kernel instability on unload.
- A system restart is required after applying settings for kernel and registry changes to take effect.
- Diagnostics logs are automatically written on each launch in the `logs` folder next to the executable.

---

## Requirements

- Windows 10 or Windows 11 (64-bit).
- Build variants:
  - `DEVICE.TWEAKER.exe` (~4 MB) — standard executable, requires [.NET 8 Desktop Runtime](https://dotnet.microsoft.com/download/dotnet/8.0).
  - `DEVICE.TWEAKER.NET.FRAMEWORK.exe` (~150 MB) — standalone build with embedded .NET 8 runtime (runs out of the box on clean Windows).

---

## Usage Guide

1. Download the latest release from [GitHub Releases](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases).
2. Run the executable as **Administrator**.
3. Review detected devices and current core assignments.
4. Click **AUTO-OPTIMIZATION** for automatic CPU layout or adjust affinities manually.
5. Click **APPLY**.
6. Restart your computer.

---

## Building from Source

Requires the [.NET 8.0 SDK](https://dotnet.microsoft.com/download/dotnet/8.0).

Build both flavors with a single command:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\build.ps1 -Flavor both -Configuration Release -SkipImodDriverBuild
```

Binaries and SHA-256 hashes will be placed in `bin\ReleasePackages\<version>\`.

---

## Project Structure

- `Affinity` — CPU Affinity, interrupt steering, RSS queues, and ReservedCpuSets mask.
- `Core` — system services, backup creation, restore logic, and IMOD implementation.
- `Devices` — detection of PCI/PCI-e devices, USB hosts, and NDIS network adapters.
- `GUI` — WinForms user interface, custom controls, and modal dialogs.
- `IMOD` — `DTIMOD.sys` kernel driver and KDU loader.
- `Localization` — Russian and English UI strings.
- `Tweaks` — CPU topology auto-optimization and reset logic.

---

## Developer

- **Telegram:** [@arsenzaa](https://t.me/arsenzaa)
- **GitHub:** [arsenzaaa](https://github.com/arsenzaaa)
- **Repository:** [DEVICE-TWEAKER](https://github.com/arsenzaaa/DEVICE-TWEAKER)

## License

This project is licensed under the [GNU General Public License v3.0](LICENSE).  
KDU components in `IMOD/KDU` are licensed under the MIT License.
