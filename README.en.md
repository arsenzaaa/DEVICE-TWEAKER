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

**DEVICE TWEAKER** is a low-level Windows 10 and 11 utility engineered for deterministic device interrupt management, input pipeline arbitration, and Windows NT kernel latency optimization.

The tool provides Message Signaled Interrupts configuration (**MSI / MSI-X**), hardware topology-aware core affinity routing (**CPU Affinity**), kernel core isolation via **`ReservedCpuSets`**, network stack queue partitioning (**RSS**, **NIC ITR**), and direct physical MMIO programming of USB interrupt moderation registers (**xHCI IMOD** at a 250 ns quantum via a custom `DTIMOD.sys` kernel driver).

---

## Windows Kernel Latency & Interrupt Architecture

In the Windows NT architecture, hardware interrupt dispatching operates on a strict priority hierarchy of Interrupt Request Levels (**IRQL**):

```
[ Hardware Device: USB / GPU / NIC ]
                  │
                  ▼  (Hardware IRQ)
    [ Interrupt Service Routine (ISR) ]  ──►  Executes at DIRQL (disables core interrupts)
                  │
                  ▼  (Queues deferred call)
    [ DPC Queue (WDF01000.sys) ]         ──►  Executes at DISPATCH_LEVEL
                  │
                  ▼  (Signals OS input subsystem thread)
    [ Subsystem Input Thread ]
       ├─ Windows 10: CSRSS (Win32k Raw Input Thread)
       └─ Windows 11: DWM   (Kernel Sensor Thread / Master Input Thread)
                  │
                  ▼  (Delivers raw input via buffer or IPI)
    [ Game / User-Mode Process ]
       ├─ GameThread / App Input Thread
       └─ RenderThread / RHIThread
```

### 1. Render Thread Preemption (Execution Preemption)
Interrupt Service Routines (**ISR**) and Deferred Procedure Calls (**DPC**) execute at `DIRQL` and `DISPATCH_LEVEL`, which strictly supersede all user-mode execution (`PASSIVE_LEVEL`).

When interrupts from a high-polling mouse (1000–8000 Hz), Ethernet controller, or GPU are scheduled on the same physical core running the game's critical loop (`RenderThread` or `GameThread`), the scheduler forcibly preempts game execution to drain driver DPC queues. This stalls the frame delivery pipeline, producing severe frame time variance (**Frame Time jitter**) and irregular pacing.

### 2. Queue Congestion on CPU 0
By default, the Windows scheduler directs system timers, storage NVMe/SATA I/O, system bus interrupts, and standard PnP controllers to logical core 0. Leaving graphics controllers and gaming peripherals on `CPU 0` forces their ISR/DPC handlers into a shared system queue, introducing unpredictable dispatch delays.

### 3. Hardware USB Moderation (xHCI IMOD)
The Intel xHCI specification implements hardware interrupt moderation (**IMOD**, controlled by `IMODI` / `IMODC` registers) inside USB host controllers. By default, Windows configures an interval of ~50 μs (200 units at a 250 ns quantum) per interrupter. The controller deliberately holds back interrupts to batch incoming data packets.

For high-rate gaming mice (1000–8000 Hz), this batching introduces artificial delivery jitter. **DEVICE TWEAKER** communicates through `DTIMOD.sys` directly with the controller's physical MMIO space, writing `IMODI = 0 μs` (Zero Moderation) for immediate interrupt dispatch upon packet arrival. Conversely, audio endpoints can receive calibrated moderation intervals to reduce overall DPC load without buffer underruns.

### 4. Topology Penalties (CCD0 3D V-Cache vs E-Cores)
- **AMD Ryzen Dual-CCD (X3D):** CCD0 features a high-density 3D V-Cache slice with low access latency, while CCD1 provides standard cache at higher boost clocks. Inter-CCD communication via the Infinity Fabric interconnect incurs a 60–80 ns round-trip latency and invalidates L3 cache lines. Routing device interrupts to CCD1 while a game executes on CCD0 causes constant cross-die memory synchronization.
- **Intel Hybrid Architecture (12–14th Gen):** Efficiency cores (E-Cores) lack Hyper-Threading, feature narrower execution pipelines, and exhibit higher C-state exit latencies. Handling real-time hardware interrupts on E-Cores introduces dispatch latency spikes.

### 5. Network Stack: RSS, NIC ITR & NetAdapterCx
High-throughput Ethernet controllers generate massive interrupt rates. Without explicit **Receive Side Scaling (RSS)** partitioning, incoming network packets share `WDF01000.sys` execution time with mouse input. Disabling network moderation (`NIC ITR = 0 / Off` on supported Intel and Realtek controllers) and deactivating **Energy Efficient Ethernet (EEE)** prevents PHY sleep states and eliminates packet buffering delays.

### 6. Kernel Core Isolation (ReservedCpuSets)
The `ReservedCpuSets` registry value in `HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\kernel` configures a CPU bitmask instructing the Windows NT scheduler to withhold unpinned threads and background system services from designated cores. Binding critical hardware interrupts to reserved cores insulates driver execution from background operating system noise.

---

## Core Features

- **Bus-Level Device Classification:** Dedicated management blocks for GPUs, USB host controllers (with CHIP 0 / CHIP 1 hardware pathing via PCI ID), network adapters, storage controllers, and audio endpoints.
- **MSI / MSI-X & IRQ Priority:** Transition devices from legacy line-based interrupts to Message Signaled Interrupts (MSI), eliminate message limits (`MessageNumberLimit`), and enforce `IRQ Priority = High`.
- **Hardware-Aware Affinity Engine:**
  - Physical core identification and SMT / Hyper-Threading logical pair mapping.
  - AMD Ryzen chiplet awareness (CCD0 with 3D V-Cache vs CCD1) and CCX cluster grouping.
  - Intel Performance (P-Core) and Efficiency (E-Core) segregation.
  - Direct hardware core performance rating queries via kernel ETW events (`EventLogReader`, Event ID 53) in <60 ms without WMI overhead.
- **Low-Level Kernel Driver `DTIMOD.sys`:**
  - Direct physical MMIO reading and writing with 250 ns precision.
  - PCI BAR boundary checking and operational register bounds verification.
  - Zero Moderation configuration (`0 μs`) for mouse controllers.
- **Network Stack Arbitration:**
  - Receive Side Scaling queue tuning (`*NumRssQueues`, `*RssBaseProcNumber`).
  - Direct hardware throttle rate disabling (`NIC ITR = 0 / Off`) for Intel (I210, I211, I225, I226, I350) and Realtek NICs.
  - Disabling Energy Efficient Ethernet (EEE).
- **1-Click Topology Auto-Optimization:**
  - Complete offloading of graphics and input peripherals from congested `CPU 0`.
  - Pinning the GPU to an adjacent pair of dedicated physical P-Cores.
  - Core separation between mouse and network interrupts to prevent DPC queue collisions.
  - Restricting latency-sensitive devices to CCD0 on AMD Ryzen X3D processors.
- **Bus Power Policy Control:** Fast toggling of selective suspend and device power-saving states for USB controllers and NDIS network adapters.
- **State Protection & Rollback:**
  - Mandatory automated backup generation before applying any registry or hardware modifications.
  - Write-protected factory snapshot (**ORIGINAL STATE**).
  - Pre-write parameter validation and automatic atomic rollback on write failures.
- **Live Bilingual Localization:** Instant English / Russian switching without application restart.

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

## Driver Architecture & Kernel Safety

- **Driverless Hardware Discovery:** The **REFRESH** action queries device status strictly through Windows SetupAPI and registry hives without loading any kernel modules.
- **Isolated Kernel Driver (`DTIMOD.sys`):** Loaded on-demand via KDU exclusively when clicking **CHECK** or writing IMOD register values. The driver remains resident until reboot to prevent kernel instability from dynamic driver unloading.
- **Parameter Application:** PCI/MSI interrupt routing and NDIS driver parameters are parsed by the OS kernel during device initialization, requiring a system reboot for changes to take full effect.
- **Structured Logging:** Detailed diagnostic execution logs are recorded on every run in the `logs` folder next to the executable.

---

## Build Variants & System Requirements

- **Operating System:** Windows 10 or Windows 11 (64-bit, Home, Pro, Enterprise, LTSC).
- **Available Builds:**
  - `DEVICE.TWEAKER.exe` (~4 MB) — standard build, requires [.NET 8 Desktop Runtime](https://dotnet.microsoft.com/download/dotnet/8.0).
  - `DEVICE.TWEAKER.NET.FRAMEWORK.exe` (~150 MB) — standalone self-contained build with embedded .NET 8 runtime (runs out of the box on clean Windows installations).

---

## Building from Source

Compilation requires the [.NET 8.0 SDK](https://dotnet.microsoft.com/download/dotnet/8.0).

Build both application variants and generate SHA-256 checksums with a single command from the repository root:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File ".\DEVICE TWEAKER\build.ps1" -Flavor both -Configuration Release -SkipImodDriverBuild
```

Compiled binaries and checksum files are generated in `DEVICE TWEAKER\bin\ReleasePackages\<version>\`.

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
