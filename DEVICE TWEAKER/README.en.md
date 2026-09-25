<div align="center">

<img src="./DEVICE%20TWEAKER/assets/banner.svg" alt="DEVICE TWEAKER" width="460">

[![Release](https://img.shields.io/github/v/release/arsenzaaa/DEVICE-TWEAKER?include_prereleases&style=for-the-badge&color=007acc&label=Release)](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%20%7C%2011%20x64-0078D4?style=for-the-badge&logo=windows&logoColor=white)](https://microsoft.com/windows)
[![Telegram](https://img.shields.io/badge/Telegram-arsenzaa-2CA5E0?style=for-the-badge&logo=telegram&logoColor=white)](https://t.me/arsenzaa)

[Русский](./README.md) | **English**

</div>

**DEVICE TWEAKER** is a low-level utility for Windows 10 and Windows 11 designed for fine-grained tuning of hardware interrupts (MSI / MSI-X), deterministic device queue steering (CPU Affinity), and direct hardware register manipulation of USB interrupt moderation (xHCI IMOD).

It replaces fragmented legacy utilities (**MSI Utility v3**, **Microsoft Interrupt Affinity Policy Tool**, **GoInterruptPolicy**, as well as **RWEverything** invocations and fragile scripts blocked by modern anti-cheats). All functionality is unified within a single application featuring automated processor topology discovery, a custom signed kernel driver, and single-click restoration to pristine factory defaults.

<p align="center">
  <img src="./DEVICE%20TWEAKER/assets/screenshots/main_interface_en.png" alt="DEVICE TWEAKER Main Interface" width="850">
</p>

---

## Architecture & Features

### 1. MSI / MSI-X Vector Interrupts (Message Signaled Interrupts)
- **Eliminating Shared IRQ Line Conflicts:** Migrates devices from legacy Line-based (INTx) mode to MSI / MSI-X vector mode. Each device and queue is assigned a dedicated interrupt vector, and interrupts are delivered via direct PCIe Memory Write transactions into the Local APIC address space of the target core without polling foreign drivers.
- **Enabling MSI / MSI-X Mode:** Writes `MSISupported = 1` in the device registry key (`Device Parameters\Interrupt Management\MessageSignaledInterruptProperties`).
- **Unlocking Vector Limits (MessageNumberLimit = 0):** Removes the artificial Windows clamp on vector count. Setting `0` (Unlocked mode) deletes the `MessageNumberLimit` registry value, allowing the `pci.sys` bus driver to allocate the full vector pool requested in the hardware MSI-X table (up to 2048 vectors per PCI-SIG specifications).
- **Interrupt Dispatch Priority:** Configures `DevicePriority = 3` (`High`) under `Affinity Policy` for preferential mouse and GPU queue handling at the Windows kernel level.

### 2. Interrupt Queue Steering (CPU Affinity)
- **Precise Queue Routing:** Binds target logical processors via `DevicePolicy = 4` (`IrqPolicySpecifiedProcessors`) and the `AssignmentSetOverride` bitmask (supporting up to 64 logical processors).
- **Automated Topology-Aware Optimization (AUTO):**
  - **AMD Ryzen (Dual-CCD / 3D V-Cache):** Anchors GPU and mouse queues to physical cores on the primary compute die (CCD0) with fast L3 cache, eliminating cross-CCD Infinity Fabric transit penalties (~20 ns intra-CCD vs 70-90 ns inter-CCD across the I/O Die).
  - **Intel Core (Hybrid):** Directs high-frequency device queues strictly to Performance cores (P-Cores). Efficient cores (E-Cores) and sibling Hyper-Threading threads are excluded from real-time queues.
  - **Mouse & Network Isolation:** Isolates the primary mouse controller on a dedicated physical core with the highest ACPI CPPC2 ranking, while routing the network adapter to a separate core to prevent DPC conflicts.
  - **CPPC2 Core Ranking:** Reads Windows ETW events (Event ID 55 from `Microsoft-Windows-Kernel-Processor-Power`, `MaximumPerformancePercent`) to identify top-performing physical cores instantly without overhead.

<details>
<summary>View queue allocation examples for AMD Ryzen 9 9950X3D and Intel Core i9-14900K</summary>

<p align="center">
  <img src="./DEVICE%20TWEAKER/assets/screenshots/showcase_amd_9950x3d_dual_ccd_en.png" alt="AMD Ryzen 9 9950X3D Dual-CCD" width="850">
</p>

<p align="center">
  <img src="./DEVICE%20TWEAKER/assets/screenshots/showcase_intel_14900k_hybrid_en.png" alt="Intel Core i9-14900K Hybrid" width="850">
</p>

</details>

### 3. System-Wide Core Isolation (ReservedCpuSets)
- **Kernel-Level Core Reservation:** Configures the Windows setting `[HKLM\System\CurrentControlSet\Control\Session Manager\kernel] ReservedCpuSets`.
- **How It Works:** Excludes selected processors from the generic thread scheduler pool. System services, background tasks, and unpinned DPCs are barred from executing on reserved cores, preserving dedicated compute cycles for game loops and deterministic interrupt handling.

### 4. Hardware xHCI IMOD Control (DTIMOD.sys Driver)
- **xHCI Interrupter Registers:** Per Intel xHCI specification (Section 5.5.2), each hardware interrupter features a 32-bit `IMOD` register with a hardware clock resolution of 250 nanoseconds.
- **Removing Windows Default Moderation Delay:** Standard `USBXHCI.SYS` often enforces a 50 µs (`0xC8`) or 1 ms (`0xFA0`) moderation interval, holding packets in the buffer. Setting `IMOD = 0` removes artificial throttling, delivering interrupts to the PCIe bus immediately upon packet completion.
- **Safe Custom Kernel Driver (Ring 0):** Custom signed driver `DTIMOD.sys` loads on demand to map physical MMIO memory via `MmMapIoSpace`. It requires no disabling of `VulnerableDriverBlocklist` in Windows 11 and is fully compatible with competitive anti-cheats (Vanguard, EAC, BattlEye).
- **Physical Controller Hierarchy:** Distinguishes direct CPU Root Complex controllers (CHIP 0 / Direct CPU) from chipset hubs (CHIP 1), enabling clean separation of mice on direct CPU lanes with zero moderation while routing audio to chipset hubs at 1 ms to prevent dropouts.

<p align="center">
  <img src="./DEVICE%20TWEAKER/assets/screenshots/showcase_amd_imod_table_en.png" alt="xHCI IMOD Hardware Register Table" width="850">
</p>

### 5. Network Stack Optimization (NIC ITR and RSS)
- **Network Interrupt Throttle Rate (ITR / EITR / IntrMit):** Direct inspection and disabling of hardware moderation registers on Intel (I210, I211, I225, I226, I350, I219) and Realtek (RTL8111, RTL8125, RTL8126) controllers.
- **Receive Side Scaling (RSS):** Granular configuration of `*NumRssQueues` and `*RssBaseProcNumber` under the adapter registry class. Restricting RSS queues and routing them away from CPU 0 ensures deterministic packet processing without competing with the game loop.
- **Disabling Energy Efficient Ethernet (EEE):** Prevents physical transceivers (PHY) from dropping into low-power states during idle periods.

### 6. Background Input Throttling (Windows 11 RawMouseThrottleDuration)
- **RawMouseThrottleDuration Management:** Configures the native Windows 11 parameter under `HKCU\Control Panel\Mouse` (available in Windows 11 build 22621.1928 / KB5028185 and newer).
- **Relieving Background Overhead:** Throttles raw mouse reporting for background windows (1 to 20 ms range), eliminating unnecessary CPU interrupts from 1000-8000 Hz mice during background tasks without affecting the active game.

### 7. Backup & Factory State Protection
- **Immutable Initial Snapshot:** Automatically captures a pristine baseline of original Windows settings upon initial launch, protected from overwrites.
- **Differential Checkpoints:** Creates safe registry rollback points prior to each APPLY action.
- **Single-Click Recovery:** Allows step-by-step undo of individual adjustments or a full return to factory Windows defaults.

<p align="center">
  <img src="./DEVICE%20TWEAKER/assets/screenshots/restore_dialog.png" alt="Restore Checkpoint Manager" width="700">
</p>

---

## Deliverables & Flavors

- **Supported OS:** Windows 10 and Windows 11 (64-bit, build 19041 and newer).
- **Executables:**
  - `DEVICE.TWEAKER.exe` (~4 MB) - compact binary, requires [.NET 8 Desktop Runtime](https://dotnet.microsoft.com/download/dotnet/8.0).
  - `DEVICE.TWEAKER.SelfContained.exe` (~150 MB) - fully standalone build with embedded .NET 8 runtime, runs on clean systems without installing dependencies.

---

## Result Verification Methodology

To verify the impact of interrupt steering and moderation tuning objectively, the following Windows telemetry tools are recommended:

1. **ISR & DPC Distribution (ETW / Windows Performance Recorder):**
   ```cmd
   wpr -start CPU.light -start GPU.light
   :: Run a 1-2 minute gaming benchmark session
   wpr -stop trace.etl
   ```
   Open `trace.etl` in **Windows Performance Analyzer (WPA)** to inspect DPC/ISR distribution across cores (*DPC/ISR Duration by Module and Function*) and confirm absence of DPC spikes on dedicated gaming cores.

2. **Frametime Stability Monitoring:**
   Use **CapFrameX** or **PresentMon** to record frametime distribution (p95, p99, p99.9) and variance before and after applying configurations.

3. **USB Polling Interval Inspection:**
   Use **MouseTester** to plot polling intervals (*Interval vs Time*) and verify elimination of packet batching when xHCI IMOD is set to 0.

---

## Building from Source

Building requires the [.NET 8.0 SDK](https://dotnet.microsoft.com/download/dotnet/8.0).

Build both flavors and produce the release package with SHA-256 sums:

```powershell
powershell -ExecutionPolicy Bypass -File ".\publish-variants.ps1" -Flavor both
```

Compiled binaries and the `SHA256SUMS.txt` manifest are placed in:
```
.\bin\ReleasePackages\v0.0.4-alpha.2\
```

---

## License & Contacts

- Author: [@arsenza](https://t.me/arsenzaa)
- Official Channel & Feedback: [Telegram @arsenzaa](https://t.me/arsenzaa)
- License: [GNU General Public License v3.0](./LICENSE)
