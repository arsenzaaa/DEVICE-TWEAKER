<div align="center">

<img src="./DEVICE%20TWEAKER/assets/banner.svg" alt="DEVICE TWEAKER" width="460">

[![Release](https://img.shields.io/github/v/release/arsenzaaa/DEVICE-TWEAKER?include_prereleases&style=for-the-badge&color=007acc&label=Release)](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%20%7C%2011%20x64-0078D4?style=for-the-badge&logo=windows&logoColor=white)](https://microsoft.com/windows)
[![Telegram](https://img.shields.io/badge/Telegram-arsenzaa-2CA5E0?style=for-the-badge&logo=telegram&logoColor=white)](https://t.me/arsenzaa)

[Русский](./README.md) | **English**

</div>

**DEVICE TWEAKER** is a low-level tuning utility for Windows 10 and Windows 11 engineered for fine-grained configuration of Message Signaled Interrupts (MSI / MSI-X), deterministic interrupt routing (CPU Affinity), direct hardware register manipulation of USB interrupt moderation (xHCI IMOD) via a custom kernel driver, and network stack optimization (RSS / NIC ITR).

By default, the Windows operating system (HAL and kernel) dispatches hardware interrupt vectors and associated Deferred Procedure Calls (DPCs) without considering processor topology or microarchitecture, concentrating workloads on system core CPU 0 or directing mouse, GPU, and network queues to the same physical cores. This causes hardware contention for execution pipelines, preemption of critical game rendering threads, micro-stutters, and frametime spikes.

DEVICE TWEAKER replaces the fragmented ecosystem of legacy utilities, such as **MSI Utility v2 / v3**, **Microsoft Interrupt Affinity Policy Tool (intaffinity)**, **GoInterruptPolicy**, as well as **RWEverything** invocations and fragile startup scripts for xHCI IMOD that are blocked by modern anti-cheats (Vanguard, Easy Anti-Cheat, BattlEye) due to the requirement of disabling the Vulnerable Driver Blocklist. All these critical tasks, including RSS queue balancing, network adapter moderation (NIC ITR), processor core isolation via ReservedCpuSets, and Windows 11 background raw mouse input throttle management, are unified within a single native utility featuring automated hardware topology analysis and guaranteed rollback safety (ORIGINAL STATE).

<p align="center">
  <img src="./DEVICE%20TWEAKER/assets/screenshots/main_interface_en.png" alt="DEVICE TWEAKER Main Interface" width="850">
</p>

---

## Architecture & Core Features

### 1. MSI / MSI-X Vector Interrupts (Message Signaled Interrupts)
- **Eliminating Shared Interrupt Conflicts (IRQ Sharing):** In legacy Line-based (INTx) mode, multiple physical devices share a single IRQ line. When an interrupt occurs, the operating system must sequentially poll the Interrupt Service Routines (ISRs) of all registered devices at elevated hardware priority (DIRQL). MSI and MSI-X replace shared lines with targeted memory write transactions (PCIe Memory Write) directed to the Local APIC address space of designated processor cores.
- **Enabling MSI / MSI-X Mode:** Enforcing the `MSISupported = 1` registry configuration on target devices (`Device Parameters\Interrupt Management\MessageSignaledInterruptProperties`).
- **Unlocking Vector Allocation Limits (MessageNumberLimit):** By default, Windows drivers frequently apply artificial vector clamping. Setting `MessageNumberLimit = 0` (or removing the key) unlocks access to the full pool of hardware interrupt vectors and queues.
- **Hardware Interrupt Priority (DevicePriority):** Configuring `DevicePriority = 3` (`High`) in `Device Parameters\Interrupt Management\Affinity Policy` guarantees prioritized kernel servicing for latency-critical mouse and GPU queues.

### 2. Queue Allocation & Core Pinning (CPU Affinity)
- **Precise Core Binding:** Assigning target logical processors via the Windows device policy `DevicePolicy = 4` (`IrqPolicySpecifiedProcessors`) and the `AssignmentSetOverride` bitmask (supporting configurations of up to 64 logical processors).
- **Topology-Aware Automated Mode (AUTO):**
  - **Mouse and Controller:** Allocating an isolated physical core with lowest latency exclusively to input peripherals, eliminating queue contention.
  - **Graphics Card (GPU):** Binding graphics stack interrupts to physical cores on the 3D V-Cache die (CCD0 on AMD Ryzen), completely freeing the secondary compute die (CCD1). This avoids cross-CCD Infinity Fabric latency penalties (~20 ns intra-CCD vs. 70-90 ns inter-CCD).
  - **Network Adapters & Keyboards:** Distributing network queues and keyboard controllers to independent cores, preventing interrupt preemption on active game rendering threads.
  - **Intel Hybrid Architecture:** Restricting high-frequency peripherals exclusively to high-IPC Performance cores (P-Cores). Efficient cores (E-Cores) are completely excluded from real-time interrupt processing due to their reduced frequency and higher inter-core latency.
  - **Hyper-Threading / SMT Thread Exclusion:** Bypassing sibling logical threads on active physical cores to prevent hardware execution unit contention (ALU, AGU, FPU) and L1/L2 cache evictions.
  - **CPPC Silicon Binning:** Reading hardware performance rankings (ETW Event ID 55 from `Microsoft-Windows-Kernel-Processor-Power`) to prioritize the highest-performing physical cores on the die.

### 3. System-Wide Core Isolation via ReservedCpuSets
- **Shielding Cores from Background Noise:** Configuring the Windows kernel setting `[HKLM\System\CurrentControlSet\Control\Session Manager\kernel] ReservedCpuSets`.
- **Operating Principle:** Removes selected logical processors from the generic Windows thread scheduler pool. System services, background processes, and unpinned DPCs (ntoskrnl, symcryptk, tcpip) are barred from executing on reserved cores, preserving dedicated compute cycles for latency-critical tasks.

### 4. Hardware xHCI IMOD Control via Custom Driver DTIMOD.sys
- **Direct Ring 0 Hardware Register Access:** Per Intel xHCI specification (Section 5.5.2), every hardware interrupter features a 32-bit `IMOD` (Interrupt Moderation Register) with a hardware clock resolution of 250 nanoseconds.
- **The Issue with Windows Default Moderation:** The standard `USBXHCI.SYS` driver enforces a 50 µs moderation delay (`0xC8`, 200) or 1 ms (`0xFA0`, 4000). While beneficial for office workloads, an 8000 Hz or 4000 Hz gaming mouse delivers packets every 125 to 250 µs. A 50 µs moderation window causes packet batching and input latency jitter.
- **Custom Signed Kernel Driver:** Unlike deprecated tools (such as RWEverything) that require disabling Windows Vulnerable Driver Blocklists and trigger modern game anti-cheat flags, DEVICE TWEAKER loads its own custom signed driver `DTIMOD.sys` on demand to map physical controller memory via `MmMapIoSpace`.
- **Granular Interrupter Tuning:** Full zero-latency mode (`0x0`, 0 ns) for mice, alongside optimized moderation (`0xFA0`, 1 ms) for USB audio interfaces to save CPU cycles without buffer dropouts.
- **Physical USB Topology Mapping:** Identifies physical hardware paths: CPU Root Complex (CHIP 0 / Direct CPU) versus chipset hubs (CHIP 1 / CHIP 1+), enabling clean separation of mice and keyboards across distinct physical controllers.

### 5. Network Stack Optimization (NIC ITR and RSS)
- **Network Interrupt Throttle Rate (ITR / EITR / IntrMit):** Direct inspection and disabling of hardware moderation registers on Intel (I210, I211, I225, I226, I350, I219) and Realtek (RTL8111, RTL8125, RTL8126) controllers.
- **Receive Side Scaling (RSS):** Granular configuration of `*NumRssQueues` and `*RssBaseProcNumber` under the adapter registry class (`HKLM\...\Class\{4d36e972-e325-11ce-bfc1-08002be10318}`) to distribute packet processing across designated cores rather than bottlenecking on CPU 0.
- **Disabling Energy Efficient Ethernet (EEE):** Prevents network PHY chips from dropping into low-power states during active connections.

### 6. Background Input Throttling (Windows 11)
- **RawMouseThrottleDuration Management:** Configures the native Windows 11 input stack parameter under `HKCU\Control Panel\Mouse` (1 to 20 ms range).
- **Relieving Background CPU Overhead:** Automatically reduces raw mouse input reporting for background windows (down to 50 Hz at setting 20) during high-rate (1-8 kHz) mouse operation without affecting the active game.

### 7. Bus Power Management
- **Disabling USB Selective Suspend:** Eliminates wake latencies by preventing USB ports from entering low-power idle states (`SelectiveSuspendEnabled = 0`).
- **Controller Power Policies:** Disables aggressive power savings on network adapters and system buses.

### 8. Backup & Rollback System (ORIGINAL STATE)
- **Immutable Baseline Snapshot:** Automatically captures pristine system state on initial launch into an immutable backup.
- **Differential Checkpoints:** Creates safe registry rollback points prior to any write operation.
- **One-Click Restoration:** Allows seamless step-by-step undo or full reversion to factory Windows defaults.

---

## Optimization Scenario Showcase

<details>
<summary><b>Expand Optimization Scenario Gallery</b></summary>
<br>

### GPU Queue Allocation (CPU Affinity)
Binding graphics stack interrupts to physical cores on the low-latency 3D V-Cache die (CCD0), isolating the secondary die (CCD1). Dynamic vendor badge accents (GeForce green for NVIDIA, Radeon red for AMD, Intel blue for Arc).

<p align="center">
  <img src="./DEVICE%20TWEAKER/assets/screenshots/showcase_gpu_affinity_en.png" alt="GPU Core Allocation" width="850">
</p>

### Hardware xHCI IMOD Register Table
Direct hardware interrupter register control with 250 ns granularity via custom kernel driver DTIMOD.sys, displaying live hardware values, latency metrics, and controller topology (CHIP 0 / CHIP 1).

<p align="center">
  <img src="./DEVICE%20TWEAKER/assets/screenshots/showcase_amd_imod_table_en.png" alt="xHCI IMOD Register Table" width="850">
</p>

### Hybrid Architecture (P-Cores / E-Cores)
Separation of Performance and Efficient cores on Intel processors. Total exclusion of E-Cores and sibling Hyper-Threading threads from processing high-rate peripheral interrupts.

<p align="center">
  <img src="./DEVICE%20TWEAKER/assets/screenshots/showcase_intel_14900k_hybrid_en.png" alt="Intel Hybrid Core Allocation" width="850">
</p>

### Recovery System (ORIGINAL STATE)
Single-click restoration to previous tuning stages or full rollback to the clean baseline Windows configuration captured prior to any modifications.

<p align="center">
  <img src="./DEVICE%20TWEAKER/assets/screenshots/restore_dialog.png" alt="Backup and Restore Dialog" width="850">
</p>

</details>

---

## Universal Hardware Compatibility

DEVICE TWEAKER is engineered on fundamental PCI Express standards, the xHCI specification, and the Windows NT kernel architecture, ensuring universal hardware compatibility:

- **Any x64 Multi-Core Processor:** AMD Ryzen, Threadripper, and EPYC families (including single-CCD and dual-CCD 3D V-Cache models as well as monolithic APUs); Intel Core 10th through 14th generations, Core Ultra, and Xeon (hybrid P/E and monolithic architectures).
- **Any PCI Express Device:** Graphics cards from all vendors (NVIDIA GeForce / RTX, AMD Radeon, Intel Arc), NVMe M.2 and SATA AHCI storage controllers, capture cards, and discrete sound cards.
- **Any xHCI USB Controller:** Standard USB 3.0 / 3.1 / 3.2 / USB4 host controllers embedded in AMD and Intel platforms, as well as discrete controllers from ASMedia, VIA, and Renesas.
- **Any Network Adapter:** Ethernet and Wi-Fi adapters supporting NDIS 6.x drivers, MSI-X, and RSS (with native hardware ITR/EITR/IntrMit register support for Intel and Realtek).

---

## Security & Custom Driver DTIMOD.sys

- **Standard Read-Only Mode:** Default application execution, hardware tree scanning, and the **REFRESH** command operate strictly through Windows SetupAPI and registry interfaces in Ring 3 without loading kernel code.
- **On-Demand Driver Loading:** The custom driver `DTIMOD.sys` loads strictly when reading interrupter registers or applying IMOD values to perform physical memory mapping via `MmMapIoSpace`.
- **Kernel Panic Prevention (BSOD):** The driver remains resident in memory until a scheduled system reboot, preventing bus driver crashes associated with dynamic hot-unloading.
- **Registry Policy Activation:** Windows interrupt policies (`DevicePolicy`, `DevicePriority`, `MessageSignaledInterruptProperties`, `ReservedCpuSets`) take effect upon a standard system restart when the Windows HAL re-initializes APIC interrupt dispatch tables.

---

## Distribution Variants

- **Operating System:** Windows 10 or Windows 11 (64-bit, build 19041 and newer).
- **Binaries:**
  - `DEVICE.TWEAKER.exe` (~4 MB) - compact standard executable, requires [.NET 8 Desktop Runtime](https://dotnet.microsoft.com/download/dotnet/8.0).
  - `DEVICE.TWEAKER.NET.FRAMEWORK.exe` (~150 MB) - fully self-contained build with bundled .NET 8 runtime, ready for clean systems without prerequisites.

---

## Building from Source

Building the project requires the [.NET 8.0 SDK](https://dotnet.microsoft.com/download/dotnet/8.0).

Build both application flavors and generate release packages with SHA-256 checksum manifests using:

```powershell
powershell -ExecutionPolicy Bypass -File ".\DEVICE TWEAKER\publish-variants.ps1" -Flavor both
```

*(If executing from within the project directory, use `.\publish-variants.ps1`)*

Compiled release packages and the checksum manifest `SHA256SUMS.txt` are created in:
```
.\DEVICE TWEAKER\bin\ReleasePackages\v0.0.4-alpha.2\
```
And individual publishing folders:
```
.\DEVICE TWEAKER\bin\Publish\
```

---

## License & Contact

- Author: [@arsenza](https://t.me/arsenzaa)
- Official Channel & Feedback: [Telegram @arsenzaa](https://t.me/arsenzaa)
- Licensed under the [GNU General Public License v3.0](./LICENSE).
