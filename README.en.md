<div align="right"><a href="./README.md">Русский</a></div>

<p align="center"><img src="./assets/branding/mark.svg" width="96" height="96" alt=""></p>

<h1 align="center">DEVICE TWEAKER</h1>

<p align="center"><b>Device interrupts on Windows 10 / 11 x64, tuned to your CPU topology.</b></p>

<p align="center">MSI Utility V3 + Interrupt Affinity Policy Tool (IAPT) + ReservedCpuSets + RSS + USB IMOD / NIC ITR in one interface, with intelligent auto-optimization.</p>

<p align="center"><b><a href="https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest">Download 0.0.4</a></b> · <a href="https://arsenzaaa.github.io/DEVICE-TWEAKER/en/">Documentation</a> · <a href="./CHANGELOG.en.md">Changelog</a> · <a href="https://t.me/arsenzaa">Telegram</a></p>

<p align="center"><a href="./assets/screenshots/showcase_intel_14900k_full_en.png"><img src="./assets/screenshots/framed/showcase_intel_14900k_full_en.png" alt="DEVICE TWEAKER: a USB controller on Intel Core i9-14900K, IMOD and interrupt affinity" width="100%"></a></p>

## <samp>[ WHY ]</samp>

ISR / DPC run above any thread and do not know which cores the game is using. A driver on the core of an important thread (GAME / RENDER / RHI) interrupts it -> the thread waits or moves to another core and loses its L1 / L2 -> frame time grows and stutters appear. Two drivers on one core delay each other's ISR / DPC; on one SMT-SET they share cache and EXECUTION-UNITS.

DEVICE TWEAKER spreads interrupts across physical cores: drivers do not share cores or SMT-SETs, and core 0 stays free when possible. Everything needed for that is in one window: MSI, affinity, priority, ReservedCpuSets, RSS, and interrupt moderation for USB and network.

## <samp>[ AUTO-OPTIMIZATION ]</samp>

**AUTO-OPTIMIZATION** reads the topology (P-Core / E-Core, SMT / HT, CPPC, CCD / CCX), finds the device roles and builds a plan:

- **GPU** -> two different physical cores. On one core the GPU's ISR interrupts its own DPC -> latency spikes. Two SMT siblings of one core do not count as two cores.
- **USB controller with the mouse / keyboard / gamepad, network, audio** -> their own cores, by placement rules.
- **Several CCDs** -> a target CCD is chosen.
- **Storage** -> MSI and high priority, affinity stays with Windows: on some NVMe drives `SpecifiedProcessors` stops Windows from booting.
- **Integrated GPU** -> affinity goes back to Windows. **Wi-Fi** is left untouched.

If the topology cannot be detected reliably, nothing is written. The original state is saved before the first run, and **RESTORE → ORIGINAL** brings it back.

## <samp>[ INTERRUPTS ]</samp>

**MSI MODE** — message-signaled interrupts instead of a line. Without MSI, a trace of an xHCI controller shows its DPC interrupted by two ISRs; with MSI it is not interrupted. **MSI LIMIT** — `Unlimited` removes the message limit.

**IRQ PRIORITY** affects the priority of the hardware interrupt. It is not the game's priority and not the DPC's priority.

**INTERRUPT AFFINITY** sets the logical processors for the device's IRQ. The list shows which physical core each thread belongs to, P / E-core, CCD and CPPC rank. The `SpecCPU` policy uses the checked mask; the other policies leave the choice to Windows.

```text
HKLM\SYSTEM\CurrentControlSet\Enum\PCI\<device>\Device Parameters\Interrupt Management
  MessageSignaledInterruptProperties  MSISupported           1 = MSI / MSI-X
                                      MessageNumberLimit     removed for Unlimited
  Affinity Policy                     DevicePriority         3 = High
                                      DevicePolicy           4 = SpecifiedProcessors
                                      AssignmentSetOverride  CPU mask
```

**IRQ MODE** shows what the device uses now: Line / MSI / MSI-X. **HW** shows what it supports. **READ IRQ MODE** confirms the mode from the `MSI Enable` / `MSI-X Enable` bits in PCI configuration space through DTIMOD.sys.

## <samp>[ USB IMOD ]</samp>

IMOD is the interrupt moderation interval of an xHCI **interrupter**, not of a single mouse. 250 ns steps: `0xC8` = 50 µs (Windows default), `0xFA0` = 1 ms, `0x0` = moderation off.

A device gets its interrupter when it connects, and it may change after a restart -> the program reads the device → interrupter map from the xHCI contexts in memory every time: the TRB of its transfer ring, or the Slot Context if the ring cannot be read.

- **Devices** — values by role. The AUTO profile: mouse `0x0`, audio `0xFA0` (fewer, longer interrupts -> less load), keyboard and gamepad `0xC8`.
- **XHCI** — one value for the whole controller.
- **Interrupters** — values by index.

Registers are read back after every write. **CHECK** shows the actual values and the time they were read.

## <samp>[ NETWORK ]</samp>

On NDIS adapters RSS distributes the interrupts, so the base CPU and queues are set instead of a mask. Processors are taken one per physical core, SMT siblings are skipped, and the queue count is limited by the driver's values and the MSI-X vectors.

```text
HKLM\SYSTEM\CurrentControlSet\Control\Class\{4d36e972-e325-11ce-bfc1-08002be10318}\<adapter>
  *RssBaseProcNumber  *NumRssQueues  *MaxRssProcessors  *RssMaxProcNumber
```

**NIC ITR** — the controller's own moderation registers: EITR on Intel I210 / I211, I225 / I226, I350, 82576 / 82580, ITR on I219, IntrMit on Realtek RTL8111 / 8168, RTL8125 / 8126, including compatible Killer adapters. **SAVE** turns off the driver's own moderation (`*InterruptModeration`, on Intel also `ITR`) -> otherwise the driver overwrites the register.

## <samp>[ STARTUP ]</samp>

IMOD and NIC ITR registers keep their value until a restart, sleep or controller reset. A saved profile is applied by the Task Scheduler task `\DEVICE TWEAKER\IMOD`, running as SYSTEM: at sign-in, after resume from sleep and when a USB / HID device connects.

The task runs the readable `ApplyIMOD.ps1` from `%ProgramData%\DEVICE TWEAKER\Startup` and works without the EXE. SHA-256 of the runtime, driver and loader is checked before each run; only administrators and SYSTEM can modify the folder.

## <samp>[ MORE ]</samp>

- **ReservedCpuSets** -> Windows does not schedule ordinary threads on reserved cores. A driver with a mask on such a core keeps running there, and threads are protected from it. `Session Manager\kernel`, CPU mask.
- **Mouse Throttle** -> `RawMouseThrottleDuration` in `HKCU\Control Panel\Mouse`: the Raw Input rate for background windows on Windows 11, 1000 / value Hz.
- **Polling Rate** -> mouse and keyboard from USB data, plus a Raw Input event-rate measurement.
- **Power saving** -> USB controllers, root USB hubs, wired adapters and USB selective suspend in the power plan. Values are read back through WMI after writing.
- **Backups and logs** -> a verified backup before every apply in the `Backups` folder, a detailed log in `logs`: what was applied, what was skipped and why.

Registry settings take effect after a restart, IMOD and NIC ITR registers immediately. Every setting is described in detail in the [documentation](https://arsenzaaa.github.io/DEVICE-TWEAKER/en/settings/).

## <samp>[ INSTALLATION ]</samp>

- [`DEVICE.TWEAKER.SelfContained.exe`](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest/download/DEVICE.TWEAKER.SelfContained.exe) — .NET included.
- [`DEVICE.TWEAKER.exe`](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest/download/DEVICE.TWEAKER.exe) — requires the [.NET 8 Desktop Runtime x64](https://dotnet.microsoft.com/download/dotnet/8.0).

Run as administrator -> **AUTO-OPTIMIZATION**, or manual settings and **APPLY** -> restart. **CHECK** in the IMOD or NIC ITR block shows the actual register values.

> **DTIMOD.sys and KDU.** IMOD and NIC ITR registers are not reachable from user mode; they are written by the custom DTIMOD.sys driver, and only to memory that Windows assigned to the selected device. The driver has no public signature and is loaded by [KDU](./src/IMOD/KDU/README.md) -> antivirus software flags such loaders as HackTool. HVCI and the Microsoft Vulnerable Driver Blocklist block the loading, and FACEIT and Vanguard may conflict with it: **SYSTEM CHECK** shows their state. All other settings work without the driver.

SHA-256 hashes of the files are in `SHA256SUMS.txt` on the [release page](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest), of the sources in `SOURCE-SHA256SUMS.txt`:

```powershell
(Get-FileHash .\DEVICE.TWEAKER.SelfContained.exe -Algorithm SHA256).Hash
Select-String -Path .\SHA256SUMS.txt -Pattern 'DEVICE.TWEAKER.SelfContained.exe'
```

## <samp>[ ROLLBACK ]</samp>

**RESTORE → ORIGINAL** returns the registry to its state before the first auto-optimization and removes the startup task. The same menu has every backup and **RESET WINDOWS DEFAULT**.

## <samp>[ FEEDBACK ]</samp>

Report bugs in [issues](https://github.com/arsenzaaa/DEVICE-TWEAKER/issues/new?template=bug_report.yml) with the log from the `logs` folder, your Windows version and PC configuration. To help test auto-optimization, send a [hardware report](https://github.com/arsenzaaa/DEVICE-TWEAKER/issues/new?template=hardware_report.yml) with three logs: before applying, right after and after a restart. Updates are posted in [Telegram](https://t.me/arsenzaa).

<p align="center"><sub><a href="https://arsenzaaa.github.io/DEVICE-TWEAKER/en/">Documentation</a> · <a href="./docs/DEVELOPMENT.en.md">Development and build</a> · <a href="./docs/THIRD_PARTY_NOTICES.md">Third-party components</a> · <a href="./.github/SECURITY.md">Security</a> · <a href="./LICENSE">GPL-3.0</a></sub></p>
