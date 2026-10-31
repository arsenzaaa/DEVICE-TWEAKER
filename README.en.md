<div align="right"><a href="./README.md">Русский</a></div>

<p align="center"><img src="./assets/branding/mark.svg" width="96" height="96" alt=""></p>

<h1 align="center">DEVICE TWEAKER</h1>

<p align="center"><b>Device interrupts on Windows 10 / 11 x64, tuned to your CPU topology.</b></p>

<p align="center">MSI Utility V3 + Interrupt Affinity Policy Tool (IAPT) + ReservedCpuSets + RSS + USB IMOD / NIC ITR in one interface, with intelligent auto-optimization.</p>

<p align="center"><b><a href="https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest">Download 0.0.4</a></b> · <a href="https://arsenzaaa.github.io/DEVICE-TWEAKER/en/">Site and guide</a> · <a href="./CHANGELOG.en.md">Changelog</a> · <a href="https://t.me/arsenzaa">Telegram</a></p>

<p align="center"><a href="./assets/screenshots/showcase_intel_14900k_full_en.png"><img src="./assets/screenshots/framed/showcase_intel_14900k_full_en.png" alt="DEVICE TWEAKER: a USB controller on Intel Core i9-14900K, IMOD and interrupt affinity" width="100%"></a></p>

## <samp>[ WHY ]</samp>

ISRs / DPCs run above any thread and do not know which cores the game is using. A driver on the core of an important thread (GAME / RENDER / RHI) interrupts it -> the thread waits or moves to another core and loses its L1 / L2 -> frame time grows and stutters appear. Two drivers on one core delay each other's ISRs / DPCs; on one SMT-SET they share cache and EXECUTION-UNITS.

DEVICE TWEAKER spreads interrupts across physical cores: drivers share neither cores nor SMT-SETs, and core 0 stays free when possible. **AUTO-OPTIMIZATION** builds a plan from the topology (P-Core / E-Core, SMT / HT, CPPC, CCD / CCX) and device roles. Every device can also be tuned by hand: MSI, affinity, priority, ReservedCpuSets, RSS, USB IMOD, NIC ITR, Mouse Throttle, power saving.

## <samp>[ GUIDE ]</samp>

How it works in Windows and what the program does, device type by device type. The same chapters with search are on the [site](https://arsenzaaa.github.io/DEVICE-TWEAKER/en/guide/).

- `[IRQ]` **[Interrupts: ISR, DPC and cores](./docs/guide/01-interrupts.en.md)** — How ISRs and DPCs work, why they outrank any game thread, and what happens when drivers share a core with GAME / RENDER / RHI.
- `[MSI]` **[MSI and MSI-X](./docs/guide/02-msi.en.md)** — Line versus message: why disabling MSI MODE adds extra and longer ISRs, why MessageNumberLimit should be removed, and how the program reads a device's real mode.
- `[CPU]` **[Affinity, policies and interrupt priority](./docs/guide/03-affinity.en.md)** — Interrupt Affinity policies, the AssignmentSetOverride mask, what IRQ Priority really does, Interrupt Steering, and the alternative "all cores except 0" method.
- `[GPU]` **[Graphics card](./docs/guide/04-gpu.en.md)** — Why a graphics card needs two different physical cores, how the NVIDIA driver differs from AMD, and what to do when ISRs / DPCs still pile up on one core.
- `[USB]` **[USB controllers and IMOD](./docs/guide/05-usb-imod.en.md)** — How an xHCI controller works: interrupters, Event Rings and the IMOD register. Why a mouse gets a different interrupter on every connection, what IMOD to give audio, and how to spread devices across controllers.
- `[NET]` **[Network: RSS, NDIS and NIC ITR](./docs/guide/06-network.en.md)** — NDIS versus NetAdapterCx, why a mask does not work on NDIS and RSS is needed instead, how many queues a game really uses, the NDIS thread and MMCSS, NIC ITR and the driver's own moderation.
- `[CPU]` **[CPU topology and ReservedCpuSets](./docs/guide/07-cpu-topology.en.md)** — SMT-SETs, P- and E-cores, CCDs and CPPC ranks: how to read the CPU list in the program. ReservedCpuSets, the mask format and two strategies for CPUs with two CCDs or E-cores.
- `[I/O]` **[Input: Polling Rate, Raw Input and Mouse Throttle](./docs/guide/08-input.en.md)** — The input path of a mouse, keyboard and controller on Windows 10 and 11, background-window throttling and RawMouseThrottleDuration, Raw Input Buffer and dedicated input threads, and what the Polling Rate measurement really measures.
- `[PWR]` **[Device power saving](./docs/guide/09-power.en.md)** — USB selective suspend, letting Windows turn a device off, and the Device Parameters values: what actually works, where it is stored, and how the program verifies the write.
- `[LOG]` **[Verification: traces and logs](./docs/guide/10-verification.en.md)** — How to check the result by facts rather than the FPS of one run: an xperf / WPR trace, where to look at ISRs and DPCs in WPA and MXA, what to send in a hardware report and how to read the program log.
- `[SYS]` **[DTIMOD.sys, KDU and startup](./docs/guide/11-driver.en.md)** — Why the program has its own driver, where it may write, why KDU loads it and what antivirus software, HVCI and anti-cheats think about that. How IMOD / NIC ITR startup works without the EXE.

## <samp>[ INSTALLATION ]</samp>

- [`DEVICE.TWEAKER.SelfContained.exe`](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest/download/DEVICE.TWEAKER.SelfContained.exe) — .NET included.
- [`DEVICE.TWEAKER.exe`](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest/download/DEVICE.TWEAKER.exe) — requires the [.NET 8 Desktop Runtime x64](https://dotnet.microsoft.com/download/dotnet/8.0).

Run as administrator -> **AUTO-OPTIMIZATION**, or manual settings and **APPLY** -> restart. **CHECK** in the IMOD or NIC ITR block shows the actual register values.

> **DTIMOD.sys and KDU.** IMOD and NIC ITR are written by the custom DTIMOD.sys driver, and only to memory that Windows assigned to the selected device. The driver has no public signature and is loaded by [KDU](./src/IMOD/KDU/README.md) -> antivirus software flags such loaders as HackTool, HVCI and the Microsoft Vulnerable Driver Blocklist block the loading, and FACEIT and Vanguard may conflict with it. All other settings work without the driver. Details in [chapter 11](./docs/guide/11-driver.en.md).

SHA-256 hashes of the files are in `SHA256SUMS.txt` on the [release page](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest), of the sources in `SOURCE-SHA256SUMS.txt`:

```powershell
(Get-FileHash .\DEVICE.TWEAKER.SelfContained.exe -Algorithm SHA256).Hash
Select-String -Path .\SHA256SUMS.txt -Pattern 'DEVICE.TWEAKER.SelfContained.exe'
```

## <samp>[ ROLLBACK ]</samp>

Before every apply the program creates a verified backup in the `Backups` folder next to the EXE. **RESTORE → ORIGINAL** returns the registry to its state before the first auto-optimization and removes the startup task. The same menu has every backup and **RESET WINDOWS DEFAULT**.

## <samp>[ FEEDBACK ]</samp>

Report bugs in [issues](https://github.com/arsenzaaa/DEVICE-TWEAKER/issues/new?template=bug_report.yml) with the log from the `logs` folder, your Windows version and PC configuration. To help with auto-optimization, send a [hardware report](https://github.com/arsenzaaa/DEVICE-TWEAKER/issues/new?template=hardware_report.yml) with three logs: before applying, right after and after a restart. Updates are posted in [Telegram](https://t.me/arsenzaa).

<p align="center"><sub><a href="https://arsenzaaa.github.io/DEVICE-TWEAKER/en/">Site</a> · <a href="./docs/SETTINGS.en.md">Settings guide</a> · <a href="./docs/DEVELOPMENT.en.md">Development and build</a> · <a href="./docs/THIRD_PARTY_NOTICES.md">Third-party components</a> · <a href="./.github/SECURITY.md">Security</a> · <a href="./LICENSE">GPL-3.0</a></sub></p>
