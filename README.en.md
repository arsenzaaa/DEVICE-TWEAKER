<p align="center">
  <a href="https://github.com/arsenzaaa/DEVICE-TWEAKER">
    <picture>
      <source media="(prefers-color-scheme: dark)" srcset="./assets/readme/wordmark-dark.svg">
      <img src="./assets/readme/wordmark-light.svg" alt="DEVICE TWEAKER" width="560">
    </picture>
  </a>
</p>

<p align="center">Device interrupts on Windows 10 / 11 x64, tuned to your CPU topology.</p>

<p align="center">
  <a href="https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest"><b>Download 0.0.4</b></a> ·
  <a href="./docs/guide/README.en.md">Guide</a> ·
  <a href="./CHANGELOG.en.md">Changelog</a> ·
  <a href="https://t.me/arsenzaa">Telegram</a> ·
  <a href="./README.md">Русский</a>
</p>

<p align="center"><a href="./assets/screenshots/showcase_intel_14900k_full_en.png"><picture><source srcset="./assets/screenshots/readme/showcase_intel_14900k_full_en.png 1x, ./assets/screenshots/framed/showcase_intel_14900k_full_en.png 2x"><img src="./assets/screenshots/readme/showcase_intel_14900k_full_en.png" alt="DEVICE TWEAKER: a USB controller on Intel Core i9-14900K, IMOD and interrupt affinity" width="838"></picture></a></p>

---

## <samp>[ ABOUT ]</samp>

ISRs / DPCs run above any thread and do not know which cores the game is using. A driver on the core of an important thread (GAME / RENDER / RHI) interrupts it -> the thread waits or moves to another core and loses its L1 / L2 -> frame time grows. Two drivers on one core delay each other's ISRs / DPCs; on one SMT-SET they share cache and EXECUTION-UNITS.

DEVICE TWEAKER spreads interrupts across physical cores: MSI Utility V3 + Interrupt Affinity Policy Tool (IAPT) + ReservedCpuSets + RSS + USB IMOD / NIC ITR in one interface, with intelligent auto-optimization.

## <samp>[ FEATURES ]</samp>

- **Auto-optimization.** A plan built from the topology (P / E-cores, SMT / HT, CPPC, CCD / CCX) and device roles: the GPU on two different physical cores, drivers sharing neither cores nor SMT-SETs, core 0 kept free when possible. [Chapter 01](./docs/guide/01-interrupts.en.md)
- **MSI and MSI-X.** MSI Mode, MSI Limit, the current IRQ Mode and the device's capabilities shown separately; **READ IRQ MODE** reads the real mode from PCI configuration space. [Chapter 02](./docs/guide/02-msi.en.md)
- **Interrupt Affinity and priority.** Policy, mask and IRQ Priority per device, with a CPU list showing physical cores, CCDs and CPPC ranks. [Chapter 03](./docs/guide/03-affinity.en.md)
- **USB IMOD.** Moderation of each xHCI interrupter by device role; the device -> interrupter map is read from controller memory on every apply. [Chapter 05](./docs/guide/05-usb-imod.en.md)
- **RSS and NIC ITR.** Base CPU and queues of NDIS adapters, one per physical core, and the moderation registers of Intel and Realtek controllers. [Chapter 06](./docs/guide/06-network.en.md)
- **ReservedCpuSets, Mouse Throttle, Polling Rate, power saving.** [Chapters 07–09](./docs/guide/README.en.md)
- **Startup without the EXE.** A Task Scheduler task applies the IMOD / NIC ITR profile at sign-in, after sleep and when USB devices connect; the script is readable and its files are checked by SHA-256. [Chapter 11](./docs/guide/11-driver.en.md)
- **Backups and logs.** A verified backup before every apply; **RESTORE -> ORIGINAL** returns the state before the first auto-optimization; the log shows what was applied, what was skipped and why.

## <samp>[ INSTALLATION ]</samp>

Download one of the files from the [latest release](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest):

- `DEVICE.TWEAKER.SelfContained.exe` — .NET included;
- `DEVICE.TWEAKER.exe` — requires the [.NET 8 Desktop Runtime x64](https://dotnet.microsoft.com/download/dotnet/8.0).

Run as administrator -> **AUTO-OPTIMIZATION**, or manual settings and **APPLY** -> restart. **CHECK** in the IMOD or NIC ITR block shows the actual register values.

> [!IMPORTANT]
> IMOD and NIC ITR are written by the custom DTIMOD.sys driver, and only to memory that Windows assigned to the selected device. The driver has no public signature and is loaded by [KDU](./src/IMOD/KDU/README.md), so antivirus software flags the program as HackTool, and HVCI and the Microsoft Vulnerable Driver Blocklist block the loading. All other settings work without the driver. Details in [chapter 11](./docs/guide/11-driver.en.md).

SHA-256 hashes of the files are in `SHA256SUMS.txt` on the release page:

```powershell
(Get-FileHash .\DEVICE.TWEAKER.SelfContained.exe -Algorithm SHA256).Hash
```

## <samp>[ GUIDE ]</samp>

How it works in Windows and what the program does.

1. [Interrupts: ISR, DPC and cores](./docs/guide/01-interrupts.en.md)
2. [MSI and MSI-X](./docs/guide/02-msi.en.md)
3. [Affinity, policies and interrupt priority](./docs/guide/03-affinity.en.md)
4. [Graphics card](./docs/guide/04-gpu.en.md)
5. [USB controllers and IMOD](./docs/guide/05-usb-imod.en.md)
6. [Network: RSS, NDIS and NIC ITR](./docs/guide/06-network.en.md)
7. [CPU topology and ReservedCpuSets](./docs/guide/07-cpu-topology.en.md)
8. [Input: Polling Rate, Raw Input and Mouse Throttle](./docs/guide/08-input.en.md)
9. [Device power saving](./docs/guide/09-power.en.md)
10. [Verification: traces and logs](./docs/guide/10-verification.en.md)
11. [DTIMOD.sys, KDU and startup](./docs/guide/11-driver.en.md)

A reference for every interface element is in the [settings guide](./docs/SETTINGS.en.md).

## <samp>[ FEEDBACK ]</samp>

Report bugs in [issues](https://github.com/arsenzaaa/DEVICE-TWEAKER/issues/new?template=bug_report.yml) with the log from the `logs` folder next to the EXE. Results on your hardware go to a [hardware report](https://github.com/arsenzaaa/DEVICE-TWEAKER/issues/new?template=hardware_report.yml): three logs, before applying, right after and after a restart. Updates are posted in [Telegram](https://t.me/arsenzaa).

Building from source: [docs/DEVELOPMENT.en.md](./docs/DEVELOPMENT.en.md). License: [GPL-3.0](./LICENSE).
