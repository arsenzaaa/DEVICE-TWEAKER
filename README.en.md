<div align="right"><a href="./README.md">Русский</a></div>

<h1 align="center"><img src="./assets/icon.png" width="128" height="128" alt="DEVICE TWEAKER"></h1>

<p align="center">Device interrupt tuning for Windows 10 / 11 x64, matched to the CPU topology</p>

<p align="center"><sub>MSI · affinity · IRQ priority · ReservedCpuSets · RSS · USB&nbsp;IMOD · NIC&nbsp;ITR</sub></p>

<p align="center"><a href="https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest"><b>Download 0.0.4</b></a> · <a href="./docs/SETTINGS.en.md">Settings guide</a> · <a href="./CHANGELOG.en.md">Changelog</a></p>

<a href="./assets/screenshots/showcase_intel_14900k_full_en.png"><img src="./assets/screenshots/showcase_intel_14900k_full_en.png" alt="DEVICE TWEAKER: a USB controller on Intel Core i9-14900K" width="100%"></a>

<table>
  <tr>
    <td width="33%"><a href="./assets/screenshots/showcase_amd_imod_full_en.png"><img src="./assets/screenshots/thumbs/showcase_amd_imod_full_en.png" alt="USB IMOD on AMD Ryzen 9 9950X3D"></a></td>
    <td width="33%"><a href="./assets/screenshots/showcase_gpu_full_en.png"><img src="./assets/screenshots/thumbs/showcase_gpu_full_en.png" alt="GPU interrupts"></a></td>
    <td width="33%"><a href="./assets/screenshots/showcase_nic_full_en.png"><img src="./assets/screenshots/thumbs/showcase_nic_full_en.png" alt="NIC ITR and RSS"></a></td>
  </tr>
  <tr>
    <td><sub>USB IMOD, Ryzen 9 9950X3D</sub></td>
    <td><sub>GPU, a pair of physical cores</sub></td>
    <td><sub>NIC ITR and RSS, Intel I226-V</sub></td>
  </tr>
</table>

<sub>Screenshots show test devices and values.</sub>

## What it changes

| Setting | Written to |
| --- | --- |
| MSI Mode, MSI Limit | `MSISupported`, `MessageNumberLimit` in `Interrupt Management\MessageSignaledInterruptProperties` |
| IRQ Priority | `DevicePriority` in `Interrupt Management\Affinity Policy` |
| Policy, Interrupt Affinity | `DevicePolicy`, `AssignmentSetOverride` in the same key |
| RSS | `*RssBaseProcNumber`, `*NumRssQueues`, `*MaxRssProcessors`, `*RssMaxProcNumber` in the adapter's class key |
| ReservedCpuSets | `ReservedCpuSets` in `Session Manager\kernel` |
| Mouse Throttle | `RawMouseThrottleDuration` in `HKCU\Control Panel\Mouse`, Windows 11 |
| Power saving | `PnPCapabilities`, `SelectiveSuspendEnabled`, `EnhancedPowerManagementEnabled`, USB selective suspend in the power plan |
| USB IMOD | the IMOD register of every xHCI interrupter |
| NIC ITR | the controller's moderation registers: EITR / ITR on Intel, IntrMit on Realtek |

Registry settings take effect after a restart. IMOD and NIC ITR registers change immediately and hold until a restart or sleep. A verified backup is made before settings are applied. IMOD and NIC ITR registers are read back after every write. The log records what was applied, what was skipped and why.

## Auto-optimization

Builds a plan from the CPU topology and device roles.

- A discrete GPU gets two different physical cores. Two SMT siblings of one core do not count as two cores.
- The controller with the mouse and keyboard, the network adapter and audio follow their own rules. On CPUs with several CCDs a target CCD is chosen.
- Storage controllers get MSI and high priority; their affinity stays with Windows. An integrated GPU's affinity goes back to Windows; Wi-Fi is left untouched.
- If the topology is uncertain, nothing is written.

The original state is saved before the first auto-optimization and can be restored at any time.

## USB IMOD and NIC ITR

IMOD is the interrupt moderation interval of an xHCI interrupter in 250 ns steps: `0xC8` = 50 µs, `0x0` turns moderation off. The value is set for the whole controller, per interrupter or per device role. The device → interrupter map is read from the xHCI contexts in memory.

NIC ITR is supported on Intel I210 / I211, I219, I225 / I226, I350, 82576 / 82580 and Realtek RTL8111 / 8168, RTL8125 / 8126. Saving a profile turns off the driver's own moderation in the registry (`*InterruptModeration`), which would otherwise overwrite the register.

The registers are written by the project's own driver, **DTIMOD.sys**, and only into memory Windows allocated to the selected device. Register values last until a restart or sleep, so the saved profile is applied by the Task Scheduler task `\DEVICE TWEAKER\IMOD`: at sign-in, after resume from sleep and when a USB device is attached. The task runs the open `ApplyIMOD.ps1` script from `%ProgramData%\DEVICE TWEAKER\Startup` and works without the EXE.

## Install

| File | Requires |
| --- | --- |
| [DEVICE.TWEAKER.SelfContained.exe](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest/download/DEVICE.TWEAKER.SelfContained.exe) | nothing, .NET is built in |
| [DEVICE.TWEAKER.exe](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest/download/DEVICE.TWEAKER.exe) | [.NET 8 Desktop Runtime x64](https://dotnet.microsoft.com/download/dotnet/8.0) |

Run as administrator. DTIMOD.sys and the KDU loader are built into both. SHA-256 checksums are in `SHA256SUMS.txt` on the [release page](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest).

1. **AUTO-OPTIMIZATION** builds the plan and applies it. If a suitable xHCI controller is found, it offers to include USB IMOD. For manual settings, use **APPLY**.
2. Restart the PC.
3. **CHECK** in the IMOD or NIC ITR block reads the actual register values.

## Rollback

**RESTORE → ORIGINAL** returns the registry to the state before the first auto-optimization and removes the startup task. The same menu restores any backup or resets to Windows defaults. Backups are kept in the `Backups` folder next to the EXE.

## FAQ

<details>
<summary>Why a driver, and why an antivirus may flag KDU</summary>
<br>

IMOD and NIC ITR registers cannot be written from user mode. DTIMOD.sys has no public signature, so [KDU](./src/IMOD/KDU/README.md) loads it through a signed vulnerable driver, and antiviruses classify such loaders as HackTool. The DTIMOD sources and the pinned KDU version are in this repository, and the binaries are checked by SHA-256. The driver is only needed for IMOD, NIC ITR and **READ IRQ MODE**.
</details>

<details>
<summary>HVCI, Vulnerable Driver Blocklist, FACEIT and Vanguard</summary>
<br>

Memory Integrity (HVCI) and the Microsoft Vulnerable Driver Blocklist block loading through KDU; the FACEIT and Vanguard anti-cheats may conflict with it. Before USB IMOD is enabled, the program shows their state in the **SYSTEM CHECK**. Registry settings do not depend on the driver.
</details>

## Feedback

Report bugs in [issues](https://github.com/arsenzaaa/DEVICE-TWEAKER/issues/new?template=bug_report.yml) with the log from the `logs` folder next to the EXE, your Windows build and PC configuration. For auto-optimization, three logs help most: before applying, right after, and after a restart. News: [Telegram](https://t.me/arsenzaa).

[Settings guide](./docs/SETTINGS.en.md) · [Development](./docs/DEVELOPMENT.en.md) · [Third-party components](./docs/THIRD_PARTY_NOTICES.md) · [Security](./.github/SECURITY.md) · [GPL-3.0](./LICENSE)
