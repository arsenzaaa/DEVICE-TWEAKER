<h1 align="center">DEVICE TWEAKER</h1>

<p align="center">
  <strong>Device interrupt configuration for Windows 10/11 x64.</strong><br>
  MSI Utility V3 + IAPT + ReservedCpuSets + IMOD / ITR<br>
  in one interface with intelligent auto-optimization.
  <br><a href="./README.md">Русский</a>
</p>

## Download

| Build | Requirements |
| :--- | :--- |
| [**Standalone**](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/download/v0.0.4/DEVICE.TWEAKER.SelfContained.exe) | Includes .NET; no separate installation needed. |
| [**Standard**](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/download/v0.0.4/DEVICE.TWEAKER.exe) | Requires the [.NET 8 Desktop Runtime x64](https://dotnet.microsoft.com/download/dotnet/8.0). |

Both builds include DTIMOD.sys. Files and SHA-256 checksums are available in [release 0.0.4](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.4).

## Interface

![DEVICE TWEAKER / USB IMOD / Intel Core i9-14900K](./assets/screenshots/showcase_intel_14900k_full_en.png)

[All screenshots](./SCREENSHOTS.en.md)

## Features

**Intelligent auto-optimization**

Automatically assigns devices to CPUs based on P-Core / E-Core, SMT / HT, CPPC rankings, AMD CCD / CCX and each device's role. A discrete GPU gets two distinct physical cores; USB, network and audio devices use their own allocation rules.

**MSI / IRQ / Interrupt Affinity**

- Supports MSI Mode / MSI Limit / IRQ Priority / IRQ Policy / Interrupt Affinity settings.
- Displays the current interrupt mode in IRQ Mode. Supported modes are listed separately in HW.
- Supports CPU reservation through ReservedCpuSets.

**USB IMOD / NIC ITR / RSS**

- Reads and displays IMOD on supported xHCI controllers. Set one value for the controller or individual values per interrupter.
- Reads and configures NIC ITR on supported network adapters. Queue and RX/TX information depends on the adapter model.
- Saves IMOD / NIC ITR profiles for the next Windows sign-in.
- Configures the base CPU, queue count and RSS/IRQ mode on wired NDIS network adapters.

**Device detection and Polling Rate**

Automatically detects GPUs, USB controllers, network adapters, audio and storage controllers. Shows which USB controller hosts the mouse, keyboard, gamepad, microphone and audio. Displays mouse / keyboard Polling Rate information and measures Raw Input event frequency.

**Other settings**

- Supports RawMouseThrottleDuration settings on compatible Windows 11 builds.
- Supports power-saving settings for USB controllers, root hubs and wired network adapters.
- Creates backups, restores a selected backup, deletes old backups and resets settings.
- Writes detailed logs: what was detected, applied or skipped, and why.

[Settings explained](./DEVICE%20TWEAKER/docs/SETTINGS.en.md)

## Usage

1. Download a build and run it as administrator.
2. For intelligent auto-optimization, click **AUTO-OPTIMIZATION**. If eligible xHCI controllers are detected, the program offers to include USB IMOD. NIC ITR is configured manually.
3. For manual configuration, choose the settings and click **APPLY**. A backup is created before applying; intelligent auto-optimization preserves the original settings.
4. Restart the PC after applying. To restore a selected backup, use **RESTORE**.

<details>
<summary><strong>Separate IMOD / NIC ITR buttons</strong></summary>

- IMOD: **SET** applies settings and saves values that differ from the defaults for the next Windows sign-in.
- NIC ITR: **SET** changes current values; **SAVE** stores a profile for the next Windows sign-in.
- **CHECK** reads current IMOD / NIC ITR values.
- **REFRESH** detects devices again. The main **APPLY** button also applies USB IMOD alongside device settings.

</details>

<details>
<summary><strong>How IRQ Mode is read</strong></summary>

MSI Mode shows the registry setting; IRQ Mode shows the current interrupt mode.

On opening, the program uses an already loaded DTIMOD.sys or Windows resources without loading the driver. If Windows resources cannot distinguish MSI from MSI-X, it displays MSI/MSI-X.

The shared **READ IRQ MODE** button reads all displayed PCI device modes through DTIMOD.sys, loading the driver if needed. Supported modes are listed separately in HW.

</details>

## Bugs / testing

Report bugs or other problems in [issues](https://github.com/arsenzaaa/DEVICE-TWEAKER/issues/new?template=bug_report.yml). Include a log, the program version, Windows build and PC configuration. Describe what you did and what happened.

<details>
<summary><strong>Logs for testing intelligent auto-optimization</strong></summary>

1. Before intelligent auto-optimization or after resetting settings.
2. After applying it.
3. After restarting the PC and reopening the program.

</details>

## Development

[Build](./DEVICE%20TWEAKER/README.en.md) · [QA / test panel](./dev/qa/README.md) · [Changelog](./CHANGELOG.en.md) · [GPL-3.0](./LICENSE)

Project updates and discussion: [Telegram](https://t.me/arsenzaa).
