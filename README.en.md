# DEVICE TWEAKER

[Download](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.4) · [Screenshots](./SCREENSHOTS.en.md) · [Changelog](./CHANGELOG.en.md) · [Русский](./README.md) · [Telegram](https://t.me/arsenzaa)

DEVICE TWEAKER is a utility for configuring device interrupts on Windows 10/11 x64.

MSI Utility V3 + Interrupt Affinity Policy Tool (IAPT) + ReservedCpuSets + IMOD / ITR controls in one interface with intelligent auto-optimization.

![DEVICE TWEAKER interface](./assets/screenshots/showcase_intel_14900k_full_en.png)

[More screenshots](./SCREENSHOTS.en.md)

## Features

- Automatically detects GPUs / USB controllers / network adapters / audio controllers / storage controllers.
- Automatically identifies which USB controller hosts the mouse / keyboard / gamepad / microphone / audio. If multiple devices share a controller, all detected roles are displayed.
- Automatically detects and displays P-Core / E-Core, SMT / HT, AMD CCD / CCX, and available CPPC rankings. Shows which physical core each logical thread belongs to.
- Supports intelligent auto-optimization with a planning engine that accounts for P-Core / E-Core, SMT / HT, CPPC rankings, CCD / CCX, device count, and device roles. A discrete GPU gets two distinct physical cores; USB / network / audio devices have separate allocation rules.
- Supports MSI Mode / MSI Limit / IRQ Priority / IRQ Policy / Interrupt Affinity settings.
- Displays the MSI Mode registry setting separately from the current interrupt mode in IRQ Mode. On opening, it uses an already loaded DTIMOD.sys or Windows resources without loading the driver. If Windows resources cannot distinguish MSI from MSI-X, it shows MSI/MSI-X. The READ IRQ MODE button at the top reads the modes of all displayed PCI devices through DTIMOD.sys, loading the driver if needed. HW lists the modes supported by the device.
- Supports reserving selected CPUs through ReservedCpuSets.
- Supports RSS settings on NDIS network adapters: base CPU, queue count, and RSS/IRQ mode.
- Reads and displays IMOD on supported xHCI USB controllers. Set one value for the whole controller or different values for individual interrupters. Uses the custom DTIMOD.sys driver for reads and writes, then reads the values again after writing. Supports saving custom values for the next Windows sign-in.
- Reads and displays NIC ITR on supported network adapters. Queue and RX/TX information depends on the controller family. Change current values and save a profile for automatic reapplication at the next Windows sign-in.
- Displays mouse / keyboard Polling Rate information and supports measuring Raw Input event frequency.
- Supports RawMouseThrottleDuration settings on compatible Windows 11 builds, with the current value displayed in the interface.
- Supports power-saving settings for USB controllers / root USB hubs / wired network adapters.
- Supports creating backups, restoring a selected backup, deleting old backups, and settings reset.
- Writes detailed logs of device detection and settings changes (what was applied / what was skipped / why).

## Download

The [release](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.4) has two builds:

- **DEVICE.TWEAKER.exe** requires the [.NET 8 Desktop Runtime x64](https://dotnet.microsoft.com/download/dotnet/8.0).
- **DEVICE.TWEAKER.SelfContained.exe** includes .NET, so no separate installation is needed.

Both builds include the DTIMOD.sys driver and KDU loader. The SHA-256 hash of each EXE is listed in SHA256SUMS.txt.

## Usage

Download the build you need and run the program as administrator.

For intelligent auto-optimization, click **AUTO-OPTIMIZATION**. If eligible xHCI USB controllers are detected, the program offers to include IMOD tuning before applying settings. NIC ITR is configured manually.

For manual configuration, choose the settings and click **APPLY**. This button applies interrupt, ReservedCpuSets, RSS, Mouse Throttle, power-saving, and USB IMOD settings.

IMOD / NIC ITR also have separate buttons:

- IMOD: **SET** applies settings and saves values that differ from the defaults for the next Windows sign-in.
- NIC ITR: **SET** changes current values; **SAVE** stores a profile for the next Windows sign-in.
- **CHECK** reads current IMOD / NIC ITR values.

**REFRESH** detects devices again. To roll back settings, click **RESTORE** and select the backup you need.

After **APPLY** / **AUTO-OPTIMIZATION**, restart the PC to finish applying the settings.

The program creates a backup before manual apply and reset. Intelligent auto-optimization preserves the original settings.

Detailed settings description: [SETTINGS.en.md](./DEVICE%20TWEAKER/docs/SETTINGS.en.md).

## Bugs / testing

Report bugs or other problems in [issues](https://github.com/arsenzaaa/DEVICE-TWEAKER/issues/new?template=bug_report.yml). Include a log, the program version, Windows build, and PC configuration. Also describe what you did and what happened.

To help test intelligent auto-optimization, send logs in this order:

1. Before intelligent auto-optimization (or after settings reset).
2. After intelligent auto-optimization.
3. After restarting the PC and reopening the program.

Project updates and discussion: [Telegram](https://t.me/arsenzaa).

## Source code

[Build and checks](./DEVICE%20TWEAKER/README.en.md) · [Changelog](./CHANGELOG.en.md) · [GPL-3.0](./LICENSE)
