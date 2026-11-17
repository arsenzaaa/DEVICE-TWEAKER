# DEVICE TWEAKER

**English** · [Русский](./README.ru.md)

**DEVICE TWEAKER** combines the controls familiar from MSI Utility V3 and Interrupt Affinity Policy Tool with ReservedCpuSets and USB IMOD in one Windows utility.

**[Download for Windows](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest/download/DEVICE.TWEAKER.SelfContained.exe)** · [Guide](./docs/guide/README.md) · [Screenshots](./docs/SCREENSHOTS.md) · [Changelog](./CHANGELOG.md)

## Features

### Interrupts and CPU affinity

Each device card provides MSI Mode, MSI Limit, IRQ Priority and interrupt-affinity controls, together with the current IRQ mode. A separate panel configures ReservedCpuSets.

### Topology-aware automatic optimization

Automatic optimization assigns cores to graphics cards, input devices and network adapters using physical-core topology, SMT, P/E cores, CCDs and available CPPC rankings. Every assignment can be adjusted manually.

### USB and network tuning

USB IMOD supports controller-wide values, device-role profiles and individual interrupter settings. Network controls include RSS processor and queue selection, plus NIC ITR on supported Intel and Realtek adapters. IMOD and NIC ITR include register readback and startup profiles.

### Backups and diagnostics

Original device settings and later backups are available through RESTORE. Operation logs record the build, device, requested values and completed steps. Additional controls cover mouse event-rate measurement, Mouse Throttle and device power saving.

## Getting started

Requires **Windows 10/11 x64** and administrator access.

The **[self-contained build](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest/download/DEVICE.TWEAKER.SelfContained.exe)** includes .NET and needs no installation. The smaller [framework-dependent build](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest/download/DEVICE.TWEAKER.exe) requires [.NET 10 Desktop Runtime x64](https://dotnet.microsoft.com/download/dotnet/10.0).

1. Run the EXE as administrator and select a device.
2. Adjust its settings and click **APPLY**, or use **AUTO-OPTIMIZATION** to select settings automatically.
3. Restart Windows after changing MSI, Interrupt Affinity or ReservedCpuSets. See the [verification guide](./docs/guide/10-verification.md) for checking the result.

Open **RESTORE** to access saved backups. Logs are stored in the `logs` folder next to the EXE. All release files and checksums are on the [releases page](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest).

USB IMOD, NIC ITR and hardware PCI reads use DTIMOD.sys through KDU. The other settings work independently of this driver. DTIMOD can conflict with anti-cheat software; loading requirements and startup behavior are described in the [driver guide](./docs/guide/11-driver.md).

## Documentation

- [User guide](./docs/guide/README.md) - interrupt handling, CPU topology and device configuration.
- [Settings reference](./docs/SETTINGS.md) - interface controls, applying changes and restoring settings.
- [Development](./docs/DEVELOPMENT.md) - building from source and running focused checks with the QA workbench.

## Feedback and support

Report bugs and suggest improvements through [Issues](https://github.com/arsenzaaa/DEVICE-TWEAKER/issues). Include the relevant operation log after removing personal paths and names. Use the [hardware report form](https://github.com/arsenzaaa/DEVICE-TWEAKER/issues/new?template=hardware_report.yml) to share results from your system. Project news is posted in the [Telegram channel](https://t.me/arsenzaa).

Support the development of DEVICE TWEAKER on [Boosty](https://boosty.to/arsenza).

[Contributing](./.github/CONTRIBUTING.md) · [GPL-3.0](./LICENSE) · [Third-party components](./docs/THIRD_PARTY_NOTICES.md) · [Security](./.github/SECURITY.md)
