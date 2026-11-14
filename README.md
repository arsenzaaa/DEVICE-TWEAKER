# DEVICE TWEAKER

**English** · [Русский](./README.ru.md)

**DEVICE TWEAKER** brings the MSI and interrupt-affinity controls familiar from MSI Utility V3 and Interrupt Affinity Policy Tool together with ReservedCpuSets and USB IMOD in one Windows utility. Configure devices manually or let automatic optimization select CPU cores based on your processor's topology.

[Download](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest/download/DEVICE.TWEAKER.SelfContained.exe) · [Guide](./docs/guide/README.md) · [Screenshots](./docs/SCREENSHOTS.md) · [Changelog](./CHANGELOG.md)

## Features

Configure MSI Mode, MSI Limit, IRQ Priority and the CPU cores assigned to device interrupts. Saved settings are shown separately from the current IRQ mode, so you can check how the device driver uses them.

**Automatic optimization** selects cores for the GPU, input devices and network adapters. It accounts for physical cores and SMT siblings, P/E cores, CCDs and available CPPC rankings. You can also configure each device manually and manage ReservedCpuSets separately.

USB IMOD can be configured for an entire controller, by device role or by interrupter index. Network controls include RSS processor and queue selection, with NIC ITR available for supported Intel and Realtek adapters. IMOD and NIC ITR values can be read back from the hardware and saved in startup profiles.

DEVICE TWEAKER preserves original device settings and supports restoring saved backups. Detailed logs record the application build, devices, requested values and the outcome of each stage. The utility also includes device power settings, mouse event-rate measurement and Mouse Throttle controls.

## Getting started

Requires **Windows 10/11 x64** and administrator access. The [self-contained build](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest/download/DEVICE.TWEAKER.SelfContained.exe) includes .NET and runs without installation. The [smaller framework-dependent build](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest/download/DEVICE.TWEAKER.exe) requires [.NET 10 Desktop Runtime x64](https://dotnet.microsoft.com/download/dotnet/10.0).

1. Run the EXE as administrator and select a device.
2. Adjust its settings and click **APPLY**, or use **AUTO-OPTIMIZATION** to select settings automatically.
3. Restart Windows after changing MSI or Interrupt Affinity. See the [verification guide](./docs/guide/10-verification.md) for checking the result.

Saved backups are available under **RESTORE**. Operation logs are stored in the `logs` folder next to the EXE. All release files and their checksums are available on the [download page](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest).

IMOD, NIC ITR and PCI configuration reads use DTIMOD.sys, loaded through KDU. The driver may conflict with anti-cheat software, including EAC, Vanguard and FACEIT. Closing the application does not unload it; see the [driver guide](./docs/guide/11-driver.md) for loading requirements and startup profile behavior.

## Documentation

The [guide](./docs/guide/README.md) explains interrupt handling and device configuration. The [settings reference](./docs/SETTINGS.md) covers interface fields, applying changes and restoring settings. Build instructions and developer checks are covered in the [development guide](./docs/DEVELOPMENT.md).

## Feedback and support

Report bugs and suggest improvements through [Issues](https://github.com/arsenzaaa/DEVICE-TWEAKER/issues). You can share results from your hardware using the [hardware report form](https://github.com/arsenzaaa/DEVICE-TWEAKER/issues/new?template=hardware_report.yml), or get in touch on [Telegram](https://t.me/arsenzaa).

Support the development of DEVICE TWEAKER on [Boosty](https://boosty.to/arsenza).

[Contributing](./.github/CONTRIBUTING.md) · [GPL-3.0](./LICENSE) · [Third-party components](./docs/THIRD_PARTY_NOTICES.md) · [Security](./.github/SECURITY.md)
