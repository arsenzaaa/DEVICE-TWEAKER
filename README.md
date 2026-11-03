<h1 align="center">
  <picture>
    <source media="(prefers-color-scheme: dark)" srcset="./assets/branding/wordmark-dark.svg">
    <img src="./assets/branding/wordmark-light.svg" width="320" height="97" alt="DEVICE TWEAKER">
  </picture>
</h1>

<p align="center">Device interrupt configuration for Windows.</p>

<p align="center"><a href="https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest/download/DEVICE.TWEAKER.SelfContained.exe"><strong>Download</strong></a> · <a href="./docs/README.md">Documentation</a> · <a href="./CHANGELOG.md">Changelog</a> · <a href="./docs/SCREENSHOTS.md">Screenshots</a></p>

<p align="center"><strong>English</strong> · <a href="./README.ru.md">Русский</a></p>

## Features

- **Device interrupts:** MSI Mode and Limit, IRQ Priority, CPU affinity, active IRQ mode and PCI capabilities. AUTO selects settings using the CPU topology.
- **USB:** xHCI IMOD by controller, device role or interrupter, with register readback and startup profiles.
- **Network:** RSS CPU allocation and queue counts, plus NIC ITR for supported Intel and Realtek adapters.

Also includes [ReservedCpuSets](./docs/guide/07-cpu-topology.md), [mouse event-rate measurement and Mouse Throttle](./docs/guide/08-input.md), [device power saving](./docs/guide/09-power.md), and settings backup and restore.

## Getting started

1. Download the [self-contained build](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest/download/DEVICE.TWEAKER.SelfContained.exe) and run the EXE as administrator. Requires Windows 10/11 x64; no installation is needed.
2. Select a device, edit the values and use **APPLY**. Use **AUTO-OPTIMIZATION** for automatic selection. The original settings are saved before writing.
3. Restart Windows after changing MSI or Interrupt Affinity. For IMOD/NIC ITR, **CHECK** reads the registers. Follow the [verification guide](./docs/guide/10-verification.md) to check the result.

Choose a saved backup under **RESTORE**. Operation logs are stored in `logs` next to the EXE.

The [framework-dependent build](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/download/v0.0.4/DEVICE.TWEAKER.exe) requires [.NET 10 Desktop Runtime x64](https://dotnet.microsoft.com/download/dotnet/10.0). Files and checksums are available on the [release page](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest).

IMOD, NIC ITR and PCI configuration reads use DTIMOD.sys loaded through KDU. Support depends on the hardware, device driver and Windows security settings. See the [driver guide](./docs/guide/11-driver.md) for loading requirements.

## Documentation

- [User guide](./docs/guide/README.md) - Windows interrupts and device configuration.
- [Settings reference](./docs/SETTINGS.md) - fields, controls, applying and restoring settings.
- [Build and development](./docs/DEVELOPMENT.md) - sources, dependencies and checks.
- [Contributing](./.github/CONTRIBUTING.md) - fixes, suggestions and hardware results.

## Support

[Report a bug](https://github.com/arsenzaaa/DEVICE-TWEAKER/issues/new?template=bug_report.yml) · [Hardware report](https://github.com/arsenzaaa/DEVICE-TWEAKER/issues/new?template=hardware_report.yml) · [Telegram](https://t.me/arsenzaa) · [Boosty](https://boosty.to/arsenza)

Include the application version, Windows build, device and steps to reproduce. Remove personal paths and identifiers before attaching a log.

---

[GPL-3.0](./LICENSE) · [Third-party components](./docs/THIRD_PARTY_NOTICES.md) · [Security](./.github/SECURITY.md)
