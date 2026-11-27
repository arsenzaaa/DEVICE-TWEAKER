# DEVICE TWEAKER

**English** · [Русский](./README.ru.md)

DEVICE TWEAKER is a Windows 10/11 x64 utility for configuring hardware interrupts and low-level device settings. It covers MSI/MSI-X, interrupt affinity, ReservedCpuSets, USB xHCI IMOD, RSS, and NIC ITR, with topology-aware auto-optimization, live readback, and automatic backups.

**[Download for Windows](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest/download/DEVICE.TWEAKER.SelfContained.exe)** · [Guide](./docs/guide/README.md) · [Screenshots](./docs/SCREENSHOTS.md) · [Changelog](./CHANGELOG.md)

[![DEVICE TWEAKER main window](./assets/screenshots/intel-affinity.en.png)](./docs/SCREENSHOTS.md)

## Features

- **Interrupts:** configure MSI/MSI-X, MSI Limit, IRQ Priority, CPU affinity masks, and ReservedCpuSets, with live PCI IRQ mode readback.
- **Auto-optimization:** place GPU, input, and network interrupts on optimal physical cores using CPU topology (P/E cores, CCD/CCX, SMT), or configure any device manually with **APPLY**.
- **USB xHCI IMOD:** set moderation intervals per controller, device role, or interrupter index with live register readback.
- **Networking:** configure RSS queues and base CPU, plus read, apply, or save hardware NIC ITR on supported Intel and Realtek adapters.
- **Input, power, rollback:** measure live mouse polling rate, configure Mouse Throttle and device power saving, and restore settings from automatic backups.

## Getting started

Requires **Windows 10/11 x64** and administrator access. The main download includes .NET. A smaller build requiring .NET 10 Desktop Runtime x64 is available on the [release page](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest).

1. Run the EXE as administrator and select a device from the list.
2. Adjust its settings and click **APPLY**, or click **AUTO-OPTIMIZATION** to configure devices automatically.
3. Restart Windows after changing MSI, interrupt affinity, or ReservedCpuSets. Verify the result with the [verification guide](./docs/guide/10-verification.md).

USB IMOD, NIC ITR, and PCI configuration readback use the DTIMOD driver. See the [driver guide](./docs/guide/11-driver.md) for loading requirements and possible anti-cheat conflicts.

## Documentation

- [Settings reference](./docs/SETTINGS.md)
- [Build from source and QA](./docs/DEVELOPMENT.md)

## Feedback and support

[Report a bug](https://github.com/arsenzaaa/DEVICE-TWEAKER/issues/new?template=bug_report.yml) · [Hardware report](https://github.com/arsenzaaa/DEVICE-TWEAKER/issues/new?template=hardware_report.yml) · [News on Telegram](https://t.me/arsenzaa) · [Support on Boosty](https://boosty.to/arsenza)

[Contributing](./.github/CONTRIBUTING.md) · [GPL-3.0](./LICENSE) · [Third-party components](./docs/THIRD_PARTY_NOTICES.md) · [Security](./.github/SECURITY.md)
