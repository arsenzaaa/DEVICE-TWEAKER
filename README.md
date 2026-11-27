# DEVICE TWEAKER

**English** · [Русский](./README.ru.md)

Windows hardware interrupt and device tuning in one place — MSI/MSI-X, interrupt affinity, ReservedCpuSets, USB IMOD, RSS, and NIC ITR, with CPU-topology auto-optimization, live readback, and safe backups.

**[Download for Windows](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest/download/DEVICE.TWEAKER.SelfContained.exe)** · [Guide](./docs/guide/README.md) · [Screenshots](./docs/SCREENSHOTS.md) · [Changelog](./CHANGELOG.md)

## Features

- **Interrupts & affinity:** configure MSI/MSI-X, MSI Limit, IRQ Priority, CPU affinity masks, and ReservedCpuSets, with live IRQ mode readback across PCI devices.
- **Topology-aware auto-optimization:** automatically place GPU, input, and network interrupts on optimal physical cores (P/E cores, CCD/CCX, SMT) — or fine-tune any device manually with **APPLY**.
- **USB IMOD & network tuning:** set xHCI IMOD per controller, device role, or interrupter with full readback; configure RSS queues/base CPU and read, apply, or save NIC ITR on supported Intel and Realtek adapters.
- **Input, power & safe rollback:** measure real-time mouse polling rate, toggle Mouse Throttle and power saving, and restore previous settings from automatic backups.

## Getting started

Requires **Windows 10/11 x64** and administrator access. The main download includes .NET. The smaller build requires .NET 10 Desktop Runtime x64 and is available on the [release page](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest).

1. Run the EXE as administrator and select a device from the list.
2. Adjust its settings and click **APPLY**, or click **AUTO-OPTIMIZATION** to configure devices automatically.
3. Restart Windows after changing MSI, interrupt affinity or ReservedCpuSets. Check the result with the [verification guide](./docs/guide/10-verification.md).

The application uses DTIMOD for USB IMOD, NIC ITR and reading PCI configuration. The [driver guide](./docs/guide/11-driver.md) covers loading requirements and possible anti-cheat conflicts.

## Documentation

- [Settings reference](./docs/SETTINGS.md)
- [Build from source and QA](./docs/DEVELOPMENT.md)

## Feedback and support

[Report a bug](https://github.com/arsenzaaa/DEVICE-TWEAKER/issues/new?template=bug_report.yml) · [Hardware report](https://github.com/arsenzaaa/DEVICE-TWEAKER/issues/new?template=hardware_report.yml) · [News on Telegram](https://t.me/arsenzaa) · [Support on Boosty](https://boosty.to/arsenza)

[Contributing](./.github/CONTRIBUTING.md) · [GPL-3.0](./LICENSE) · [Third-party components](./docs/THIRD_PARTY_NOTICES.md) · [Security](./.github/SECURITY.md)
