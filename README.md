# DEVICE TWEAKER

**English** · [Русский](./README.ru.md)

A utility for configuring MSI, interrupt affinity, ReservedCpuSets and USB IMOD on Windows. It brings together controls familiar from MSI Utility V3 and Interrupt Affinity Policy Tool and adds auto-optimization based on CPU topology.

**[Download for Windows](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest/download/DEVICE.TWEAKER.SelfContained.exe)** · [Guide](./docs/guide/README.md) · [Screenshots](./docs/SCREENSHOTS.md) · [Changelog](./CHANGELOG.md)

## Features

- Adjust MSI Mode, MSI Limit, IRQ Priority and interrupt affinity. Read the active IRQ mode and configure ReservedCpuSets in a separate panel.
- Select cores for graphics cards, input devices and network adapters with **AUTO-OPTIMIZATION**. Adjust individual device settings with **APPLY**.
- Configure USB IMOD by controller, device role or interrupter index. Choose the processors and the number of queues for RSS. Read, apply or save NIC ITR on supported Intel and Realtek adapters.
- Measure the mouse event rate, adjust Mouse Throttle and power saving, and restore backed-up settings.

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
