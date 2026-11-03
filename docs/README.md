# Documentation

[DEVICE TWEAKER](../README.md) / Documentation · [Русский](./README.ru.md)

Device configuration, Windows interrupt handling and DEVICE TWEAKER development.

## Getting started

- [Screenshots](./SCREENSHOTS.md) - GPU, USB and network settings.
- [Settings reference](./SETTINGS.md) - fields, controls, applying changes and restoring settings.
- [Guide](./guide/README.md) - 11 chapters from ISR/DPC and MSI to IMOD and RSS.
- [Logs and verification](./guide/10-verification.md) - saved settings, registers and system behaviour.

## By device

- **[Graphics card](./guide/04-gpu.md)** - physical cores, SMT and affinity policy.
- **[USB](./guide/05-usb-imod.md)** - xHCI controllers, interrupters and IMOD for mouse, keyboard and audio.
- **[Network](./guide/06-network.md)** - RSS, NDIS/NetAdapterCx and NIC ITR.
- **[Processor](./guide/07-cpu-topology.md)** - P/E cores, CPPC, CCD/CCX and ReservedCpuSets.
- **[Input](./guide/08-input.md)** - Polling Rate, Raw Input and Mouse Throttle.
- **[Power saving](./guide/09-power.md)** - USB selective suspend and device power controls.

## Driver and startup

[DTIMOD.sys, KDU and the PowerShell profile](./guide/11-driver.md): which features require the driver, how memory writes are restricted and how the profile works without the main EXE.

## For developers

[Build and code structure](./DEVELOPMENT.md) · [Local QA](../dev/qa/README.md) · [GUI workbench](../dev/qa/WORKBENCH.md) · [Prepare a release](./RELEASE.md) · [Contribute](../.github/CONTRIBUTING.md)

---

[Changelog](../CHANGELOG.md) · [Release notes](./releases/README.md) · [Third-party components](./THIRD_PARTY_NOTICES.md) · [Security policy](../.github/SECURITY.md)
