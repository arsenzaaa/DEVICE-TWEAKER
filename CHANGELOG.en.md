# Changelog

[Русский](./CHANGELOG.md) · [DEVICE TWEAKER](./README.en.md)

## [0.0.3](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.3) · October 1, 2026

Changes compared with the `v0.0.2` source.

### Interrupts and CPU placement

- Added RSS queue count control through `*NumRssQueues` alongside the existing base CPU setting. Saved values and active `Get-NetAdapterRss` state are shown separately.
- Reworked automatic placement around physical cores, P/E-cores, SMT/HT, CPPC, and AMD CCD/CCX. The GPU is assigned two physical cores. An assignment is skipped with an explanation when CPU topology or device role cannot be determined reliably.
- Separated the `MSISupported` setting, the interrupt mode estimate based on allocated IRQ numbers, and PCI hardware capability. The IRQ estimate remains a heuristic and cannot determine whether the driver uses MSI or MSI-X.
- Added an explicit **Unlimited** display and numeric validation for MSI Limit. An empty Interrupt Affinity mask removes `DevicePolicy` and `AssignmentSetOverride` to restore Windows policy.

### Hardware registers and input

- Added `DTIMOD.sys` and xHCI IMOD controls for the controller and individual interrupters, USB device mapping, register readback after a write, and saved profiles for later application.
- Added NIC ITR controls for recognized Intel and Realtek controllers, including register reads, queue settings, and saved profiles. Hardware writes are unavailable for unknown PCI IDs.
- Added mouse and keyboard event rate estimates based on Raw Input and `RawMouseThrottleDuration` controls on supported Windows 11 builds. This setting concerns throttling of background Raw Input listeners, not USB polling rate.
- Refined HDMI/DisplayPort and S/PDIF audio detection. USB controllers with a custom Interrupt Affinity policy remain visible without a detected HID role.

### Interface and restore

- Added RU/EN switching without restart. Reworked device cards, IMOD/ITR panels, and category navigation.
- Added backups with a storage location choice and restore of a selected snapshot. Manual apply, IMOD/ITR writes, and reset require a validated backup; automatic optimization requires an original settings snapshot.
- Reports distinguish applied, skipped, and failed actions. Ordinary device refresh does not start the hardware access driver; **CHECK** rereads registers.

## [0.0.2](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.2) · January 14, 2026

Changes compared with the `v0.0.1` source.

- Added UI scaling according to Windows DPI.
- Refined P/E-core classification using `EfficiencyClass` and SMT presence for CPU display and placement.
- Reworked CPU allocation order across device roles in automatic optimization.
- Removed duplicate USB role log entries for alternative keys of the same controller.

## [0.0.1](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.1) · January 14, 2026

- First public release: a Windows 10/11 x64 GUI for MSI/MSI-X, Interrupt Affinity, and device settings.
