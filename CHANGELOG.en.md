# Changelog

[Русский](./CHANGELOG.md) · [DEVICE TWEAKER](./README.en.md)

## [0.0.4](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.4) · October 1, 2026

- Corrected the Russian `Affinity Mask` help text and added a test for translation consistency.
- Fixed GUI checks and the release script under Windows PowerShell 5.1 so Cyrillic text is read correctly.
- Release builds are tied to a tag and source commit: tests run, both EXEs are built, and their SHA-256 hashes are checked before a draft release is created.
- The main screenshot shows expanded IMOD values on Intel Core i9-14900K. The gallery shows AMD with the IMOD table, Intel with the GPU, and AMD with the network adapter. Devices and values are simulated; no settings were written to the system.
- Removed graphic buttons from the start of the README, leaving a plain heading and text links.
- Updated the application version, executable metadata, and current-release links. Device-tuning behavior did not change.

## [0.0.3](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.3) · October 1, 2026

Changes compared with the `v0.0.2` source.

The `v0.0.3` tag was moved to a later commit after the EXEs were published. As a result, the source at that tag does not exactly match the published EXEs. This release remains available as an archive; the [v0.0.4](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.4) description identifies its exact source commit.

### Interrupts and CPU placement

- Added RSS queue count control through `*NumRssQueues` alongside the existing base CPU setting. Saved values and active `Get-NetAdapterRss` state are shown separately.
- Reworked automatic placement around physical cores, P/E-cores, SMT/HT, CPPC, and AMD CCD/CCX. The GPU is assigned two physical cores. An assignment is skipped with an explanation when CPU topology or device role cannot be determined reliably.
- Separated the `MSISupported` setting, the interrupt mode estimate based on allocated IRQ numbers, and PCI hardware capability. The IRQ estimate remains a heuristic and cannot determine whether the driver uses MSI or MSI-X.
- Added an explicit **Unlimited** display and MSI Limit validation against detected device capabilities before writing. An empty mask for a regular device sets `DevicePolicy` to `MachineDefault` and removes `AssignmentSetOverride`.
- Corrected the `AllClose` and `All` mappings for `DevicePolicy` when reading and writing values: `1` and `3`, respectively. Version 0.0.2 had these labels reversed.

### Hardware registers and input

- Moved xHCI IMOD control from WinIO to `DTIMOD.sys` and KDU. Added per-interrupter values, USB device mapping, and register readback to the existing controller-wide control. Saved-profile reapplication was adapted to the new driver.
- Added NIC ITR controls for recognized Intel and Realtek controllers, including register reads, queue settings, and saved profiles. Hardware writes are unavailable for unknown PCI IDs.
- Added mouse and keyboard event rate estimates based on Raw Input and `RawMouseThrottleDuration` controls on supported Windows 11 builds. This setting concerns throttling of background Raw Input listeners, not USB polling rate.
- Refined HDMI/DisplayPort and S/PDIF audio detection. USB controllers with a custom Interrupt Affinity policy remain visible without a detected HID role.
- Added power settings for USB controllers and their root hubs, including USB selective suspend. Wired network adapters gained control over whether Windows may turn off the device to save power.

### Interface and restore

- Added RU/EN switching without restart. Reworked device cards, IMOD/ITR panels, and category navigation.
- Added backups with a storage location choice, restore of a selected snapshot, deletion of old backups, and settings reset. Manual apply, IMOD/ITR writes, and reset require a validated backup; automatic optimization requires an original settings snapshot.
- Reports distinguish applied, skipped, and failed actions. Ordinary device refresh does not start the hardware access driver; **CHECK** rereads registers.

## [0.0.2](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.2) · January 15, 2026

Changes compared with the `v0.0.1` source.

- Added UI scaling according to Windows DPI.
- Refined P/E-core classification using `EfficiencyClass` and SMT presence for CPU display and placement.
- Reworked CPU allocation order across device roles in automatic optimization.
- Removed duplicate USB role log entries for alternative keys of the same controller.

## [0.0.1](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.1) · January 14, 2026

- First public release for Windows 10/11 x64: MSI Mode and Limit, IRQ Priority, Interrupt Affinity, ReservedCpuSets, RSS base CPU, and xHCI IMOD through WinIO.
