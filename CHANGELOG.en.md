# Changelog

[Русский](./CHANGELOG.md) · [DEVICE TWEAKER](./README.en.md)

## [0.0.4](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.4) · October 1, 2026

No new device settings were added in this release. Changes:

- Fixed the Russian Affinity Mask help text and added a translation check.
- Fixed GUI checks and the release script under Windows PowerShell 5.1 so Cyrillic text is read correctly.
- Tied release builds to a tag and source commit. Tests run, both EXEs are built, and their SHA-256 hashes are checked before a draft release is created.
- Updated the application version and release information.
- Recaptured the main screenshot with IMOD details open on Intel Core i9-14900K. The gallery shows AMD IMOD, Intel GPU, and AMD NIC ITR. All values are simulated; no settings were written to the system.
- Removed graphic buttons from the README and kept plain text links.

Device-tuning behavior did not change in this release.

## [0.0.3](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.3) · October 1, 2026

A larger update after 0.0.2:

- Moved xHCI IMOD from WinIO to DTIMOD.sys and KDU. You can inspect values per interrupter, configure the whole controller or individual indexes, map USB devices, and read registers back after writing. Saved-profile reapplication uses the new driver too.
- Added NIC ITR for recognized Intel and Realtek controllers: register reads and writes, queue information, and a profile for the next Windows sign-in. Writing is unavailable for unknown PCI IDs.
- Reworked auto-optimization around physical cores, P/E cores, SMT/HT, CPPC, and AMD CCD/CCX. The GPU gets two different physical cores. If the device role or CPU topology cannot be identified reliably, the assignment is skipped with a reason in the report.
- Added RSS queue count alongside base CPU. Saved settings and active Get-NetAdapterRss state appear separately.
- Separated the MSI Mode setting, the estimate based on allocated IRQ numbers, and hardware capability. The IRQ estimate is only a hint and cannot distinguish MSI from MSI-X.
- Added Unlimited for MSI Limit and validation of numeric values against detected device capabilities. An empty mask for a regular device restores MachineDefault and removes AssignmentSetOverride.
- Fixed the AllClose and All mappings to DevicePolicy values 1 and 3. Their labels were reversed in 0.0.2.
- Added mouse and keyboard event-rate estimates from Raw Input and RawMouseThrottleDuration controls on supported Windows 11 builds. This setting affects background Raw Input listeners, not USB polling rate.
- Added power settings for USB controllers and root hubs. Wired network adapters gained control over whether Windows may turn them off to save power.
- Improved HDMI/DisplayPort and S/PDIF audio detection. USB controllers with a custom Interrupt Affinity policy stay visible even without a detected HID role.
- Added RU/EN switching without restart and reworked device display, IMOD/ITR panels, and category navigation.
- Reworked backups: choose a folder, restore a specific backup, delete old backups, or reset settings. Manual apply, IMOD/ITR writes, and reset require a validated backup; auto-optimization saves the original settings first.
- Reports now show applied, skipped, and failed actions. A normal device refresh does not load the driver; CHECK rereads registers.

Archive note: the v0.0.3 tag was moved after the EXEs had been published, so the source at that tag does not exactly match the published files. The [v0.0.4](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.4) description records its exact source commit.

## [0.0.2](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.2) · January 15, 2026

After 0.0.1:

- Added UI scaling according to Windows DPI.
- Improved P/E-core detection using EfficiencyClass and SMT for the CPU list and auto-optimization.
- Changed CPU allocation order across device roles in auto-optimization.
- Removed duplicate USB-role log lines for alternative keys of one controller.

## [0.0.1](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.1) · January 14, 2026

First public release for Windows 10/11 x64: MSI Mode and Limit, IRQ Priority, Interrupt Affinity, ReservedCpuSets, RSS base CPU, and xHCI IMOD through WinIO. The dark interface lists devices and lets you search them.
