# Changelog

[Русский](./CHANGELOG.md) · [DEVICE TWEAKER](./README.en.md)

## [0.0.4](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.4)

A major update since 0.0.3.

- Added IRQ Mode, separate from the MSI Mode setting and device capabilities. Opening the app reads Windows data without loading the driver. The shared READ IRQ MODE button checks PCI MSI / MSI-X Enable bits through DTIMOD.sys.
- Fixed USB selective suspend restore: settings are written through the Windows Power API to the saved power plan, with separate AC and DC values. The active plan stays selected; written values are verified.
- Fixed mouse rate display: a confirmed Raw Input rate replaces the maximum USB endpoint rate. Keyboard readings account for key activity. Duplicate polling labels have been removed.
- Updated IMOD / NIC ITR handling for saved profiles. Concurrent requests stage the driver file sequentially; logs distinguish file preparation from connection to a loaded driver.
- Refined RSS diagnostics: skipped settings, unsupported features, read failures and apply failures are reported separately. Successful reads require returned data. Increased the active RSS query timeout.
- Reworked result, restore and system check dialogs. Backup paths are separate from messages; deleting backups keeps the restore menu open.
- Improved input fields, dropdowns and keyboard navigation. Fixed clipped text, white flashes during refresh and an error when closing the window.
- Updated RU/EN wording, backup labels and the changed-settings indicator. Fixed language switching for dynamic text and the Power Saving state.
- Ordered USB roles: mouse, keyboard, gamepad, microphone, audio, camera. Updated category colors and device capability display.
- Added a Windows settings and anti-cheat service summary before USB IMOD configuration.
- Added application build identity, IRQ Mode data sources, RSS query duration and query results to logs.
- Separated restore, CPU planning, MSI / affinity writes, IMOD, NIC ITR and polling from MainForm. CPU allocation rules are preserved.
- Moved checks and the test panel to dev/qa. Production EXEs exclude the panel. The separate developer build does not apply Windows settings; the workbench records screenshots, transitions and reproducible scenario results.
- Updated the README, screenshot gallery and build documentation. Git preserves exact file bytes for SHA-256 manifest checks.

## [0.0.3](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.3) · October 1, 2026

A larger update after 0.0.2:

- Moved xHCI IMOD from WinIO to DTIMOD.sys and KDU. You can inspect values per interrupter, configure the whole controller or individual indexes, map USB devices, and read registers back after writing. Saved-profile reapplication uses the new driver too.
- Added NIC ITR for recognized Intel and Realtek controllers: register reads and writes, queue information, and a profile for the next Windows sign-in. Writing is unavailable for unknown PCI IDs.
- Reworked intelligent auto-optimization around physical cores, P/E cores, SMT/HT, CPPC, and AMD CCD/CCX. A discrete GPU gets two different physical cores. If the device role or CPU topology cannot be identified reliably, the assignment is skipped with a reason in the report.
- Added RSS queue count alongside base CPU. Saved settings and active Get-NetAdapterRss state appear separately.
- Separated the MSI Mode setting, the estimate based on allocated IRQ numbers, and hardware capability. The IRQ estimate is only a hint and cannot distinguish MSI from MSI-X.
- Added Unlimited for MSI Limit and validation of numeric values against detected device capabilities. An empty mask for a regular device restores MachineDefault and removes AssignmentSetOverride.
- Fixed the AllClose and All mappings to DevicePolicy values 1 and 3. Their labels were reversed in 0.0.2.
- Added mouse and keyboard event-rate estimates from Raw Input and RawMouseThrottleDuration controls on supported Windows 11 builds. This setting affects background Raw Input listeners, not USB polling rate.
- Added power settings for USB controllers and root hubs. Wired network adapters gained control over whether Windows may turn them off to save power.
- Improved HDMI/DisplayPort and S/PDIF audio detection. USB controllers with a custom Interrupt Affinity policy stay visible even without a detected HID role.
- Added RU/EN switching without restart and reworked device display, IMOD/ITR panels, and category navigation.
- Reworked backups: choose a folder, restore a specific backup, delete old backups, or reset settings. Manual apply, IMOD/ITR writes, and reset require a validated backup; intelligent auto-optimization saves the original settings first.
- Reports now show applied, skipped, and failed actions. A normal device refresh does not load the driver; CHECK rereads registers.


## [0.0.2](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.2) · January 15, 2026

After 0.0.1:

- Added UI scaling according to Windows DPI.
- Improved P/E-core detection using EfficiencyClass and SMT for the CPU list and intelligent auto-optimization.
- Changed CPU allocation order across device roles in intelligent auto-optimization.
- Removed duplicate USB-role log lines for alternative keys of one controller.

## [0.0.1](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.1) · January 14, 2026

First public release for Windows 10/11 x64: MSI Mode and Limit, IRQ Priority, Interrupt Affinity, ReservedCpuSets, RSS base CPU, and xHCI IMOD through WinIO. The dark interface lists devices and lets you search them.
