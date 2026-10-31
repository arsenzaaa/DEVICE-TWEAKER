# Changelog

[Русский](./CHANGELOG.md) · [Project page](./README.en.md)

Notable changes to DEVICE TWEAKER by version. Every setting is explained in the [settings guide](./docs/SETTINGS.en.md).

## [0.0.4] — October 31, 2026

Changes since 0.0.3.

### Highlights

- IMOD / NIC ITR startup via Task Scheduler: applied at sign-in, after resume from sleep and when USB devices are attached. The open script works without the EXE.
- DTIMOD.sys writes only to memory Windows allocated to the selected device.
- RSS is spread across physical cores; NIC ITR turns off the driver's own moderation.
- The program and the startup script share one register-write path and one interrupter map.
- New IRQ Mode, the READ IRQ MODE button, and a verified backup before settings are applied.

### Interrupt modes

- Added IRQ Mode. The MSI Mode registry setting, current interrupt mode and supported HW modes are displayed separately.
- On opening, the interrupt mode is read from allocated Windows resources or an already loaded DTIMOD.sys. Opening and normal list refresh do not load the driver. The interface indicates when the data cannot distinguish MSI from MSI-X.
- Added one **READ IRQ MODE** button for all PCI devices in the list. It reads MSI / MSI-X Enable through DTIMOD.sys, loading the driver if needed.
- IRQ Mode reads run in the background. Results from an earlier request are discarded after the device list is refreshed.

### IMOD and NIC ITR

- **CHECK** in the IMOD block reveals the values and table. Viewing the result no longer requires enabling the IMOD apply checkbox.
- Background IMOD reads no longer stage or replace the driver file. They use an already loaded DTIMOD.sys.
- IMOD / NIC ITR startup now uses Task Scheduler: the `\DEVICE TWEAKER\IMOD` task runs as SYSTEM and applies the settings at sign-in, after resume from sleep, and when a USB/HID device is attached. Previously, startup ran only at sign-in.
- Startup files live in `%ProgramData%\DEVICE TWEAKER\Startup` and only administrators and SYSTEM can modify them. `ApplyIMOD.ps1` stays an open script that works without the EXE; values are kept in `profile.json`, the runtime, driver and loader SHA-256 are verified before each run, and logs are kept for 14 days. The previous startup entry in AppData and the Run key is migrated automatically.
- The program and the startup script share one register-write and interrupter-map code path. Fixed a mismatch: the startup script took a device's interrupter from the Slot Context instead of the TRBs of its transfer ring, so after a restart it could assign roles differently from what the program showed.
- Current IMOD values show the read time; after a device is attached or removed, the map is marked "devices changed".
- **SAVE** for NIC ITR turns off the driver's own adaptive moderation in the registry (`*InterruptModeration`, and `ITR` on Intel), which otherwise overwrote the register. **DELETE** restores the driver value. The Interrupt Moderation line shows the state.
- The program and the startup script no longer load DTIMOD.sys at the same time.
- Backups store the startup profile; backups with the old script are converted on restore without losing values.

### Backups and restoration

- Fixed USB Selective Suspend restoration: Windows Power API replaces direct registry writes. AC and battery values are restored separately to the power plan saved in the backup.
- Power-plan values are verified by reading them back after restoration; the active plan is not switched.
- Backup dialogs separate the filename from the message. The full path can be viewed and copied, and a separate button opens the backup folder.
- Deleting a backup no longer closes the restore menu or opens a separate result dialog.

### Power saving

- Power-saving writes through WMI now read the value back from Windows. Unsupported functions, query / write errors and mismatched values are reported separately.

### Device information

- Fixed mouse Polling Rate: after a confirmed Raw Input measurement, the interface shows the measured frequency even when it is below the USB endpoint maximum. Low keyboard activity does not replace the device frequency.
- Removed duplicate polling labels. USB roles are ordered as mouse, keyboard, gamepad, microphone, audio and camera.

### RSS

- RSS diagnostics distinguish unsupported functions, missing data, read failures, apply failures and skipped settings. Saved and active parameters are marked as matching only after a confirmed read.
- Removed unwarranted warnings for Wi-Fi and adapters where Windows selects RSS CPUs. Wi-Fi settings remain skipped. Active RSS queries have a longer timeout.
- Fixed RSS CPU selection: processors are taken one per physical core, skipping SMT siblings, and `*MaxRssProcessors` and `*RssMaxProcNumber` are written together with the base CPU.
- RSS queue counts are limited to the values the driver offers and to the number of MSI-X vectors.

### Interface / navigation

- Reworked result, confirmation, backup-location and system-check dialogs. The main message is separated from technical details.
- Fixed clipped field text, dialog wrapping and block layout. Current and saved value labels have been clarified.
- Improved dropdowns and keyboard navigation. Home / End / Page Up / Page Down and Ctrl+Z in an editor are handled by that editor.
- Removed the white block flash during IRQ Mode reads. Fixed resource disposal order when closing the window.
- Updated RU/EN wording, backup-reason labels and the modified-settings indicator. Fixed language switching for dynamic strings and Power Saving.
- Updated device-category colours to make GPU and network easier to distinguish.
- Russian filter buttons no longer lose their spaces after switching language.

### Diagnostics

- Added an application build identifier while keeping version 0.0.4. Logs can distinguish different EXEs of the same version.
- Logs now include the IRQ Mode read source and time, RSS query duration and detailed read / apply statuses.
- Added a system check before USB IMOD tuning: Blocklist / HVCI / Microsoft Defender and FACEIT / Vanguard service states.
- Registry, WMI and USB topology read failures are logged as warnings instead of being skipped silently. Numbers, dates and register values in logs are formatted the same way under any Windows locale.

### Development and testing

- Extracted the CPU planner, MSI / affinity writer, restoration, NIC ITR and polling logic from MainForm. Existing CPU allocation rules are retained.
- Moved checks and the test panel to dev/qa. The panel is included only in a separate developer build, not the normal EXE.
- Expanded the development workbench with reproducible scenarios, GUI and transition captures, RU/EN, layout-scale models, test files and frame comparison. Autonomous runs use a hidden desktop without applying Windows settings.
- Added test-sensitivity checks using deliberate faults and SHA-256 verification of sources / built components. Updated build documentation, the program description and screenshots.
- DTIMOD.sys writes only to memory Windows allocated to the named PCI device; the former general physical-memory write was removed. The range check runs on the driver model together with a deliberate fault.

## [0.0.3] — October 1, 2026

A larger update after 0.0.2:

- Moved xHCI IMOD from WinIO to DTIMOD.sys and KDU. You can inspect values per interrupter, configure the whole controller or individual indexes, map USB devices, and read registers back after writing. Saved-profile reapplication uses the new driver too.
- Added NIC ITR for recognized Intel and Realtek controllers: register reads and writes, queue information, and a profile for the next Windows sign-in. Writing is unavailable for unknown PCI IDs.
- Reworked intelligent auto-optimization around physical cores, P/E cores, SMT/HT, CPPC, and AMD CCD/CCX. A discrete GPU gets two different physical cores. If the device role or CPU topology cannot be identified reliably, the assignment is skipped with a reason in the report.
- Added RSS queue count alongside base CPU. Saved settings and active Get-NetAdapterRss state appear separately.
- Separated the MSI Mode setting, the estimate based on allocated IRQ numbers, and hardware capability. The IRQ estimate is only a hint and cannot distinguish MSI from MSI-X.
- Added Unlimited for MSI Limit and validation of numeric values against detected device capabilities. An empty mask for a regular device restores MachineDefault and removes AssignmentSetOverride.
- Fixed the AllClose and All mappings to DevicePolicy values 1 and 3. Their labels were reversed in 0.0.2.
- Added mouse and keyboard event-rate estimates from Raw Input and RawMouseThrottleDuration controls on supported Windows 11 builds. This setting affects background Raw Input listeners, not USB polling rate.
- Added power settings for USB controllers and root hubs. Wired network adapters gained control over whether Windows may turn them off to save power.
- Improved HDMI/DisplayPort and S/PDIF audio detection. USB controllers with a custom Interrupt Affinity policy stay visible even without a detected HID role.
- Added RU/EN switching without restart and reworked device display, IMOD/ITR panels, and category navigation.
- Reworked backups: choose a folder, restore a specific backup, delete old backups, or reset settings. Manual apply, IMOD/ITR writes, and reset require a validated backup; intelligent auto-optimization saves the original settings first.
- Reports now show applied, skipped, and failed actions. A normal device refresh does not load the driver; CHECK rereads registers.

## [0.0.2] — January 15, 2026

After 0.0.1:

- Added UI scaling according to Windows DPI.
- Improved P/E-core detection using EfficiencyClass and SMT for the CPU list and intelligent auto-optimization.
- Changed CPU allocation order across device roles in intelligent auto-optimization.
- Removed duplicate USB-role log lines for alternative keys of one controller.

## [0.0.1] — January 14, 2026

First public release for Windows 10/11 x64: MSI Mode and Limit, IRQ Priority, Interrupt Affinity, ReservedCpuSets, RSS base CPU, and xHCI IMOD through WinIO. The dark interface lists devices and lets you search them.

[0.0.4]: https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.4
[0.0.3]: https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.3
[0.0.2]: https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.2
[0.0.1]: https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.1
