# <samp>[ CHANGELOG ]</samp>

[Русский](./CHANGELOG.md) / [DEVICE TWEAKER](./README.en.md)

## <samp>[ 0.0.4 ]</samp>

[Download 0.0.4](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.4). Changes since 0.0.3.

### <samp>MSI / IRQ Mode</samp>

- Added IRQ Mode. The MSI Mode registry setting, current interrupt mode and supported HW modes are displayed separately.
- On opening, the interrupt mode is read from allocated Windows resources or an already loaded DTIMOD.sys. Opening and normal list refresh do not load the driver. The interface indicates when the data cannot distinguish MSI from MSI-X.
- Added one READ IRQ MODE button for all displayed PCI devices. It reads MSI / MSI-X Enable through DTIMOD.sys, loading the driver if needed.
- Reading runs in the background. Results from an earlier request are not inserted into a refreshed device list.

### <samp>USB IMOD / NIC ITR</samp>

- CHECK in the IMOD block reveals the values and table. Viewing the result no longer requires enabling the IMOD apply checkbox.
- Background IMOD reads use only an available driver; they do not stage or replace its file.
- When restoring a saved IMOD / NIC ITR script, its code is updated to the current template while retaining device values. Required loader files are prepared before restoration.

### <samp>Backups / power saving</samp>

- Fixed USB Selective Suspend restoration: Windows Power API replaces direct registry writes. AC and battery values are restored separately to the power plan saved in the backup.
- Power-plan values are read back after restoration. An inactive plan is not activated.
- Improved WMI power-saving verification: Windows is queried again after a write. Unsupported functions, query failures, write failures and readback mismatches are distinguished.
- Backup dialogs separate the filename from the message. The full path can be viewed and copied, and a separate button opens the backup folder.
- The restore menu stays open after deleting backups. The extra deletion-result dialog has been removed.

### <samp>Devices / Polling Rate / RSS</samp>

- Fixed mouse frequency display: a confirmed Raw Input measurement replaces the USB endpoint maximum. Low keyboard activity does not replace the device frequency.
- Removed duplicate polling labels. USB roles are ordered as mouse, keyboard, gamepad, microphone, audio and camera.
- RSS now distinguishes unsupported functions, missing data, read failures, apply failures and skipped settings. A match with active settings is reported only from retrieved values.
- Removed unwarranted warnings for Wi-Fi and the default Windows CPU selection. Wi-Fi remains skipped. Active RSS queries have a longer timeout.

### <samp>Interface / RU / EN</samp>

- Reworked result, confirmation, backup-location and system-check dialogs. The main message is separated from technical details.
- Fixed clipped field text, dialog wrapping and block layout. Current and saved value labels have been clarified.
- Improved dropdowns and keyboard navigation. Home / End / Page Up / Page Down and Ctrl+Z in an editor are handled by that editor.
- Removed the white block flash during IRQ Mode reads. Fixed resource disposal order when closing the window.
- Updated RU/EN wording, backup-reason labels and the modified-settings indicator. Fixed language switching for dynamic strings and Power Saving.
- Updated device-category colours to make GPU and network easier to distinguish.

### <samp>Logs / system check</samp>

- Added an application build identifier while keeping version 0.0.4. Logs can distinguish different EXEs of the same version.
- Logs now include the IRQ Mode read source and time, RSS query duration and detailed read / apply statuses.
- Added a system check before USB IMOD tuning: Blocklist / HVCI / Microsoft Defender and FACEIT / Vanguard service states.

### <samp>Source / development</samp>

- Extracted the CPU planner, MSI / affinity writer, restoration, NIC ITR and polling logic from MainForm. Existing CPU allocation rules are retained.
- Moved checks and the test panel to dev/qa. The panel is included only in a separate developer build, not the normal EXE.
- Expanded the development workbench with reproducible scenarios, GUI and transition captures, RU/EN, layout-scale models, test files and frame comparison. Autonomous runs use a hidden desktop without applying Windows settings.
- Added test-sensitivity checks using deliberate faults and SHA-256 verification of sources / built components. Updated build documentation, the program description and screenshots.

## <samp>[ 0.0.3 ]</samp>

[Release](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.3) · October 1, 2026

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


## <samp>[ 0.0.2 ]</samp>

[Release](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.2) · January 15, 2026

After 0.0.1:

- Added UI scaling according to Windows DPI.
- Improved P/E-core detection using EfficiencyClass and SMT for the CPU list and intelligent auto-optimization.
- Changed CPU allocation order across device roles in intelligent auto-optimization.
- Removed duplicate USB-role log lines for alternative keys of one controller.

## <samp>[ 0.0.1 ]</samp>

[Release](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.1) · January 14, 2026

First public release for Windows 10/11 x64: MSI Mode and Limit, IRQ Priority, Interrupt Affinity, ReservedCpuSets, RSS base CPU, and xHCI IMOD through WinIO. The dark interface lists devices and lets you search them.
