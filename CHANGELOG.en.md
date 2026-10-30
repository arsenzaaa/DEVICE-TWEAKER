# <samp>[ CHANGELOG ]</samp>

[Русский](./CHANGELOG.md) / [DEVICE TWEAKER](./README.en.md)

## <samp>[ 0.0.4 ]</samp>

[Download 0.0.4](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.4). Changes since 0.0.3.

<table width="100%">
<tr><th align="left" width="860"><strong>Interrupt modes</strong></th></tr>
<tr><td width="860">

<ul>
<li>Added IRQ Mode. The MSI Mode registry setting, current interrupt mode and supported HW modes are displayed separately.</li>
<li>On opening, the interrupt mode is read from allocated Windows resources or an already loaded DTIMOD.sys. Opening and normal list refresh do not load the driver. The interface indicates when the data cannot distinguish MSI from MSI-X.</li>
<li>Added one <strong>READ IRQ MODE</strong> button for all PCI devices in the list. It reads MSI / MSI-X Enable through DTIMOD.sys, loading the driver if needed.</li>
<li>IRQ Mode reads run in the background. Results from an earlier request are discarded after the device list is refreshed.</li>
</ul>

</td></tr>
</table>

<table width="100%">
<tr><th align="left" width="860"><strong>IMOD and NIC ITR</strong></th></tr>
<tr><td width="860">

<ul>
<li><strong>CHECK</strong> in the IMOD block reveals the values and table. Viewing the result no longer requires enabling the IMOD apply checkbox.</li>
<li>Background IMOD reads no longer stage or replace the driver file. They use an already loaded DTIMOD.sys.</li>
<li>IMOD / NIC ITR startup now uses Task Scheduler: the <code>\DEVICE TWEAKER\IMOD</code> task runs as SYSTEM and applies the settings at sign-in, after resume from sleep, and when a USB/HID device is attached. Previously, startup ran only at sign-in.</li>
<li>Startup files live in <code>%ProgramData%\DEVICE TWEAKER\Startup</code> and only administrators and SYSTEM can modify them. <code>ApplyIMOD.ps1</code> stays an open script that works without the EXE; values are kept in <code>profile.json</code>, the runtime, driver and loader SHA-256 are verified before each run, and logs are kept for 14 days. The previous startup entry in AppData and the Run key is migrated automatically.</li>
<li>The program and the startup script share one register-write and interrupter-map code path. Fixed a mismatch: the startup script took a device's interrupter from the Slot Context instead of the TRBs of its transfer ring, so after a restart it could assign roles differently from what the program showed.</li>
<li>Current IMOD values show the read time; after a device is attached or removed, the map is marked "devices changed".</li>
<li><strong>SAVE</strong> for NIC ITR turns off the driver's own adaptive moderation in the registry (<code>*InterruptModeration</code>, and <code>ITR</code> on Intel), which otherwise overwrote the register. <strong>DELETE</strong> restores the driver value. The Interrupt Moderation line shows the state.</li>
<li>The program and the startup script no longer load DTIMOD.sys at the same time.</li>
<li>Backups store the startup profile; backups with the old script are converted on restore without losing values.</li>
</ul>

</td></tr>
</table>

<table width="100%">
<tr><th align="left" width="860"><strong>Backups and restoration</strong></th></tr>
<tr><td width="860">

<ul>
<li>Fixed USB Selective Suspend restoration: Windows Power API replaces direct registry writes. AC and battery values are restored separately to the power plan saved in the backup.</li>
<li>Power-plan values are verified by reading them back after restoration; the active plan is not switched.</li>
<li>Backup dialogs separate the filename from the message. The full path can be viewed and copied, and a separate button opens the backup folder.</li>
<li>Deleting a backup no longer closes the restore menu or opens a separate result dialog.</li>
</ul>

</td></tr>
</table>

<table width="100%">
<tr><th align="left" width="860"><strong>Power saving</strong></th></tr>
<tr><td width="860">

<ul>
<li>Power-saving writes through WMI now read the value back from Windows. Unsupported functions, query / write errors and mismatched values are reported separately.</li>
</ul>

</td></tr>
</table>

<table width="100%">
<tr><th align="left" width="860"><strong>Device information</strong></th></tr>
<tr><td width="860">

<ul>
<li>Fixed mouse Polling Rate: after a confirmed Raw Input measurement, the interface shows the measured frequency even when it is below the USB endpoint maximum. Low keyboard activity does not replace the device frequency.</li>
<li>Removed duplicate polling labels. USB roles are ordered as mouse, keyboard, gamepad, microphone, audio and camera.</li>
</ul>

</td></tr>
</table>

<table width="100%">
<tr><th align="left" width="860"><strong>RSS</strong></th></tr>
<tr><td width="860">

<ul>
<li>RSS diagnostics distinguish unsupported functions, missing data, read failures, apply failures and skipped settings. Saved and active parameters are marked as matching only after a confirmed read.</li>
<li>Removed unwarranted warnings for Wi-Fi and adapters where Windows selects RSS CPUs. Wi-Fi settings remain skipped. Active RSS queries have a longer timeout.</li>
<li>Fixed RSS CPU selection: processors are taken one per physical core, skipping SMT siblings, and <code>*MaxRssProcessors</code> and <code>*RssMaxProcNumber</code> are written together with the base CPU.</li>
<li>RSS queue counts are limited to the values the driver offers and to the number of MSI-X vectors.</li>
</ul>

</td></tr>
</table>

<table width="100%">
<tr><th align="left" width="860"><strong>Interface / navigation</strong></th></tr>
<tr><td width="860">

<ul>
<li>Reworked result, confirmation, backup-location and system-check dialogs. The main message is separated from technical details.</li>
<li>Fixed clipped field text, dialog wrapping and block layout. Current and saved value labels have been clarified.</li>
<li>Improved dropdowns and keyboard navigation. Home / End / Page Up / Page Down and Ctrl+Z in an editor are handled by that editor.</li>
<li>Removed the white block flash during IRQ Mode reads. Fixed resource disposal order when closing the window.</li>
<li>Updated RU/EN wording, backup-reason labels and the modified-settings indicator. Fixed language switching for dynamic strings and Power Saving.</li>
<li>Updated device-category colours to make GPU and network easier to distinguish.</li>
<li>Russian filter buttons no longer lose their spaces after switching language.</li>
</ul>

</td></tr>
</table>

<table width="100%">
<tr><th align="left" width="860"><strong>Diagnostics</strong></th></tr>
<tr><td width="860">

<ul>
<li>Added an application build identifier while keeping version 0.0.4. Logs can distinguish different EXEs of the same version.</li>
<li>Logs now include the IRQ Mode read source and time, RSS query duration and detailed read / apply statuses.</li>
<li>Added a system check before USB IMOD tuning: Blocklist / HVCI / Microsoft Defender and FACEIT / Vanguard service states.</li>
<li>Registry, WMI and USB topology read failures are logged as warnings instead of being skipped silently. Numbers, dates and register values in logs are formatted the same way under any Windows locale.</li>
</ul>

</td></tr>
</table>

<table width="100%">
<tr><th align="left" width="860"><strong>Development and testing</strong></th></tr>
<tr><td width="860">

<ul>
<li>Extracted the CPU planner, MSI / affinity writer, restoration, NIC ITR and polling logic from MainForm. Existing CPU allocation rules are retained.</li>
<li>Moved checks and the test panel to dev/qa. The panel is included only in a separate developer build, not the normal EXE.</li>
<li>Expanded the development workbench with reproducible scenarios, GUI and transition captures, RU/EN, layout-scale models, test files and frame comparison. Autonomous runs use a hidden desktop without applying Windows settings.</li>
<li>Added test-sensitivity checks using deliberate faults and SHA-256 verification of sources / built components. Updated build documentation, the program description and screenshots.</li>
<li>DTIMOD.sys writes only to memory Windows allocated to the named PCI device; the former general physical-memory write was removed. The range check runs on the driver model together with a deliberate fault.</li>
</ul>

</td></tr>
</table>

## <samp>[ PREVIOUS VERSIONS ]</samp>

<details>
<summary><samp>0.0.3</samp></summary>

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

</details>

<details>
<summary><samp>0.0.2</samp></summary>

[Release](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.2) · January 15, 2026

After 0.0.1:

- Added UI scaling according to Windows DPI.
- Improved P/E-core detection using EfficiencyClass and SMT for the CPU list and intelligent auto-optimization.
- Changed CPU allocation order across device roles in intelligent auto-optimization.
- Removed duplicate USB-role log lines for alternative keys of one controller.

</details>

<details>
<summary><samp>0.0.1</samp></summary>

[Release](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.1) · January 14, 2026

First public release for Windows 10/11 x64: MSI Mode and Limit, IRQ Priority, Interrupt Affinity, ReservedCpuSets, RSS base CPU, and xHCI IMOD through WinIO. The dark interface lists devices and lets you search them.

</details>
