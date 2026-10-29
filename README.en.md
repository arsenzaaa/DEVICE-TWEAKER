<p align="right"><a href="./README.md">Русский</a></p>

<h1>
  <a href="https://github.com/arsenzaaa/DEVICE-TWEAKER#gh-dark-mode-only"><picture><source media="(max-width: 600px)" srcset="./assets/branding/hero-mobile-dark.svg"><img src="./assets/branding/hero-dark.svg" alt="DEVICE TWEAKER / Windows 10/11 x64 / 0.0.4" width="1080"></picture></a>
  <a href="https://github.com/arsenzaaa/DEVICE-TWEAKER#gh-light-mode-only"><picture><source media="(max-width: 600px)" srcset="./assets/branding/hero-mobile-light.svg"><img src="./assets/branding/hero-light.svg" alt="DEVICE TWEAKER / Windows 10/11 x64 / 0.0.4" width="1080"></picture></a>
</h1>

<p align="center"><strong>DEVICE TWEAKER is a utility for configuring device interrupts on Windows 10/11 x64.</strong></p>

MSI Utility V3 + Interrupt Affinity Policy Tool (IAPT) + ReservedCpuSets + IMOD / ITR controls in one interface with intelligent auto-optimization.

<p align="center">
  <a href="https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.4#gh-dark-mode-only"><img src="./assets/branding/download-en-dark.svg" alt="Download 0.0.4" width="170"></a>
  <a href="https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.4#gh-light-mode-only"><img src="./assets/branding/download-en-light.svg" alt="Download 0.0.4" width="170"></a>
  <a href="./src/docs/SETTINGS.en.md#gh-dark-mode-only"><img src="./assets/branding/settings-en-dark.svg" alt="Settings" width="140"></a>
  <a href="./src/docs/SETTINGS.en.md#gh-light-mode-only"><img src="./assets/branding/settings-en-light.svg" alt="Settings" width="140"></a>
</p>

## <samp>[ INTERFACE ]</samp>

<p align="center"><a href="./assets/screenshots/showcase_intel_14900k_full_en.png"><img src="./assets/screenshots/showcase_intel_14900k_full_en.png" alt="USB IMOD / Intel Core i9-14900K" width="860"></a></p>

<table width="100%">
<tr><th align="left" width="860"><samp>More configurations</samp></th></tr>
<tr><td width="860"><details><summary><samp>USB IMOD on AMD</samp></summary>
<p><a href="./assets/screenshots/showcase_amd_imod_full_en.png"><img src="./assets/screenshots/showcase_amd_imod_full_en.png" alt="USB IMOD on AMD" width="100%"></a></p>
</details></td></tr>
<tr><td width="860"><details><summary><samp>GPU interrupt settings</samp></summary>
<p><a href="./assets/screenshots/showcase_gpu_full_en.png"><img src="./assets/screenshots/showcase_gpu_full_en.png" alt="GPU interrupt settings" width="100%"></a></p>
</details></td></tr>
<tr><td width="860"><details><summary><samp>NIC ITR / RSS</samp></summary>
<p><a href="./assets/screenshots/showcase_nic_full_en.png"><img src="./assets/screenshots/showcase_nic_full_en.png" alt="NIC ITR / RSS" width="100%"></a></p>
</details></td></tr>
</table>

<sub>Screenshots show test devices and values.</sub>

## <samp>[ FEATURES ]</samp>

<table width="100%">
<tr><th align="left" width="860"><strong>Intelligent auto-optimization</strong></th></tr>
<tr><td width="860">
<p>Intelligent auto-optimization accounts for P-Core / E-Core, SMT / HT, CPPC rankings, CCD / CCX, device count, and device roles. A discrete GPU gets two distinct physical cores; USB / network / audio devices have separate allocation rules.</p><p align="center"><a href="./src/docs/SETTINGS.en.md#gh-dark-mode-only"><img src="./assets/branding/cpu-allocation-en-dark.svg" alt="Two physical cores for a discrete GPU" width="360"></a><a href="./src/docs/SETTINGS.en.md#gh-light-mode-only"><img src="./assets/branding/cpu-allocation-en-light.svg" alt="Two physical cores for a discrete GPU" width="360"></a></p>
</td></tr>
</table>

<table width="100%">
<tr><th align="left" width="860"><strong>Interrupt settings</strong></th></tr>
<tr><td width="860">
<p>Supports MSI Mode / MSI Limit / IRQ Priority / IRQ Policy / Interrupt Affinity settings.</p>
<p>Displays the MSI Mode registry setting separately from the current interrupt mode in IRQ Mode. HW lists the modes supported by the device.</p>
<p>On opening, it uses an already loaded DTIMOD.sys or Windows resources without loading the driver. If Windows resources cannot distinguish MSI from MSI-X, it shows MSI/MSI-X.</p>
<p><strong>READ IRQ MODE</strong> at the top reads the modes of all displayed PCI devices through DTIMOD.sys, loading the driver if needed.</p>
</td></tr>
</table>

<table width="100%">
<tr><th align="left" width="860"><strong>USB IMOD</strong></th></tr>
<tr><td width="860">
<p>Reads and displays IMOD on supported xHCI USB controllers. <strong>Devices</strong> sets a profile by USB role, <strong>XHCI</strong> uses one value for the controller, and <strong>Interrupters</strong> allows values by index.</p>
<p>Uses the custom DTIMOD.sys driver for reads and writes, then reads the values again after writing. Supports saving custom values for the next Windows sign-in.</p>
</td></tr>
</table>

<table width="100%">
<tr><th align="left" width="860"><strong>Network settings</strong></th></tr>
<tr><td width="860">
<p>Reads and displays NIC ITR on supported Intel / Realtek network adapters, including compatible Killer models. Queue and RX/TX information depends on the controller family. Change current values and save a profile for automatic reapplication at the next Windows sign-in.</p>
<p>Supports RSS settings on wired NDIS network adapters: base CPU, queue count, and RSS/IRQ mode.</p>
</td></tr>
</table>

<table width="100%">
<tr><th align="left" width="860"><strong>Device detection</strong></th></tr>
<tr><td width="860">
<p>Automatically detects GPUs / USB controllers / network adapters / audio controllers / storage controllers.</p>
<p>Automatically identifies which USB controller hosts the mouse / keyboard / gamepad / microphone / audio / camera. If multiple devices share a controller, all detected roles are displayed.</p>
</td></tr>
</table>

<table width="100%">
<tr><th align="left" width="860"><strong>CPU topology</strong></th></tr>
<tr><td width="860">
<p>Automatically detects and displays P-Core / E-Core, SMT / HT, AMD CCD / CCX, and available CPPC rankings. Shows which physical core each logical thread belongs to.</p>
</td></tr>
</table>

<table width="100%">
<tr><th align="left" width="860"><strong>CPU reservation</strong></th></tr>
<tr><td width="860">
<p>Supports reserving selected CPUs through ReservedCpuSets.</p>
</td></tr>
</table>

<table width="100%">
<tr><th align="left" width="860"><strong>Polling Rate</strong></th></tr>
<tr><td width="860">
<p>Displays mouse / keyboard Polling Rate information and supports measuring Raw Input event frequency. Keyboard event frequency depends on key activity.</p>
</td></tr>
</table>

<table width="100%">
<tr><th align="left" width="860"><strong>Mouse Throttle</strong></th></tr>
<tr><td width="860">
<p>Supports RawMouseThrottleDuration settings on compatible Windows 11 builds, with the current value displayed in the interface.</p>
</td></tr>
</table>

<table width="100%">
<tr><th align="left" width="860"><strong>Power saving</strong></th></tr>
<tr><td width="860">
<p>Supports power-saving settings for USB controllers / root USB hubs / wired network adapters.</p>
</td></tr>
</table>

<table width="100%">
<tr><th align="left" width="860"><strong>Interface / navigation</strong></th></tr>
<tr><td width="860">
<p>Supports device search, category navigation and RU/EN switching without restarting.</p>
</td></tr>
</table>

<table width="100%">
<tr><th align="left" width="860"><strong>Backups / logs</strong></th></tr>
<tr><td width="860">
<p>Supports creating backups, restoring a selected backup, deleting old backups, and settings reset.</p>
<p>Writes detailed logs of device detection and settings changes (what was applied / what was skipped / why).</p>
</td></tr>
</table>

<table width="100%">
<tr><th align="left" width="860"><strong>Download</strong></th></tr>
<tr><td width="860">

<p>The <a href="https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.4">release</a> has two builds:</p>
<p><strong>Standard build</strong><br/><sub>DEVICE.TWEAKER.exe</sub><br/>Requires the <a href="https://dotnet.microsoft.com/download/dotnet/8.0">.NET 8 Desktop Runtime x64</a>.</p><p><strong>Standalone build</strong><br/><sub>DEVICE.TWEAKER.SelfContained.exe</sub><br/>Includes .NET, so no separate installation is needed.</p>
<p>Both builds include the DTIMOD.sys driver and KDU loader. The SHA-256 hash of each EXE is listed in SHA256SUMS.txt.</p>

</td></tr>
</table>

<table width="100%">
<tr><th align="left" width="860"><strong>Usage</strong></th></tr>
<tr><td width="860">

<p>Download the build you need and run the program as administrator.</p>
<p>For intelligent auto-optimization, click <strong>AUTO-OPTIMIZATION</strong>. If eligible xHCI USB controllers are detected, the program offers to include IMOD tuning before applying settings. NIC ITR is configured manually.</p>
<p>For manual configuration, choose the settings and click <strong>APPLY</strong>. This button applies interrupt, ReservedCpuSets, RSS, Mouse Throttle, power-saving, and USB IMOD settings.</p>
<p>IMOD / NIC ITR also have separate buttons:</p>
<ul>
<li>IMOD: <strong>SET</strong> applies settings and saves values that differ from the defaults for the next Windows sign-in.</li>
<li>NIC ITR: <strong>SET</strong> changes current values; <strong>SAVE</strong> stores a profile for the next Windows sign-in.</li>
<li><strong>CHECK</strong> reads current IMOD / NIC ITR values.</li>
<li>IMOD: <strong>DELETE</strong> removes the shared IMOD / NIC ITR startup entry and returns the IMOD fields to their default value. Restart the PC to finish resetting.</li>
</ul>
<p><strong>REFRESH</strong> detects devices again. To roll back settings, click <strong>RESTORE</strong> and select the backup you need.</p>
<p>After <strong>APPLY</strong> / <strong>AUTO-OPTIMIZATION</strong>, restart the PC to finish applying the settings.</p>
<p>The program creates a backup before manual apply and reset. Intelligent auto-optimization preserves the original settings.</p>
<p>Detailed settings description: <a href="./src/docs/SETTINGS.en.md">SETTINGS.en.md</a>.</p>

</td></tr>
</table>

<table width="100%">
<tr><th align="left" width="860"><strong>Bugs / testing</strong></th></tr>
<tr><td width="860">

<p>Report bugs or other problems in <a href="https://github.com/arsenzaaa/DEVICE-TWEAKER/issues/new?template=bug_report.yml">issues</a>. Include a log, the program version, Windows build, and PC configuration. Also describe what you did and what happened.</p>
<p>To help test intelligent auto-optimization, send logs in this order:</p>
<ol>
<li>Before intelligent auto-optimization (or after settings reset).</li>
<li>After intelligent auto-optimization.</li>
<li>After restarting the PC and reopening the program.</li>
</ol>
<p>Project updates and discussion: <a href="https://t.me/arsenzaa">Telegram</a>.</p>

</td></tr>
</table>

<table width="100%">
<tr><th align="left" width="860"><strong>Source</strong></th></tr>
<tr><td width="860">

<p><a href="./src/README.en.md">Build and checks</a> · <a href="./CHANGELOG.en.md">Changelog</a> · <a href="./LICENSE">GPL-3.0</a></p>
</td></tr>
</table>
