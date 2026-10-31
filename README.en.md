<p align="right"><a href="./README.md">Русский</a></p>

<h1>
<picture>
<source media="(max-width: 600px) and (prefers-color-scheme: dark)" srcset="./assets/branding/hero-mobile-dark.svg">
<source media="(max-width: 600px)" srcset="./assets/branding/hero-mobile-light.svg">
<source media="(prefers-color-scheme: dark)" srcset="./assets/branding/hero-dark.svg">
<img src="./assets/branding/hero-light.svg" alt="DEVICE TWEAKER — Windows 10 / 11 x64: MSI, affinity, RSS, IMOD, ITR" width="1080">
</picture>
</h1>

<p align="center"><strong>DEVICE TWEAKER configures device interrupts on Windows 10/11 x64.</strong></p>

<p align="center">MSI Utility V3 + Interrupt Affinity Policy Tool (IAPT) + ReservedCpuSets + RSS + IMOD / ITR control in one interface, with intelligent auto-optimization for your CPU topology.</p>

<p align="center">
<a href="https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest"><picture><source media="(prefers-color-scheme: dark)" srcset="./assets/branding/button-download-en-dark.svg"><img src="./assets/branding/button-download-en-light.svg" alt="Download 0.0.4" width="179" height="40"></picture></a>
<a href="./docs/SETTINGS.en.md"><picture><source media="(prefers-color-scheme: dark)" srcset="./assets/branding/button-settings-en-dark.svg"><img src="./assets/branding/button-settings-en-light.svg" alt="Settings guide" width="122" height="40"></picture></a>
<a href="./CHANGELOG.en.md"><picture><source media="(prefers-color-scheme: dark)" srcset="./assets/branding/button-changelog-en-dark.svg"><img src="./assets/branding/button-changelog-en-light.svg" alt="Changelog" width="132" height="40"></picture></a>
</p>

## <samp>[ INTERFACE ]</samp>

<p align="center"><a href="./assets/screenshots/showcase_intel_14900k_full_en.png"><img src="./assets/screenshots/showcase_intel_14900k_full_en.png" alt="DEVICE TWEAKER: USB IMOD on Intel Core i9-14900K" width="100%"></a></p>

<table>
<tr><th align="left" colspan="3"><samp>More configurations</samp></th></tr>
<tr>
<td width="33%" valign="top"><a href="./assets/screenshots/showcase_amd_imod_full_en.png"><img src="./assets/screenshots/thumbs/showcase_amd_imod_full_en.png" alt="USB IMOD on AMD Ryzen 9 9950X3D"></a><br><sub>USB IMOD, Ryzen 9 9950X3D</sub></td>
<td width="33%" valign="top"><a href="./assets/screenshots/showcase_gpu_full_en.png"><img src="./assets/screenshots/thumbs/showcase_gpu_full_en.png" alt="GPU interrupts on two physical cores"></a><br><sub>GPU on two physical cores</sub></td>
<td width="33%" valign="top"><a href="./assets/screenshots/showcase_nic_full_en.png"><img src="./assets/screenshots/thumbs/showcase_nic_full_en.png" alt="NIC ITR and RSS on Intel I226-V"></a><br><sub>NIC ITR and RSS, Intel I226-V</sub></td>
</tr>
</table>

<sub>Screenshots show test devices and values.</sub>

## <samp>[ FEATURES ]</samp>

<table>
<tr><th align="left">Intelligent auto-optimization</th></tr>
<tr><td>
<p>Accounts for P-Core / E-Core, SMT / HT, CPPC rankings, CCD / CCX, device count and device roles. A discrete GPU gets two different physical cores: two SMT siblings of one core do not count as two.</p>
<p align="center"><a href="./docs/SETTINGS.en.md#intelligent-auto-optimization-and-restore"><picture><source media="(max-width: 600px) and (prefers-color-scheme: dark)" srcset="./assets/branding/cpu-allocation-mobile-en-dark.svg"><source media="(max-width: 600px)" srcset="./assets/branding/cpu-allocation-mobile-en-light.svg"><source media="(prefers-color-scheme: dark)" srcset="./assets/branding/cpu-allocation-en-dark.svg"><img src="./assets/branding/cpu-allocation-en-light.svg" alt="GPU: CPU 4 on core 2 and CPU 8 on core 4, SMT siblings CPU 5 and CPU 9 stay free" width="560"></picture></a></p>
<p>The controller with the mouse and keyboard, the network and audio have their own placement rules; on CPUs with several CCDs a target CCD is chosen. Storage gets MSI and high priority while affinity stays with Windows. Integrated GPU affinity goes back to Windows, Wi-Fi is left untouched.</p>
<p>If the topology cannot be detected reliably, nothing is written. The original state is saved before the first auto-optimization, and <strong>RESTORE → ORIGINAL</strong> brings it back.</p>
</td></tr>
</table>

<table>
<tr><th align="left">Interrupt settings</th></tr>
<tr><td>
<p>MSI Mode / MSI Limit / IRQ Priority / IRQ Policy / Interrupt Affinity for each device.</p>
<p>The MSI Mode registry setting and the current interrupt mode in IRQ Mode are shown separately; HW lists the modes the device supports. On opening, the program takes the mode from Windows resources or from an already loaded DTIMOD.sys and does not load the driver. If Windows resources cannot tell MSI from MSI-X, it shows MSI/MSI-X.</p>
<p><strong>READ IRQ MODE</strong> at the top reads MSI / MSI-X Enable of every listed PCI device through DTIMOD.sys and loads the driver if needed.</p>
</td></tr>
</table>

<table>
<tr><th align="left">USB IMOD</th></tr>
<tr><td>
<p>Reads and changes IMOD, the interrupt moderation interval of each xHCI interrupter, in 250 ns steps: <code>0xC8</code> = 50 µs, <code>0x0</code> turns moderation off.</p>
<p><strong>Devices</strong> sets values by USB role, <strong>XHCI</strong> uses one value for the controller, <strong>Interrupters</strong> sets them by index. The device → interrupter map is read from the xHCI contexts in memory. Values are read back after every write.</p>
</td></tr>
</table>

<table>
<tr><th align="left">Network settings</th></tr>
<tr><td>
<p>NIC ITR on Intel I210 / I211, I219, I225 / I226, I350, 82576 / 82580 and Realtek RTL8111 / 8168, RTL8125 / 8126, including compatible Killer adapters. Queue and RX / TX details depend on the controller family. <strong>SAVE</strong> turns off the driver's own moderation in the registry (<code>*InterruptModeration</code>), which would otherwise overwrite the register; <strong>DELETE</strong> restores the driver value.</p>
<p>RSS on wired NDIS adapters: base CPU, queue count and RSS / IRQ mode. Processors are taken one per physical core, SMT siblings are skipped. The queue count is limited to the values offered by the driver and the number of MSI-X vectors.</p>
</td></tr>
</table>

<table>
<tr><th align="left">IMOD / NIC ITR startup</th></tr>
<tr><td>
<p>Register values last until a restart or sleep. A saved profile is applied by the Task Scheduler task <code>\DEVICE TWEAKER\IMOD</code>, running as SYSTEM: at sign-in, after resume from sleep and when a USB device is connected.</p>
<p>The task runs the readable script <code>ApplyIMOD.ps1</code> from <code>%ProgramData%\DEVICE TWEAKER\Startup</code> and works without the EXE. SHA-256 of the runtime, driver and loader is checked before each run. Only administrators and SYSTEM can modify the files; logs are kept for 14 days.</p>
</td></tr>
</table>

<table>
<tr><th align="left">Device detection</th></tr>
<tr><td>
<p>Finds GPUs / USB controllers / network adapters / audio controllers / storage controllers and lists them by category.</p>
<p>Identifies which USB controller hosts the mouse / keyboard / gamepad / microphone / audio / camera. If several devices share one controller, all detected roles are shown.</p>
</td></tr>
</table>

<table>
<tr><th align="left">CPU topology / ReservedCpuSets</th></tr>
<tr><td>
<p>Detects P-Core / E-Core, SMT / HT, AMD CCD / CCX and CPPC rankings when available. Shows which physical core each logical thread belongs to. Selected CPUs can be reserved through ReservedCpuSets.</p>
</td></tr>
</table>

<table>
<tr><th align="left">Polling Rate / Mouse Throttle</th></tr>
<tr><td>
<p>Shows mouse and keyboard Polling Rate and measures the Raw Input event rate; after a measurement the measured rate is displayed. Keyboard event rate depends on key activity. RawMouseThrottleDuration can be set on compatible Windows 11 builds, with the current value shown in the interface.</p>
</td></tr>
</table>

<table>
<tr><th align="left">Power saving</th></tr>
<tr><td>
<p>Power saving of USB controllers / root USB hubs / wired network adapters and USB selective suspend in the power plan. Values are read back through WMI after writing: an unsupported feature, an error and a value mismatch are reported separately.</p>
</td></tr>
</table>

<table>
<tr><th align="left">Backups / logs</th></tr>
<tr><td>
<p>A verified backup is created before applying, including the startup profile. Backups can be restored or deleted, and settings can be reset to Windows defaults with <strong>RESET WINDOWS DEFAULT</strong>. Backups are stored in the <code>Backups</code> folder next to the EXE.</p>
<p>Detailed logs of device detection and applied settings — what was applied, what was skipped and why — are written to the <code>logs</code> folder next to the EXE.</p>
</td></tr>
</table>

<table>
<tr><th align="left">Interface / navigation</th></tr>
<tr><td>
<p>Device search, category filters, keyboard navigation and RU / EN switching without a restart. Settings that were changed but not yet applied are marked by an indicator.</p>
</td></tr>
</table>

## <samp>[ WHAT THE PROGRAM CHANGES ]</samp>

| Setting | Where it is written |
| :--- | :--- |
| MSI Mode, MSI Limit | `MSISupported`, `MessageNumberLimit` in `Interrupt Management\MessageSignaledInterruptProperties` |
| IRQ Priority | `DevicePriority` in `Interrupt Management\Affinity Policy` |
| IRQ Policy, Interrupt Affinity | `DevicePolicy`, `AssignmentSetOverride` in the same key |
| RSS | `*RssBaseProcNumber`, `*NumRssQueues`, `*MaxRssProcessors`, `*RssMaxProcNumber` in the adapter class key |
| ReservedCpuSets | `ReservedCpuSets` in `Session Manager\kernel` |
| Mouse Throttle | `RawMouseThrottleDuration` in `HKCU\Control Panel\Mouse` |
| Power saving | `PnPCapabilities`, `SelectiveSuspendEnabled`, `EnhancedPowerManagementEnabled`, USB selective suspend in the power plan |
| USB IMOD | the IMOD register of each xHCI interrupter |
| NIC ITR | EITR / ITR on Intel, IntrMit on Realtek |

Registry settings take effect after a restart. IMOD and NIC ITR registers change immediately and are read back after writing.

## <samp>[ INSTALLATION ]</samp>

<table>
<tr><th align="left">Download</th></tr>
<tr><td>
<p>The <a href="https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest">release</a> has two builds:</p>
<ul>
<li><a href="https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest/download/DEVICE.TWEAKER.SelfContained.exe"><code>DEVICE.TWEAKER.SelfContained.exe</code></a> — self-contained, nothing else to install;</li>
<li><a href="https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest/download/DEVICE.TWEAKER.exe"><code>DEVICE.TWEAKER.exe</code></a> — standard build, requires the <a href="https://dotnet.microsoft.com/download/dotnet/8.0">.NET 8 Desktop Runtime x64</a>.</li>
</ul>
<p>Both builds include the DTIMOD.sys driver and the KDU loader. SHA-256 hashes are listed in <code>SHA256SUMS.txt</code> for the files and in <code>SOURCE-SHA256SUMS.txt</code> for the sources.</p>
</td></tr>
</table>

<table>
<tr><th align="left">Usage</th></tr>
<tr><td>
<p>Run the program as administrator.</p>
<p><strong>AUTO-OPTIMIZATION</strong> builds a plan for your CPU topology and applies it. If an eligible xHCI controller is found, the program offers to enable USB IMOD. The <strong>SYSTEM CHECK</strong> button in that dialog shows the state of HVCI, the Blocklist, Microsoft Defender and the FACEIT / Vanguard services. NIC ITR is configured manually.</p>
<p>For manual configuration, choose the settings and click <strong>APPLY</strong>. It applies interrupt, ReservedCpuSets, RSS, Mouse Throttle, power-saving and USB IMOD settings.</p>
<p>IMOD / NIC ITR also have separate buttons:</p>
<ul>
<li>IMOD: <strong>SET</strong> applies the values and saves those that differ from the defaults for startup.</li>
<li>NIC ITR: <strong>SET</strong> changes the current values, <strong>SAVE</strong> stores a profile for startup.</li>
<li><strong>CHECK</strong> reads the actual register values.</li>
<li>IMOD: <strong>DELETE</strong> removes the shared IMOD / NIC ITR startup task and returns the IMOD fields to their default value. Restart the PC to finish the reset.</li>
</ul>
<p>Restart the PC after <strong>APPLY</strong> / <strong>AUTO-OPTIMIZATION</strong>. <strong>REFRESH</strong> detects devices again.</p>
<p>To roll back, use <strong>RESTORE</strong>: <strong>ORIGINAL</strong> returns the registry to its state before the first auto-optimization and removes the startup task. The same menu has every backup and a reset to Windows defaults.</p>
</td></tr>
</table>

<table>
<tr><th align="left">Driver / anti-cheats</th></tr>
<tr><td>
<p>IMOD and NIC ITR registers are not reachable from user mode. They are written by the custom DTIMOD.sys driver, and only to memory that Windows assigned to the selected device. The driver has no public signature, so it is loaded by <a href="./src/IMOD/KDU/README.md">KDU</a> through a signed vulnerable driver, and antivirus software flags such loaders as HackTool. The DTIMOD sources and the pinned KDU version are in the repository; binaries are checked by SHA-256.</p>
<p>Memory integrity (HVCI) and the Microsoft Vulnerable Driver Blocklist block loading through KDU; the FACEIT and Vanguard anti-cheats may conflict with it. The driver is only needed for IMOD, NIC ITR and <strong>READ IRQ MODE</strong>; registry settings do not depend on it.</p>
</td></tr>
</table>

## <samp>[ FEEDBACK ]</samp>

<table>
<tr><th align="left">Bugs / testing</th></tr>
<tr><td>
<p>Report bugs and other problems in <a href="https://github.com/arsenzaaa/DEVICE-TWEAKER/issues/new?template=bug_report.yml">issues</a>. Attach the log from the <code>logs</code> folder next to the EXE, the program version, Windows build and PC configuration. Also describe what you did and what happened.</p>
<p>To help test intelligent auto-optimization, send a <a href="https://github.com/arsenzaaa/DEVICE-TWEAKER/issues/new?template=hardware_report.yml">hardware report</a> with logs in this order:</p>
<ol>
<li>Before auto-optimization or after a settings reset.</li>
<li>Right after applying it.</li>
<li>After restarting the PC and opening the program again.</li>
</ol>
<p>Project updates and discussion: <a href="https://t.me/arsenzaa">Telegram</a>.</p>
</td></tr>
</table>

<p align="center"><samp><a href="./docs/SETTINGS.en.md">Settings guide</a> · <a href="./docs/DEVELOPMENT.en.md">Development</a> · <a href="./CHANGELOG.en.md">Changelog</a> · <a href="./docs/THIRD_PARTY_NOTICES.md">Third-party components</a> · <a href="./.github/SECURITY.md">Security</a> · <a href="./LICENSE">GPL-3.0</a></samp></p>
