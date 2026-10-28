<p align="right"><a href="./README.md">Русский</a></p>

<h1>
  <a href="https://github.com/arsenzaaa/DEVICE-TWEAKER#gh-dark-mode-only"><picture><source media="(max-width: 600px)" srcset="./assets/branding/hero-mobile-dark.svg"><img src="./assets/branding/hero-dark.svg" alt="DEVICE TWEAKER / Windows 10/11 x64 / 0.0.4" width="1080"></picture></a>
  <a href="https://github.com/arsenzaaa/DEVICE-TWEAKER#gh-light-mode-only"><picture><source media="(max-width: 600px)" srcset="./assets/branding/hero-mobile-light.svg"><img src="./assets/branding/hero-light.svg" alt="DEVICE TWEAKER / Windows 10/11 x64 / 0.0.4" width="1080"></picture></a>
</h1>

<p align="center"><strong>Device interrupt configuration for Windows 10/11 x64.</strong><br>
Manual settings and intelligent auto-optimization.
</p>

<p align="center">
  <a href="https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.4#gh-dark-mode-only"><img src="./assets/branding/download-en-dark.svg" alt="Download 0.0.4" width="170"></a>
  <a href="https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.4#gh-light-mode-only"><img src="./assets/branding/download-en-light.svg" alt="Download 0.0.4" width="170"></a>
  <a href="./DEVICE%20TWEAKER/docs/SETTINGS.en.md#gh-dark-mode-only"><img src="./assets/branding/settings-en-dark.svg" alt="Settings" width="140"></a>
  <a href="./DEVICE%20TWEAKER/docs/SETTINGS.en.md#gh-light-mode-only"><img src="./assets/branding/settings-en-light.svg" alt="Settings" width="140"></a>
</p>

<p align="center">
  <a href="./SCREENSHOTS.en.md"><img src="./assets/screenshots/showcase_intel_14900k_full_en.png" alt="DEVICE TWEAKER / USB IMOD / Intel Core i9-14900K" width="860"></a>
</p>

<p align="center"><a href="./SCREENSHOTS.en.md">USB IMOD / GPU / NIC ITR: all screenshots</a></p>

## Intelligent auto-optimization

Automatically assigns devices to CPUs based on P-Core / E-Core, SMT / HT, CPPC rankings, AMD CCD / CCX and each device's role. A discrete GPU gets two distinct physical cores; USB, network and audio devices use their own allocation rules.

## Settings

**Interrupts / CPU**

MSI Mode / MSI Limit, IRQ Priority / Policy, Interrupt Affinity, ReservedCpuSets and current IRQ Mode detection.

**USB / network**

Read and configure IMOD on supported xHCI controllers and NIC ITR on supported network adapters. Wired NDIS adapters have base CPU, queue count and RSS/IRQ mode settings.

IMOD / NIC ITR profiles can be saved for the next Windows sign-in.

**Devices / diagnostics**

USB role detection, mouse / keyboard Polling Rate information and Raw Input event frequency measurement. Backups and detailed logs of applied settings.

Mouse Throttle is available on compatible Windows 11 builds, along with USB / wired network power-saving settings.

<details>
<summary><strong>Getting started, backups and checking the result</strong></summary>

1. Run the program as administrator.
2. Click **AUTO-OPTIMIZATION**, or choose settings manually and click **APPLY**. Auto-optimization offers USB IMOD separately. NIC ITR is configured manually.
3. Restart the PC after applying. To restore a selected backup, use **RESTORE**.

A backup is created before manual apply; intelligent auto-optimization preserves the original settings.

Parameter and button descriptions are available in the [settings guide](./DEVICE%20TWEAKER/docs/SETTINGS.en.md).

## Bugs / testing

Report bugs or other problems in [issues](https://github.com/arsenzaaa/DEVICE-TWEAKER/issues/new?template=bug_report.yml). Include a log, program version, Windows build and PC configuration. Describe what you did and what happened.

To test intelligent auto-optimization, send logs from before applying, after applying, and after restarting the PC and reopening the program.

</details>

## Documentation

- [Changelog](./CHANGELOG.en.md)
- [Build and developer tools](./DEVICE%20TWEAKER/README.en.md)

Updates and discussion: [Telegram](https://t.me/arsenzaa).
