# Changelog

[DEVICE TWEAKER](./README.md) / Changelog · [Русский](./CHANGELOG.ru.md)

## [0.0.4](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.4) - 2026-10-31

Source comparison: [v0.0.3 → v0.0.4](https://github.com/arsenzaaa/DEVICE-TWEAKER/compare/b82e8d095ac3303e0bcd8607ac291f0d3dc29bc1...v0.0.4).

### Changed

- Moved IMOD/NIC ITR startup to Task Scheduler, adding resume and USB-device connection events alongside logon.

### Added

- The application and startup script share register-access code. The runtime SHA-256 is checked before compilation; driver and loader files are checked when loading DTIMOD.
- Added IRQ Mode and a command to read IRQ information for all PCI devices. The active interrupt mode and PCI capabilities are displayed separately from the MSI registry setting.
- RSS allocation selects one logical processor per physical core. Available queue counts account for driver and device limits.
- Saving NIC ITR disables supported driver moderation settings; deleting the profile can restore their INF defaults.
- Added diagnostics for HVCI, Vulnerable Driver Blocklist, Microsoft Defender and FACEIT/Vanguard services before driver operations.

### Fixed

- Removed the general physical-memory write command from DTIMOD.sys. Its replacement validates the write range against the selected device resources.
- USB Selective Suspend restore uses Windows Power API for AC and DC values and verifies the result.
- The mouse display shows the measured event rate, including when it is below the USB descriptor maximum.
- Automatic backup cleanup no longer deletes the selected restore file or the rollback copy during restoration. Device power-restore errors are reported separately.
- The application no longer overwrites a startup profile when reading or parsing it fails.
- After a driver-loading attempt with an uncertain outcome, the application and startup script block loading through the fallback mechanism.
- DTIMOD.sys rejects incomplete hardware-resource lists, including malformed data after a valid write range.

### Interface

- Backup selection uses neutral colors for the selected row, file locations and protected original settings.
- IMOD readback includes the last read time in the current value's tooltip and marks the device map as stale after device changes.
- Dual-CCD status is displayed as On/Off.
- Apply, save, restore and reset run long operations in the background. Refresh preserves pending edits and scroll position; a failed device scan keeps the previous cards on screen.
- RSS settings, power settings and driver details are read in the background. Failed or timed-out reads preserve the previous device list; late results are ignored.
- Layout changes preserve card state, focus and the text caret. After a partial apply, unconfirmed settings remain marked as pending.
- Moved English and Russian strings to resource files. Clarified messages for failed settings reads, corrupt profiles and invalid MSI Limit values.

### Build

- Moved application sources to src and QA tools to dev/qa. The test panel is included only in DeveloperQa builds and excluded from the normal EXE.
- Updated the application and build tools to .NET 10. The framework-dependent build requires .NET 10 Desktop Runtime x64; the self-contained build includes the runtime.
- The release workflow builds a selected source commit without publishing by default. Replacing 0.0.4 is a separate action with backup and download verification.

## [0.0.3](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.3) - 2026-10-01

Source comparison: [v0.0.2 → v0.0.3](https://github.com/arsenzaaa/DEVICE-TWEAKER/compare/39f016b6ab5e5edda4fbb83c115aae6cff2ee4c7...b82e8d095ac3303e0bcd8607ac291f0d3dc29bc1).

### Added

- Replaced WinIO with DTIMOD.sys loaded through KDU.
- Added NIC ITR readback, writes and saved profiles for supported Intel and Realtek adapters.
- Expanded RSS controls with queue counts and runtime state from Get-NetAdapterRss.
- Added mouse event-rate measurement through Raw Input and Mouse Throttle for supported Windows 11 builds.
- Added USB and network-adapter power-saving controls.
- Added CPPC and AMD CCD/CCX information to CPU topology and AUTO planning.
- Added RU/EN switching without restart, backup selection for restore and operation-result reports.

### Fixed

- Corrected AllClose and All mappings to Windows policy values; added Unlimited and capability-aware validation for MSI Limit.

## [0.0.2](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.2) - 2026-01-15

Source comparison: [v0.0.1 → v0.0.2](https://github.com/arsenzaaa/DEVICE-TWEAKER/compare/7e003dcf55e843881372c3308b5df7032c049295...39f016b6ab5e5edda4fbb83c115aae6cff2ee4c7).

### Changed

- Added UI scaling based on Windows DPI.
- P/E-core classification for the CPU list and AUTO now combines EfficiencyClass with SMT availability.
- Changed the order in which AUTO allocates CPUs to devices.

### Fixed

- Removed duplicate USB-role log entries for alternative keys of the same controller.

## [0.0.1](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.1) - 2026-01-14

Initial release. Sources: [v0.0.1](https://github.com/arsenzaaa/DEVICE-TWEAKER/tree/7e003dcf55e843881372c3308b5df7032c049295).

- MSI Mode and Limit, IRQ Priority, Interrupt Affinity and ReservedCpuSets controls.
- RSS base CPU selection and xHCI IMOD configuration through WinIO.
