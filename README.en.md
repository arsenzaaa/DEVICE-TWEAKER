# DEVICE TWEAKER

**[Download](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest)** · [Русский](./README.md) · [Changelog](./CHANGELOG.en.md) · [Author's Telegram](https://t.me/arsenzaa)

**DEVICE TWEAKER** is a Windows 10/11 x64 utility. It brings MSI Utility V3, Interrupt Affinity Policy Tool, and ReservedCpuSets controls together with xHCI IMOD, NIC ITR, and RSS settings in one interface.

What you can configure:

- MSI Mode, MSI Limit, and IRQ Priority set interrupt parameters for a selected device; Interrupt Affinity selects CPUs for its IRQs.
- xHCI IMOD configures the USB controller or individual interrupters and checks values after writing.
- NIC ITR reads and sets registers on supported network adapters, configures queues, and saves a profile.
- RSS sets the receive queue count and base CPU; saved settings and current adapter state appear separately.
- Auto-optimization places devices on cores according to P/E cores, SMT/HT, CPPC, AMD CCD/CCX, and device role.
- Raw Input estimates mouse and keyboard event rates; supported Windows 11 builds also offer `RawMouseThrottleDuration` controls.
- Backups, restore, and a log showing what was applied, what was skipped, and why.

## Download and run

The current release is [v0.0.4](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.4). See the [changelog](./CHANGELOG.en.md) for release history.

Two EXE builds are available:

- `DEVICE.TWEAKER.exe` is the compact build and requires the [.NET 8 Desktop Runtime x64](https://dotnet.microsoft.com/download/dotnet/8.0).
- `DEVICE.TWEAKER.SelfContained.exe` includes the .NET runtime.

Run the EXE as administrator. The release includes `SHA256SUMS.txt` for both builds. Hardware IMOD/ITR access uses the bundled `DTIMOD.sys` and KDU; if register access is unavailable, the reason is recorded in the session log.

## Interface

USB controller on a simulated Intel Core i9-14900K system. IMOD details are expanded; the values are test data, and no settings were written to the system.

![DEVICE TWEAKER: USB controller on Intel Core i9-14900K with expanded IMOD table](./assets/screenshots/showcase_intel_14900k_full_en.png)

[More screenshots: per-interrupter IMOD values, GPU, and network adapter](./SCREENSHOTS.en.md).

## Interrupts

### MSI / MSI-X

**MSI Mode** writes `MSISupported` for the selected device. [MSI/MSI-X](https://learn.microsoft.com/en-us/windows-hardware/drivers/kernel/enabling-message-signaled-interrupts-in-the-registry) use messages instead of an interrupt line; MSI-X permits multiple vectors with different CPU affinities. MSI-X capability does not mean the driver uses every available vector. **WMI hint** estimates the mode from the device's allocated IRQ numbers: above 999 suggests MSI/MSI-X; other numbers suggest Line. This is a heuristic, not a check of the interrupt resource type assigned to the driver, and cannot distinguish MSI from MSI-X. **HW** shows device capability; `MSISupported` is the registry setting.

**MSI Limit** sets `MessageNumberLimit`. **Unlimited** removes the registry limit; numeric input is checked against MSI/MSI-X capabilities when they can be detected. **[IRQ Priority](https://learn.microsoft.com/en-us/windows-hardware/drivers/ddi/wdm/ne-wdm-_irq_priority)** sets `DevicePriority`, which Windows considers when assigning hardware interrupt priority. It does not change game priority, set DPC priority, or guarantee a fixed latency.

### Interrupt Affinity

**[Interrupt Affinity](https://learn.microsoft.com/en-us/windows-hardware/drivers/kernel/interrupt-affinity-and-priority)** selects the logical processors eligible to service a device's IRQs. The CPU list identifies SMT/HT siblings, Intel P/E cores, AMD CCD/CCX, and available CPPC rankings. The selection can therefore be made by physical core and sibling thread, not just by CPU number.

**SpecCPU** writes `DevicePolicy` and `AssignmentSetOverride`. Clearing the CPU selection for a regular device writes `DevicePolicy=0` (`MachineDefault`) and removes `AssignmentSetOverride`; NDIS adapters follow the selected RSS/IRQ mode. The current mask covers processor group 0, up to 64 logical processors. It controls interrupt routing; a driver can queue [DPCs](https://learn.microsoft.com/en-us/windows-hardware/drivers/kernel/organization-of-dpc-queues) on another CPU or change [MSI-X](https://learn.microsoft.com/en-us/windows-hardware/drivers/network/changing-the-cpu-affinity-of-msi-x-table-entries) vector affinity at runtime. Actual ISR/DPC placement must be verified with a trace, not inferred from the registry mask.

**Policy** also offers `MachineDefault`, `All`, `AllClose`, `Single`, and [`SpreadMessages`](https://learn.microsoft.com/en-us/windows-hardware/drivers/ddi/wdm/ne-wdm-_irq_device_policy). Only `SpecCPU` uses the CPUs selected in the list as a mask. The other policies leave processor selection to Windows; `SpreadMessages` can place different messages on different CPUs when the device and driver use multiple MSI-X vectors. The policy itself does not create additional vectors.

## USB xHCI IMOD

**[IMOD](https://cdrdv2-public.intel.com/625472/625472_xHCI_Rev1_2b.pdf)** sets the xHCI interrupt-moderation interval. The register belongs to a controller *interrupter* with its own Event Ring, not to an individual mouse or keyboard. Each step is 250 ns: `0xC8` is 50 µs, `0xFA0` is 1 ms, and `0x0` disables moderation.

**Devices** builds a profile for detected mice, keyboards, gamepads, and USB audio, then attempts to map them to interrupters through the xHCI topology. If a reliable map is unavailable, it applies one controller-wide value. **XHCI** writes one value to all interrupters; **Interrupters** opens per-index control. The details show the controller's first 64 interrupters.

**SET** applies the configuration and reads the registers back after writing. Custom settings are saved for reapplication when the user next signs in to Windows. **CHECK** rereads the hardware; **REFRESH** only updates the device list. **DELETE** removes the shared startup configuration for IMOD and saved NIC ITR profiles, sets the UI to the program's default IMOD value (`0xC8`), and writes it to the registers if the driver is already loaded. Otherwise, current registers stay unchanged; Windows and the controller driver initialize them on the next boot. DELETE does not create a separate automatic backup.

The IMOD interval is not added to every USB report. An event after the interrupter has been idle and one arriving during an active countdown take different paths. Neither the register value nor mouse Polling Rate measures end-to-end input latency.

## Network adapters

### RSS

**[RSS](https://learn.microsoft.com/en-us/windows-hardware/drivers/network/introduction-to-receive-side-scaling)** distributes incoming packet processing across CPUs using a hash and an indirection table. **RSS Queues** sets the requested receive-queue count; the base CPU is the starting point for processor selection. **NDIS Mode** applies RSS, Interrupt Affinity, or both.

The written `*NumRssQueues` value is not the number of queues used by a particular workload. DEVICE TWEAKER displays saved parameters separately from the active `Get-NetAdapterRss` state and logs mismatches. [`NdisRssProfileBalanced`](https://learn.microsoft.com/en-us/windows-hardware/drivers/network/standardized-inf-keywords-for-rss) does not allow manual base-CPU configuration. The adapter is not restarted automatically after an RSS write. RSS controls receive processing; Interrupt Affinity controls eligible CPUs for IRQs. Neither replaces the other.

### NIC ITR

**NIC ITR** accesses hardware moderation registers on supported network controllers: Intel I210/I211, I219, I225/I226, I350, 82576/82580, Realtek RTL8111/8168, RTL8125/8126, and compatible Killer models. Hardware writes are disabled for unknown PCI IDs.

**SET** writes the register now, **SAVE** stores a profile for application at the next Windows sign-in without changing the current register, and **CHECK** rereads the hardware. Register formats and interval units differ by controller family: an [Intel EITR](https://cdrdv2-public.intel.com/333016/333016%20-%20I210_Datasheet_v_3_7.pdf) value is not a ready-made Realtek profile. A shorter interval can increase both interrupt rate and CPU load.

## Additional settings

**ReservedCpuSets** defines a system-wide mask of CPUs that Windows tries to avoid for ordinary scheduling. It is not hard isolation or a substitute for device Interrupt Affinity.

**Power Saving** controls available settings for the USB controller and its root hubs, including [USB selective suspend](https://learn.microsoft.com/en-us/windows-hardware/drivers/usbcon/usb-selective-suspend). For a wired network adapter, the switch changes whether Windows may turn off the device to save power. Availability depends on the driver.

On supported Windows 11 builds, **Mouse Throttle** changes `RawMouseThrottleDuration`, a throttle interval for background [Raw Input](https://learn.microsoft.com/en-us/windows/win32/inputdev/raw-input) listeners to which Windows applies this mechanism. Some background registrations can bypass it. The setting does not change the mouse's USB polling rate or target the foreground window. The program estimates event rates from Raw Input intervals and shows them with the device information; this does not directly measure USB polling frequency.

The **MOUSE**, **KEYBOARD**, **USB**, **GPU**, **NETWORK**, and **STORAGE** buttons jump to matching devices in the full list. Text search hides nonmatching devices.

## Auto-optimization and restore

**AUTO-OPTIMIZATION** selects physical cores from topology and CPPC rankings. On multi-CCD systems it selects a target CCD separately; two SMT threads of one core are not counted as two independent cores. A GPU is planned on two distinct physical cores, while input devices, wired network adapters, and audio have separate sharing rules. Assignments without a suitable core or a reliably detected role are skipped with a reason in the report. If CPU topology cannot be determined reliably, affinity masks are not written. For NDIS adapters, the RSS/IRQ mode is chosen using the detected runtime state. IMOD is offered as a separate step; NIC ITR remains manual.

A validated backup is created before the main manual apply, IMOD/ITR writes, and the full reset. Auto-optimization requires an original-settings snapshot; an extra backup is optional. No write begins if a required snapshot cannot be created. Saved backups can be selected under **RESTORE**.

## Build from source

Windows x64 and the .NET 8 SDK are required:

```powershell
dotnet build ".\DEVICE TWEAKER\DeviceTweakerCS.csproj" -c Release
dotnet test ".\DEVICE TWEAKER\tests\DeviceTweaker.Tests\DeviceTweaker.Tests.csproj" -c Release
```

See the [build guide](./DEVICE%20TWEAKER/README.en.md) for details. Source code is licensed under [GPL-3.0](./LICENSE).
