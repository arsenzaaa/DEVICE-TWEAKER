# How DEVICE TWEAKER settings work

[Project page](../../README.en.md) · [Русский](./SETTINGS.md)

This page explains what the program writes, what the interface shows, and when it rereads values.

## Interrupts

### MSI / MSI-X

**MSI Mode** writes `MSISupported` for the selected device. [MSI/MSI-X](https://learn.microsoft.com/en-us/windows-hardware/drivers/kernel/enabling-message-signaled-interrupts-in-the-registry) use messages instead of an interrupt line; MSI-X permits multiple vectors with different CPU affinities. MSI-X capability does not mean the driver uses every available vector.

The **READ IRQ MODE** button at the top reads all PCI devices in the list through DTIMOD.sys, loading the driver if needed. Cards hidden by filters are included. The query reads Windows allocated interrupt resources and the `MSI Enable` / `MSI-X Enable` bits in PCI configuration space. Reading these bits successfully distinguishes MSI from MSI-X even with one IRQ. The query applies no settings.

**HW** shows device capability; `MSISupported` is the registry setting.

Once loaded, DTIMOD stays in memory until Windows restarts. Closing the application closes its connection. IMOD / ITR values remain in the registers until the device driver changes them, the controller resets, or Windows restarts. Background IRQ Mode reads use an already loaded DTIMOD or Windows resources without loading the driver.

- **MSI Mode**. Registry: `MSISupported`.

  Saved setting; a change selected in the dropdown is written when settings are applied.

- **IRQ Mode**. Windows resources; also PCI configuration through DTIMOD.

  Current Line / MSI / MSI-X mode; MSI/MSI-X when the data cannot distinguish the two.

- **HW**. Device hardware capabilities.

  Supported modes.

A failed IRQ Mode read is not replaced with the registry value. MSI Mode and IRQ Mode may differ before a restart or because of how the device driver uses the setting.

The registry is sufficient to configure MSI. IRQ Mode is a separate check of the current state. [MSI-X table entries and message routing can change at runtime](https://learn.microsoft.com/en-us/windows-hardware/drivers/kernel/dynamically-configuring-msi-x), so the result describes the state when it was read.

`HW: Line/MSI/MSI-X` lists the device's supported modes. Parent-controller capabilities are not substituted. IRQ count comes from WMI and does not identify the mode.

Before reading, IRQ Mode shows "not read". On opening or refreshing the list, the application queries an already loaded DTIMOD. If it is absent, it reads Windows resources without loading the driver. Windows resources distinguish MSI from MSI-X when the message allocation and hardware capabilities allow it. Otherwise, the result is `MSI/MSI-X`, with the reason in the tooltip. The same label remains if DTIMOD confirms message interrupts but cannot read PCI configuration. "No IRQ" means no allocated interrupt resources. Read errors and conflicting evidence are shown separately and logged. Refresh discards the previous result and reads again.

**MSI Limit** sets `MessageNumberLimit`. **Unlimited** removes the registry limit; numeric input is checked against MSI/MSI-X capabilities when they can be detected. **[IRQ Priority](https://learn.microsoft.com/en-us/windows-hardware/drivers/ddi/wdm/ne-wdm-_irq_priority)** sets `DevicePriority`, which Windows considers when assigning hardware interrupt priority. It does not change game priority, set DPC priority, or guarantee a fixed latency.

### Interrupt Affinity

**[Interrupt Affinity](https://learn.microsoft.com/en-us/windows-hardware/drivers/kernel/interrupt-affinity-and-priority)** selects the logical processors eligible to service a device's IRQs. The CPU list identifies SMT/HT siblings, Intel P/E cores, AMD CCD/CCX, and available CPPC rankings. The selection can therefore be made by physical core and sibling thread, not just by CPU number.

**SpecCPU** writes `DevicePolicy` and `AssignmentSetOverride`. Clearing the CPU selection for a regular device writes `DevicePolicy=0` (`MachineDefault`) and removes `AssignmentSetOverride`; NDIS adapters follow the selected RSS/IRQ mode.

The current mask covers processor group 0, up to 64 logical processors. It controls interrupt routing; a driver can queue [DPCs](https://learn.microsoft.com/en-us/windows-hardware/drivers/kernel/organization-of-dpc-queues) on another CPU or change [MSI-X](https://learn.microsoft.com/en-us/windows-hardware/drivers/network/changing-the-cpu-affinity-of-msi-x-table-entries) vector affinity at runtime. Actual ISR/DPC placement must be verified with a trace, not inferred from the registry mask.

**Policy** also offers `MachineDefault`, `All`, `AllClose`, `Single`, and [`SpreadMessages`](https://learn.microsoft.com/en-us/windows-hardware/drivers/ddi/wdm/ne-wdm-_irq_device_policy). Only `SpecCPU` uses the CPUs selected in the list as a mask. The other policies leave processor selection to Windows; `SpreadMessages` can place different messages on different CPUs when the device and driver use multiple MSI-X vectors. The policy itself does not create additional vectors.

## USB xHCI IMOD

Before including USB IMOD, the program reads Windows settings and the status of Defender, FACEIT and Vanguard services. A detected anti-cheat or a possible restriction on loading DTIMOD.sys appears beside the apply question. **SYSTEM CHECK** opens two sections: Windows settings and services. "No registry setting" means the registry value is absent, rather than that the feature is off. This check does not load the driver.

The AUTO profile sets mouse IMOD to 0 regardless of polling rate, disabling interrupt moderation. This is a profile rule, not a latency measurement or a change to USB polling frequency. Manual IMOD is configured separately.

**[IMOD](https://cdrdv2-public.intel.com/625472/625472_xHCI_Rev1_2b.pdf)** sets the xHCI interrupt-moderation interval. The register belongs to a controller *interrupter* with its own Event Ring, not to an individual mouse or keyboard. Each step is 250 ns: `0xC8` is 50 µs, `0xFA0` is 1 ms, and `0x0` disables moderation.

**Devices** builds a profile for detected mice, keyboards, gamepads, and USB audio, then attempts to map them to interrupters through the xHCI topology. If a reliable map is unavailable, it applies one controller-wide value. **XHCI** writes one value to all interrupters; **Interrupters** opens per-index control. The details show the controller's first 64 interrupters.

The device map is read from the xHCI contexts in memory: for each attached device, the interrupter named in the TRBs of its transfer ring is used, or the Slot Context target if the ring cannot be read. The read time is shown next to the current value. After a device is attached or removed, the map is marked "devices changed"; **CHECK** reads it again.

The buttons have different jobs:

- **SET** writes the selected configuration and reads the registers back. The program saves custom values to the startup task (see below).
- **CHECK** rereads the registers. **REFRESH** only updates the device list.
- **DELETE** removes the shared startup configuration for IMOD and saved NIC ITR profiles and restores the network driver's moderation that **SAVE** turned off. The UI switches to the default IMOD value (0xC8). With the driver loaded, the program can also reset IMOD registers. USB IMOD writes are skipped when only saved NIC ITR profiles remain in the configuration. Restart Windows to finish resetting.

**DELETE** does not create a separate automatic backup.

The main **APPLY** button also applies USB IMOD along with the other device settings.

The IMOD interval is not added to every USB report. An event after the interrupter has been idle and one arriving during an active countdown take different paths. Neither the register value nor mouse Polling Rate measures end-to-end input latency.

### IMOD / NIC ITR at startup

Register values last until a restart, sleep, or controller reset, so a separate open script applies the saved settings. It works without `DEVICE TWEAKER.exe`: the program can be removed and the startup task keeps working.

- Folder `%ProgramData%\DEVICE TWEAKER\Startup`: `ApplyIMOD.ps1` (a readable PowerShell script), `DeviceTweakerImodRuntime.cs` (register code shared with the program, compiled at run time), `profile.json` (saved values and SHA-256 of every file), `Loader\` (KDU for loading DTIMOD.sys), and `logs\` (logs for 14 days). Only administrators and SYSTEM can modify the folder.
- The Task Scheduler task `\DEVICE TWEAKER\IMOD` runs the script as SYSTEM at sign-in, after resume from sleep, and when a USB/HID device is attached. Before running, the script checks the runtime, driver, and loader hashes against `profile.json`; a mismatch stops it.
- The script and the program share one DTIMOD load mutex, so they never load the driver at the same time.
- The previous startup entry in `%AppData%` and the `Run` key is migrated automatically the first time the program starts, keeping the saved values.

## Network adapters

### RSS

**[RSS](https://learn.microsoft.com/en-us/windows-hardware/drivers/network/introduction-to-receive-side-scaling)** distributes incoming packet processing across CPUs using a hash and an indirection table. **RSS Queues** sets the requested receive-queue count; the base CPU is the starting point for processor selection. **NDIS Mode** applies RSS, Interrupt Affinity, or both.

Queue choices are limited to the values the driver offers and to the number of MSI-X vectors: without MSI-X, all queues arrive on one interrupt. RSS processors are taken one per physical core, skipping SMT siblings; `*MaxRssProcessors` and `*RssMaxProcNumber` are written together with the base CPU.

The written `*NumRssQueues` value is not the number of queues used by a particular workload. DEVICE TWEAKER displays saved parameters separately from the active `Get-NetAdapterRss` state and logs mismatches. [`NdisRssProfileBalanced`](https://learn.microsoft.com/en-us/windows-hardware/drivers/network/standardized-inf-keywords-for-rss) does not allow manual base-CPU configuration. The adapter is not restarted automatically after an RSS write. RSS controls receive processing; Interrupt Affinity controls eligible CPUs for IRQs. Neither replaces the other.

### NIC ITR

**NIC ITR** accesses hardware moderation registers on supported network controllers: Intel I210/I211, I219, I225/I226, I350, 82576/82580, Realtek RTL8111/8168, RTL8125/8126, and compatible Killer models. Hardware writes are disabled for unknown PCI IDs.

**SET** writes the register now, **SAVE** stores a profile in the startup task without changing the current register, and **CHECK** rereads the hardware.

The driver's own adaptive moderation would overwrite the register, so **SAVE** turns it off in the registry (`*InterruptModeration`, and `ITR` on Intel). The change takes effect after a restart. The card's **Interrupt Moderation** line shows the current state; **DELETE** in the IMOD block restores the driver default. Register formats and interval units differ by controller family: an [Intel EITR](https://cdrdv2-public.intel.com/333016/333016%20-%20I210_Datasheet_v_3_7.pdf) value is not a ready-made Realtek profile. A shorter interval can increase both interrupt rate and CPU load.

## Additional settings

**ReservedCpuSets** defines a system-wide mask of CPUs that Windows tries to avoid for ordinary scheduling. It is not hard isolation or a substitute for device Interrupt Affinity.

**Power Saving** controls available settings for the USB controller and its root hubs, including [USB selective suspend](https://learn.microsoft.com/en-us/windows-hardware/drivers/usbcon/usb-selective-suspend). For a wired network adapter, the switch changes whether Windows may turn off the device to save power. Availability depends on the driver.

On supported Windows 11 builds, **Mouse Throttle** changes `RawMouseThrottleDuration`, a throttle interval for background [Raw Input](https://learn.microsoft.com/en-us/windows/win32/inputdev/raw-input) listeners to which Windows applies this mechanism. Some background registrations can bypass it. The setting does not change the mouse's USB polling rate or target the foreground window.

**Raw Input measurements** estimate event rates from the intervals between processed events. Mouse measurements need a stream of movement events; message delivery and processing by the program also affect the result. Keyboard events represent key presses and releases, so pressing several keys or holding a key does not reliably reveal the keyboard's USB polling rate. Raw Input event rate, USB polling rate, and the keyboard's internal scan rate are different quantities.

The **MOUSE**, **KEYBOARD**, **USB**, **GPU**, **NETWORK**, and **STORAGE** buttons jump to matching devices in the full list. Text search hides nonmatching devices.

## Intelligent auto-optimization and restore

**AUTO-OPTIMIZATION** selects physical cores from topology and CPPC rankings. On multi-CCD systems it selects a target CCD separately; two SMT threads of one core are not counted as two independent cores. A discrete GPU gets two distinct physical cores; an integrated GPU's affinity is returned to Windows defaults. Input devices, wired network adapters, and audio have separate sharing rules.

Assignments without a suitable core or a reliably detected role are skipped with a reason in the report. If CPU topology cannot be determined reliably, affinity masks are not written. For NDIS adapters, the RSS/IRQ mode is chosen using the detected runtime state. USB IMOD tuning can be included when starting intelligent auto-optimization; NIC ITR remains manual.

A validated backup is created before the main manual apply, IMOD / NIC ITR **SET**, NIC ITR **SAVE**, backup restoration, and the main reset. Intelligent auto-optimization requires an original-settings snapshot; an extra backup is optional. No write begins if a required snapshot cannot be created. Saved backups can be selected under **RESTORE**.

The main reset preserves MSI Mode / MSI Limit and Wi-Fi settings.
