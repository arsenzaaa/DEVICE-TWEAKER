# Settings reference

[Documentation](./README.md) / Settings · [Русский](./SETTINGS.ru.md)

Settings, stored values and readback for each control.

[Interface](#interface) · [Interrupts](#interrupts) · [USB xHCI IMOD](#usb-xhci-imod) · [Network adapters](#network-adapters) · [Additional settings](#additional-settings) · [Intelligent auto-optimization and restore](#intelligent-auto-optimization-and-restore)

## Interface

Toggles, scrolling and button hover use short transitions. Animations work independently of Windows visual effects. High Contrast uses immediate transitions.

The interface language is saved in `%LOCALAPPDATA%\DEVICE TWEAKER\settings.json`.

## Interrupts

### MSI / MSI-X

**MSI Mode** selects the saved interrupt setting. Changing the dropdown creates a pending edit; **APPLY** writes `MSISupported` to the device's registry key. Restart Windows before checking whether its driver uses the new setting. Disabling MSI is refused when PCI capabilities explicitly report message interrupts without line-based INTx support.

**IRQ Mode** reports the current mode: Line, MSI, MSI-X, or MSI/MSI-X when the available data cannot distinguish the two message modes. A saved **Enabled** setting alongside **Line** can mean a reboot is still needed or the device driver did not use the setting. [Chapter 2](./guide/02-msi.md) explains the difference.

**READ IRQ MODE** queries every PCI device in the list, including cards hidden by filters. It reads allocated Windows interrupt resources and, through DTIMOD, the PCI `MSI Enable` / `MSI-X Enable` bits. Reading these bits distinguishes MSI from MSI-X even with one IRQ. The command can load DTIMOD; it does not apply settings. Progress and the final readable/total count appear on the button, with a summary in its tooltip.

Opening the program or pressing **REFRESH** also starts background reads, using Windows resources or an already loaded DTIMOD. These reads do not load the driver. The old result is cleared before rereading; unavailable data, read errors and conflicting evidence are shown separately and logged. **No IRQ** means no allocated interrupt resources. A successful read describes the state at that moment: [MSI-X routing can change at runtime](https://learn.microsoft.com/en-us/windows-hardware/drivers/kernel/dynamically-configuring-msi-x).

**HW** lists the device's supported interrupt modes and message limit. **IRQ** shows its allocated Windows interrupt-resource count.

**MSI Limit** sets `MessageNumberLimit`. **Unlimited** removes the custom registry limit so Windows and the driver can select the count. Numeric input is checked against detected capabilities. See [enabling MSI](https://learn.microsoft.com/en-us/windows-hardware/drivers/kernel/enabling-message-signaled-interrupts-in-the-registry).

**IRQ Priority** writes `DevicePriority`, the hardware interrupt-priority preference used by Windows. Values and policy behavior are covered in the [affinity chapter](./guide/03-affinity.md#irq-priority).

### Interrupt Affinity

**[Interrupt Affinity](https://learn.microsoft.com/en-us/windows-hardware/drivers/kernel/interrupt-affinity-and-priority)** selects the logical processors eligible to service a device's IRQs. The CPU list identifies SMT/HT siblings, Intel P/E cores, AMD CCD/CCX, and available CPPC rankings. The selection can therefore be made by physical core and sibling thread, not just by CPU number.

**SpecCPU** writes `DevicePolicy` and `AssignmentSetOverride`. Clearing the CPU selection for a regular device writes `DevicePolicy=0` (`MachineDefault`) and removes `AssignmentSetOverride`; NDIS adapters follow the selected RSS/IRQ mode.

The mask covers processor group 0, up to 64 logical processors. [Chapter 3](./guide/03-affinity.md) describes processor selection; [chapter 10](./guide/10-verification.md) covers ISR/DPC traces.

**Policy** also offers `MachineDefault`, `All`, `AllClose`, `Single`, and [`SpreadMessages`](https://learn.microsoft.com/en-us/windows-hardware/drivers/ddi/wdm/ne-wdm-_irq_device_policy). Only `SpecCPU` uses the CPUs selected in the list as a mask. The other policies leave processor selection to Windows; `SpreadMessages` can place different messages on different CPUs when the device and driver use multiple MSI-X vectors. The policy itself does not create additional vectors.

## USB xHCI IMOD

USB IMOD uses DTIMOD. The application requests confirmation once per session; AUTO includes it in the USB IMOD prompt. **DRIVER INFO** displays HVCI, the driver-blocklist registry setting and FACEIT/Vanguard service states. Loading and compatibility are described in the [driver chapter](./guide/11-driver.md).

AUTO sets mouse IMOD to `0x0`, disabling moderation regardless of polling rate. Manual values are configured separately.

**[IMOD](https://cdrdv2-public.intel.com/625472/625472_xHCI_Rev1_2b.pdf)** sets the xHCI interrupt-moderation interval. The register belongs to a controller *interrupter* with its own Event Ring, not to an individual mouse or keyboard. Each step is 250 ns: `0xC8` is 50 µs, `0xFA0` is 1 ms, and `0x0` disables moderation.

**Devices** builds a profile for detected mice, keyboards, gamepads, and USB audio, then attempts to map them to interrupters through the xHCI topology. If a reliable map is unavailable, it applies one controller-wide value. **XHCI** writes one value to all interrupters; **Interrupters** opens per-index control. The details show the controller's first 64 interrupters.

The current-value tooltip shows the last read time. After a device is attached or removed, the map is marked "devices changed"; **CHECK** rereads it.

### Commands

- **SET** writes the selected configuration and reads the registers back. The program saves custom values to the startup task (see below).
- **CHECK** rereads the registers and loads DTIMOD if needed. **REFRESH** updates the device list and available current state; hardware reads use an already loaded DTIMOD.
- **DELETE** removes the shared startup configuration for IMOD and saved NIC ITR profiles and restores the network driver's moderation that **SAVE** turned off. The UI switches to the default IMOD value (0xC8). With the driver loaded, the program can also reset IMOD registers. USB IMOD writes are skipped when only saved NIC ITR profiles remain in the configuration. Restart Windows to finish resetting.

**DELETE** does not create a separate automatic backup.

The main **APPLY** button also applies USB IMOD along with the other device settings.

The [USB chapter](./guide/05-usb-imod.md) explains interrupter mapping, moderation intervals and device-role profiles.

### IMOD / NIC ITR at startup

The saved profile is stored in `%ProgramData%\DEVICE TWEAKER\Startup`. The `\DEVICE TWEAKER\IMOD` task runs as SYSTEM at sign-in, after resume and on relevant USB/HID events. It applies settings independently of the EXE. The previous AppData/Run startup layout is migrated with its saved values.

Files, hash checks, permissions and logs are described in the [driver chapter](./guide/11-driver.md#saved-profile).

## Network adapters

### RSS

**[RSS](https://learn.microsoft.com/en-us/windows-hardware/drivers/network/introduction-to-receive-side-scaling)** distributes incoming packet processing across CPUs using a hash and an indirection table. **RSS Queues** sets the requested receive-queue count; the base CPU is the starting point for processor selection. **NDIS Mode** applies RSS, Interrupt Affinity, or both.

Queue choices follow the adapter's advertised driver options. RSS processor selection uses one logical processor per physical core, with `*MaxRssProcessors` and `*RssMaxProcNumber` defining the requested range. For adapters without MSI-X, multiple advertised queues can share one processor. AUTO requests one queue on one physical core; SMT does not double this count. The [network chapter](./guide/06-network.md) explains these settings.

The card displays saved parameters and active `Get-NetAdapterRss` state, with differences recorded in the log. The [Balanced profile](https://learn.microsoft.com/en-us/windows-hardware/drivers/network/standardized-inf-keywords-for-rss) manages processor placement automatically. RSS writes take effect through the adapter driver; the application leaves adapter restart to the user.

### NIC ITR

**NIC ITR** accesses hardware moderation registers on supported network controllers: Intel I210/I211, I219, I225/I226, I350, 82576/82580, Realtek RTL8111/8168, RTL8125/8126, and compatible Killer models. Hardware writes are disabled for unknown PCI IDs.

**SET** writes the register now, **SAVE** stores a profile in the startup task without changing the current register, and **CHECK** rereads the hardware.

The driver's own adaptive moderation would overwrite the register, so **SAVE** turns it off in the registry (`*InterruptModeration`, and `ITR` on Intel). The change takes effect after a restart. The card's **Interrupt Moderation** line shows the current state; **DELETE** in the IMOD block restores the driver default. Register formats and interval units differ by controller family: an [Intel EITR](https://cdrdv2-public.intel.com/333016/333016%20-%20I210_Datasheet_v_3_7.pdf) value is not a ready-made Realtek profile. A shorter interval can increase both interrupt rate and CPU load.

## Additional settings

**ReservedCpuSets** defines the system-wide CPU reservation mask. Selection and storage format are described in the [CPU chapter](./guide/07-cpu-topology.md#reservedcpusets).

**Power Saving** controls available settings for the USB controller and its root hubs, including [USB selective suspend](https://learn.microsoft.com/en-us/windows-hardware/drivers/usbcon/usb-selective-suspend). For a wired network adapter, the switch changes whether Windows may turn off the device to save power. Availability depends on the driver.

On supported Windows 11 builds, **Mouse Throttle** changes `RawMouseThrottleDuration`, a throttle interval for background [Raw Input](https://learn.microsoft.com/en-us/windows/win32/inputdev/raw-input) listeners to which Windows applies this mechanism. Some background registrations can bypass it. The setting does not change the mouse's USB polling rate or target the foreground window.

**Raw Input measurements** estimate the mouse event rate while the mouse is moving. The input chapter explains [event delivery and polling rate](./guide/08-input.md).

The **MOUSE**, **KEYBOARD**, **GAMEPAD**, **USB**, **GPU**, **NETWORK**, **AUDIO**, and **STORAGE** category buttons jump to matching devices in the full list. Text search hides nonmatching devices.

## Intelligent auto-optimization and restore

**AUTO-OPTIMIZATION** selects physical cores from topology and CPPC rankings. On multi-CCD systems it selects a target CCD separately; two SMT threads of one core are not counted as two independent cores. A discrete GPU gets two distinct physical cores; an integrated GPU's affinity is returned to Windows defaults. Input devices, wired network adapters, and audio have separate sharing rules.

Assignments without a suitable core or a reliably detected role are skipped with a reason in the report. If CPU topology cannot be determined reliably, affinity masks are not written. For NDIS adapters, the RSS/IRQ mode is chosen using the detected runtime state. USB IMOD tuning can be included when starting intelligent auto-optimization; NIC ITR remains manual.

A validated backup is created before the main manual apply, IMOD / NIC ITR **SET**, NIC ITR **SAVE**, backup restoration, and the main reset. Intelligent auto-optimization requires an original-settings snapshot; an extra backup is optional. No write begins if a required snapshot cannot be created. Saved backups can be selected under **RESTORE**.

Restoring a backup with an outdated NIC ITR layout skips the incompatible startup entry and reports a warning. The other supported settings are restored. Save NIC ITR again for that adapter; see [NIC ITR compatibility](./guide/06-network.md#nic-itr).

The main reset preserves MSI Mode / MSI Limit and Wi-Fi settings.
