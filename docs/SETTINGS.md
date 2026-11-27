# Settings reference

[Documentation](./README.md) / Settings · [Русский](./SETTINGS.ru.md)

[Interface](#interface) · [Interrupts](#interrupts) · [USB xHCI IMOD](#usb-xhci-imod) · [Network adapters](#network-adapters) · [Additional settings](#additional-settings) · [Auto-optimization and restore](#auto-optimization-and-restore)

## Interface

Toggles, scrolling and button hover use short transitions driven by a shared frame clock that stops while idle, working independently of Windows visual effects and switching to immediate updates under High Contrast.

The interface language is stored in `%LOCALAPPDATA%\DEVICE TWEAKER\settings.json`.

## Interrupts

### MSI / MSI-X

Select a value in **MSI Mode** and click **APPLY** to write `MSISupported` to the device's registry key, then reboot before checking whether the driver adopted it. If PCI capabilities explicitly report message-signaled interrupts without line-based INTx support, disabling MSI is refused.

**IRQ Mode** reports `Line`, `MSI`, or `MSI-X` (or `MSI/MSI-X` when Windows resources confirm message interrupts without PCI config readback to tell them apart). See [MSI and MSI-X](./guide/02-msi.md).

**READ IRQ MODE** queries every PCI device in the list (including cards hidden by filters), combining Windows allocated interrupt resources with PCI `MSI Enable` and `MSI-X Enable` bits via DTIMOD so MSI and MSI-X can be distinguished even with a single allocated IRQ. It can load DTIMOD, never modifies device settings, displays progress and resolved/total counts on the button, and lists unavailable results in its tooltip.

Startup and **REFRESH** also run background reads using Windows resources and any already loaded DTIMOD without loading the driver. Previous readings are cleared before rereading, and both the card and log separate unavailable data, read errors and conflicting states. **No IRQ** means no allocated interrupt resources. A read reflects the state at that moment, since [MSI-X routing can change dynamically at runtime](https://learn.microsoft.com/en-us/windows-hardware/drivers/kernel/dynamically-configuring-msi-x).

**HW** shows the device's PCI interrupt capabilities and hardware message ceiling; **IRQ** shows the interrupt-resource count reported by WMI. Configured limits appear separately in **MSI Limit**.

**MSI Limit** sets `MessageNumberLimit`. **Unlimited** deletes the custom value so Windows and the driver negotiate the vector count normally; manual numbers are validated against detected capabilities ([Enabling MSI in the registry](https://learn.microsoft.com/en-us/windows-hardware/drivers/kernel/enabling-message-signaled-interrupts-in-the-registry)).

**IRQ Priority** writes `DevicePriority`, the DIRQL preference hint used by Windows when assigning IDT vectors (affecting ISR preemption level only, not DPC priority). See [IRQ Priority](./guide/03-affinity.md#irq-priority).

### Interrupt Affinity

**CPU Affinity** picks which logical processors in processor group 0 (up to 64 LPs) may service the device's interrupts. The list shows SMT siblings, P-/E-cores on hybrid CPUs, CCD/CCX domains and CPPC ranks.

**Policy** controls `DevicePolicy`: use **SpecCPU** (`4`) for the checked mask. On devices with a manual affinity picker, clearing all checkboxes and clicking **APPLY** deletes the custom mask and returns to **MachineDefault**. NDIS adapters use their dedicated **RSS / IRQ / BOTH** mode selector. See [Chapter 3](./guide/03-affinity.md) for all policy values and registry layout.

## USB xHCI IMOD

IMOD configures the xHCI interrupt-moderation interval (250 ns steps; Windows `usbxhci.sys` typically initializes controllers to `0xC8` / 50 µs vs `0xFA0` / 1 ms in the xHCI spec) through [DTIMOD](./guide/11-driver.md). **XHCI** applies one interval across the controller, **Devices** assigns per-role values when the xHCI slot map is available, and **Interrupters** sets values by index. See [Chapter 5](./guide/05-usb-imod.md) for role priority, polling-rate trade-offs and AUTO defaults.

Hover over the current value for the last read timestamp. If USB peripherals are connected or removed, the map is flagged "devices changed"; click **CHECK** to reread it.

### Commands

- **SET** writes the selected controller's registers, refreshes displayed readings, automatically expands the interrupter readback details table, and saves custom settings for startup once the write succeeds.
- **CHECK** rereads hardware registers (loading DTIMOD if needed). **REFRESH** updates the device list and live state without loading DTIMOD on its own.
- **DELETE PROFILE** backs up and removes the shared USB IMOD and NIC ITR startup profile; restart Windows to clear active MMIO register overrides. See [Delete or restore](./guide/05-usb-imod.md#delete-or-restore) for network moderation restore and driver-file cleanup.

The toolbar **APPLY** button also applies pending USB IMOD edits alongside other device settings.

### IMOD / NIC ITR at startup

The startup bundle lives in `%ProgramData%\DEVICE TWEAKER\Startup`. Task `\DEVICE TWEAKER\IMOD` runs as `SYSTEM` at sign-in, after resume and on matching USB/HID events, applying register values independently of the GUI EXE. Legacy AppData/Run layouts migrate automatically when saving settings or restoring a backup.

See [Saved profile](./guide/11-driver.md#saved-profile) for hash verification, directory ACLs and log retention.

## Network adapters

### RSS

**[RSS](https://learn.microsoft.com/en-us/windows-hardware/drivers/network/introduction-to-receive-side-scaling)** hashes incoming packet flows across CPUs via an indirection table. **RSS Queues** sets the requested receive-queue count; the base CPU defines where processor assignment begins; **NDIS Mode** applies **RSS**, **IRQ** (`Affinity Policy`), or **BOTH**.

Available queue counts follow the driver's advertised options, and processor selection takes one logical CPU per physical core (skipping SMT siblings). See [What RSS does](./guide/06-network.md#what-rss-does) for why AUTO chooses the minimum advertised queue count and how queues relate to CPU count.

Each NDIS card shows the saved queue setting alongside a live active-RSS summary, while the log compares requested values with the adapter's runtime state. When writing an RSS processor range, DEVICE TWEAKER sets `*RSSProfile = 4` (`NUMAScalingStatic`) so Windows does not default to the [Balanced profile](https://learn.microsoft.com/en-us/windows-hardware/drivers/network/standardized-inf-keywords-for-rss) on hybrid CPUs and ignore `*RssBaseProcNumber` / `*RssMaxProcNumber` (the original value is backed up and cleared on reset). Reboot Windows so the driver loads updated registry values.

**Affinity (RSS mask)** displays the CPUs chosen by DEVICE TWEAKER; check the active RSS summary or `Get-NetAdapterRss` to confirm what the adapter actually adopted.

### NIC ITR

**NIC ITR** accesses hardware moderation registers on supported Ethernet controllers: Intel I210/I211, I219, I225/I226, I350, 82576/82580, Realtek RTL8111/8168, RTL8125/8126 and compatible Killer adapters. Hardware writes are blocked for unknown PCI IDs.

**SET** writes the register immediately, **SAVE** writes a per-instance startup profile, and **CHECK** reads current hardware values.

**SAVE** disables supported driver-level moderation settings (`*InterruptModeration` and applicable Intel `ITR` keys) so adaptive moderation does not overwrite manual registers on adapter start; **Interrupt Moderation** in the card reflects configured driver values. Use **DELETE PROFILE** in the USB IMOD block to remove the startup profile and restore [applicable INF defaults](./guide/05-usb-imod.md#delete-or-restore).

## Additional settings

**ReservedCpuSets** sets the system-wide soft CPU reservation bitmask (`REG_BINARY` under `Session Manager\Kernel`), which also constrains system threads. See [ReservedCpuSets](./guide/07-cpu-topology.md#reservedcpusets) for details and the warning against mixing reserved and unreserved CPUs in a single hard process affinity mask.

**Power Saving** controls USB controller and root-hub power management (`SelectiveSuspendEnabled`, `EnhancedPowerManagementEnabled`, `PnPCapabilities` bit `0x08`, WMI `MSPower_DeviceEnable` and global [USB selective suspend](https://learn.microsoft.com/en-us/windows-hardware/drivers/usbcon/usb-selective-suspend)) as well as wired NIC power-down permission (`PnPCapabilities` / WMI), depending on driver support.

**Mouse Throttle** configures `HKCU\Control Panel\Mouse\RawMouseThrottleDuration` for eligible background Raw Input listeners on Windows 11 build `22621.1928+` (default 8 ms, valid range 1–20 ms; turning it off removes the custom value). See [Mouse Throttle](./guide/08-input.md#mouse-throttle) for background-app caveats.

**Raw Input measurements** report live mouse event delivery rates on a dedicated background message thread while moving the mouse ([Chapter 8](./guide/08-input.md)).

The **MOUSE**, **KEYBOARD**, **GAMEPAD**, **USB**, **GPU**, **NETWORK**, **AUDIO**, and **STORAGE** category buttons jump to matching cards in the list; text search filters out nonmatching devices.

## Auto-optimization and restore

**AUTO-OPTIMIZATION** allocates physical cores from Windows topology and CPPC rankings: on multi-CCD CPUs it targets a single CCD, avoids pairing SMT siblings as separate cores, gives a discrete GPU two neighboring physical cores, returns integrated graphics to Windows defaults, and applies dedicated placement rules for input, wired Ethernet and audio.

Devices lacking a suitable core or reliably identified role are skipped with a reason in the report; if CPU topology cannot be determined reliably, affinity masks are not written. On NDIS adapters, AUTO selects RSS, IRQ or BOTH from driver capabilities and current state. You choose whether to include USB IMOD before starting; NIC ITR remains manual.

AUTO clears [ReservedCpuSets](./guide/07-cpu-topology.md#reservedcpusets) when the reservation panel is available.

DEVICE TWEAKER creates and verifies a backup before manual **APPLY**, USB IMOD or NIC ITR **SET**, NIC ITR **SAVE**, **RESTORE** and reset. AUTO requires a backup of original settings (the extra snapshot is optional) and aborts before writing if a required backup fails. Select any snapshot under **RESTORE**.

Backups store exact Windows registry paths for the device and driver instances present when captured. Restore them only on the same Windows installation and hardware configuration.

When restoring a backup with an outdated NIC ITR register layout, the incompatible startup entry is skipped with a warning while all other supported settings are restored; save NIC ITR again for that adapter ([NIC ITR](./guide/06-network.md#nic-itr)).

**RESET WINDOWS DEFAULT** preserves MSI Mode / MSI Limit and Wi-Fi settings.
