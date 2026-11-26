# Settings reference

[Documentation](./README.md) / Settings · [Русский](./SETTINGS.ru.md)

[Interface](#interface) · [Interrupts](#interrupts) · [USB xHCI IMOD](#usb-xhci-imod) · [Network adapters](#network-adapters) · [Additional settings](#additional-settings) · [Auto-optimization and restore](#auto-optimization-and-restore)

## Interface

Toggles, scrolling and button hover use short transitions. Animations work independently of Windows visual effects. High Contrast uses immediate transitions.

The interface language is saved in `%LOCALAPPDATA%\DEVICE TWEAKER\settings.json`.

## Interrupts

### MSI / MSI-X

Choose a value in **MSI Mode** and click **APPLY** to write `MSISupported` to the device's registry key. Restart Windows before checking whether its driver uses the new setting. Disabling MSI is refused when PCI capabilities explicitly report message interrupts without line-based INTx support.

**IRQ Mode** reports Line, MSI or MSI-X. It shows MSI/MSI-X when it can confirm message-signaled interrupts but cannot distinguish the two modes. See [MSI and MSI-X](./guide/02-msi.md) for examples.

**READ IRQ MODE** queries every PCI device in the list, including cards hidden by filters. It reads allocated Windows interrupt resources and, through DTIMOD, the PCI `MSI Enable` / `MSI-X Enable` bits. Reading these bits distinguishes MSI from MSI-X even with one IRQ. The command can load DTIMOD; it does not apply settings. The button shows progress and the number of devices with a result out of the total checked. Its tooltip lists how many results were unavailable.

Opening the program or pressing **REFRESH** also starts background reads, using Windows resources or an already loaded DTIMOD. These reads do not load the driver. These background checks clear the old result before rereading. The card and log distinguish unavailable data, read failures and conflicting information. **No IRQ** means no allocated interrupt resources. A successful read describes the state at that moment: [MSI-X routing can change at runtime](https://learn.microsoft.com/en-us/windows-hardware/drivers/kernel/dynamically-configuring-msi-x).

**HW** lists the device's supported interrupt modes and maximum message count. **IRQ** shows the interrupt-resource count from WMI. The configured limit is shown separately in **MSI Limit**.

**MSI Limit** sets `MessageNumberLimit`. **Unlimited** removes the custom registry limit so Windows and the driver can select the count. Numeric input is checked against detected capabilities. See [enabling MSI](https://learn.microsoft.com/en-us/windows-hardware/drivers/kernel/enabling-message-signaled-interrupts-in-the-registry).

**IRQ Priority** writes `DevicePriority`, the hardware interrupt-priority preference used by Windows. Values and policy behavior are covered in the [affinity chapter](./guide/03-affinity.md#irq-priority).

### Interrupt Affinity

**CPU Affinity** selects the logical processors eligible to service a device's interrupts. The list shows SMT siblings and available topology information, including P-/E-cores on hybrid CPUs, CCD/CCX and CPPC rank. The selection covers processor group 0, up to 64 logical processors.

**Policy** controls how Windows selects processors. Use **SpecCPU** for the mask shown in the card. On devices with an affinity picker, clearing the selection and clicking **APPLY** removes the custom mask and restores **MachineDefault**. NDIS adapters use the selected RSS/IRQ mode. The [affinity chapter](./guide/03-affinity.md) lists the policies and registry values.

## USB xHCI IMOD

IMOD sets the xHCI interrupt moderation interval and uses [DTIMOD](./guide/11-driver.md). **XHCI** uses one value for the whole controller. **Devices** uses values for detected roles when the device map is available. **Interrupters** lets you set values by index. See the [USB chapter](./guide/05-usb-imod.md) for interval units, role mapping and AUTO values.

Hover over the current value to see when it was read. If a device is connected or removed, the map is marked "devices changed". Click **CHECK** to read it again.

### Commands

- **SET** writes the selected controller's registers and refreshes the displayed values. Custom settings are saved for startup after a successful write.
- **CHECK** rereads the registers and loads DTIMOD if needed. **REFRESH** updates the device list and available current state; hardware reads use an already loaded DTIMOD.
- **DELETE PROFILE** backs up and removes the shared USB IMOD and NIC ITR startup profile. Restart Windows to clear active register changes. See [Delete or restore](./guide/05-usb-imod.md#delete-or-restore) for network moderation and driver-file cleanup.

The main **APPLY** button also applies USB IMOD along with the other device settings.

### IMOD / NIC ITR at startup

The saved profile is stored in `%ProgramData%\DEVICE TWEAKER\Startup`. The `\DEVICE TWEAKER\IMOD` task runs as SYSTEM at sign-in, after resume and on relevant USB/HID events. It applies settings independently of the EXE. A profile in the previous AppData/Run layout is converted when you save new settings or restore a backup.

Files, hash checks, permissions and logs are described in the [driver chapter](./guide/11-driver.md#saved-profile).

## Network adapters

### RSS

**[RSS](https://learn.microsoft.com/en-us/windows-hardware/drivers/network/introduction-to-receive-side-scaling)** distributes incoming packet processing across CPUs using a hash and an indirection table. **RSS Queues** sets the requested receive-queue count; the base CPU is the starting point for processor selection. **NDIS Mode** applies RSS, Interrupt Affinity, or both.

Queue choices follow the adapter driver's advertised options. RSS processor selection uses one logical CPU per physical core. AUTO queue selection and the distinction between CPUs and queues are covered in the [network chapter](./guide/06-network.md#what-rss-does).

The card shows the saved queue setting and a summary of active RSS. The log compares the requested settings with the adapter's current values. With the [Balanced profile](https://learn.microsoft.com/en-us/windows-hardware/drivers/network/standardized-inf-keywords-for-rss), Windows chooses the processor range automatically, while the queue count and maximum processor count remain configurable. Restart Windows to let the adapter driver read the saved values.

**Affinity (RSS mask)** shows the processors selected by DEVICE TWEAKER. Check the active RSS state to see what the adapter adopted.

### NIC ITR

**NIC ITR** accesses hardware moderation registers on supported network controllers: Intel I210/I211, I219, I225/I226, I350, 82576/82580, Realtek RTL8111/8168, RTL8125/8126, and compatible Killer models. Hardware writes are disabled for unknown PCI IDs.

**SET** writes the register now. **SAVE** stores a startup profile for the adapter. **CHECK** reads the current register values.

**SAVE** turns off supported driver-moderation settings (`*InterruptModeration` and, on applicable Intel drivers, `ITR`). The driver reads them when the adapter starts. **Interrupt Moderation** in the card shows the configured driver values. Use **DELETE PROFILE** in the USB IMOD block to remove the saved profile and restore [applicable driver defaults](./guide/05-usb-imod.md#delete-or-restore).

## Additional settings

**ReservedCpuSets** defines the system-wide CPU reservation mask. Selection and storage format are described in the [CPU chapter](./guide/07-cpu-topology.md#reservedcpusets).

**Power Saving** controls available settings for the USB controller and its root hubs, including [USB selective suspend](https://learn.microsoft.com/en-us/windows-hardware/drivers/usbcon/usb-selective-suspend). For a wired network adapter, the switch changes whether Windows may turn off the device to save power. Availability depends on the driver.

**Mouse Throttle** sets the interval Windows applies to eligible background Raw Input listeners. Switching it off removes the custom registry value. See [Mouse Throttle](./guide/08-input.md#mouse-throttle) for the supported Windows builds and interval units.

**Raw Input measurements** estimate the mouse event rate while the mouse is moving. The input chapter explains [event delivery and polling rate](./guide/08-input.md).

The **MOUSE**, **KEYBOARD**, **GAMEPAD**, **USB**, **GPU**, **NETWORK**, **AUDIO**, and **STORAGE** category buttons jump to matching devices in the full list. Text search hides nonmatching devices.

## Auto-optimization and restore

**AUTO-OPTIMIZATION** selects physical cores from topology and CPPC rankings. On multi-CCD systems it selects a target CCD separately; two SMT threads of one core are not counted as two independent cores. A discrete GPU gets two distinct physical cores; an integrated GPU's affinity is returned to Windows defaults. Input devices, wired network adapters, and audio have separate sharing rules.

Assignments without a suitable core or a reliably detected role are skipped with a reason in the report. If CPU topology cannot be determined reliably, affinity masks are not written. On NDIS adapters, AUTO chooses RSS, IRQ, or BOTH based on the adapter's current settings and capabilities. Choose whether to include USB IMOD before starting. NIC ITR is configured manually.

Auto-optimization clears [ReservedCpuSets](./guide/07-cpu-topology.md#reservedcpusets) when the reservation panel is available.

The application creates and verifies a backup before manual **APPLY**, USB IMOD or NIC ITR **SET**, NIC ITR **SAVE**, restore and reset. Auto-optimization requires a backup of the original settings. The extra backup is optional. If a required backup fails, the operation stops before writing. Choose a backup under **RESTORE**.

Backups store Windows registry paths for the device and driver instances present at the time. Use them on the Windows installation and hardware configuration where they were created.

Restoring a backup with an outdated NIC ITR layout skips the incompatible startup entry and reports a warning. The other supported settings are restored. Save NIC ITR again for that adapter; see [NIC ITR compatibility](./guide/06-network.md#nic-itr).

The main reset preserves MSI Mode / MSI Limit and Wi-Fi settings.
