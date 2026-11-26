# Network: RSS, NDIS and NIC ITR

[Documentation](../README.md) / [Guide](./README.md) / 06 · [Русский](./06-network.ru.md)

## Driver and affinity mode

For an NDIS adapter, choose **RSS** to configure receive-processing CPUs and queues, **IRQ** for interrupt affinity, or **BOTH** to configure both when RSS is available. Set the relevant fields and click **APPLY**. The RSS controls are available in NDIS cards. [NetAdapterCx drivers can also support RSS](https://learn.microsoft.com/en-us/windows-hardware/drivers/netcx/netadaptercx-receive-side-scaling-rss-).

## What RSS does

RSS maps received flows to processors using a hash and an indirection table. A driver with [one hardware receive queue](https://learn.microsoft.com/en-us/windows-hardware/drivers/network/rss-with-a-single-hardware-receive-queue) can schedule receive processing on other CPUs after the interrupt. With MSI-X, a supported adapter can route receive interrupts to the CPU that processes the data. Check the interrupt and receive paths in the trace when tuning both.

**RSS Queues** sets the requested number of receive queues. The program uses one logical CPU per physical core when choosing processors, within the adapter's limits. For example, starting at CPU 8 with two queues and at least two available cores can select CPUs 8 and 10 if CPU 9 is CPU 8's SMT sibling. Several queues can share a processor, as on some [Intel I219 adapters](https://www.intel.com/content/www/us/en/support/articles/000017506/network-and-io.html).

AUTO requests the smallest count the driver advertises. For example, if the manual options are two and four queues, AUTO selects two. If the driver provides no list, AUTO requests one. A single connection can stay on one queue even when several are enabled.

After **APPLY**, the log compares the saved queue count, processor count and both ends of the CPU range, including processor groups, with the active RSS configuration. If the values differ, check again after Windows restarts. To read the active configuration in PowerShell, run:

```powershell
Get-NetAdapterRss -Name "<adapter name>" |
    Format-List *
```

More CPUs can add overhead without improving throughput. Select a count for the workload: [Setting the number of RSS processors](https://learn.microsoft.com/en-us/windows-hardware/drivers/network/setting-the-number-of-rss-processors). Compare queue counts using the traffic pattern of the application you are tuning.

## NIC ITR

ITR controls hardware interrupt moderation on supported controllers. **SET** applies a value now, **SAVE** writes the startup profile, and **CHECK** reads registers again. Writes are blocked for unknown PCI IDs.

**SAVE** binds the entry to that adapter's PCI instance. Older entries identified an adapter model. At startup, the script skips an older entry if it matches more than one connected adapter. Save the desired values from each adapter's card to set them independently.

Intel EITR/ITR and Realtek IntrMit profiles use different registers and units. The range depends on the adapter family. Use **CHECK** to read the current values.

Adaptive driver moderation can overwrite a manually applied ITR. **SAVE** disables supported driver-moderation settings in the registry. The driver reads the settings at adapter start. Use [DELETE PROFILE](./05-usb-imod.md#delete-or-restore) to remove the saved profile and restore applicable defaults.

If the card reports incomplete moderation setup after **SAVE**, the startup profile was written. See the log for the driver setting that could not be read or changed.

NIC ITR selects the register bank from PCI configuration and checks it against the memory ranges allocated by Windows. It rejects inaccessible or inconsistent resources and register spans overlapping an MSI-X table or pending-bit array.

When you install or restore a profile, the application checks each NIC ITR entry against the current register layout for its PCI ID. An incompatible entry is skipped. The restore dialog and log identify the affected adapter. Save NIC ITR again for the affected adapter. The original backup is retained.

## Verification

Wi-Fi is excluded from AUTO and from resets of wired network settings. Compare active RSS, ITR readback, packet loss, CPU load and latency under the same traffic. Less moderation can increase interrupt rate and CPU load.

The [verification chapter](./10-verification.md) covers packet-processing traces and before/after comparisons.

---

[Previous: USB and IMOD](./05-usb-imod.md) · [Guide](./README.md) · [Next: CPU topology](./07-cpu-topology.md)
