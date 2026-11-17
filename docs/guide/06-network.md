# Network: RSS, NDIS and NIC ITR

[Documentation](../README.md) / [Guide](./README.md) / 06 · [Русский](./06-network.ru.md)

## Driver and affinity mode

DEVICE TWEAKER distinguishes NDIS and NetAdapterCx and exposes the settings available for the adapter. NDIS supports RSS, IRQ and combined modes in the application. Feature availability and processing paths depend on the NIC and driver.

RSS distributes receive processing across CPUs, while Interrupt Affinity controls interrupt assignment. BOTH configures the two mechanisms together.

For an NDIS adapter, choose **RSS** to configure receive-processing CPUs and queues, **IRQ** for interrupt affinity, or **BOTH** for both when RSS is supported. Select the CPUs and queue count, then click **APPLY**. NetAdapterCx cards expose the controls available for that adapter type. NIC ITR has separate SET/SAVE commands, described below.

## What RSS does

RSS distributes receive processing across CPUs using hashing and an indirection table. Capabilities vary: Microsoft describes single-queue, multiple-queue and MSI hardware support. See [Introduction to receive-side scaling](https://learn.microsoft.com/en-us/windows-hardware/drivers/network/introduction-to-receive-side-scaling).

The driver can schedule receive DPCs on target processors. See [Receive-side scaling support](https://learn.microsoft.com/en-us/windows-hardware/drivers/network/receive-scaling). Multiple queues do not imply that one connection will use them all.

The application selects physical cores with SMT awareness, queue count and available adapter limits. The card compares saved parameters with the active RSS state and reports differences.

**RSS Queues** is a receive-queue count. DEVICE TWEAKER selects eligible physical cores for RSS, skipping SMT siblings and respecting known adapter limits. More queues can share the available cores. For example, if CPU 9 is CPU 8's SMT sibling, two queues select CPU 8 and CPU 10. AUTO uses one queue on one physical core when configuring RSS.

Queue count can differ from processor count. [Intel I219 and related adapters](https://www.intel.com/content/www/us/en/support/articles/000017506/network-and-io.html) support up to two receive queues serviced by one CPU core. The application retains such driver-advertised queue options; without advertised options it uses a single-queue fallback for known non-MSI-X adapters.

```powershell
Get-NetAdapterRss -Name "<adapter name>" |
    Format-List *
```

More CPUs can add overhead without improving throughput. Select a count for the workload: [Setting the number of RSS processors](https://learn.microsoft.com/en-us/windows-hardware/drivers/network/setting-the-number-of-rss-processors). Compare queue counts using the traffic pattern of the application you are tuning.

## NIC ITR

ITR controls hardware interrupt moderation on supported controllers. **SET** applies a value now, **SAVE** writes the startup profile, and **CHECK** reads registers again. Writes are blocked for unknown PCI IDs.

Intel EITR/ITR and Realtek IntrMit profiles use different registers and units. The selected family's description defines the range; CHECK shows the current register values.

Adaptive driver moderation can overwrite a manually applied ITR. The application accounts for available driver settings and returns control to the driver when the saved profile is removed. Operation results retain details and unconfirmed steps.

NIC ITR selects the register bank from PCI configuration and checks it against the memory ranges allocated by Windows. It rejects inaccessible or inconsistent resources and register spans overlapping an MSI-X table or pending-bit array.

Profile installation and backup restoration check saved NIC ITR entries against the current register layout for each PCI ID. Incompatible entries are skipped with a warning; the log records both layouts. Save NIC ITR again for the affected adapter to update its startup entry. Other supported entries and the original backup are preserved.

## Verification

Wi-Fi is excluded from AUTO and supported wired-setting resets. Compare active RSS, ITR readback, packet loss, CPU load and latency under the same traffic. Less moderation can increase interrupt rate and CPU load.

The [verification chapter](./10-verification.md) covers packet-processing traces and before/after comparisons.

---

[Previous: USB and IMOD](./05-usb-imod.md) · [Guide](./README.md) · [Next: CPU topology](./07-cpu-topology.md)
