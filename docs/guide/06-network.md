# Network: RSS, NDIS and NIC ITR

[Documentation](../README.md) / [Guide](./README.md) / 06 · [Русский](./06-network.ru.md)

## Driver and affinity mode

DEVICE TWEAKER distinguishes NDIS and NetAdapterCx and exposes the settings available for the adapter. NDIS supports RSS, IRQ and combined modes in the application. Feature availability and processing paths depend on the NIC and driver.

RSS and Interrupt Affinity are not mutually exclusive: one distributes network processing, the other controls interrupt assignment. Their interaction depends on the NIC and driver.

For an NDIS adapter, choose **RSS** to configure receive-processing CPUs and queues, **IRQ** for interrupt affinity, or **BOTH** for both when RSS is supported. Select the CPUs and queue count, then click **APPLY**. NetAdapterCx cards expose the controls available for that adapter type. NIC ITR has separate SET/SAVE commands, described below.

## What RSS does

RSS distributes receive processing across CPUs using hashing and an indirection table. Capabilities vary: Microsoft describes single-queue, multiple-queue and MSI hardware support. See [Introduction to receive-side scaling](https://learn.microsoft.com/en-us/windows-hardware/drivers/network/introduction-to-receive-side-scaling).

The driver can schedule receive DPCs on target processors. See [Receive-side scaling support](https://learn.microsoft.com/en-us/windows-hardware/drivers/network/receive-scaling). Multiple queues do not imply that one connection will use them all.

The application selects physical cores with SMT awareness, queue count and available adapter limits. Saved parameters are compared with active RSS state; a mismatch is not presented as confirmed success.

```powershell
Get-NetAdapterRss -Name "<adapter name>" |
    Format-List *
```

More CPUs can add overhead without improving throughput. Select a count for the workload: [Setting the number of RSS processors](https://learn.microsoft.com/en-us/windows-hardware/drivers/network/setting-the-number-of-rss-processors). Compare queue counts using the traffic pattern of the application you are tuning.

## NIC ITR

ITR controls hardware interrupt moderation on supported controllers. **SET** applies a value now, **SAVE** writes the startup profile, and **CHECK** reads registers again. Writes are blocked for unknown PCI IDs.

Intel EITR/ITR and Realtek IntrMit profiles use different registers and units. Follow the selected family's description and range instead of copying another adapter's numeric profile. Matching readback confirms a register value, not better network latency.

Adaptive driver moderation can overwrite a manually applied ITR. The application accounts for available driver settings and returns control to the driver when the saved profile is removed. Operation results retain details and unconfirmed steps.

NIC ITR selects the register bank from PCI configuration and checks it against the memory ranges allocated by Windows. It rejects inaccessible or inconsistent resources and register spans overlapping an MSI-X table or pending-bit array.

The NIC layouts for Killer E3000 (`10EC:3000`, RTL8125 family) and RTL8118AS (`10EC:8162`, RTL8168 family) have been corrected. When a saved NIC ITR entry no longer matches the supported layout for its PCI ID, profile installation and backup restoration omit that entry. Restore reports the omission as a warning and records the old and current layout in the log. Other supported entries are retained; the backup file is unchanged. Save NIC ITR again for the affected adapter to create a current startup entry.

## Verification

Wi-Fi is excluded from AUTO and supported wired-setting resets. Compare active RSS, ITR readback, packet loss, CPU load and latency under the same traffic. Less moderation can increase interrupt rate and CPU load.

MMCSS and `NetworkThrottlingIndex` are outside the application's settings. Use a trace to inspect network worker-thread placement and priority for the installed driver.

---

[Previous: USB and IMOD](./05-usb-imod.md) · [Guide](./README.md) · [Next: CPU topology](./07-cpu-topology.md)
