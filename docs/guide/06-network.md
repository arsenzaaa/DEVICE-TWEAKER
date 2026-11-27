# Network: RSS, NDIS and NIC ITR

[Documentation](../README.md) / [Guide](./README.md) / 06 · [Русский](./06-network.ru.md)

## NDIS vs NetAdapterCx and affinity modes

Network drivers in Windows follow two main architectures:

- **NDIS** miniports handle incoming traffic in DPCs (or in an NDIS receive worker thread when MMCSS network throttling is active) and distribute receive processing through **RSS** registry parameters (`*RssBaseProcNumber`, `*NumRssQueues`, `*MaxRssProcessors`, `*RssMaxProcNumber`, `*RssBaseProcGroup`, `*RSSMaxProcGroup`, plus `*RSSProfile = 4` / `NUMAScalingStatic` so `ndis.sys` does not fall back to `Balanced` on hybrid CPUs and ignore the processor range). On many RSS-capable NDIS adapters, writing only an `Affinity Policy` mask either leaves receive DPCs on their original cores or shifts only the ISR, which is why NDIS cards in DEVICE TWEAKER let you choose **RSS**, **IRQ** or **BOTH**.
- **NetAdapterCx** (WDF) drivers run packet queues in dedicated `ExecutionContext` system threads inside `System`. DEVICE TWEAKER exposes the RSS controls on NDIS cards; [NetAdapterCx drivers can also support RSS](https://learn.microsoft.com/en-us/windows-hardware/drivers/netcx/netadaptercx-receive-side-scaling-rss-), though their interrupt and thread affinity behave differently.

## What RSS does

RSS hashes incoming packet headers (IP addresses and ports) and maps each flow through an indirection table to a target CPU. Even an adapter with [a single hardware receive queue](https://learn.microsoft.com/en-us/windows-hardware/drivers/network/rss-with-a-single-hardware-receive-queue) can hand off packet processing to other CPUs after the initial interrupt; with **MSI-X**, a multi-queue adapter can deliver each queue's interrupt directly to the CPU that processes its data.

**RSS Queues** sets the requested queue count within the options advertised by the driver. When choosing processors from the base CPU upward, DEVICE TWEAKER picks **one logical CPU per physical core** (skipping SMT siblings). For example, starting at CPU 8 with two queues selects CPUs 8 and 10 when CPU 9 is CPU 8's SMT sibling. Some controllers (including certain [Intel I219 adapters](https://www.intel.com/content/www/us/en/support/articles/000017506/network-and-io.html)) map multiple queues onto one processor or offer limited hardware RSS support regardless of registry settings.

Because a single TCP or UDP flow (such as an online match connected to one game server) hashes to a **single** RSS queue, enabling 8–16 queues does not speed up that connection; it only spreads cache state and adds scheduling overhead ([Setting the number of RSS processors](https://learn.microsoft.com/en-us/windows-hardware/drivers/network/setting-the-number-of-rss-processors)). That is why **AUTO-OPTIMIZATION** picks the smallest queue count the driver advertises (falling back to 1 when the driver exposes no list).

After **APPLY**, the log compares the saved queue count, processor count and both ends of the CPU range (including processor groups) against the adapter's live RSS state. If the active state has not updated yet, reboot and recheck in PowerShell:

```powershell
Get-NetAdapterRss -Name "<adapter name>" |
    Format-List *
```

## NIC ITR

**NIC ITR** reads and writes hardware interrupt-moderation registers directly on supported Ethernet controllers (Intel `EITR`/`ITR` and Realtek `IntrMit` families use different registers, units and valid ranges). **SET** writes the register immediately, **SAVE** stores a startup profile, and **CHECK** reads current hardware values. Writes are blocked for unrecognized PCI IDs.

**SAVE** binds each profile entry to that adapter's exact PCI device instance. Legacy model-based entries are skipped at startup if they match more than one connected adapter; save from each adapter's card to configure them independently.

Because adaptive driver moderation can overwrite manual register values at runtime, **SAVE** disables supported driver-level moderation settings (`*InterruptModeration` and applicable Intel `ITR` keys), which the driver reads on adapter start. If the card reports incomplete moderation setup after **SAVE**, the startup profile was written, and the log identifies which driver setting could not be read or updated. Use [DELETE PROFILE](./05-usb-imod.md#delete-or-restore) to remove the startup profile and restore applicable INF defaults.

Before touching MMIO, DEVICE TWEAKER resolves the register bank from PCI BAR configuration, validates it against Windows-allocated memory resources, and rejects inaccessible ranges or spans overlapping an MSI-X table or Pending Bit Array (PBA). When installing a profile or restoring a backup, entries whose stored register layout no longer matches the current definition for that PCI ID are skipped with a warning while retaining the original backup; save NIC ITR again for the reported adapter.

## Verification and trade-offs

Wi-Fi adapters are excluded from AUTO and wired network resets. Lowering or disabling interrupt moderation reduces per-packet coalescing delay, but under sustained high-throughput traffic (downloads, streams, multi-gigabit transfers) it significantly increases ISR/DPC frequency, which can hurt CPU efficiency and input responsiveness. Always compare active RSS, ITR readback, packet loss, CPU load and trace behavior under the same traffic pattern; see [Chapter 10](./10-verification.md).

---

[Previous: USB and IMOD](./05-usb-imod.md) · [Guide](./README.md) · [Next: CPU topology](./07-cpu-topology.md)
