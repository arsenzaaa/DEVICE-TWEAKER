# Network: RSS, NDIS and NIC ITR

## Two kinds of drivers

How to configure an adapter depends on its driver:

- **NetAdapterCx** — data goes through WDF01000 ISR / DPC and the NetAdapterCx system threads (separate for receive and transmit). Interrupts move with an ordinary mask.
- **NDIS** (a classic miniport) — everything goes through NDIS: ISR / DPC and, with MMCSS active, the NDIS thread. Receive and transmit are not split. **An Interrupt Affinity mask does not work on NDIS**; RSS distributes the interrupts instead.

## RSS

**RSS** (Receive Side Scaling) spreads incoming packets over queues and queues over processors. The settings live in the network adapter class key, in your adapter's subkey (look for `DriverDesc` with its name):

```text
HKLM\SYSTEM\CurrentControlSet\Control\Class\{4d36e972-e325-11ce-bfc1-08002be10318}\00NN
  *RSS                REG_SZ   1 = RSS on
  *NumRssQueues       REG_SZ   number of queues
  *RssBaseProcNumber  REG_SZ   first processor
  *MaxRssProcessors   REG_SZ   how many processors may be used
  *RssMaxProcNumber   REG_SZ   last processor
```

The values are strings (REG_SZ), as in the driver INF.

- **Queues.** Without MSI-X all queues arrive as one interrupt, so RSS across cores requires MSI-X. Online-match traffic is small: measurements show at most two queues actually work, and at a few hundred kbit/s only one. One or two queues are enough for a game.
- **Processors** are taken one per physical core: SMT siblings are skipped, otherwise two queues end up on one SMT-SET.
- `*NumRssQueues` is a request to the driver. The active state is reported by `Get-NetAdapterRss` and may differ.
- The `NdisRssProfileBalanced` profile has no manual base CPU.

## The NDIS thread and MMCSS

With MMCSS active, NDIS work moves from the DPC into an **NDIS thread** at priority 8, and the NDIS DPC time drops noticeably. The total network processing time stays about the same, but a game thread can preempt that thread, while it cannot preempt a DPC. If FPS is the goal, the NDIS thread is preferable to the DPC.

- The NDIS thread always runs **on the network adapter's core** and cannot be moved, so the adapter's core is that thread's core.
- With MMCSS off or `NetworkThrottlingIndex = FFFFFFFF`, all the work stays in the DPC.
- The thread priority is set by `HKLM\SYSTEM\CurrentControlSet\Services\NDIS\Parameters\ReceiveWorkerThreadPriority` (default 8).

NDIS is the more convenient choice with FACEIT / Vanguard and MMCSS active: you know exactly which core the NDIS thread will be on.

## NIC ITR

**ITR** (Interrupt Throttling Rate) is the network controller's own hardware interrupt moderation, the counterpart of USB IMOD. The implementation depends on the family: some chips set receive and transmit separately, others share one interval.

```text
Intel I210 / I211, I225 / I226, I350, 82576 / 82580   EITR
Intel I219                                            ITR
Realtek RTL8111 / 8168, RTL8125 / 8126                IntrMit
```

Compatible Killer adapters behave like their base chips. An Intel EITR value is not a ready Realtek profile: the units differ.

There is no single rule here. One approach turns moderation off: the network produces frequent short interrupts, and for latency they are better not batched. The other uses Medium / Adaptive: too many interrupts under load hurt input. A shorter interval → more interrupts and more CPU load. That is why DEVICE TWEAKER configures NIC ITR **manually only**.

## The driver's moderation

The adapter driver has its own adaptive moderation (`*InterruptModeration`, on Intel also `ITR`), and it overwrites the register. **SAVE** turns it off in the registry; otherwise a saved ITR value will not hold. The change takes effect after a restart, the current state is shown in the **Interrupt Moderation** line, and **DELETE** restores the driver value.

## What DEVICE TWEAKER does

- For NDIS it shows the **RSS / IRQ / both** mode, the stored and active RSS parameters separately, and logs mismatches. AUTO-OPTIMIZATION picks the mode from the adapter's state.
- The queue count is limited by the driver's values and the MSI-X vectors; processors are taken one per physical core.
- NIC ITR: **SET** writes the register now, **SAVE** stores the profile for startup, **CHECK** re-reads the registers. Writing is blocked for unknown PCI IDs.
- Wi-Fi is left untouched.

**Conclusion:** on NDIS use RSS instead of a mask, one or two queues on separate physical cores, MSI-X required. NIC ITR is a deliberate trade between latency and load, and only after the driver's moderation is off.

---

[← Chapter 05](./05-usb-imod.en.md) · [Contents](./README.en.md) · [Chapter 07 →](./07-cpu-topology.en.md)
