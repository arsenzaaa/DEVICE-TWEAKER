# DTIMOD.sys, KDU and startup

[Documentation](../README.md) / [Guide](./README.md) / 11 · [Русский](./11-driver.ru.md)

## When DTIMOD is needed

Only **USB IMOD**, **NIC ITR** and hardware PCI configuration queries require `DTIMOD.sys`. Standard settings (MSI Mode/Limit, Interrupt Affinity, RSS, ReservedCpuSets, Mouse Throttle and device power policies) use regular Windows registry, WMI and Power APIs. [IRQ Mode](./02-msi.md) combines Windows allocated interrupt resources with PCI configuration bits (`MSI Enable` / `MSI-X Enable`) whenever DTIMOD is available.

## Loading through KDU and compatibility boundaries

`DTIMOD.sys` is signed with a test certificate and has no public Microsoft WHQL signature for standard Driver Signature Enforcement (DSE). The bundled [KDU](../../src/IMOD/KDU/README.md) loader maps it via a signed vulnerable helper driver. Because of this mechanism:

- **HVCI (Memory Integrity)**, the **Microsoft Vulnerable Driver Blocklist** and antivirus tools can block loading; Microsoft Defender frequently flags KDU as `HackTool`. Failure details are written to the operation report and application log.
- **Kernel anti-cheats** may reject DTIMOD or kick an active game session while the driver is resident in memory (EAC can report `Forbidden Tool`; compatibility with Vanguard and FACEIT is not established). Closing DEVICE TWEAKER only closes its handle to the driver. **Reboot Windows** to unload DTIMOD or replace an older resident build.

DEVICE TWEAKER requests confirmation once per session before loading the driver or saving/restoring a driver-backed profile (in AUTO, this confirmation is combined with the USB IMOD prompt). **DRIVER INFO** checks HVCI status, the driver-blocklist registry key and FACEIT/Vanguard service states; it reports system configuration only and does not load DTIMOD or test anti-cheat compatibility.

The fallback service name is `DeviceTweakerImod` (**DEVICE TWEAKER IMOD**). Both KDU and service-based loading verify `DTIMOD.sys` against its expected SHA-256 hash before loading.

## Saved profile

Hardware MMIO register overrides do not survive a reboot or controller reset, so the scheduled task `\DEVICE TWEAKER\IMOD` runs as `SYSTEM` at sign-in, after resume from sleep and on matching USB/HID device events to reapply saved USB IMOD and NIC ITR values independently of the GUI EXE:

```text
%ProgramData%\DEVICE TWEAKER\Startup
  ApplyIMOD.ps1
  DeviceTweakerImodRuntime.cs
  profile.json
  Loader\
  logs\
```

`profile.json` stores per-instance register settings together with expected component hashes. Before compiling the shared C# runtime, `ApplyIMOD.ps1` verifies its SHA-256; when loading DTIMOD, it also verifies the driver and loader binaries. A shared mutex serializes loading between the GUI and the startup script. The startup directory ACL restricts writes to Administrators and `SYSTEM`, and startup logs keep the latest 14 files.

When reading a current-format profile, DEVICE TWEAKER updates outdated startup files automatically and skips [NIC ITR entries whose register layout changed](./06-network.md#nic-itr). If reading or parsing `profile.json` fails, existing files are left untouched rather than overwritten. See [Delete or restore](./05-usb-imod.md#delete-or-restore) for removing the profile and staged driver file.

If you manually disable `\DEVICE TWEAKER\IMOD` in Task Scheduler, saving a profile or updating the task preserves its disabled state. Startup logs distinguish skipped adapters from apply failures, and any write failure returns a nonzero task exit code.

## Controller lifecycle during register access

The target PCI controller must remain present and powered throughout register access, including during startup-profile application (unplugging a USB peripheral normally leaves its xHCI host controller online).

DTIMOD validates every requested MMIO address against the controller's allocated memory ranges, but it does **not** synchronize register access with Windows PnP or power state transitions: hardware resources can change during [surprise removal](https://learn.microsoft.com/en-us/windows-hardware/drivers/kernel/handling-an-irp-mn-surprise-removal-request) or [stop/rebalance](https://learn.microsoft.com/en-us/windows-hardware/drivers/kernel/handling-an-irp-mn-stop-device-request--windows-2000-and-later-).

## Scoped register access and USB mapping

Every MMIO request identifies a specific PCI device (by instance ID and PDO) and is restricted to its allocated memory ranges within the **xHCI (`0C0330`)** and **Ethernet (`0200`)** device classes; generic physical-memory read/write IOCTLs and I/O port access are not exposed. Read-only PCI configuration queries remain available for IRQ Mode detection on other PCI classes.

To build the USB device-to-interrupter map without exposing arbitrary kernel memory, DTIMOD decodes xHCI controller contexts internally via [MmCopyMemory](https://learn.microsoft.com/en-us/windows-hardware/drivers/ddi/ntddk/nf-ntddk-mmcopymemory) and returns only `(root port, USB address, interrupter)` tuples. Device-object ACLs and validation rules are documented in the [security policy](../../.github/SECURITY.md).

## Release verification

Each release includes `SHA256SUMS.txt` covering both EXEs, and the release description cites the exact Git source SHA they were built from. For post-change validation, register readback and ETW traces, see [Chapter 10](./10-verification.md).

---

[Previous: Verifying results](./10-verification.md) · [Guide](./README.md)
