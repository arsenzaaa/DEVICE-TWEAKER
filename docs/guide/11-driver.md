# DTIMOD.sys, KDU and startup

[Documentation](../README.md) / [Guide](./README.md) / 11 · [Русский](./11-driver.ru.md)

## Operations that need DTIMOD

USB IMOD, NIC ITR and hardware PCI queries use DTIMOD. MSI, interrupt affinity, RSS, ReservedCpuSets, Mouse Throttle and device power settings use Windows interfaces. [IRQ Mode](./02-msi.md) combines Windows resource information with additional PCI data when the driver is available.

## Loading through KDU

DTIMOD has no public Microsoft signature for ordinary loading. The bundled [KDU](../../src/IMOD/KDU/README.md) loader uses a vulnerable signed helper driver to load it. HVCI, the Windows driver blocklist and antivirus software can block this loading method; Microsoft Defender may identify KDU as HackTool. Loading errors are recorded in the operation details and log.

Anti-cheat software can block DTIMOD or disconnect a game while it is loaded. EAC can report `Forbidden Tool`; compatibility with Vanguard and FACEIT is not established. Closing DEVICE TWEAKER closes its connection to DTIMOD. Restart Windows to unload the driver or replace a previously loaded build.

The application asks for confirmation once per session before driver use or saving/restoring a driver profile. AUTO includes this in its USB IMOD prompt. **DRIVER INFO** reads HVCI, the driver-blocklist registry setting and FACEIT/Vanguard service states. This is a system-status report, not a driver-loading or compatibility test.

The Windows service used by the fallback is `DeviceTweakerImod` (**DEVICE TWEAKER IMOD**). Both loaders verify the driver file before loading it.

## Saved profile

The `\DEVICE TWEAKER\IMOD` task runs as SYSTEM at sign-in, after resume and on relevant USB/HID device events. It reapplies the saved USB IMOD and NIC ITR settings independently of the application.

```text
%ProgramData%\DEVICE TWEAKER\Startup
  ApplyIMOD.ps1
  DeviceTweakerImodRuntime.cs
  profile.json
  Loader\
  logs\
```

`profile.json` stores register settings and component hashes. The script checks the shared runtime's SHA-256 before compilation, then checks DTIMOD and the loader files when loading the driver. The application and script share a load mutex. Only administrators and SYSTEM can modify the directory; startup logs retain the latest 14 files.

When loading a profile in the current format, the application can update outdated startup files and drop [NIC ITR entries with an incompatible register layout](./06-network.md#nic-itr). If reading or parsing the profile fails, the existing configuration stays in place. See [Delete or restore](./05-usb-imod.md#delete-or-restore) for profile and driver-file removal.

A task disabled in Windows Task Scheduler stays disabled when the profile is saved or the task is updated. The startup log records skipped adapters separately from application errors. An error during application sets the task result to a nonzero exit code.

## Controller lifecycle during register access

The target PCI controller must remain connected and powered during register access, including startup-profile application. Reconnecting a USB peripheral normally leaves its host controller present.

DTIMOD validates addresses against the controller's allocated MMIO ranges. It does not coordinate access with Windows PnP or power transitions: resources may change during [surprise removal](https://learn.microsoft.com/en-us/windows-hardware/drivers/kernel/handling-an-irp-mn-surprise-removal-request) or [stop/rebalance](https://learn.microsoft.com/en-us/windows-hardware/drivers/kernel/handling-an-irp-mn-stop-device-request--windows-2000-and-later-).

## Register access and USB mapping

Register operations identify the PCI device and use its allocated MMIO ranges. They support xHCI (`0C0330`) and Ethernet (`0200`) classes. Read-only PCI configuration queries also supply IRQ information for other device classes.

The USB map contains root-port, USB-address and interrupter tuples decoded from xHCI contexts. The driver reads system-memory contexts with [MmCopyMemory](https://learn.microsoft.com/en-us/windows-hardware/drivers/ddi/ntddk/nf-ntddk-mmcopymemory). Driver access permissions and request validation are described in the [security policy](../../.github/SECURITY.md).

## Release files and verification

`SHA256SUMS.txt` identifies the EXEs in a release; the source SHA in the release description identifies their code. Register readback, startup logs and device checks are described in [chapter 10](./10-verification.md).

---

[Previous: Verifying results](./10-verification.md) · [Guide](./README.md)
