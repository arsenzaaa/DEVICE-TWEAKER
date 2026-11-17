# DTIMOD.sys, KDU and startup

[Documentation](../README.md) / [Guide](./README.md) / 11 · [Русский](./11-driver.ru.md)

## Operations that need DTIMOD

DTIMOD is used for USB IMOD, NIC ITR and additional PCI configuration reads. MSI, interrupt affinity, RSS, ReservedCpuSets, Mouse Throttle and device power settings work without it. [IRQ Mode](./02-msi.md) can obtain some information from Windows resources; the driver supplies additional device information.

## Loading through KDU

DTIMOD has no public Microsoft signature for ordinary loading. The project uses a pinned [KDU](../../src/IMOD/KDU/README.md) version that relies on a vulnerable signed driver. This is a highly privileged hardware tool; antivirus, HVCI, Vulnerable Driver Blocklist and anti-cheat software can prevent it from working.

DTIMOD can conflict with Easy Anti-Cheat, Riot Vanguard, FACEIT Anti-Cheat and other protection software. Successful driver loading does not confirm compatibility: a game may disconnect you while the driver is loaded. EAC may report `Forbidden Tool` while DTIMOD is loaded. Compatibility with Vanguard and FACEIT has not been established.

The application shows a notice once per session before driver use or saving/restoring a driver profile. AUTO includes the notice in its IMOD question. **DRIVER INFO** reads available Windows settings and selected service information without loading DTIMOD. It does not test compatibility, and absence of a registry value or service does not prove that the system permits the driver. Details are written to the log. The application does not offer to disable Windows protections after a loading error.

The application checks pinned file hashes and blocks fallback loading after an uncertain load outcome. Inspect the operation details for the reason. Restart Windows to replace an already loaded older driver or unload DTIMOD; closing the EXE only closes its connection.

## Saved profile

The `\DEVICE TWEAKER\IMOD` SYSTEM task has logon, resume and relevant PnP triggers. Its files are stored here:

```text
%ProgramData%\DEVICE TWEAKER\Startup
  ApplyIMOD.ps1
  DeviceTweakerImodRuntime.cs
  profile.json
  Loader\
  logs\
```

The runtime is shared with Windows PowerShell 5.1. Its SHA-256 is checked before compilation; DTIMOD and loader files are checked when loading the driver. A shared mutex prevents simultaneous loading by the application and startup script. Administrator/SYSTEM permissions protect the directory. Startup logs retain the latest 14 files.

The task runs independently of the EXE. Profile read or parse errors stop changes to the profile. **DELETE** in the IMOD block removes the shared startup configuration for USB IMOD and NIC ITR, including the scheduled task.

When an older NIC ITR entry no longer matches its supported register layout, restoring the backup skips that entry with a warning. Save NIC ITR again for the affected adapter. See [NIC profile compatibility](./06-network.md#nic-itr).

## Controller lifecycle during register access

Keep the target PCI controller present and running during register access, including startup-profile application. Unplugging a mouse or keyboard while its USB controller remains present is a different event.

DTIMOD checks addresses against the controller's allocated MMIO ranges, but does not synchronize access with Windows PnP and power transitions. Resources can change during [surprise removal](https://learn.microsoft.com/en-us/windows-hardware/drivers/kernel/handling-an-irp-mn-surprise-removal-request) or [stop/rebalance](https://learn.microsoft.com/en-us/windows-hardware/drivers/kernel/handling-an-irp-mn-stop-device-request--windows-2000-and-later-). Concurrent controller removal or power transitions are not covered by that range check.

## Register access and USB mapping

Register reads and writes identify the PCI device, stay within its allocated MMIO ranges and are restricted to xHCI (`0C0330`) and Ethernet (`0200`) classes. Read-only PCI configuration queries also support IRQ inspection on other device classes.

For the exact USB map, a separate xHCI query returns root-port, USB-address and interrupter tuples. Its internal walk reads controller contexts and transfer-ring data from system memory through [MmCopyMemory](https://learn.microsoft.com/en-us/windows-hardware/drivers/ddi/ntddk/nf-ntddk-mmcopymemory). The request accepts no caller-supplied RAM address and returns decoded topology. Further access boundaries are described in the [security policy](../../.github/SECURITY.md).

## Verification and recovery

Compare an EXE's SHA-256 with `SHA256SUMS.txt` from the same release package. The source commit in the release description identifies the corresponding code. Matching hashes establish package consistency, not hardware behavior.

Hardware changes need readback and device checks, including sleep and reconnect. Local QA does not perform these: it uses models and real write blocking. Preserve the original profile and a recovery backup before experimenting.

---

[Previous: Verifying results](./10-verification.md) · [Guide](./README.md)
