# DTIMOD.sys, KDU and startup

[Documentation](../README.md) / [Guide](./README.md) / 11 · [Русский](./11-driver.ru.md)

## Operations that need DTIMOD

DTIMOD.sys provides hardware access to supported USB IMOD/NIC ITR registers and additional PCI information. Register reads and writes are bound to the exact PCI device, checked against its allocated MMIO ranges and restricted to xHCI (`0C0330`) and Ethernet (`0200`) classes. The former general physical-memory read and write requests are rejected. Read-only PCI configuration queries remain available for IRQ inspection on other device classes.

The exact USB map uses a separate xHCI query. The driver follows the controller's DCBAA, device contexts and endpoint/TRB data and returns only root-port, USB-address and interrupter tuples. These structures are in system memory, so the internal walk uses [MmCopyMemory](https://learn.microsoft.com/en-us/windows-hardware/drivers/ddi/ntddk/nf-ntddk-mmcopymemory), rather than mapping RAM as non-cached device memory. The query does not accept a caller-supplied RAM address or return raw context bytes. This remains privileged hardware access, not an isolation boundary against an administrator controlling the device.

MSI, affinity, RSS, ReservedCpuSets, Mouse Throttle and available power settings do not require DTIMOD. Some IRQ Mode information comes from Windows resources without the driver; an available driver supplies additional evidence. Incomplete mode detection is distinct from a failed setting write.

## Controller lifecycle during register access

DTIMOD checks each device-scoped register access against a snapshot of the target device's allocated MMIO ranges. It does not coordinate register reads or writes with that controller's PnP and power transitions. The range check does not guarantee that resources remain valid during controller stop, removal or reassignment. An object reference preserves the device object, not its hardware resources: Windows can release resources during [surprise removal](https://learn.microsoft.com/en-us/windows-hardware/drivers/kernel/handling-an-irp-mn-surprise-removal-request) or [stop/rebalance](https://learn.microsoft.com/en-us/windows-hardware/drivers/kernel/handling-an-irp-mn-stop-device-request--windows-2000-and-later-).

Avoid disabling, restarting or removing the target PCI controller during register operations, including saved-profile application outside the GUI. Unplugging a USB peripheral while its host controller remains present is a different event. Safe handling of concurrent controller removal or power transitions is not established for this driver.

## Loading through KDU

DTIMOD has no public Microsoft signature for ordinary loading. The project uses a pinned [KDU](../../src/IMOD/KDU/README.md) version that relies on a vulnerable signed driver. This is a highly privileged hardware tool; antivirus, HVCI, Vulnerable Driver Blocklist and anti-cheat software can prevent it from working.

DTIMOD can conflict with Easy Anti-Cheat, Riot Vanguard, FACEIT Anti-Cheat and other protection software. Successful driver loading does not confirm compatibility: a game may disconnect you while the driver is loaded. EAC disconnects with `Forbidden Tool` have been reproduced with DTIMOD loaded on a tested system. Compatibility with Vanguard and FACEIT has not been established.

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

The NIC layouts for Killer E3000 (`10EC:3000`, RTL8125 family) and RTL8118AS (`10EC:8162`, RTL8168 family) have been corrected. When a saved NIC ITR entry no longer matches the supported layout for its PCI ID, profile installation and backup restoration omit that entry. Restore reports the omission as a warning and records the old and current layout in the log. Other supported entries are retained; the backup file is unchanged. Save NIC ITR again for the affected adapter to create a current startup entry.

## Verification and recovery

Compare an EXE's SHA-256 with `SHA256SUMS.txt` from the same release package. The source commit in the release description identifies the corresponding code. Matching hashes establish package consistency, not hardware behavior.

Hardware changes need readback and device checks, including sleep and reconnect. Local QA does not perform these: it uses models and real write blocking. Preserve the original profile and a recovery backup before experimenting.

---

[Previous: Verifying results](./10-verification.md) · [Guide](./README.md)
