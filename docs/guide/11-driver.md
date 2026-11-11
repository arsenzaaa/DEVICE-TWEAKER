# DTIMOD.sys, KDU and startup

[Documentation](../README.md) / [Guide](./README.md) / 11 · [Русский](./11-driver.ru.md)

## Operations that need DTIMOD

DTIMOD.sys provides hardware access to supported USB IMOD/NIC ITR registers and additional PCI information. Writes are restricted to checked MMIO ranges of the selected device; it exposes no general arbitrary physical-memory write interface.

MSI, affinity, RSS, ReservedCpuSets, Mouse Throttle and available power settings do not require DTIMOD. Some IRQ Mode information comes from Windows resources without the driver; an available driver supplies additional evidence. Incomplete mode detection is distinct from a failed setting write.

## Controller removal during register access

The driver checks the target device's allocated MMIO range before a write. This check does not establish full synchronization with PCI controller removal or resource reassignment. Holding a reference keeps the device object alive; hardware resources can be released earlier during [surprise removal](https://learn.microsoft.com/en-us/windows-hardware/drivers/kernel/handling-an-irp-mn-surprise-removal-request).

Avoid disabling, restarting or removing the target PCI controller during an application or startup-profile register operation. Unplugging a mouse while its USB controller remains present is a different event. The native QA model can demonstrate access after modeled resource invalidation; it does not reproduce Windows PnP or establish that a kernel crash occurs.

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

## Verification and recovery

Compare an EXE's SHA-256 with `SHA256SUMS.txt` from the same release package. `SOURCE-SHA256SUMS.txt` identifies the corresponding sources. Matching hashes establish package consistency, not hardware behavior.

Hardware changes need readback and device checks, including sleep and reconnect. Local QA does not perform these: it uses models and real write blocking. Preserve the original profile and a recovery backup before experimenting.

---

[Previous: Verifying results](./10-verification.md) · [Guide](./README.md)
