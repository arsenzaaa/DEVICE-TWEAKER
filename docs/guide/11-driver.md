# DTIMOD.sys, KDU and startup

[Documentation](../README.md) / [Guide](./README.md) / 11 · [Русский](./11-driver.ru.md)

## Operations that need DTIMOD

DTIMOD.sys provides hardware access to supported USB IMOD/NIC ITR registers and additional PCI information. Writes are restricted to checked MMIO ranges of the selected device; it exposes no general arbitrary physical-memory write interface.

MSI, affinity, RSS, ReservedCpuSets, Mouse Throttle and available power settings do not require DTIMOD. Some IRQ Mode information comes from Windows resources without the driver; an available driver supplies additional evidence. Incomplete mode detection is distinct from a failed setting write.

## Loading through KDU

DTIMOD has no public Microsoft signature for ordinary loading. The project uses a pinned [KDU](../../src/IMOD/KDU/README.md) version that relies on a vulnerable signed driver. This is a highly privileged hardware tool; antivirus, HVCI, Vulnerable Driver Blocklist and anti-cheat software can prevent it from working.

**SYSTEM CHECK** displays available protection/service information without loading the driver. Failure to find a service is not an anti-cheat compatibility guarantee. Do not disable protections merely to obtain a green status; other features remain available without the loader.

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

The runtime is shared with Windows PowerShell 5.1. File hashes are checked before application; a shared mutex prevents simultaneous loading by the application and startup script. Administrator/SYSTEM permissions protect the directory. Startup logs retain the latest 14 files.

The task runs independently of the EXE. Profile read or parse errors stop changes to the profile. **DELETE** in the IMOD block removes the shared startup configuration for USB IMOD and NIC ITR, including the scheduled task.

## Verification and recovery

Compare an EXE's SHA-256 with `SHA256SUMS.txt` from the same release package. `SOURCE-SHA256SUMS.txt` identifies the corresponding sources. Matching hashes establish package consistency, not hardware behavior.

Hardware changes need readback and device checks, including sleep and reconnect. Local QA does not perform these: it uses models and real write blocking. Preserve the original profile and a recovery backup before experimenting.

---

[Previous: Verifying results](./10-verification.md) · [Guide](./README.md)
