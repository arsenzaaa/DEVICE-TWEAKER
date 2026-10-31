# DTIMOD.sys, KDU and startup

## Why a driver

IMOD registers (xHCI controller) and NIC ITR registers (network controller) live in device memory (MMIO) and are not reachable from user mode. Reading PCI configuration space for **READ IRQ MODE** also has to happen in the kernel. So the program has its own driver, **DTIMOD.sys**.

Everything else — MSI, affinity, priority, RSS, ReservedCpuSets, Mouse Throttle, power saving — is written to the registry and **works without the driver**.

## Where the driver may write

DTIMOD.sys writes **only to memory that Windows assigned to the selected PCI device** (its BAR). The driver has no general physical-memory write. The range check is tested on a model of the driver together with a deliberate fault. Registers are read back after every write, and the interface shows the actual value.

## Why KDU

DTIMOD.sys has no public Microsoft signature. It is loaded by [KDU](../../src/IMOD/KDU/README.md), a loader that uses a signed vulnerable driver to load an unsigned one. As a result:

- **antivirus software** flags such loaders as HackTool;
- **HVCI** (memory integrity) and the **Microsoft Vulnerable Driver Blocklist** block the loading;
- **FACEIT** and **Vanguard** may conflict with a loaded driver.

Before enabling USB IMOD the program shows the state of HVCI, the Blocklist, Microsoft Defender and the FACEIT / Vanguard services — the **SYSTEM CHECK** button. The check loads nothing.

The DTIMOD sources and the pinned KDU 1.4.5.2512 are in the repository, binaries are checked by SHA-256, and QA does not start until the built driver matches its sources.

Once loaded, DTIMOD stays in memory until Windows restarts. If an older driver version is loaded, the program asks for a restart instead of writing through it.

## Startup without the EXE

Register values last until a restart, sleep or controller reset. To reapply a saved profile the program registers a Task Scheduler task:

```text
Task        \DEVICE TWEAKER\IMOD   runs as SYSTEM
Triggers    sign-in, resume from sleep, USB / HID device connected
Folder      %ProgramData%\DEVICE TWEAKER\Startup   (only administrators and SYSTEM may modify it)
  ApplyIMOD.ps1                 readable PowerShell script
  DeviceTweakerImodRuntime.cs   register code shared with the program
  profile.json                  saved values and SHA-256 of every file
  Loader\                       KDU to load DTIMOD.sys
  logs\                         logs for 14 days
```

Before each run the script checks the hashes of the runtime, driver and loader against `profile.json`; a mismatch stops the run. The program and the script share one load mutex and never load the driver at the same time. The EXE is not needed for startup: the program can be deleted and the task remains. **DELETE** in the IMOD block removes it.

## Verifying the files

SHA-256 hashes of both EXEs are in `SHA256SUMS.txt` on the release page, of the sources in `SOURCE-SHA256SUMS.txt`:

```powershell
(Get-FileHash .\DEVICE.TWEAKER.SelfContained.exe -Algorithm SHA256).Hash
Select-String -Path .\SHA256SUMS.txt -Pattern 'DEVICE.TWEAKER.SelfContained.exe'
```

**Conclusion:** the driver is only needed for IMOD, NIC ITR and **READ IRQ MODE**, writes only to the selected device's memory, and its sources are open. If HVCI, the Blocklist or an anti-cheat matters more, everything else works without it.

---

[← Chapter 10](./10-verification.en.md) · [Contents](./README.en.md)
