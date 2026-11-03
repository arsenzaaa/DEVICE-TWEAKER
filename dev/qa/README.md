# Developer checks

[Development guide](../../docs/DEVELOPMENT.md) · [Русский](./README.ru.md)

## Opening the test panel manually

For the "reproduce, change, inspect and compare" cycle, use the [development workbench](WORKBENCH.md): `Invoke-DevSession.ps1`, scenarios, native captures, transition frames, test files and review records. `Invoke-Qa.ps1` runs the full suite.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\dev\qa\Open-TestPanel.ps1
```

The command builds `local/tools/TestPanel/DEVICE TWEAKER QA.exe` from a source copy and opens the panel. This build is always read-only. **Ctrl+Shift+T** and **Ctrl+Alt+Shift+T** reopen it; normal EXEs do not include the panel. `-BuildOnly` builds and checks the EXE without opening windows.

## Full hidden QA

The native model also requires Visual Studio C++ Build Tools and the Windows SDK. It compiles DTIMOD functions with substituted system calls; it does not run the SYS.

From the repository root:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\dev\qa\Invoke-Qa.ps1
```

Requires Windows, .NET 10 SDK and an available window station for GUI checks. By default, the entire run uses a separate invisible Windows desktop. Windows, dialogs and child processes remain there; there is no desktop switch, physical input or visible fallback. `-Monitor` sets capture geometry on that desktop. `-VisibleGui` is an explicit interactive-debugging option, not part of autonomous QA.

Before building, the runner copies sources, QA scripts, tests and embedded resources. Git supplies tracked and new nonignored files in `src` and `dev/qa`, so new files are included without editing a list. Builds and tests use this copy inside the run directory rather than shared project bin/obj. `inputs.json` records input SHA-256 values, checked before and after execution. The final report includes the hashes of the tested EXEs/DLLs.

## Two builds

The normal `DEVICE TWEAKER.exe` contains no test panel, Showcase, hardware presets or Scenario Lab.

The same project with `-p:DeveloperQa=true` includes `dev/qa/Panel`, shares application code and produces `DEVICE TWEAKER QA.exe` in `src/bin/DeveloperQa/Release/net10.0-windows/win-x64/`. This build always blocks settings writes, including manual launches. Leaving emulation does not enable writes.

```powershell
dotnet build ".\src\DeviceTweakerCS.csproj" -c Release -p:DeveloperQa=true -p:BuildImodDriver=false
```

The normal EXE supports `--qa-read-only`. This mode is immutable for the process: registry, WMI, power-plan, startup and real-backup changes, driver loading and driver access are blocked. Logs and reports go to the run directory. Each new write path needs its own blocking check.

## Coverage

- **Build:** normal and developer versions, with warnings treated as errors.
- **xUnit:** masks, parsing, localization, profiles and blocking of actual write methods.
- **WMI / RSS:** the shared query/write/readback sequence against a fake backend, unavailable capabilities, failures and readback mismatches. Actual RSS PowerShell code runs with substituted Get-NetAdapter / Get-NetAdapterRss functions. Wi-Fi skips, incomplete reads and mismatches have separate cases.
- **Application identity:** logs and crash reports include the assembly MVID and running EXE SHA-256. Package checks compare that hash with the file; the version remains 0.0.4.
- **IRQ Mode:** parsing PCI MSI/MSI-X Enable and allocated CM_RESOURCE_LIST data; one IRQ for either mode, malformed data, read failures and device mismatches. QA does not open DTIMOD; live PCI reads require manual verification.
- **Recorded resources:** six saved Windows 10 responses covering USB/GPU/NIC/NVMe/SATA. Both raw and translated lists, types and malformed tails are checked from JSON without driver access.
- **DTIMOD preparation:** concurrent branches share a lock. A separate test file uses real exclusive file handles.
- **DTIMOD artifact:** sources and SYS must match `IMOD/DTIMOD.build.json`. A changed source or binary stops QA; corrupted copies verify rejection. After driver changes, use `publish-variants.ps1 -BuildDriverOnly`.
- **DTIMOD lifecycle:** the GUI/script policy covers reads without loading, timeout without SCM fallback, protocol/build identity, competing operations and an abandoned mutex. Only fake backends and separate test mutexes are used.
- **Native model:** actual DriverEntry/dispatch/MMIO functions compile with fake kernel APIs. Checks cover absent unload, protocol response, sizes, alignment, failed mapping and unmapping. Writes must stay within device-assigned memory: boundaries, ports, address overflow and malformed resource lists are checked. Two deliberate defects, restored unload and an out-of-range write, must be caught by model assertions.
- **KDU:** the EXE/DLL pair, versions, hashes, build manifest and 283 inputs of the pinned upstream. Source, EXE and database substitutions run in separate copies.
- **Apply:** shared DeviceSettingsWriter MSI, limit, priority, binary mask and Windows Default operations against an in-memory registry.
- **DTIMOD rename:** shared migration policy against a fake SCM, covering loaded/missing/unrelated services and failed deletion.
- **Startup:** `profile.json` and app settings, migration of the older AppData script and saved backups without losing values, and Task Scheduler XML (SYSTEM, logon, resume, USB/HID connection). Actual `ApplyIMOD.ps1` runs in PowerShell 5.1 up to profile loading and compiles the shared runtime. No task is registered and no driver is loaded.
- **Restore:** shared BackupRestoreService version-1 JSON, initial snapshot, writes, readback and rollback.
- **Power plan:** shared restore component with separate AC/DC values, saved GUID, active/inactive schemes and modeled API failures.
- **IMOD / NIC ITR:** shared application/startup register-writing code: addresses, masks, ordering and modeled read/write/compare errors. xHCI layout and interrupter mapping use fake controller memory: endpoint TRB takes precedence over Slot Context, hubs are skipped and RTSOFF outside BAR is rejected.
- **CPU:** shared scheduler against synthetic SMT, P/E and CCD/CCX topologies, including GPU/input isolation.
- **Polling:** synthetic intervals, confirmation and rate changes; low keyboard activity must not replace the endpoint rate.
- **Reports:** concurrent JSON reads/updates and exclusion of the test panel from normal builds.
- **Sensitivity:** nineteen corrupted C# copies cover the I225 address, NIC readback, IMOD count, backup readback, GPU pair selection, premature polling downgrade, wrong AC/DC, activation of an inactive scheme, removing read-only through the environment, using MSI-X table size as Enable, ignoring a one-pixel difference, disposing fonts before controls, failing the second Windows resource list, concurrent DTIMOD preparation without a lock, loading from background reads, SCM after KDU timeout, accepting an incompatible protocol, losing saved NIC ITR values and choosing an interrupter from Slot Context instead of TRB. The relevant assertions must fail; a build failure does not count.
- **GUI:** EN/RU, geometry, dialogs, navigation, test panel and virtual scenarios.
- **Workbench:** two scenarios run twice, with native captures, a transition, test JSON changes/restoration, a real fake-storage exception and exact checkpoint comparison. Design review is recorded separately.
- **Normal-build controls:** dropdown key routing, selection and popup closing, Ctrl+Z/Home/End in editors and RU→EN→RU with source strings preserved. Eight normal/hover/focus/disabled captures use no physical input.
- **Source inventory:** Roslyn lists active production C# files, methods and UI strings for manual inspection. This list alone does not verify all logic.
- **Hidden launch:** independent script/EXE rejection of an incorrect desktop, exact error propagation and descendant cleanup after timeout.
- **Layout:** six EN/RU × 100/125/150% scenarios with actual WinForms controls, rendered button captions and RESTORE dialog. These are developer-build layout models; Windows DPI settings and monitor transitions are not tested.
- **Logs:** session format/completion and rejection of unexpected ERROR/FATAL entries. Injected errors are accepted only in a successfully verified scenario; dialog samples are labelled separately.
- **Resources:** twenty refreshes in the normal read-only EXE, confirmation of every press and REFRESH.DONE, handle/GDI/USER settling and memory measurements.
- **Evidence:** recheck TRX, mandatory stages/test classes, nineteen C# regressions, two native regressions, six scaling scenarios, completed cycles and file hashes.

If the local published SELF-CONTAINED directory contains backups, QA reads them through the shared restore component against in-memory storage. Originals are unchanged. Without that directory, synthetic JSON is used; TRX identifies the cases.

Virtual APPLY/BACKUP/RESTORE actions are labelled as model checks. Actual hardware application, driver readback and live Raw Input measurements remain `SKIPPED`.

## Results

`src/tmp/qa/run_*/` contains `qa.json`, `inputs.json`, copied inputs, TRX, logs, mutation reports, GUI screenshots, `layout/layout.json`, `resources/resources.json`, `desktop-contracts/contracts.json`, `isolation/desktop.json` and `evidence.json`. `desktop.json` records the desktop name, active-desktop checks and descendant cleanup confirmation.

A missing result, skipped required tests, changed source snapshot, nonzero worker exit or unconfirmed cycle is not a successful check. `evidence.json` checks artifact completeness and consistency; result validity still depends on the integrity of the checks and reports. Screenshots require visual inspection: image/geometry checks cannot judge the entire design.

Layout `main.png` / `restore.png` captures use native PrintWindow within display geometry. Taller windows are explicitly marked `fullWindow=false` and show only the viewport. The invisible desktop is not displayed. Separate `*.layout.png` images contain full DrawToBitmap diagnostic rendering; native text and DWM frames can differ, so these are not presented as full visual captures. Large white regions fail native-capture validation.

`PASS` applies to the listed checks, not every hardware configuration. `-SkipGui` produces `PARTIAL`; a failed mandatory stage produces `FAIL` and a nonzero exit code. The script closes only its own processes and neither publishes nor runs Git commands.

Check an existing local EXE package separately:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\dev\qa\Test-ReleasePackage.ps1 -PackageDirectory ".\src\bin\ReleasePackages\v0.0.4"
```

This checks SHA-256, versions, both EXEs in probe mode, the embedded driver and absence of the test panel. It does not open the main window or apply settings.
