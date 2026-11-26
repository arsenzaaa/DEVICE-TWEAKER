# QA coverage

[Commands](../README.md) · [Workbench](../WORKBENCH.md) · [Русский](README.ru.md)

The integration runner calls these checks in sequence. Use a focused check when its scope matches the change. C++ Build Tools and the Windows SDK are required for native models.

## Coverage

- **Build:** normal and developer versions, with warnings treated as errors.
- **xUnit:** masks, parsing, localization, profiles and blocking of actual write methods.
- **WMI / RSS:** the shared query/write/readback sequence against a fake backend, unavailable capabilities, failures and readback mismatches. Actual RSS PowerShell code runs with substituted Get-NetAdapter / Get-NetAdapterRss functions. Wi-Fi skips, incomplete reads and mismatches have separate cases.
- **Application identity:** logs and crash reports include the assembly MVID and running EXE SHA-256. Package checks compare that hash with the file.
- **IRQ Mode:** parsing PCI MSI/MSI-X Enable and allocated CM_RESOURCE_LIST data; one IRQ for either mode, malformed data, read failures and device mismatches. QA does not open DTIMOD; live PCI reads require manual verification.
- **Recorded resources:** six saved Windows 10 responses covering USB/GPU/NIC/NVMe/SATA. Both raw and translated lists, types and malformed tails are checked from JSON without driver access.
- **DTIMOD preparation:** concurrent branches share a lock. A separate test file uses real exclusive file handles.
- **DTIMOD artifact:** sources and SYS must match `IMOD/DTIMOD.build.json`. A changed source or binary stops QA; corrupted copies verify rejection. After driver changes, use `publish-variants.ps1 -BuildDriverOnly`.
- **DTIMOD lifecycle:** the GUI/script policy covers reads without loading, timeout without SCM fallback, protocol/build identity, competing operations and an abandoned mutex. Only fake backends and separate test mutexes are used.
- **Native model:** actual DriverEntry/dispatch/MMIO functions compile with fake kernel APIs. Checks cover unload-handler registration, protocol response, sizes, alignment, failed mapping and unmapping. Writes must stay within device-assigned memory: boundaries, ports, address overflow and malformed resource lists are checked. Three injected defects must be caught by model assertions: removing the unload-handler assignment, allowing a write past the device range, and bypassing the device-write range guard.
- **KDU:** the EXE/DLL pair, versions, hashes and inputs listed in the [upstream manifest](../../../src/IMOD/KDU/upstream.json). Source, EXE and database substitutions run in separate copies.
- **Apply:** shared DeviceSettingsWriter MSI, limit, priority, binary mask and Windows Default operations against an in-memory registry.
- **Driver identity:** application/startup names, pinned SYS payload, rejection of a loaded legacy device and immutable QA blocking. No service migration. The shared loading policy covers existing/KDU/service outcomes and uncertain KDU completion.
- **Startup:** `profile.json` and app settings, migration of the older AppData script and saved backups without losing values, and Task Scheduler XML (SYSTEM, logon, resume, USB/HID connection). Actual `ApplyIMOD.ps1` runs in PowerShell 5.1 up to profile loading and compiles the shared runtime. No task is registered and no driver is loaded.
- **Restore:** shared BackupRestoreService version-1 JSON, initial snapshot, writes, readback and rollback.
- **Power plan:** shared restore component with separate AC/DC values, saved GUID, active/inactive schemes and modeled API failures.
- **IMOD / NIC ITR:** shared application/startup register-writing code: addresses, masks, ordering and modeled read/write/compare errors. xHCI layout and interrupter mapping use fake controller memory: endpoint TRB takes precedence over Slot Context, hubs are skipped and RTSOFF outside BAR is rejected.
- **CPU:** shared scheduler against synthetic SMT, P/E and CCD/CCX topologies, including GPU/input isolation.
- **Polling:** synthetic intervals, confirmation and rate changes; low keyboard activity must not replace the endpoint rate.
- **Reports:** concurrent JSON reads/updates and exclusion of the test panel from normal builds.
- **Modes:** `Invoke-Qa.ps1 -Mode Fast` (default) runs `Release` build (`TreatWarningsAsErrors`), xUnit (`DeviceTweaker.Tests`) and driver/KDU artifact checks. `-Mode Full` additionally builds `DeveloperQa`/`Inspect` and runs the streamlined `Test-Workbench.ps1` scenarios (single pass per scenario; use `Invoke-DevSession.ps1 -CompareWith` for before/after visual diffs), `Smoke-SafeGui.ps1` (read-only main window & single-instance smoke on `DEVICE TWEAKER QA.exe` without `TEST_ADMIN`), `Test-UiControls.ps1`, `Test-GuiLayout.ps1`, `Test-GuiResourceStability.ps1` (production `DEVICE TWEAKER.exe`), and log quality/error audits. Add `-WithNative` to run `Test-NativeDriverModel.ps1`.
- **Cache retention:** `Build-Cache.ps1` retains only the newest 3 completed QA builds under `local/qa-builds`.
- **Evidence:** rechecks `unit.trx`, mandatory stages, test classes, artifact hashes, log integrity, layout cases and resource stability cycles.

Restore checks use synthetic JSON by default. To additionally inspect your own backup samples against in-memory storage, explicitly pass `-BackupSamples <directory>` to `Invoke-Qa.ps1`. Originals are unchanged. Real backups are not required, bundled or selected automatically; their paths may appear in local TRX reports.

Virtual APPLY/BACKUP/RESTORE actions are labelled as model checks. Actual hardware application, driver readback and live Raw Input measurements remain `SKIPPED`.

## Native lifecycle diagnostics

`Test-NativeDriverModel.ps1 -ProbeRemovalRace` adds two deterministic interleavings after the resource query: target removal and reassignment of the old address range to another model owner. Real C driver functions execute against fake kernel APIs; the second case maps a separate memory buffer belonging to that owner. No SYS is loaded and no hardware is accessed.

`removal-probe.json` reports `EXPOSURE_OBSERVED` when the model observes a write after invalidation, or `NOT_OBSERVED` otherwise. Neither result validates Windows PnP synchronization. This diagnostic is separate from the ordinary native regression `PASS`; it is not included in the default run. Retain the model assumptions and review them against the Windows device/resource lifetime contract.

`Test-NativeDriverModel.ps1 -OutputDirectory .\local\native-lifecycle -ProbeDeviceLifecycle` exercises 27 cases using the production dispatch, register access and device-write functions. Injected resource invalidation occurs after query, before mapping, after mapping or after access. Cases cover reads and writes, removal/stop/reassignment assumptions, failures, cancellation, changed ranges and cleanup. The fake APIs do not implement Windows notification ordering, power management or the PCI open/query path. A failed resource query is a tested input, not a guarantee that Windows rejects every query on a stopped device.

`lifecycle-probe.json` records each case and hashes of the source and diagnostic inputs. `UNCOORDINATED_ACCESS_OBSERVED` means that register access was attempted after injected invalidation; permissive fake mappings do not predict the outcome on absent hardware. Assertion success validates the diagnostic expectations, not a PnP safety fix. Negative timing controls inject a change after access; those earlier accesses must not be reported as stale. Run this diagnostic separately from `-ProbeRemovalRace` and `-InjectRegression`.

## Results

`local/qa-runs/run_*/` contains `qa.json`, `unit.trx`, logs and `evidence.json`. In `-Mode Full` it additionally stores `inputs.json`, copied inputs, GUI screenshots, `workbench/workbench.json`, `layout/layout.json`, `resources/resources.json` and `isolation/desktop.json`. `desktop.json` records the desktop name, active-desktop checks and descendant cleanup confirmation.

A missing result, skipped required tests, changed source snapshot, nonzero worker exit or unconfirmed cycle is not a successful check. `evidence.json` checks artifact completeness and consistency; result validity still depends on the integrity of the checks and reports. Screenshots require visual inspection: image/geometry checks cannot judge the entire design.

Layout `main.png` / `restore.png` captures use native PrintWindow within display geometry. Taller windows are explicitly marked `fullWindow=false` and show only the viewport. The invisible desktop is not displayed. Separate `*.layout.png` images contain full DrawToBitmap diagnostic rendering; native text and DWM frames can differ, so these are not presented as full visual captures. Large white regions fail native-capture validation.

`PASS` applies to the listed checks, not every hardware configuration. `-SkipGui` switches `-Mode Full` to `-Mode Fast`; a failed mandatory stage produces `FAIL` and a nonzero exit code. The script closes only its own processes and does not commit, push or publish.

Check an existing local EXE package separately:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\dev\qa\Test-ReleasePackage.ps1 -PackageDirectory ".\src\bin\ReleasePackages\v0.1.0"
```

This checks SHA-256, versions, both EXEs in probe mode, the embedded driver and absence of the test panel. It does not open the main window or apply settings.

The native model compiles the production device-write worker and writer as well as dispatch and range parsing. It exercises successful writes, out-of-range requests, cancellation, invalid width/version/IRQL, PCI/resource failures, malformed lists, map failures and reference/work-item cleanup. PCI lookup, resource delivery, the work queue and memory mapping are fake; kernel scheduling and physical devices are outside this model.

`Test-BuildUserData.ps1` checks clean-build preservation using temporary EXE-folder fixtures and a fake publisher. It covers a failure while moving the second folder, a failure at the next EXE flavor, publisher failure and successful publishing. It never starts the app, signs a driver or uses real backups/logs; CI runs it separately.

The workbench suite includes `refresh-reuse-lifecycle` (and `refresh-reuse` / `refresh-reuse-lifecycle-amd` remain available as standalone `Invoke-DevSession.ps1` scenarios). They exercise retained control identity, fresh settings and editor baselines, drafts/focus/viewport, virtual APPLY targeting, device removal/reconnection, role changes, discovery faults, topology changes during discovery/readback, and clearing stale RSS selection. These are synthetic read-only scenarios, not physical hotplug validation.

QA launchers allow one independent workload per Windows session. Preparation is serialized before snapshots/builds. Invoke-DevSession creates the bounded private job before cache inspection or compilation; the scenario and compiler descendants share that job. Nested probes remain inside their existing private job. Private workers and their descendants run below normal priority with a job commit limit of at most 3 GiB (also bounded by host RAM). Sustained low physical memory or commit headroom stops only the owned job. `isolation/resources.json` records the limit, peak job commit and minimum observed headroom during the run. This is a host safeguard, not a performance benchmark or a guarantee against every source of host load. Direct manual build commands are outside this launcher.

For smaller hardware-matrix runs, set the `audit-hardware-matrix` step's `device` field to comma-separated profile IDs, for example `system-13,system-14,cpu-00`. IDs come from the catalog indices recorded in matrix reports. Unknown, duplicate or empty IDs fail the run; omitting the filter retains the full matrix. Run groups sequentially in separate processes.
