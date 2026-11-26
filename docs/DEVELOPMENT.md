# Developing DEVICE TWEAKER

[Documentation](./README.md) / Development · [Русский](./DEVELOPMENT.ru.md)

## Requirements

- Windows 10 / 11 x64, [.NET 10 SDK](https://dotnet.microsoft.com/download/dotnet/10.0), Windows PowerShell 5.1.
- To rebuild `DTIMOD.sys`: Visual Studio with C++ Build Tools, the Windows SDK and the Windows Driver Kit. The native QA model needs C++ Build Tools and Windows SDK, without WDK. The built driver is already in the repository, so a regular application build does not need them.

## Build and test

`global.json` requires SDK 10.0.404 or a newer stable 10.0 feature band. Install a matching SDK before building; CI reads the same file. `src/Startup/DeviceTweakerImodRuntime.cs` is also compiled by Windows PowerShell 5.1 and must keep C# 5 syntax.

From the repository root:

```powershell
dotnet build .\src\DeviceTweakerCS.csproj -c Release -p:TreatWarningsAsErrors=true
dotnet test .\dev\qa\tests\DeviceTweaker.Tests\DeviceTweaker.Tests.csproj -c Release
```

Warnings are errors, in CI as well. `src/.editorconfig` enables CA1304 / CA1305 / CA1310: string formatting and comparisons must not depend on the Windows locale.

## Release EXEs

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\src\build.ps1 -Flavor both -Configuration Release -SkipImodDriverBuild
```

- `-Flavor both` / `with-net` / `without-net` · both builds, only the self-contained one, or only the standard one
- `-SkipImodDriverBuild` · use the existing `IMOD/DTIMOD.sys`
- `-TrustImodDriverCert` · install the driver certificate on a test machine
- `-NoClean` · keep intermediate files

Output:

- `src/bin/Publish/DEVICE TWEAKER/DEVICE TWEAKER.exe` · requires .NET 10 Desktop Runtime x64
- `src/bin/Publish/DEVICE TWEAKER (SELF-CONTAINED)/DEVICE TWEAKER (SELF-CONTAINED).exe` · compressed EXE with .NET included
- `src/bin/ReleasePackages/v<version>/` · both release EXEs and `SHA256SUMS.txt`

A clean build keeps the `Backups` and `logs` folders next to the EXEs, including when preparation or publishing fails. Close running published copies before rebuilding them.

## DTIMOD driver

The sources are `src/IMOD/Driver/imod_driver.c` and `.h`. After changing them:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\src\publish-variants.ps1 -BuildDriverOnly
```

This builds and signs `IMOD/DTIMOD.sys` and updates `DTIMOD.sys.sha256` and `DTIMOD.build.json` with the source hashes. Without the WDK MSBuild tasks it calls `cl` / `link` directly. QA stops until the binary matches the sources. Test-certificate scripts are in `src/Scripts`; private keys, `.pfx` and `.p12` files never go into the repository.

WDK signing targets can create test certificates; `-TrustImodDriverCert` explicitly changes system trust stores. Keep driver signing in an isolated development environment, separate from GUI checks. Building does not require loading the driver.

The KDU loader is pinned to version 1.4.5.2512; its sources and manifests are in [`src/IMOD/KDU`](../src/IMOD/KDU/README.md).

## IMOD / NIC ITR startup

`src/Startup/ApplyIMOD.ps1` and `src/Startup/DeviceTweakerImodRuntime.cs` are embedded in the EXE and copied to `%ProgramData%\DEVICE TWEAKER\Startup` when a profile is saved. The runtime is shared by the program and the script and is also compiled by Windows PowerShell 5.1, so it must stay within C# 5 syntax.

## QA

See [Developer checks](../dev/qa/README.md) for unit tests, the native driver model and GUI scenarios. GUI automation uses a hidden desktop in read-only mode.

GitHub CI builds the app and developer workbench, checks PowerShell 5.1 scripts and runs unit tests. CodeQL default setup covers **C#** and **GitHub Actions**. The [native driver model](../dev/qa/Checks/README.md) covers selected C driver paths.

## Choosing checks

Build and test the behavior you changed. Run the full suite for changes that affect several components.

For GUI work, reproduce a scenario first, then run it after the change:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\dev\qa\Invoke-DevSession.ps1 -Scenario .\dev\qa\Scenarios\device-flow.json -OutputRoot .\local\dev-sessions
# Add -CompareWith <previous run directory> for comparison.
```

Inspect PNGs, differences and transitions. Complete review-input.json after inspection and use `dev/qa/Review-DevSession.ps1`. MATCH establishes pixel equality, not design approval. See [WORKBENCH.md](../dev/qa/WORKBENCH.md).

The workbench uses a hidden private desktop and immutable `--qa-read-only`. It must not manipulate existing user processes or the active desktop. Hardware and layout models do not replace real driver/readback, sleep, reboot or Windows DPI validation. Window-size constraints are retained in the report.

## RU/EN and UI state

Neutral English strings live in `src/Localization/Strings.resx`, Russian strings in `Strings.ru.resx`. UiText declares stable keys for core actions; UiResources loads values. Control IDs are independent of translated captions.

UiLanguage retains an adapter for dynamic diagnostic phrases. Update both cultures and check format parameters when adding text. Language/layout changes must not reread devices or lose pending/baseline state, focus or scroll position.

## Screenshots

`dev/qa/Export-Gallery.ps1` runs the selected scenarios from `dev/qa/Scenarios/Gallery/catalog.json` on the hidden desktop. The [public gallery](./SCREENSHOTS.md) contains a main-window example and complete USB and network cards. Additional scenarios are available for development checks.

Review the original PNGs before copying them into `assets/screenshots`, keeping English and Russian pairs consistent. Capture commands are in the [release guide](./RELEASE.md#maintain-the-project-page).

## Release automation

The `Release package` workflow builds an exact selected SHA, checks both EXEs and saves the package as an Actions artifact. It has read-only permissions and cannot create or replace tags or releases. Publish the accepted files separately after owner review. See [releasing](./RELEASE.md).

## Code map

Driver session ownership is in `src/Core/ImodDriverContext.cs`, with device/SCM interop in `ImodDriverNative.cs` and path policy in `ImodDriverPaths.cs`. The session receives its loader callback from the application; it has no dependency on the MainForm type. `PciInterruptBatch.cs` materializes per-device IRQ results while the session is open. MainForm retains progress, translation and control updates, including rejection of stale results after a topology change.

`DriverIdentity.cs` keeps device/service names and the loaded legacy-device preflight. `DriverRuntime.cs` contains loader selection and app/startup coordination; `ImodDriverContext` calls the selection policy exercised by unit tests. Existing-device readback does not invoke loading.

`ImodDriverFiles.cs` verifies and stages embedded hardware components. `KduDriverLoader.cs` owns the loader process, captured output and uncertain-outcome handling. `ImodReadbackService.cs` reads a complete indexed series of IMOD values; an incomplete series is not published as a device map. `ImodRoleMap.cs` assigns USB roles from captured topology, while `ImodMapText.cs` formats the result.

`StartupBundle.cs` manages startup files and task registration. `BackupStore.cs` handles snapshot files and retention; `BackupRegistry.cs` handles backed-up registry and power-plan values. MainForm coordinates these components with user choices and operation reports. The shared hardware runtime stays in `src/Startup` so the EXE and PowerShell script use the same register algorithms.

`src/Devices/NdisRssSettingsWriter.cs` owns RSS registry write and clear sequences. MainForm keeps the write authorization and operation coordination; the Windows key factory independently enforces the immutable QA write block. Tests use disposable fake registry keys. Per-value logs record completed writes and the exact failed step; active adapter readback remains separate from registry-write success.

- `src/Affinity` · CPU topology, affinity masks, ReservedCpuSets
- `src/Devices` · device discovery and classification, RSS, power saving
- `src/Core` · applying and restoring settings, IMOD, NIC ITR, startup, logs
- `src/GUI`, `src/Localization` · user interface and RU / EN strings
- `src/Interop` · Windows API calls
- `src/Startup` · startup script and the shared register runtime
- `src/IMOD` · the DTIMOD driver and the KDU loader
- `dev/qa` · tests, test panel and QA

Application logs go to the `logs` folder next to the EXE, one file per run. An unhandled error also writes a crash file and `last-crash.txt` there.

Releasing a version is described in [RELEASE.md](./RELEASE.md); contribution rules are in [CONTRIBUTING](../.github/CONTRIBUTING.md).
