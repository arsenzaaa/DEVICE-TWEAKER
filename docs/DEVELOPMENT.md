# Developing DEVICE TWEAKER

[Documentation](./README.md) / Development · [Русский](./DEVELOPMENT.ru.md)

## Requirements

- Windows 10 / 11 x64, [.NET 10 SDK](https://dotnet.microsoft.com/download/dotnet/10.0), Windows PowerShell 5.1.
- To rebuild `DTIMOD.sys` and the native model used by QA: Visual Studio with C++ Build Tools, the Windows SDK and the Windows Driver Kit. The built driver is already in the repository, so a regular build does not need them.

## Build and test

`global.json` selects a stable 10.0 SDK and allows newer feature bands within that version. `src/Startup/DeviceTweakerImodRuntime.cs` is also compiled by Windows PowerShell 5.1 and must keep C# 5 syntax.

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

- `-Flavor both` / `with-net` / `without-net` - both builds, only the self-contained one, or only the standard one
- `-SkipImodDriverBuild` - use the existing `IMOD/DTIMOD.sys`
- `-TrustImodDriverCert` - install the driver certificate on a test machine
- `-NoClean` - keep intermediate files

Output:

- `src/bin/Publish/DEVICE TWEAKER/DEVICE TWEAKER.exe` - standard build, requires the .NET 10 Desktop Runtime x64;
- `src/bin/Publish/DEVICE TWEAKER (SELF-CONTAINED)/…` - self-contained, compressed single file;
- `src/bin/ReleasePackages/v<version>/` - both EXEs and `SHA256SUMS.txt` for the release.

A clean build keeps the `Backups` and `logs` folders next to the EXEs.

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

The full local QA - builds, tests, deliberate faults, the native driver model and UI checks on a hidden desktop - is described in [`dev/qa`](../dev/qa/README.md). It is read-only and never loads the driver.

## Choosing checks

Build and run checks that protect the changed behavior. Full QA belongs at integration points for cross-cutting changes; an expensive successful run need not be repeated after every text edit.

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

Capture the read-only developer build with `dev/qa/Capture-Showcase.ps1`. The [screenshot gallery](./SCREENSHOTS.md) uses full application windows, with no added frame or shadow. Label demonstration hardware and keep the displayed controls and readback consistent. Capture commands are in the [release guide](./RELEASE.md#maintain-the-project-page).

## Release automation

The release workflow defaults to packaging without publication. Replacement uses a previously accepted build artifact and a verified archive of the former release. Run its failure-injection checks locally without GitHub access:

```powershell
python -m unittest discover -s dev/release -p test_replace_release.py -v
```

Python 3.11+ is needed only for this automation. See [releasing](./RELEASE.md).

## Code map

- `src/Affinity` - CPU topology, affinity masks, ReservedCpuSets
- `src/Devices` - device discovery and classification, RSS, power saving
- `src/Core` - applying and restoring settings, IMOD, NIC ITR, startup, logs
- `src/GUI`, `src/Localization` - user interface and RU / EN strings
- `src/Interop` - Windows API calls
- `src/Startup` - startup script and the shared register runtime
- `src/IMOD` - the DTIMOD driver and the KDU loader
- `dev/qa` - tests, test panel and QA

Application logs go to the `logs` folder next to the EXE, one file per run. An unhandled error also writes a crash file and `last-crash.txt` there.

Releasing a version is described in [RELEASE.md](./RELEASE.md); contribution rules are in [CONTRIBUTING](../.github/CONTRIBUTING.md).
