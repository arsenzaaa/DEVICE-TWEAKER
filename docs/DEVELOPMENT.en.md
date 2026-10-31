# Developing DEVICE TWEAKER

[Project page](../README.en.md) · [Settings guide](./SETTINGS.en.md) · [Русский](./DEVELOPMENT.md)

## Requirements

- Windows 10 / 11 x64, [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0), Windows PowerShell 5.1.
- To rebuild `DTIMOD.sys` and the native model used by QA: Visual Studio with C++ Build Tools, the Windows SDK and the Windows Driver Kit. The built driver is already in the repository, so a regular build does not need them.

## Build and test

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

| Parameter | Effect |
| --- | --- |
| `-Flavor both` / `with-net` / `without-net` | both builds, only the self-contained one, or only the standard one |
| `-SkipImodDriverBuild` | use the existing `IMOD/DTIMOD.sys` |
| `-TrustImodDriverCert` | install the driver certificate on a test machine |
| `-NoClean` | keep intermediate files |

Output:

- `src/bin/Publish/DEVICE TWEAKER/DEVICE TWEAKER.exe` — standard build, requires the .NET 8 Desktop Runtime x64;
- `src/bin/Publish/DEVICE TWEAKER (SELF-CONTAINED)/…` — self-contained, compressed single file;
- `src/bin/ReleasePackages/v<version>/` — both EXEs and `SHA256SUMS.txt` for the release.

A clean build keeps the `Backups` and `logs` folders next to the EXEs.

## DTIMOD driver

The sources are `src/IMOD/Driver/imod_driver.c` and `.h`. After changing them:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\src\publish-variants.ps1 -BuildDriverOnly
```

This builds and signs `IMOD/DTIMOD.sys` and updates `DTIMOD.sys.sha256` and `DTIMOD.build.json` with the source hashes. Without the WDK MSBuild tasks it calls `cl` / `link` directly. QA stops until the binary matches the sources. Test-certificate scripts are in `src/Scripts`; private keys, `.pfx` and `.p12` files never go into the repository.

The KDU loader is pinned to version 1.4.5.2512; its sources and manifests are in [`src/IMOD/KDU`](../src/IMOD/KDU/README.md).

## IMOD / NIC ITR startup

`src/Startup/ApplyIMOD.ps1` and `src/Startup/DeviceTweakerImodRuntime.cs` are embedded in the EXE and copied to `%ProgramData%\DEVICE TWEAKER\Startup` when a profile is saved. The runtime is shared by the program and the script and is also compiled by Windows PowerShell 5.1, so it must stay within C# 5 syntax.

## QA

The full local QA — builds, tests, deliberate faults, the native driver model and UI checks on a hidden desktop — is described in [`dev/qa`](../dev/qa/README.md). It is read-only and never loads the driver.

## Site and artwork

The site [arsenzaaa.github.io/DEVICE-TWEAKER](https://arsenzaaa.github.io/DEVICE-TWEAKER/en/) is built from `docs/`, `CHANGELOG.md` / `CHANGELOG.en.md` and the landing page text in `dev/site/build.py`. The `Site` workflow publishes it to GitHub Pages after changes on `main`. Locally:

```powershell
pip install markdown==3.7
python dev/site/build.py --out _site
```

The mark and diagrams in `assets/branding` are drawn by `dev/branding/render.py`; text is converted to outlines, so the SVGs look the same on every system. `dev/branding/screens.py` adds the frame and shadow to screenshots captured by `dev/qa/Capture-Showcase.ps1`.

## Code map

| Folder | Contents |
| --- | --- |
| `src/Affinity` | CPU topology, affinity masks, ReservedCpuSets |
| `src/Devices` | device discovery and classification, RSS, power saving |
| `src/Core` | applying and restoring settings, IMOD, NIC ITR, startup, logs |
| `src/GUI`, `src/Localization` | user interface and RU / EN strings |
| `src/Interop` | Windows API calls |
| `src/Startup` | startup script and the shared register runtime |
| `src/IMOD` | the DTIMOD driver and the KDU loader |
| `dev/qa` | tests, test panel and QA |
| `dev/site`, `dev/branding` | site, mark, diagrams and screenshot frames |

Application logs go to the `logs` folder next to the EXE, one file per run. An unhandled error also writes a crash file and `last-crash.txt` there.

Releasing a version is described in [RELEASE.md](./RELEASE.md) (Russian); contribution rules are in [CONTRIBUTING](../.github/CONTRIBUTING.md).
