# DEVICE TWEAKER source

[Project page](../README.en.md) · [How the settings work](./docs/SETTINGS.en.md) · [Русский](./README.md)

This directory contains the Windows Forms source. If you need a ready-to-run EXE, get it from the [current release](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.4).

## Build and test

You need Windows 10/11 x64, the .NET 8 SDK, and PowerShell 5.1 or newer. Run these commands from the repository root:

```powershell
dotnet restore ".\DEVICE TWEAKER\DeviceTweakerCS.csproj"
dotnet build ".\DEVICE TWEAKER\DeviceTweakerCS.csproj" -c Release -p:TreatWarningsAsErrors=true
dotnet test ".\DEVICE TWEAKER\tests\DeviceTweaker.Tests\DeviceTweaker.Tests.csproj" -c Release
```

The repository already contains `IMOD/DTIMOD.sys`. You only need Visual Studio C++ Build Tools and the Windows Driver Kit if you want to rebuild the driver.

To produce both EXEs and `SHA256SUMS.txt`:

```powershell
powershell -ExecutionPolicy Bypass -File ".\DEVICE TWEAKER\publish-variants.ps1" `
  -Flavor both -Configuration Release -SkipImodDriverBuild
```

See [BUILD.md](./docs/BUILD.md) and [RELEASE.md](./docs/RELEASE.md) for the build options and release process.

## Where to find the code

- `Affinity` — CPU topology and masks.
- `Devices` — device detection and classification.
- `Core` — applying settings, restore, IMOD, NIC ITR, and diagnostics.
- `GUI` and `Localization` — the interface and RU/EN strings.
- `Interop` — Windows API calls.
- `IMOD` — the driver and loader.
- `tests` — tests.

For bug reports or contributions, see [CONTRIBUTING.md](../.github/CONTRIBUTING.md).
