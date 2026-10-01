# Building DEVICE TWEAKER

[Project page](../README.en.md) | [Русский](./README.md) | [Changelog](../CHANGELOG.en.md)

## Requirements

- Windows 10/11 x64
- .NET 8 SDK
- PowerShell 5.1 or later
- Visual Studio C++ Build Tools and the Windows Driver Kit to rebuild `DTIMOD.sys`

## Directories

- `Affinity` contains CPU topology, masks, and affinity settings.
- `Core` contains settings application, restore, IMOD, NIC ITR, and diagnostics.
- `Devices` handles device discovery and classification.
- `GUI` contains the Windows Forms interface.
- `Interop` contains Windows API calls.
- `Localization` contains Russian and English strings.
- `IMOD` contains the driver, loader, and related files.
- `tests` contains unit tests.

A normal device-list refresh does not load the Ring 0 driver. MMIO access only occurs during an explicit IMOD/ITR read or write and only for a recognized hardware profile.

## Build

Run these commands from the repository root.

```powershell
dotnet restore ".\DEVICE TWEAKER\DeviceTweakerCS.csproj"
dotnet build ".\DEVICE TWEAKER\DeviceTweakerCS.csproj" -c Release -p:TreatWarningsAsErrors=true
dotnet test ".\DEVICE TWEAKER\tests\DeviceTweaker.Tests\DeviceTweaker.Tests.csproj" -c Release
```

## Release files

```powershell
powershell -ExecutionPolicy Bypass -File ".\DEVICE TWEAKER\publish-variants.ps1" `
  -Flavor both -Configuration Release -SkipImodDriverBuild
```

The script creates compact and self-contained builds plus `SHA256SUMS.txt`.

## Driver

The repository includes the test-signed `IMOD/DTIMOD.sys`. Its expected hash is stored in `IMOD/DTIMOD.sys.sha256` and checked before use.

Do not commit private keys, PFX files, local certificates, or unrelated binaries. Do not copy MMIO offsets to a new Device ID without checking its documentation and registers.

Contribution and security notes are in [CONTRIBUTING.md](../.github/CONTRIBUTING.md) and [SECURITY.md](../.github/SECURITY.md).
