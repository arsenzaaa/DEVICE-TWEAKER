# KDU source snapshot

This directory contains the KDU source needed to rebuild the loader files used by DEVICE TWEAKER.

Runtime files used by DEVICE TWEAKER:
- `IMOD/Loader/kdu.exe`
- `IMOD/Loader/drv64.dll`

Source mapping:
- `Source/Hamakaze/KDU.vcxproj` builds `kdu.exe`.
- `Source/Tanikaze/Tanikaze.vcxproj` builds `drv64.dll`.
- `Source/Utils/GenAsIo2Unlock/GenAsIo2Unlock.vcxproj` builds the helper required by `Hamakaze` post-build.
- `Source/Shared` contains common KDU headers and code.
- `Source/Taigei` is kept because it is part of the upstream solution.

Upstream project:
https://github.com/hfiref0x/KDU

Upstream files preserved here:
- `UPSTREAM_README.md`
- `UPSTREAM_CHANGELOG.txt`
- `UPSTREAM_appveyor.yml`

## Build

Requirements:
- Visual Studio 2019 or newer with C++ build tools.
- Windows SDK/WDK components required by KDU.
- The build script overrides upstream `PlatformToolset` with `v143` by default. Use `-PlatformToolset` if your Visual Studio installation requires another toolset.

Build from the DEVICE TWEAKER project root:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\third_party\KDU\build-kdu.ps1
```

The script builds:
- `Source/Utils/GenAsIo2Unlock/output/x64/Release/GenAsIo2Unlock.exe`
- `Source/Tanikaze/output/x64/Release/drv64.dll`
- `Source/Hamakaze/output/x64/Release/kdu.exe`

To copy rebuilt binaries into `IMOD/Loader`:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\third_party\KDU\build-kdu.ps1 -CopyToLoader
```

Do not replace the bundled runtime binaries without testing driver loading on a clean system.
