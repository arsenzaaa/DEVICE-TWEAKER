# KDU source

This directory contains only the KDU parts needed to rebuild the loader files used by DEVICE TWEAKER.

Runtime files used by DEVICE TWEAKER:
- `IMOD/Loader/kdu.exe`
- `IMOD/Loader/drv64.dll`

Source mapping:
- `Source/Hamakaze/KDU.vcxproj` builds `kdu.exe`.
- `Source/Tanikaze/Tanikaze.vcxproj` builds `drv64.dll`.
- `Source/Utils/GenAsIo2Unlock/GenAsIo2Unlock.vcxproj` builds the helper required by `Hamakaze` post-build.
- `Source/Shared` contains common KDU headers and code.

Upstream project:
https://github.com/hfiref0x/KDU

Runtime files currently bundled in `IMOD/Loader`:
- `kdu.exe` version `1.4.5.2512`, SHA256 `B340DAD4DDBE8607F9FDDB79F679375B4FF5080FE1A7EDB6CE015F69D3A0CD4F`
- `drv64.dll` version `1.4.5.2512`, SHA256 `D032D855A0FF0EF3E3AD6EC8DAFFB8649048F82D3F1EEB89BEE3652EE6F01F80`

## Build

Requirements:
- Visual Studio 2019 or newer with C++ build tools.
- Windows SDK/WDK components required by KDU.
- The build script overrides upstream `PlatformToolset` with `v143` by default. Use `-PlatformToolset` if your Visual Studio installation requires another toolset.

Build from the DEVICE TWEAKER project root:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\IMOD\KDU\build-kdu.ps1
```

The script builds:
- `Source/Utils/GenAsIo2Unlock/output/x64/Release/GenAsIo2Unlock.exe`
- `Source/Tanikaze/output/x64/Release/drv64.dll`
- `Source/Hamakaze/output/x64/Release/kdu.exe`

To copy rebuilt binaries into `IMOD/Loader`:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\IMOD\KDU\build-kdu.ps1 -CopyToLoader
```

Do not replace the bundled runtime binaries without testing driver loading on a clean system.
