# Third-party notices

## KDU

DEVICE TWEAKER includes a source snapshot of KDU under `third_party/KDU`.
The runtime loader files used by DEVICE TWEAKER remain in `IMOD/Loader`.

Upstream project:
https://github.com/hfiref0x/KDU

Included upstream snapshot:
- `third_party/KDU/Source/Hamakaze` builds `kdu.exe`.
- `third_party/KDU/Source/Tanikaze` builds `drv64.dll`.
- `third_party/KDU/Source/Taigei` is kept because it is part of the original KDU solution.
- `third_party/KDU/Source/Shared` contains shared KDU code used by the projects above.
- `third_party/KDU/Source/Utils/GenAsIo2Unlock` is required by the `Hamakaze` post-build step.

Current runtime binaries bundled in `IMOD/Loader`:
- `kdu.exe` version `1.4.5.2512`, SHA256 `B340DAD4DDBE8607F9FDDB79F679375B4FF5080FE1A7EDB6CE015F69D3A0CD4F`
- `drv64.dll` version `1.4.5.2512`, SHA256 `D032D855A0FF0EF3E3AD6EC8DAFFB8649048F82D3F1EEB89BEE3652EE6F01F80`

The upstream README and changelog are preserved as:
- `third_party/KDU/UPSTREAM_README.md`
- `third_party/KDU/UPSTREAM_CHANGELOG.txt`

No ownership is claimed over KDU. The files are included so the loader dependency can be audited and rebuilt from source.
