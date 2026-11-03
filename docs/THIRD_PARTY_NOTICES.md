# Third-party components

[DEVICE TWEAKER](../README.md) / [Documentation](./README.md) · [Русский](./THIRD_PARTY_NOTICES.ru.md)

## KDU

The project uses [hfiref0x/KDU](https://github.com/hfiref0x/KDU) to load `DTIMOD.sys` for USB IMOD and NIC ITR register access.

The application uses:

- `src/IMOD/Loader/kdu.exe`
- `src/IMOD/Loader/drv64.dll`

Loader version: `1.4.5.2512`.

KDU is licensed under MIT. The full license is in `src/IMOD/KDU/LICENSE.txt`.

Checksums:

```text
B340DAD4DDBE8607F9FDDB79F679375B4FF5080FE1A7EDB6CE015F69D3A0CD4F  kdu.exe
D032D855A0FF0EF3E3AD6EC8DAFFB8649048F82D3F1EEB89BEE3652EE6F01F80  drv64.dll
```

## Other components

The application targets .NET 10 for Windows and uses Windows Forms, WMI and Win32 APIs.

PCI ID tables for the `CHIP 0` / `CHIP 1` indication are adapted from [MariusHeier/cpu-direct-usb](https://github.com/MariusHeier/cpu-direct-usb) and related sources. The indication estimates the controller class; it is not a latency measurement.
