# Security policy

[DEVICE TWEAKER](../README.md) / Security · [Documentation](../docs/README.md) · [Русский](./SECURITY.ru.md)

## Supported version

Security fixes target the latest release.

## Reporting a vulnerability

Report vulnerabilities through [GitHub private vulnerability reporting](https://github.com/arsenzaaa/DEVICE-TWEAKER/security/advisories/new).

Include the affected version, Windows build, reproduction steps, impact and a minimal example with personal data and private keys removed.

Publication and credit will be coordinated after a fix is available.

## Driver access

`DTIMOD.sys` runs in kernel mode; the current artifact is signed with a test certificate. Loading requirements and compatibility are described in the [driver guide](../docs/guide/11-driver.md).

- The driver's device object is open only to SYSTEM and Administrators.
- MMIO reads and writes name the target PCI device by instance ID and PDO. The driver checks the device identity, PCI class and allocated memory range. These operations are limited to xHCI and Ethernet devices; general physical-memory read/write requests and I/O-port access are rejected.
- Read-only PCI configuration queries remain available for IRQ inspection on other device classes. A separate xHCI query reads controller contexts internally and returns decoded USB topology, not raw system-memory contents.
- The resource check does not synchronize MMIO access with controller removal or power transitions. See the [driver lifecycle limitation](../docs/guide/11-driver.md#controller-lifecycle-during-register-access).
- `%ProgramData%\DEVICE TWEAKER\Startup` is writable only by SYSTEM and Administrators. The startup script checks the runtime SHA-256 from `profile.json` before compilation. When loading DTIMOD through KDU, it also checks `DTIMOD.sys`, `kdu.exe` and `drv64.dll`; these file checks are skipped when an existing driver device is available.

Reports about bypassing any of these checks are in scope.
