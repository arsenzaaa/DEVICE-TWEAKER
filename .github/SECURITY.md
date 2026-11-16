# Security policy

[DEVICE TWEAKER](../README.md) / Security · [Documentation](../docs/README.md) · [Русский](./SECURITY.ru.md)

## Supported version

Security fixes target the latest release.

## Reporting a vulnerability

Please use [GitHub private vulnerability reporting](https://github.com/arsenzaaa/DEVICE-TWEAKER/security/advisories/new). Do not disclose a driver, privilege-boundary, arbitrary MMIO, path-validation, or restore vulnerability in a public issue before it is triaged.

Include the affected version, Windows build, reproduction steps, impact, and the smallest safe proof of concept. Do not include private keys, certificates, personal data, or destructive payloads.

Publication and credit will be coordinated after a fix is available.

## Driver trust boundary

`DTIMOD.sys` runs in kernel mode and the current artifact is test-signed, not publicly trusted. HVCI, Microsoft Vulnerable Driver Blocklist, endpoint security, or anti-cheat may block it. The application reports loading failures and does not offer to disable Windows protections. A successful load or a read-only system check does not establish anti-cheat compatibility.

- The driver's device object is open only to SYSTEM and Administrators.
- MMIO reads and writes name the target PCI device by instance ID and PDO. The driver checks the device identity, PCI class and allocated memory range. These operations are limited to xHCI and Ethernet devices; general physical-memory read/write requests and I/O-port access are rejected.
- Read-only PCI configuration queries remain available for IRQ inspection on other device classes. A separate xHCI query reads controller contexts internally and returns decoded USB topology, not raw system-memory contents.
- The resource check does not synchronize MMIO access with controller removal or power transitions. See the [driver lifecycle limitation](../docs/guide/11-driver.md#controller-lifecycle-during-register-access).
- `%ProgramData%\DEVICE TWEAKER\Startup` is writable only by SYSTEM and Administrators. The startup script checks the runtime SHA-256 from `profile.json` before compilation. When loading DTIMOD through KDU, it also checks `DTIMOD.sys`, `kdu.exe` and `drv64.dll`; these file checks are skipped when an existing driver device is available.

Reports about bypassing any of these checks are in scope.
