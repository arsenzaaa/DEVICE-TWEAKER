# Security policy

## Supported version

The current supported release is `0.0.4`. Earlier releases are unsupported.

## Reporting a vulnerability

Please use [GitHub private vulnerability reporting](https://github.com/arsenzaaa/DEVICE-TWEAKER/security/advisories/new). Do not disclose a driver, privilege-boundary, arbitrary MMIO, path-validation, or restore vulnerability in a public issue before it is triaged.

Include the affected version, Windows build, reproduction steps, impact, and the smallest safe proof of concept. Do not include private keys, certificates, personal data, or destructive payloads.

Publication and credit will be coordinated after a fix is available.

## Driver trust boundary

`DTIMOD.sys` runs in kernel mode and the current artifact is test-signed, not publicly trusted. HVCI, Microsoft Vulnerable Driver Blocklist, endpoint security, or anti-cheat may block it. The application reports a loading failure and may offer a user-confirmed blocklist change.

- The driver's device object is open only to SYSTEM and Administrators.
- Reads cover physical memory and PCI configuration space.
- A write names the target PCI device by instance ID and PDO. The driver checks that the device exists and that the address range lies inside memory Windows allocated to that device; anything else, including I/O ports, is refused. There is no general physical-memory write.
- `%ProgramData%\DEVICE TWEAKER\Startup` (startup script, runtime, profile, loader) is writable only by SYSTEM and Administrators, and the script verifies the SHA-256 of the runtime, driver, and loader recorded in `profile.json` before loading anything.

Reports about bypassing any of these checks are in scope.
