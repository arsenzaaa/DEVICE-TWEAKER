# Security policy

## Supported version

Security fixes are applied to the newest `0.0.4` preview only. Earlier previews are unsupported.

## Reporting a vulnerability

Please use [GitHub private vulnerability reporting](https://github.com/arsenzaaa/DEVICE-TWEAKER/security/advisories/new). Do not disclose a driver, privilege-boundary, arbitrary MMIO, path-validation, or restore vulnerability in a public issue before it is triaged.

Include the affected version, Windows build, reproduction steps, impact, and the smallest safe proof of concept. Do not include private keys, certificates, personal data, or destructive payloads.

Publication and credit will be coordinated after a fix is available.

## Driver trust boundary

`DTIMOD.sys` runs in kernel mode and the current artifact is test-signed, not publicly trusted. HVCI, Microsoft Vulnerable Driver Blocklist, endpoint security, or anti-cheat may block it. The application reports a loading failure and may offer a user-confirmed blocklist change.
