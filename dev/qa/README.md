# Developer checks

[Development guide](../../docs/DEVELOPMENT.md) · [Русский](README.ru.md)

Read-only hardware models and GUI checks for DEVICE TWEAKER. A clean checkout contains everything needed to reproduce them; generated captures, builds and personal data stay outside Git.

## Choose a check

For a GUI change, run one scenario, inspect its images and rerun with `-CompareWith` after editing:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\dev\qa\Invoke-DevSession.ps1 -Scenario .\dev\qa\Scenarios\device-flow.json
```

For documentation screenshots, use the gallery catalog:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\dev\qa\Export-Gallery.ps1
# One view, with both languages:
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\dev\qa\Export-Gallery.ps1 -Only gallery-usb-xhci
```

The gallery command produces an `index.html`, native PNGs and `gallery.json` under `local/gallery`. It leaves `assets/screenshots` unchanged until the images have been reviewed.

For application logic, select the relevant tests under [tests/DeviceTweaker.Tests](tests/DeviceTweaker.Tests). For a cross-cutting integration change, use the full suite below. Do not run it after each text edit or small visual adjustment.

## Hidden QA

```powershell
# Fast gate (Release build, xUnit tests, driver/KDU provenance; ~15-20s):
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\dev\qa\Invoke-Qa.ps1 -Mode Fast

# Full gate (hidden-desktop GUI smoke, workbench scenarios, 96/120/144 DPI layout, GDI/handle stability):
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\dev\qa\Invoke-Qa.ps1 -Mode Full

# Add -WithNative when driver C sources changed:
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\dev\qa\Invoke-Qa.ps1 -Mode Full -WithNative
```

See [coverage and reports](Checks/README.md) for the exact scope and requirements.

## Other entry points

- `Open-TestPanel.ps1`: manually open the read-only development panel. This is an interactive command; autonomous work uses the hidden workbench.
- `Review-DevSession.ps1`: record a visual review tied to the inspected PNG hashes.
- `Measure-IdleResources.ps1`: measure an exact EXE's CPU time and memory while idle on the hidden desktop.
- `Test-ReleasePackage.ps1`: check release EXEs, embedded dependencies and package identity.

Arguments, scenarios and examples are in the [workbench guide](WORKBENCH.md). Release preparation is described in [RELEASE.md](../../docs/RELEASE.md).

## Directory layout

- `Scenarios`: repeatable UI actions and assertions; `Scenarios/Gallery` contains documentation shots and their catalog.
- `tests`: isolated application and workbench tests with synthetic fixtures.
- `Panel`: the development-only controls, device presets and sandbox backends.
- `Workbench`: scenario execution evidence, comparison and review logic.
- `Inspect`: the console inspector and control-rendering checks.
- `Checks`: checks called by the integration runner, plus focused diagnostics.
- `Support`: shared isolation, input snapshots, build cache and evidence helpers. These are libraries, not standalone commands.

## Boundaries and results

GUI automation uses an invisible Win32 desktop and immutable read-only mode. The runner serializes workloads, limits memory and stops when host resources are insufficient. Preparing the manual test panel uses the same bounded worker; only its explicitly requested interactive launch opens a visible window. It never changes Windows settings or loads DTIMOD. Normal application builds exclude the development panel and workbench.

Models cover synthetic hardware, operation sequences and failures. They do not establish physical driver compatibility, MMIO behavior, actual monitor DPI or game performance. A matching PNG establishes repeatability; visual quality still requires inspection.

## Recovering a development checkout

Install the .NET SDK selected by `global.json` on Windows, then run `Invoke-DevSession.ps1`. Previous `local` contents, user backups, private certificates and prebuilt application EXEs are not required. Native model checks additionally require C++ Build Tools and the Windows SDK. Keep QA source and fixtures in Git; keep generated results and private data out of it.
