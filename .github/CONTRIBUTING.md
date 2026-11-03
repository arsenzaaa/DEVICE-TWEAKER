# Contributing

[DEVICE TWEAKER](../README.md) / Contributing · [Русский](./CONTRIBUTING.ru.md)

Bug reports, hardware results, documentation fixes and focused pull requests are welcome.

[Report a bug](https://github.com/arsenzaaa/DEVICE-TWEAKER/issues/new?template=bug_report.yml) · [Hardware report](https://github.com/arsenzaaa/DEVICE-TWEAKER/issues/new?template=hardware_report.yml) · [Suggest a feature](https://github.com/arsenzaaa/DEVICE-TWEAKER/issues/new?template=feature_request.yml)

## Bug reports

Include the application version, Windows build, CPU, motherboard, BIOS, device Hardware ID and driver version. Describe the steps, expected and actual results, and whether you restarted Windows. Attach a relevant log excerpt.

Remove personal paths, usernames and serial numbers before posting. A routine bug report does not need a full system dump or registry export.

## Pull requests

Keep each PR focused on one problem. Explain which hardware or Windows build you tested and how to repeat the check. Claims about lower latency need a repeatable method and raw results.

For visual changes, include before/after screenshots. Keep Russian and English documentation in sync. Documentation changes do not require hardware operations.

For code changes, build and run checks for the affected behavior using the [development guide](../docs/DEVELOPMENT.md). For documentation, check links and rendering. For register, IMOD/ITR and restore changes, cover read/write failures. For UI changes, check both languages and Windows scaling at 100%, 125% and 150%; distinguish modeled layouts from tests on Windows.

## Guide changes

Guide chapters live in `docs/guide`, with English and Russian side by side. Support claims about Windows with traces or documentation, and claims about the program with its implementation.

Contributions are distributed under [GPL-3.0](../LICENSE).
