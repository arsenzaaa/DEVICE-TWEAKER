# Input: Polling Rate, Raw Input and Mouse Throttle

[Documentation](../README.md) / [Guide](./README.md) / 08 · [Русский](./08-input.ru.md)

## USB polling and Raw Input

USB transfer rate, delivered Raw Input event rate and a game's input-processing time are different quantities. The driver, message queue and application affect when received data is used. See [Raw Input overview](https://learn.microsoft.com/en-us/windows/win32/inputdev/about-raw-input).

DEVICE TWEAKER displays available polling information and measures mouse event frequency through Raw Input. Move the mouse during measurement. The result is neither an exact USB packet rate nor end-to-end screen latency. Sparse key press/release events cannot reliably establish keyboard polling rate.

## How applications read input

Windows provides individual-message reads and [GetRawInputBuffer](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getrawinputbuffer) for accumulated data. Buffering and dedicated input threads are implemented by each application. Measure their effect in the game or program being tested.

Higher report rates increase the amount of event processing. Compare available mouse modes in the actual game, including frame-time consistency and responsiveness. DEVICE TWEAKER does not change a game's input implementation or reflash the mouse.

## Mouse Throttle

On supported Windows 11 builds, the application can configure `RawMouseThrottleDuration`:

```text
HKCU\Control Panel\Mouse
  RawMouseThrottleDuration  REG_DWORD
```

The UI accepts custom intervals from 1 to 20 ms. Its 1 kHz, 125 Hz and 50 Hz labels use the `1000 / interval` calculation. These are requested values; actual event delivery depends on Windows and the listener.

The setting concerns background Raw Input, not USB polling rate. The application exposes it only for a supported OS. Its version check accounts for high-report-rate input improvements introduced in Windows 11 22H2 with [KB5027303](https://support.microsoft.com/en-us/servicing/os/windows-11/2023/06/june-27-2023-kb5027303-os-build-22621-1928-preview).

AUTO proposes 20 ms for background receivers. Verify event delivery in the affected applications before keeping the setting.

## Applying changes

Mouse Throttle edits remain pending until **APPLY**. Writing follows a backup; a failed write is not accepted as applied. Measuring event frequency and changing the setting are separate actions.

---

[Previous: CPU topology](./07-cpu-topology.md) · [Guide](./README.md) · [Next: Power saving](./09-power.md)
