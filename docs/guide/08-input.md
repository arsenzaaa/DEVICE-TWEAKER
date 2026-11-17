# Input: Polling Rate, Raw Input and Mouse Throttle

[Documentation](../README.md) / [Guide](./README.md) / 08 · [Русский](./08-input.ru.md)

## USB polling and Raw Input

USB transfer rate, delivered Raw Input event rate and a game's input-processing time are different quantities. The driver, message queue and application affect when received data is used. See [Raw Input overview](https://learn.microsoft.com/en-us/windows/win32/inputdev/about-raw-input).

DEVICE TWEAKER displays available polling information and measures mouse event frequency through Raw Input. Move the mouse during measurement. The counter reflects event delivery to the application. Mouse movement supplies a stream of samples; keyboard press/release events are too sparse for a reliable polling-rate estimate.

## How applications read input

Windows provides individual-message reads and [GetRawInputBuffer](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getrawinputbuffer) for accumulated data. Each application implements its own buffering and input threads.

Higher report rates increase the amount of event processing. Compare available mouse modes in the actual game, including frame-time consistency and responsiveness.

## Mouse Throttle

On supported Windows 11 builds, the application can configure `RawMouseThrottleDuration`:

```text
HKCU\Control Panel\Mouse
  RawMouseThrottleDuration  REG_DWORD
```

The UI accepts custom intervals from 1 to 20 ms. Its 1 kHz, 125 Hz and 50 Hz labels use the `1000 / interval` calculation. These are requested values; actual event delivery depends on Windows and the listener.

The setting concerns background Raw Input, not USB polling rate. The switch is enabled on supported Windows 11 builds and unavailable on Windows 10. Its version check accounts for high-report-rate input improvements introduced in Windows 11 22H2 with [KB5027303](https://support.microsoft.com/en-us/servicing/os/windows-11/2023/06/june-27-2023-kb5027303-os-build-22621-1928-preview).

Turning the switch off removes the custom `RawMouseThrottleDuration` value and returns control to Windows defaults. AUTO proposes 20 ms for background receivers. Verify event delivery in the affected applications before keeping the setting.

## Applying changes

Choose an interval and click **APPLY** to save it after a backup. If writing fails, the edit remains pending.

---

[Previous: CPU topology](./07-cpu-topology.md) · [Guide](./README.md) · [Next: Power saving](./09-power.md)
