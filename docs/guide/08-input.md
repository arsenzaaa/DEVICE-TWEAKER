# Input: Polling Rate, Raw Input and Mouse Throttle

[Documentation](../README.md) / [Guide](./README.md) / 08 · [Русский](./08-input.ru.md)

## USB polling and Raw Input

USB polling determines how often the host controller asks the device for data. The driver and message queue affect when that data reaches the game. The game handles input on its own schedule. See [Raw Input overview](https://learn.microsoft.com/en-us/windows/win32/inputdev/about-raw-input).

DEVICE TWEAKER displays available polling information and measures mouse event frequency through Raw Input. Move the mouse during measurement. The counter measures mouse events delivered to DEVICE TWEAKER. Keyboard press and release events are too sparse for the same measurement.

## How applications read input

An application can read Raw Input one message at a time or drain accumulated data in batches with [GetRawInputBuffer](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getrawinputbuffer). This batching runs in the receiving application. USB IMOD controls interrupt timing at the controller.

Higher report rates increase the amount of event processing. Compare available mouse modes in the actual game, including frame-time consistency and responsiveness.

## Mouse Throttle

On supported Windows 11 builds, the application can configure `RawMouseThrottleDuration`:

```text
HKCU\Control Panel\Mouse
  RawMouseThrottleDuration  REG_DWORD
```

The UI offers presets of 1, 2, 4, 5, 8, 10 and 20 ms. The label uses `1000 / interval` to show the requested rate in Hz. The setting applies to background listeners whose event rate Windows limits. Verify the delivered rate in the receiving application, especially with the 1 and 2 ms presets.

The switch is available starting with Windows 11 build 22621.1928, which included [high-report-rate mouse improvements](https://support.microsoft.com/en-us/servicing/os/windows-11/2023/06/june-27-2023-kb5027303-os-build-22621-1928-preview).

Turning the switch off removes the custom `RawMouseThrottleDuration` value and returns control to Windows defaults. AUTO proposes 20 ms for background receivers. Verify event delivery in the affected applications before keeping the setting.

## Applying changes

Choose an interval and click **APPLY** to save it after a backup. If writing fails, the edit remains pending.

---

[Previous: CPU topology](./07-cpu-topology.md) · [Guide](./README.md) · [Next: Power saving](./09-power.md)
