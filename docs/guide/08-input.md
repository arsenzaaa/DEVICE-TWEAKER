# Input: Polling Rate, Raw Input and Mouse Throttle

[Documentation](../README.md) / [Guide](./README.md) / 08 · [Русский](./08-input.md)

## From USB polling to the game loop

A USB mouse report crosses several layers before an application reacts to it:

```text
USB polling → xHCI ISR/DPC → Windows kernel input stack (win32k / Raw Input) → application message queue
```

USB polling determines how often the xHCI host controller requests reports from the device; the input stack then queues Raw Input events for listening windows, and each game processes those events on its own loop or dedicated input thread ([Raw Input overview](https://learn.microsoft.com/en-us/windows/win32/inputdev/about-raw-input)).

DEVICE TWEAKER displays available polling metadata and measures live mouse event frequency through a dedicated background Raw Input listener (`WM_INPUT`). Move the mouse continuously during measurement — the counter reflects events actually delivered to DEVICE TWEAKER. Keystrokes are too sparse to measure keyboard polling rates the same way.

## Per-message reads vs `GetRawInputBuffer`

At 4000–8000 Hz, reading `WM_INPUT` messages one by one on the main game thread creates heavy message-pump overhead and can noticeably drop frame rates. Applications avoid this either by using a dedicated input thread or by draining queued input packets in batches with [GetRawInputBuffer](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getrawinputbuffer). That batching happens inside the application as it drains already-arrived events — unlike USB IMOD, it does not hold back hardware interrupts at the controller.

Always test higher polling rates inside the actual game you play: check frame-time stability alongside responsiveness, as not every engine handles 4–8 kHz cleanly.

## Mouse Throttle

Starting with Windows 11 22H2 preview update KB5027303 (build `22621.1928`, broadly shipped in July 2023 KB5028185), Windows introduced [background high-report-rate mouse throttling](https://support.microsoft.com/en-us/servicing/os/windows-11/2023/06/june-27-2023-kb5027303-os-build-22621-1928-preview) to prevent background listeners (overlays, peripheral utilities, monitoring tools) from burning CPU cycles on 4–8 kHz mouse streams. The throttle interval is configured via:

```text
HKCU\Control Panel\Mouse
  RawMouseThrottleDuration  REG_DWORD
```

Key details of how Windows applies this value:

- **Background listeners only:** the foreground window (the active game) is not throttled, and gamepads/controllers do not go through this mouse-throttling path.
- **Default and valid range:** Windows defaults to **8 ms (≈125 Hz)** when no custom value is set, and `win32kbase.sys` accepts **1–20 ms (`0x1–0x14`)** (out-of-range values are ignored). DEVICE TWEAKER offers presets of `1, 2, 4, 5, 8, 10` and `20 ms`, showing the nominal `1000 / interval` Hz rate. Because actual delivery depends on Windows timer scheduling and how the receiving background app pumps messages, verify the measured rate in the target app — especially at `1` and `2 ms`.
- **Disabling and AUTO:** switching Mouse Throttle off deletes `RawMouseThrottleDuration` and restores Windows defaults. **AUTO-OPTIMIZATION** selects **20 ms (≈50 Hz)** — the strongest background throttling — to minimize background input overhead during gameplay. If you rely on background overlays, capture tools or macro software that need higher mouse report rates while unfocused, test them or pick a shorter interval.

Choose an interval and click **APPLY** to write it after a backup; if writing fails, the edit remains pending.

---

[Previous: CPU topology](./07-cpu-topology.md) · [Guide](./README.md) · [Next: Power saving](./09-power.md)
