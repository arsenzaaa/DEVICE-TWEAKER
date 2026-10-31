<!-- tag: I/O | summary: The input path of a mouse, keyboard and controller on Windows 10 and 11, background-window throttling and RawMouseThrottleDuration, Raw Input Buffer and dedicated input threads, and what the Polling Rate measurement really measures. -->

# Input: Polling Rate, Raw Input and Mouse Throttle

## The input path

```text
Mouse, Windows 10       WDF01000 ISR -> DPC -> CSRSS (Win32k Raw Input Thread)    -> game thread
Mouse, Windows 11       WDF01000 ISR -> DPC -> DWM KERNEL SENSOR THREAD           -> game thread
Keyboard, 10 / 11       WDF01000 ISR -> DPC -> CSRSS (winsrvext.dll)              -> game thread
Controller, Windows 11  WDF01000 ISR -> DPC -> DWM MASTER INPUT THREAD            -> game thread
```

Windows 11 has no advantage over Windows 10 here: CSRSS already uses separate threads for the mouse and keyboard. Mouse input sometimes passes through two DWM threads; controller input may go through the DWM `WINDOWS.GAMING.INPUT` thread or straight from the DPC to the game.

The DWM and CSRSS input threads should also stay away from driver cores. That is manual thread tuning, outside what DEVICE TWEAKER does.

## How a game handles input

Quality depends on the developers. From worst to best:

```text
1. one thread runs physics and handles mouse + keyboard
2. one thread renders and handles mouse + keyboard
3. a dedicated mouse thread, keyboard mixed with physics (or the other way round)
4. one thread for both mouse and keyboard
5. separate threads for the mouse and the keyboard
```

When the input thread is mixed with other work and cannot keep up, games use:

- **RIB** (Raw Input Buffer) — accumulated messages are processed in a batch per tick. It does not add latency (that myth came from comparing RIB with IMOD) and removes FPS drops at high polling rates.
- **DIT** (Dedicated Input Thread) — a separate input thread, better than RIB for performance. Fortnite, for example, added `WindowsRawInputThread`.

Without RIB and without DIT, high polling (4–8 kHz) drops FPS at the physics-thread level. API Monitor or MXA show what a game uses.

## Background windows and Mouse Throttle

When several applications receive input, the data goes to all their input threads. Windows 11 limits the rate for **background** applications' input threads to 125 Hz; the DWM KERNEL SENSOR THREAD enforces it. The threshold is a registry value:

```text
HKCU\Control Panel\Mouse
  RawMouseThrottleDuration   REG_DWORD   interval in ms, rate = 1000 / value
```

`20` -> 50 Hz, the lowest rate; `8` -> 125 Hz, the default. The other values (`RawMouseThrottleEnabled`, `Leeway`, `Forced`) have no practical use. The setting does not change the mouse polling rate and does not affect the active window.

Controllers are not covered by this limit: their data does not pass through the KERNEL SENSOR THREAD and arrives at full rate in the DWM MASTER INPUT THREAD, `WINDOWS.GAMING.INPUT` and the GameInputSvc service. The more background threads receive it, the longer the DPC and the main input thread run.

## Polling Rate

DEVICE TWEAKER shows the mouse and keyboard Polling Rate from USB data and can **measure the Raw Input event rate**: move the mouse during the measurement, and after a confirmed measurement the measured rate is displayed.

- The Raw Input event rate is neither the USB polling rate nor the keyboard scan rate: message delivery and processing affect it.
- Keyboard events are key presses and releases, so a keyboard measurement does not determine its polling rate.

## What DEVICE TWEAKER does

- **Mouse Throttle** — presets from 50 Hz (`20`) to 1 kHz (`1`) on supported Windows 11 builds, with the current value shown.
- **AUTO-OPTIMIZATION** sets 50 Hz: background applications get input as rarely as possible.
- Every USB controller lists the input devices found on it and their rate.

**Conclusion:** background input at the minimum, high polling only if the game has RIB or DIT, and the DWM / CSRSS input threads on cores without drivers.
