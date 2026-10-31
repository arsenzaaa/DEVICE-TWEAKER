<!-- tag: USB | summary: How an xHCI controller works: interrupters, Event Rings and the IMOD register. Why a mouse gets a different interrupter on every connection, what IMOD to give audio, and how to spread devices across controllers. -->

# USB controllers and IMOD

## How xHCI is built

A mouse, keyboard or headset does not raise interrupts itself: the **xHCI controller** it is attached to does. The path of an event:

```text
device -> endpoint -> Transfer Event -> interrupter -> Event Ring -> IMOD -> MSI / MSI-X vector -> CPU
```

- An **interrupter** is a part of the controller with its own Event Ring, its own IMOD register and its own interrupt vector. There can be up to 1024, usually 8 (as many as MSI vectors). The count is stored in bits 18..8 of `HCSPARAMS1`.
- Every connected device is a **device slot**. Its **Slot Context** records the interrupter assigned to the device.
- Interrupter registers live in the controller's Runtime area: the first IMOD is at `Runtime Base + 0x24`, the following ones every `0x20`.

Hence the key point: **IMOD belongs to an interrupter, not to a mouse.** On older controllers one interrupter may serve several devices, and then IMOD applies to all of them.

## What IMOD does

Interrupt moderation sets the minimum interval between an interrupter's interrupts, in 250 ns steps:

```text
0x0      moderation off
0xC8     50 µs   (Windows 8 / 10 / 11 default; Windows 7 used 1 ms)
0xFA0    1 ms
```

Data is not lost under moderation; it **stacks up**. With 1000 Hz polling (an event every 1000 µs) and a 2000 µs IMOD, two events arrive within one interval, and the next ISR handles both in order, first come first. ISRs and DPCs become rarer but longer.

## The interrupter changes

The system assigns an interrupter to a device **on every connection**, and after a restart it may be a different one (any match is luck). So you cannot write "mouse = interrupter 2" once: per-device settings must be recalculated every time.

DEVICE TWEAKER reads the device -> interrupter map from the xHCI contexts in memory on every apply: the interrupter from the TRB of the device's transfer ring, or from the Slot Context if the ring cannot be read. The read time is shown next to the values; after a device connects or disconnects the map is marked "devices changed", and **CHECK** reads it again.

## Which value to use

- **USB audio** — the highest IMOD that does not glitch the sound. Audio ISRs / DPCs / threads then work longer at a time but less often: it is better to stall rarely and longer than often and briefly, especially if a render thread lands on that core. The AUTO profile gives audio `0xFA0` (1 ms).
- **Mouse** — at 8 kHz a report arrives every 125 µs, at 1 kHz every 1000 µs, so the default 50 µs barely merges one device's reports. The AUTO profile still turns moderation off for the mouse (`0x0`) so its interrupter never waits. The IMOD interval is not added to every report: an event after idle and an event during a running countdown take different paths.
- **Keyboard and gamepad** — the AUTO profile keeps `0xC8`.
- **Network card** — its own moderation, chapter [06](./06-network.en.md).

## Spreading devices across controllers

One controller ISR / DPC may carry data from several devices, and then they are processed in turn: the second device waits for the first. Collisions are rare but happen at high polling rates. The ideal layout is three controllers:

```text
chipset controller (PCH)    audio       high IMOD, its own core
CPU controller / second     mouse       its own core
third controller            keyboard    its own core
```

At minimum, separate the mouse and keyboard from audio so audio can get a high IMOD without touching input. DEVICE TWEAKER shows which port belongs to which controller: every controller card lists the roles found on it (mouse / keyboard / gamepad / microphone / audio / camera).

**Do not install AMD / ASMedia drivers on USB controllers.** Their DPCs are shorter (0–1 µs versus 8–16 µs for the stock WDF01000), but the data then goes through the driver's system threads instead of straight to the DWM / CSRSS input threads, and those threads take 40–50 µs. The total is worse.

## Keeping it after a restart

The IMOD register lasts until a restart, sleep or controller reset. A saved profile is applied by the Task Scheduler task `\DEVICE TWEAKER\IMOD` at sign-in, after resume from sleep and when a USB / HID device connects, re-reading the interrupter map every time. See chapter [11](./11-driver.en.md).

## What DEVICE TWEAKER does

- **Devices** — values by role: mouse, keyboard, gamepad, USB audio. Without a reliable map one value is applied to the whole controller.
- **XHCI** — one value for every interrupter of the controller.
- **Interrupters** — values by index; the details show the first 64.
- **SET** writes the registers and saves non-default values to startup, **CHECK** re-reads the registers, **DELETE** removes startup and returns `0xC8`.
- Before enabling USB IMOD the program shows the state of HVCI, the Blocklist, Defender and the FACEIT / Vanguard services (**SYSTEM CHECK**).

**Conclusion:** IMOD is an interrupter setting, not a device setting. High IMOD for audio, no moderation for the mouse, devices spread across controllers, and the interrupter map re-read on every apply.
