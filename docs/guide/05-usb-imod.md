# USB controllers and IMOD

[Documentation](../README.md) / [Guide](./README.md) / 05 · [Русский](./05-usb-imod.ru.md)

## xHCI interrupters and how IMOD works

A USB peripheral does not own an MSI/MSI-X vector of its own — all vectors belong to the PCIe **xHCI host controller**. The delivery path runs through the controller's interrupters:

```text
USB device → Endpoint (Device Slot) → Transfer Event → Event Ring → Interrupter (IMOD) → MSI/MSI-X vector → CPU
```

Each interrupter owns an Event Ring and an **IMOD** (Interrupt Moderation) MMIO register (`Runtime Base + 0x20 + 0x20 × i + 0x04`). Multiple devices can share the same interrupter depending on the controller, driver and enumeration order, so knowing the active device-to-interrupter map matters whenever you want different moderation intervals for a mouse, keyboard or USB audio interface.

IMOD sets the minimum delay between consecutive interrupts from an interrupter in **250 ns** steps:

- **`0xFA0`** (`4000 × 250 ns = 1 ms`) — the hardware reset default defined in section 5.5.2.2 of the [xHCI specification](https://www.intel.com/content/dam/www/public/us/en/documents/technical-specifications/extensible-host-controler-interface-usb-xhci.pdf).
- **`0xC8`** (`200 × 250 ns = 50 µs`) — the operating value that Microsoft's `usbxhci.sys` typically programs at controller initialization on Windows 8/10/11 (used as DEVICE TWEAKER's baseline reference). Always read the controller back to see what your driver actually set.
- **`0x0`** — disables interrupt moderation completely.

The practical effect depends heavily on polling rate and device role. At 1000 Hz (1000 µs between polls), the Windows default of 50 µs is modest; at **4000–8000 Hz** (250–125 µs period), a 50 µs moderation window is 20–40% of the polling interval and can introduce measurable jitter, while a 1 ms interval (`0xFA0`) severely bottlenecks high-rate mice. Conversely, on **USB audio** interfaces or non-critical controllers, higher moderation (`0xFA0` or more, as long as audio stays clean) cuts interrupt frequency and CPU load. Lower moderation always trades reduced batching delay for higher interrupt rates — verify on your own hardware rather than assuming `0x0` helps every device.

## Choosing a mode

- **XHCI** applies a single IMOD value across all interrupters of the selected controller.
- **Devices** assigns per-role values (Mouse, Keyboard, Gamepad, Audio, Webcam) using the decoded xHCI slot-to-interrupter map; if no reliable map is available, it falls back to the controller-wide value. When multiple configured roles share one interrupter, DEVICE TWEAKER resolves conflicts in priority order: **Mouse → Keyboard → Audio → Gamepad → Webcam**.
- **Interrupters** exposes raw values by interrupter index for direct manual control.

Click **CHECK** before editing to read live registers and expand the details table (up to 64 interrupters). Because Windows and the controller can reassign Device Slots and interrupter indices after reconnecting peripherals, sleeping or rebooting, reread the map whenever USB topology changes. Hover over the current value to see the last read timestamp.

The [gallery](../SCREENSHOTS.md#usb-controller) shows a complete USB controller card in XHCI mode.

## Applying and saving

Enter the desired values and click **SET** in the controller card. DEVICE TWEAKER creates a backup, writes that controller's registers, expands the readback details table, and saves custom values to the startup profile once the write succeeds. If a write fails or is incomplete, the existing startup profile is preserved; open **DETAILS** to inspect the failed step. Invalid values in any per-role field stop **SET** and **APPLY** before writing.

Because hardware registers reset on reboot, sleep and controller reinitialization, the [startup task](./11-driver.md#saved-profile) reapplies saved profiles at sign-in, after resume and on relevant USB/HID events.

During **AUTO-OPTIMIZATION**, choosing **SKIP** leaves any existing USB IMOD and NIC ITR profile untouched; choosing **APPLY** programs `0x0` for a detected mouse, `0xC8` for a keyboard, gamepad or webcam, and `0xFA0` for USB audio, then saves the profile on success.

## Delete or restore

**DELETE PROFILE** backs up and removes the shared USB IMOD and NIC ITR startup profile after confirmation. For network adapters present in that profile, it also restores the driver's INF defaults for supported moderation settings that are still switched **Off** (leaving settings untouched when the INF defines no usable default).

If restoring network moderation fails, the startup task is removed while the profile file is kept so you can inspect the adapter in the log and retry **DELETE PROFILE**.

Restart Windows to clear active MMIO register overrides. The command also removes the staged `DTIMOD.sys` binary if it matches this build and is no longer loaded; if the driver is still in memory, run **DELETE PROFILE** once more after rebooting to delete the file.

---

[Previous: Graphics card](./04-gpu.md) · [Guide](./README.md) · [Next: Networking, RSS and NIC ITR](./06-network.md)
