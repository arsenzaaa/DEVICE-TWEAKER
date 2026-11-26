# USB controllers and IMOD

[Documentation](../README.md) / [Guide](./README.md) / 05 · [Русский](./05-usb-imod.ru.md)

## Controller and interrupters

An xHCI controller handles the devices attached to its USB ports. Its interrupters have their own Event Rings and IMOD registers. Several USB devices can share one interrupter, so the device map matters when assigning different values to mouse, keyboard or audio.

IMOD sets the minimum interval between interrupts from an interrupter, in steps of 250 ns. `0x0` disables this limit. DEVICE TWEAKER shows `0xC8` (50 µs) as its baseline. The [xHCI specification](https://www.intel.com/content/dam/www/public/us/en/documents/technical-specifications/extensible-host-controler-interface-usb-xhci.pdf), section 5.5.2.2, lists `0xFA0` (1 ms) as the register default. Read back the controller to see the value in use after the driver has started.

## Choose a mode

Use **XHCI** to apply one value across the controller's interrupters. Select **Devices** to enter values for detected roles such as mouse, keyboard, gamepad and audio. The program maps these roles to interrupters; when a reliable map is unavailable, it uses a common controller value. **Interrupters** exposes values by index for manual configuration.

When configured roles share an interrupter, it gets a single value. DEVICE TWEAKER uses the first configured role in this order: Mouse, Keyboard, Audio, Gamepad, Webcam. For example, if a mouse and USB audio share an interrupter, its value comes from the mouse.

Click **CHECK** to read the controller before editing it. Expand the details to see device assignments and register values for up to 64 interrupters. Reread the map after reconnecting a device, resuming from sleep or rebooting because assignments may have changed. The current-value tooltip shows the last read time.

The [gallery](../SCREENSHOTS.md#usb-controller) shows a complete USB controller card in XHCI mode.

## Apply and save

Enter the values for the selected mode and click **SET** in the controller card. The program backs up the settings and writes that controller's registers. It saves custom values for startup after a successful write. If the register write fails or is partial, it keeps the previous startup profile. Open **DETAILS** if the result reports an incomplete write or a failure to save the profile.

If a role field contains an invalid value, **SET** and **APPLY** stop and identify the field to correct.

The [startup task](./11-driver.md#saved-profile) reapplies the profile at sign-in, after resume and on relevant device events.

In the AUTO prompt, **SKIP** keeps the saved USB IMOD and NIC ITR profile. **APPLY** saves the new USB settings after they are applied successfully.

AUTO proposes `0x0` for a detected mouse, `0xC8` for a keyboard, gamepad or webcam, and `0xFA0` for USB audio. With less moderation, the controller can interrupt the CPU more often. Compare interrupt activity and input or audio behavior under the same workload.

## Delete or restore

**DELETE PROFILE** backs up and removes the shared USB IMOD and NIC ITR startup profile after confirmation. For detected adapters in that profile, it restores the driver's INF defaults for supported moderation settings still set to **Off**. If the driver declares no usable default, the setting stays as it is.

If restoring network moderation fails, the startup task is removed and the profile is kept for retry. Check the reported adapter in the log, then run **DELETE PROFILE** again.

Restart Windows to clear active register changes. The command also removes the staged `DTIMOD.sys` if it matches this build and is no longer in use. If DTIMOD was loaded, run the command again after restarting to remove the file.

---

[Previous: Graphics card](./04-gpu.md) · [Guide](./README.md) · [Next: Networking, RSS and NIC ITR](./06-network.md)
