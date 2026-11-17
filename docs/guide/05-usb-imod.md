# USB controllers and IMOD

[Documentation](../README.md) / [Guide](./README.md) / 05 · [Русский](./05-usb-imod.ru.md)

## Controller and interrupters

An xHCI controller handles the devices attached to its USB ports. Its interrupters have their own Event Rings and IMOD registers. Several USB devices can share one interrupter, so the device map matters when assigning different values to mouse, keyboard or audio.

The IMOD interval uses 250 ns steps: `0x0` disables moderation, `0xC8` is 50 microseconds, and `0xFA0` is 1 millisecond. `0xC8` is the application's baseline value; a controller or its Windows driver can use a different default. See section 5.5.2.2 of the [xHCI specification](https://www.intel.com/content/dam/www/public/us/en/documents/technical-specifications/extensible-host-controler-interface-usb-xhci.pdf).

## Choose a mode

Use **XHCI** to apply one value across the controller's interrupters. Select **Devices** to enter values for detected roles such as mouse, keyboard, gamepad and audio. The program maps these roles to interrupters; when a reliable map is unavailable, it uses a common controller value. **Interrupters** exposes values by index for manual configuration.

Click **CHECK** to read the controller before editing it. Expand the details to inspect device assignments and register values. After reconnect, sleep or reboot, read the map again: an earlier interrupter index may no longer belong to the same device. The current-value tooltip includes the last read time.

See the [USB examples](../SCREENSHOTS.md#usb-controller) for all three modes.

## Apply and save

Enter the values for the selected mode and click **SET**. The program creates a backup, writes the registers, reads them back and saves the configuration for startup. The main **APPLY** button is not required for this command.

The startup task reapplies the saved configuration after logon, resume and relevant device events. The EXE can be closed. Loading requirements and startup files are described in [chapter 11](./11-driver.md).

AUTO proposes `0x0` for a detected mouse, `0xC8` for keyboard/gamepad, and `0xFA0` for USB audio/microphone. Check input, audio stability and CPU load under your normal workload. Less moderation can increase the interrupt rate; the configured interval is not a fixed delay added to every USB report.

## Delete or restore

**DELETE** removes the shared USB IMOD/NIC ITR startup configuration and restores the network-driver moderation settings managed by NIC ITR **SAVE**. USB fields return to `0xC8`. If DTIMOD is already loaded, the program can also reset IMOD registers; USB writes are skipped when the configuration contains only saved NIC ITR entries. Restart Windows to complete the reset.

DELETE does not create an extra backup. To return to a particular saved configuration, select its backup through **RESTORE**. After restoration, check the register values and normal device operation.

---

[Previous: Graphics card](./04-gpu.md) · [Guide](./README.md) · [Next: Networking, RSS and NIC ITR](./06-network.md)
