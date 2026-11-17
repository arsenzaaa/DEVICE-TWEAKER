# Device power saving

[Documentation](../README.md) / [Guide](./README.md) / 09 · [Русский](./09-power.ru.md)

## Applying a power setting

Change **Power Saving** in the device card and click **APPLY**. The program saves a backup before applying the available device power settings. USB selective suspend can suspend individual idle ports. Disabling idle power saving can increase power use, particularly on battery. Availability also depends on the device class and driver. See [USB selective suspend](https://learn.microsoft.com/en-us/windows-hardware/drivers/usbcon/usb-selective-suspend) and [HID-specific behavior](https://learn.microsoft.com/en-us/windows-hardware/drivers/hid/selective-suspend-for-hid-over-usb-devices).

AUTO disables power saving for the devices it configures. Compare responsiveness, power use, sleep and resume on your system before keeping the change.

## USB and wired networking

For USB, the application handles the controller and associated root hubs, available Device Parameters, PnPCapabilities and WMI. The power plan's global USB selective-suspend setting is applied separately from device settings.

For supported wired adapters, the application changes permission for Windows to power down the device. PnPCapabilities is handled as a bit field, preserving unrelated flags; one number is not a universal value for every driver. Wi-Fi is excluded from automatic changes.

## Confirming the outcome

Read the operation summary after applying the change. If a WMI property is unavailable or a value cannot be written or verified, open DETAILS to identify the affected step. A partially applied result can mean that saved settings changed while a live device operation failed.

After a change, check sustained use, idle, sleep and resume. Restore the original settings if the device disappears or resumes poorly.

## Restoration

Restore uses values from the selected backup rather than current UI checkboxes. USB selective suspend is restored through the Windows Power API for AC (plugged in) and DC (battery) and read back; the current power scheme is not replaced with another. If the running device does not accept a restored setting, the result reports that step separately even when registry restoration succeeded.

---

[Previous: Input and Polling Rate](./08-input.md) · [Guide](./README.md) · [Next: Verifying results](./10-verification.md)
