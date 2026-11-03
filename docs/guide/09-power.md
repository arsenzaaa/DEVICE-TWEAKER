# Device power saving

[Documentation](../README.md) / [Guide](./README.md) / 09 · [Русский](./09-power.ru.md)

## What changes

Power Saving manages available device power mechanisms. USB selective suspend can suspend individual idle ports. Microsoft recommends retaining it; drivers and device classes have additional constraints. See [USB selective suspend](https://learn.microsoft.com/en-us/windows-hardware/drivers/usbcon/usb-selective-suspend) and [HID-specific behavior](https://learn.microsoft.com/en-us/windows-hardware/drivers/hid/selective-suspend-for-hid-over-usb-devices).

AUTO disables power saving for the devices it configures. Compare responsiveness, power use, sleep and resume on your system before keeping the change.

## USB and wired networking

For USB, the application handles the controller and associated root hubs, available Device Parameters, PnPCapabilities and WMI. The power plan's global USB selective-suspend setting is applied separately from device settings.

For supported wired adapters, the application changes permission for Windows to power down the device. PnPCapabilities is handled as a bit field, preserving unrelated flags; one number is not a universal value for every driver. Wi-Fi is excluded from automatic changes.

## Confirming the outcome

A registry value, WMI state and actual device behavior are separate forms of evidence. Unsupported WMI properties, write errors and readback mismatches must be evaluated separately. A checkbox alone does not confirm a live power policy.

After a change, check sustained use, idle, sleep and resume. Restore the original settings if the device disappears or resumes poorly.

## Restoration

Restore uses values from the selected backup rather than current UI checkboxes. USB selective suspend is restored through the Windows Power API for AC/DC and read back; the current power scheme is not replaced with another. Failed live application remains a partial outcome even after registry restoration succeeds.

---

[Previous: Input and Polling Rate](./08-input.md) · [Guide](./README.md) · [Next: Verifying results](./10-verification.md)
