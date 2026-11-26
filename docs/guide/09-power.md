# Device power saving

[Documentation](../README.md) / [Guide](./README.md) / 09 · [Русский](./09-power.ru.md)

## Applying a power setting

Change **Power Saving** in the device card and click **APPLY**. The application backs up settings before writing. Disabling [USB selective suspend](https://learn.microsoft.com/en-us/windows-hardware/drivers/usbcon/usb-selective-suspend) keeps idle devices from entering this low-power state and can increase battery use.

AUTO disables power saving for the devices it configures.

## USB and wired networking

For USB, the application handles the controller and associated root hubs, available Device Parameters, PnPCapabilities and WMI. The power plan's global USB selective-suspend setting is applied separately from device settings.

For supported wired adapters, the application changes permission for Windows to power down the device. PnPCapabilities changes preserve unrelated flags. Wi-Fi is excluded from automatic changes.

## Confirming the outcome

The program uses WMI where the device exposes the relevant property. If a write or verification fails, open **DETAILS** in the result dialog to see which step failed. A live device operation can fail after the registry settings have already been saved.

After a change, check sustained use, idle, sleep and resume. Restore the original settings if the device disappears or stops responding after resume.

## Restoration

Restore applies values from the selected backup. USB selective suspend is restored through the Windows Power API for AC (plugged in) and DC (battery) in the power scheme recorded in the backup, then verified by reading it back. The current Windows scheme stays selected. If the running device does not accept a restored setting, the result reports that step separately even when registry restoration succeeded.

---

[Previous: Input and Polling Rate](./08-input.md) · [Guide](./README.md) · [Next: Verifying results](./10-verification.md)
