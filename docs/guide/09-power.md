# Device power saving

[Documentation](../README.md) / [Guide](./README.md) / 09 · [Русский](./09-power.ru.md)

## What Power Saving controls

Toggle **Power Saving** in a device card and click **APPLY** (after an automatic backup) to control whether Windows may suspend or power down that device while idle. Disabling [USB selective suspend](https://learn.microsoft.com/en-us/windows-hardware/drivers/usbcon/usb-selective-suspend) prevents idle USB controllers and hubs from entering low-power states, trading slightly higher idle power draw (and battery consumption on laptops) for avoiding wake-up latency or dropouts on sensitive peripherals.

**AUTO-OPTIMIZATION** disables power saving on the USB controllers and wired network adapters it configures; Wi-Fi adapters are left untouched.

## USB controllers and wired Ethernet

- **USB:** DEVICE TWEAKER updates the controller and its associated root hubs across available `Device Parameters` (`SelectiveSuspendEnabled`, `EnhancedPowerManagementEnabled`), class-key `PnPCapabilities` (bit `0x08`, preserving unrelated capability flags), and live WMI `MSPower_DeviceEnable` (`root\WMI`) where exposed by the driver. The active power scheme's global USB selective suspend setting (`48e6b7a6-50f5-4782-a5d4-53bb8f07e226`) is managed separately from individual device keys.
- **Wired Ethernet:** toggles the Windows permission to power down the adapter via `PnPCapabilities` bit `0x08` (keeping unrelated flags intact) and `MSPower_DeviceEnable` when supported by the driver.

## Verifying live state and sleep transitions

Registry writes persist across restarts, while live WMI updates apply immediately to the running device. Because a driver can reject a live WMI update even after the registry keys succeed, open **DETAILS** in the result dialog whenever a warning appears to see the exact failing stage.

After changing power settings, verify normal operation across sustained load, idle periods, sleep and resume. If a USB device disappears or fails to wake after sleep, restore the original backup.

## Restoring backups

**RESTORE** replays the saved device values and restores global USB selective suspend through the Windows Power API separately for AC (plugged in) and DC (battery) on the power scheme recorded in the backup, then verifies the result by reading it back without switching your active scheme. If the running device refuses the live WMI update while registry restoration succeeds, that step is reported separately.

---

[Previous: Input and Polling Rate](./08-input.md) · [Guide](./README.md) · [Next: Verifying results](./10-verification.md)
