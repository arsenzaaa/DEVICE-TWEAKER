<!-- tag: PWR | summary: USB selective suspend, letting Windows turn a device off, and the Device Parameters values: what actually works, where it is stored, and how the program verifies the write. -->

# Device power saving

## Why turn it off

Device power saving means transitions into a low-power state and back. Every wake-up delays the first event and adds interrupts and DPCs on the transitions. A gaming PC has no use for these savings on the mouse, keyboard, audio and network.

## Where it lives

**USB selective suspend** is a power-plan setting that lets the USB driver suspend idle ports. It exists separately for AC and battery.

```text
powercfg   SUB_USB   USB selective suspend setting   0 = disabled
```

**Allowing Windows to turn the device off** is the "Allow the computer to turn off this device to save power" checkbox in Device Manager. Behind it is the WMI class `MSPower_DeviceEnable`, which writes the right values itself. For a network card it is `PnPCapabilities`:

```text
HKLM\SYSTEM\CurrentControlSet\Control\Class\{4d36e972-e325-11ce-bfc1-08002be10318}\00NN
  PnPCapabilities   REG_DWORD   24 = Windows does not turn the adapter off, 16 = it does
```

The **Device Parameters** of USB controllers and hubs (`SelectiveSuspendEnabled`, `EnhancedPowerManagementEnabled` and similar) are essentially read only by USB drivers, and not every driver supports them.

## What DEVICE TWEAKER does

- **Power Saving** on a USB controller card controls power saving of the controller and its root USB hubs, including USB selective suspend in the power plan. On a wired adapter it controls whether Windows may turn the device off.
- Values are **read back through WMI** after writing. An unsupported feature, a query or write error and a value mismatch are reported separately, not as a generic failure.
- USB selective suspend is restored from a backup through the Windows Power API, separately for AC and battery, and verified by reading it back; the active power plan is not switched.

**Conclusion:** power saving is off for input devices, audio and network, and the result is confirmed by reading it back, not by a checkbox in the interface.
