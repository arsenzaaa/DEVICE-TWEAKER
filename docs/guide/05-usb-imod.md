# USB controllers and IMOD

[Documentation](../README.md) / [Guide](./README.md) / 05 · [Русский](./05-usb-imod.ru.md)

## IMOD belongs to an interrupter

USB events are handled by the host controller. An xHCI interrupter has an Event Ring and an IMOD register. Several devices can share an interrupter, so separate role values do not necessarily provide independent control of each physical device.

IMODI specifies a minimum inter-interrupt interval in 250 ns units. `0x0` disables this moderation, `0xC8` represents 50 microseconds, and `0xFA0` represents 1 millisecond. Interrupt-generation conditions are defined in section 5.5.2.2 of the [xHCI specification](https://www.intel.com/content/dam/www/public/us/en/documents/technical-specifications/extensible-host-controler-interface-usb-xhci.pdf). The interval is not a fixed addition to every USB report's latency.

`0xC8` is DEVICE TWEAKER's baseline interval, not an invariant default for every controller and Windows version.

## Card modes

| Mode | Configuration |
| :--- | :--- |
| XHCI | A common controller value. |
| Devices | Recognized role values using the current mapping. |
| Interrupters | Values by interrupter index. |

Device mappings can change after reconnect, sleep or reboot. The application rereads available xHCI information and shows mapping and diagnostics. The current value tooltip includes the last read time. With incomplete information, inspect the applied common value and limitation rather than assuming an old index remains exact.

**CHECK** reads the registers. **SET** applies values and saves the startup profile. **DELETE** removes the shared IMOD/NIC ITR startup profile and restores network-driver moderation. It resets USB controls to `0xC8`; live writes require an already loaded DTIMOD and are skipped while the configuration contains only saved NIC ITR entries. Restart Windows to complete the reset. Deletion does not create a separate automatic backup.

## AUTO policy

For detected roles, AUTO proposes `0x0` for the mouse, `0xC8` for keyboard/gamepad, and `0xFA0` for USB audio/microphone. Check audio stability and input responsiveness before keeping these values.

Less moderation can increase interrupt rate and CPU load; more moderation changes processing batches. Check input, audio, lost events and load on your controller. When comparing ports on different controllers, record the device mapping for each run.

## Startup and recovery

Registers can be reinitialized by reboot, sleep or device reset. A PowerShell task reapplies saved settings; the EXE need not stay open. See [chapter 11](./11-driver.md).

Readback confirms the value read from a register. It does not measure input latency or replace physical sleep, resume and reconnect checks.

---

[Previous: Graphics card](./04-gpu.md) · [Guide](./README.md) · [Next: Networking, RSS and NIC ITR](./06-network.md)
