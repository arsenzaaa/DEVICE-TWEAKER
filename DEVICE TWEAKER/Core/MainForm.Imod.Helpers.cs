namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private static bool IsUsbImodTarget(DeviceInfo device)
    {
        return device.Kind == DeviceKind.USB
            && device.UsbIsXhci
            && device.UsbHasDevices
            && HasUsbHidRole(device);
    }

    private static bool HasUsbHidRole(DeviceInfo device)
    {
        return HasUsbRole(device, "Keyboard")
            || HasUsbRole(device, "Mouse")
            || HasUsbRole(device, "Gamepad");
    }

    private static bool HasUsbRole(DeviceInfo device, string role)
    {
        if (device.Kind != DeviceKind.USB)
        {
            return false;
        }

        string roles = device.UsbRoles ?? string.Empty;
        if (string.IsNullOrWhiteSpace(roles))
        {
            return false;
        }

        foreach (string entry in roles.Split(',', StringSplitOptions.RemoveEmptyEntries))
        {
            if (string.Equals(entry.Trim(), role, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    private static bool ShouldShowImod(DeviceInfo device)
    {
        return IsUsbImodTarget(device);
    }
}
