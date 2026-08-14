using Microsoft.Win32;
using System.Management;

namespace DeviceTweakerCS;

/// <summary>
/// Device Manager "Allow the computer to turn off this device to save power".
/// Persistence: PnPCapabilities bit 0x08 on the device class key.
/// Live checkbox: MSPower_DeviceEnable in root\WMI.
/// </summary>
internal static class DevicePowerPolicy
{
    public const string PnPCapabilitiesName = "PnPCapabilities";

    /// <summary>
    /// When set, Device Manager unchecks "Allow the computer to turn off this device to save power".
    /// Observed: 272 (0x110, checked) vs 280 (0x118, unchecked). Difference is this bit.
    /// </summary>
    public const int DoNotTurnOffDevice = 0x08;

    public const int NicDoNotTurnOffDevice = DoNotTurnOffDevice;

    public static bool IsNicTurnOffAllowed(int pnpCapabilities)
        => (pnpCapabilities & DoNotTurnOffDevice) == 0;

    public static int ApplyNicPnPCapabilities(string classKeyPath, bool allowTurnOff)
        => ApplyPnPCapabilities(classKeyPath, allowTurnOff);

    public static int ApplyPnPCapabilities(string classKeyPath, bool allowTurnOff)
    {
        if (string.IsNullOrWhiteSpace(classKeyPath))
        {
            throw new InvalidOperationException("Device class key is missing.");
        }

        using RegistryKey? key = Registry.LocalMachine.OpenSubKey(classKeyPath, writable: true);
        if (key is null)
        {
            throw new InvalidOperationException($"Could not open HKLM\\{classKeyPath} for writing.");
        }

        int current = key.GetValue(PnPCapabilitiesName) as int? ?? 0;
        int next = allowTurnOff
            ? current & ~DoNotTurnOffDevice
            : current | DoNotTurnOffDevice;
        key.SetValue(PnPCapabilitiesName, next, RegistryValueKind.DWord);
        return next;
    }

    public static int? TryReadNicPnPCapabilities(string? classKeyPath)
    {
        if (string.IsNullOrWhiteSpace(classKeyPath))
        {
            return null;
        }

        try
        {
            using RegistryKey? key = Registry.LocalMachine.OpenSubKey(classKeyPath, writable: false);
            return key?.GetValue(PnPCapabilitiesName) as int?;
        }
        catch
        {
            return null;
        }
    }

    public static string? TryGetClassKeyPath(string instanceId)
    {
        if (string.IsNullOrWhiteSpace(instanceId))
        {
            return null;
        }

        try
        {
            using RegistryKey? enumKey = Registry.LocalMachine.OpenSubKey(
                $@"SYSTEM\CurrentControlSet\Enum\{instanceId}",
                writable: false);
            string? driver = enumKey?.GetValue("Driver") as string;
            return string.IsNullOrWhiteSpace(driver)
                ? null
                : $@"SYSTEM\CurrentControlSet\Control\Class\{driver}";
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Best-effort class-key PnPCapabilities for any PnP instance (USB controller/hub, NIC, …).
    /// </summary>
    public static bool TryApplyInstancePnPCapabilities(string instanceId, bool allowTurnOff, out int? pnpCaps)
    {
        pnpCaps = null;
        string? classKey = TryGetClassKeyPath(instanceId);
        if (string.IsNullOrWhiteSpace(classKey))
        {
            return false;
        }

        try
        {
            pnpCaps = ApplyPnPCapabilities(classKey, allowTurnOff);
            return true;
        }
        catch
        {
            return false;
        }
    }

    public static bool? TryReadDevicePowerEnable(string instanceId)
    {
        if (string.IsNullOrWhiteSpace(instanceId))
        {
            return null;
        }

        try
        {
            using ManagementObjectSearcher searcher = new(
                "root\\WMI",
                "SELECT * FROM MSPower_DeviceEnable");
            foreach (ManagementObject mo in searcher.Get())
            {
                using (mo)
                {
                    string? instanceName = mo["InstanceName"] as string;
                    if (!InstanceNameMatches(instanceName, instanceId))
                    {
                        continue;
                    }

                    object? enableObj = mo["Enable"];
                    if (enableObj is bool enable)
                    {
                        return enable;
                    }

                    if (enableObj is not null && bool.TryParse(enableObj.ToString(), out bool parsed))
                    {
                        return parsed;
                    }
                }
            }
        }
        catch
        {
            return null;
        }

        return null;
    }

    /// <summary>
    /// Sets the live Device Manager Power Management checkbox via MSPower_DeviceEnable.
    /// Must query full instances (SELECT *) — partial projections make Put() a no-op.
    /// </summary>
    public static bool TrySetDevicePowerEnable(string instanceId, bool allowTurnOff)
    {
        if (string.IsNullOrWhiteSpace(instanceId))
        {
            return false;
        }

        bool wrote = false;
        try
        {
            using ManagementObjectSearcher searcher = new(
                "root\\WMI",
                "SELECT * FROM MSPower_DeviceEnable");
            foreach (ManagementObject mo in searcher.Get())
            {
                using (mo)
                {
                    string? instanceName = mo["InstanceName"] as string;
                    if (!InstanceNameMatches(instanceName, instanceId))
                    {
                        continue;
                    }

                    mo["Enable"] = allowTurnOff;
                    mo.Put();
                    wrote = true;

                    object? enableObj = mo["Enable"];
                    bool? enable = enableObj as bool?;
                    if (enable is bool actual && actual != allowTurnOff)
                    {
                        mo["Enable"] = allowTurnOff;
                        mo.Put();
                    }
                }
            }
        }
        catch
        {
            return false;
        }

        return wrote;
    }

    private static bool InstanceNameMatches(string? instanceName, string instanceId)
    {
        if (string.IsNullOrWhiteSpace(instanceName) || string.IsNullOrWhiteSpace(instanceId))
        {
            return false;
        }

        string normalized = instanceName.Replace('#', '\\');
        return normalized.StartsWith(instanceId, StringComparison.OrdinalIgnoreCase);
    }
}
