using Microsoft.Win32;
using System.Runtime.InteropServices;

namespace DeviceTweakerCS;

/// <summary>
/// USB Selective Suspend knobs documented by Microsoft:
/// power plan (Teams Rooms / powercfg), Device Manager idle checkbox
/// (KDNET: xHCI + root hub), and Device Parameters DWORDs.
/// </summary>
internal static class UsbSelectiveSuspendPolicy
{
    public const string SelectiveSuspendEnabledName = "SelectiveSuspendEnabled";
    public const string EnhancedPowerManagementEnabledName = "EnhancedPowerManagementEnabled";

    // USB settings subgroup / USB selective suspend setting (powercfg / power plan).
    // https://learn.microsoft.com/troubleshoot/microsoftteams/teams-rooms-and-devices/usb-selective-suspend-status-unhealthy
    public static readonly Guid UsbSettingsSubgroup = new("2a737441-1930-4402-8d77-b2bebba308a3");
    public static readonly Guid UsbSelectiveSuspendSetting = new("48e6b7a6-50f5-4782-a5d4-53bb8f07e226");

    public static string DeviceParametersPath(string instanceId)
        => $@"SYSTEM\CurrentControlSet\Enum\{instanceId}\Device Parameters";

    public static string PowerPlanSettingPath(Guid schemeGuid)
        => @"SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes\" +
           $"{schemeGuid:D}\\{UsbSettingsSubgroup:D}\\{UsbSelectiveSuspendSetting:D}";

    public static IReadOnlyList<string> EnumerateRootHubs(string controllerInstanceId)
    {
        if (string.IsNullOrWhiteSpace(controllerInstanceId))
        {
            return [];
        }

        List<string> hubs = [];
        foreach ((string controllerId, string dependentId) in WmiInterop.GetUsbControllerDevicePairs(static id => id))
        {
            if (!SameController(controllerInstanceId, controllerId)
                || string.IsNullOrWhiteSpace(dependentId)
                || !dependentId.Contains("ROOT_HUB", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (!hubs.Contains(dependentId, StringComparer.OrdinalIgnoreCase))
            {
                hubs.Add(dependentId);
            }
        }

        return hubs;
    }

    public static IReadOnlyList<string> EnumerateBackupInstanceIds(string controllerInstanceId)
    {
        List<string> ids = [];
        if (!string.IsNullOrWhiteSpace(controllerInstanceId))
        {
            ids.Add(controllerInstanceId);
        }

        foreach (string hubId in EnumerateRootHubs(controllerInstanceId))
        {
            ids.Add(hubId);
        }

        return ids;
    }

    public static void ApplyControllerAndHubs(string controllerInstanceId, bool enabled)
    {
        if (string.IsNullOrWhiteSpace(controllerInstanceId))
        {
            throw new InvalidOperationException("USB controller instance id is empty.");
        }

        SetDeviceParameterDword(controllerInstanceId, SelectiveSuspendEnabledName, enabled ? 1 : 0);
        SetDeviceParameterDword(controllerInstanceId, EnhancedPowerManagementEnabledName, enabled ? 1 : 0);
        DevicePowerPolicy.TryApplyInstancePnPCapabilities(controllerInstanceId, allowTurnOff: enabled, out _);
        DevicePowerPolicy.TrySetDevicePowerEnable(controllerInstanceId, allowTurnOff: enabled);

        foreach (string hubId in EnumerateRootHubs(controllerInstanceId))
        {
            SetDeviceParameterDword(hubId, SelectiveSuspendEnabledName, enabled ? 1 : 0);
            SetDeviceParameterDword(hubId, EnhancedPowerManagementEnabledName, enabled ? 1 : 0);
            DevicePowerPolicy.TryApplyInstancePnPCapabilities(hubId, allowTurnOff: enabled, out _);
            DevicePowerPolicy.TrySetDevicePowerEnable(hubId, allowTurnOff: enabled);
        }
    }

    public static bool TryReadEnabled(string instanceId, out bool enabled)
    {
        enabled = true;
        if (string.IsNullOrWhiteSpace(instanceId))
        {
            return false;
        }

        // Prefer the live Device Manager checkbox when WMI exposes it.
        bool? wmi = DevicePowerPolicy.TryReadDevicePowerEnable(instanceId);
        if (wmi is bool live)
        {
            enabled = live;
            return true;
        }

        string? classKey = DevicePowerPolicy.TryGetClassKeyPath(instanceId);
        int? pnpCaps = DevicePowerPolicy.TryReadNicPnPCapabilities(classKey);
        if (pnpCaps is int caps)
        {
            enabled = DevicePowerPolicy.IsNicTurnOffAllowed(caps);
            return true;
        }

        int? selective = TryReadDeviceParameterDword(instanceId, SelectiveSuspendEnabledName);
        int? enhanced = TryReadDeviceParameterDword(instanceId, EnhancedPowerManagementEnabledName);
        if (selective is null && enhanced is null)
        {
            return true;
        }

        enabled = (selective is null || selective.Value != 0)
            && (enhanced is null || enhanced.Value != 0);
        return true;
    }

    public static int? TryReadDeviceParameterDword(string instanceId, string name)
    {
        try
        {
            using RegistryKey? key = Registry.LocalMachine.OpenSubKey(DeviceParametersPath(instanceId), writable: false);
            return key?.GetValue(name) as int?;
        }
        catch
        {
            return null;
        }
    }

    public static void ActivateCurrentPowerScheme()
    {
        if (!TryGetActiveScheme(out Guid scheme))
        {
            throw new InvalidOperationException("Could not read the active Windows power scheme.");
        }

        uint apply = PowerSetActiveScheme(IntPtr.Zero, ref scheme);
        if (apply != 0)
        {
            throw new InvalidOperationException($"PowerSetActiveScheme failed (0x{apply:X8}).");
        }
    }

    public static void SetPowerPlanEnabled(bool enabled)
    {
        if (!TryGetActiveScheme(out Guid scheme))
        {
            throw new InvalidOperationException("Could not read the active Windows power scheme.");
        }

        uint value = enabled ? 1u : 0u;
        Guid subgroup = UsbSettingsSubgroup;
        Guid setting = UsbSelectiveSuspendSetting;
        uint ac = PowerWriteACValueIndex(IntPtr.Zero, ref scheme, ref subgroup, ref setting, value);
        if (ac != 0)
        {
            throw new InvalidOperationException($"PowerWriteACValueIndex failed (0x{ac:X8}).");
        }

        uint dc = PowerWriteDCValueIndex(IntPtr.Zero, ref scheme, ref subgroup, ref setting, value);
        if (dc != 0)
        {
            throw new InvalidOperationException($"PowerWriteDCValueIndex failed (0x{dc:X8}).");
        }

        uint apply = PowerSetActiveScheme(IntPtr.Zero, ref scheme);
        if (apply != 0)
        {
            throw new InvalidOperationException($"PowerSetActiveScheme failed (0x{apply:X8}).");
        }
    }

    public static bool TryGetActiveScheme(out Guid scheme)
    {
        scheme = Guid.Empty;
        uint rc = PowerGetActiveScheme(IntPtr.Zero, out IntPtr ptr);
        if (rc != 0 || ptr == IntPtr.Zero)
        {
            return false;
        }

        try
        {
            scheme = Marshal.PtrToStructure<Guid>(ptr);
            return scheme != Guid.Empty;
        }
        finally
        {
            LocalFree(ptr);
        }
    }

    public static void SetDeviceParameterDword(string instanceId, string name, int value)
    {
        string path = DeviceParametersPath(instanceId);
        using RegistryKey? key = Registry.LocalMachine.CreateSubKey(path);
        if (key is null)
        {
            throw new InvalidOperationException($"Could not open HKLM\\{path} for writing.");
        }

        key.SetValue(name, value, RegistryValueKind.DWord);
    }

    private static bool SameController(string left, string right)
    {
        if (string.IsNullOrWhiteSpace(left) || string.IsNullOrWhiteSpace(right))
        {
            return false;
        }

        // Exact instance match only. Matching PCI\VEN_&DEV_ alone would apply SS
        // to sibling xHCI controllers that share the same chip ID.
        return string.Equals(left, right, StringComparison.OrdinalIgnoreCase);
    }

    [DllImport("powrprof.dll")]
    private static extern uint PowerGetActiveScheme(IntPtr userRootPowerKey, out IntPtr activePolicyGuid);

    [DllImport("powrprof.dll")]
    private static extern uint PowerWriteACValueIndex(
        IntPtr rootPowerKey,
        ref Guid schemeGuid,
        ref Guid subGroupOfPowerSettingsGuid,
        ref Guid powerSettingGuid,
        uint acValueIndex);

    [DllImport("powrprof.dll")]
    private static extern uint PowerWriteDCValueIndex(
        IntPtr rootPowerKey,
        ref Guid schemeGuid,
        ref Guid subGroupOfPowerSettingsGuid,
        ref Guid powerSettingGuid,
        uint dcValueIndex);

    [DllImport("powrprof.dll")]
    private static extern uint PowerSetActiveScheme(IntPtr userRootPowerKey, ref Guid schemeGuid);

    [DllImport("kernel32.dll")]
    private static extern IntPtr LocalFree(IntPtr hMem);
}
