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
    internal sealed record ApplyResult(
        int TargetCount,
        int PnpSucceeded,
        int WmiExposed,
        int WmiSucceeded,
        IReadOnlyList<string> Notes,
        IReadOnlyList<string> Errors)
    {
        public bool Success => Errors.Count == 0;
    }

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

    internal static bool TryParsePowerPlanTarget(string hive, string path, string name, out Guid scheme, out bool ac)
    {
        scheme = Guid.Empty;
        ac = string.Equals(name, "ACSettingIndex", StringComparison.OrdinalIgnoreCase);
        if (!string.Equals(hive, "HKLM", StringComparison.OrdinalIgnoreCase)
            || (!ac && !string.Equals(name, "DCSettingIndex", StringComparison.OrdinalIgnoreCase))) return false;
        const string root = @"SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes\";
        if (!path.StartsWith(root, StringComparison.OrdinalIgnoreCase)) return false;
        string[] parts = path[root.Length..].Split('\\');
        return parts.Length == 3 && Guid.TryParseExact(parts[0], "D", out scheme) && scheme != Guid.Empty
            && Guid.TryParse(parts[1], out Guid subgroup) && subgroup == UsbSettingsSubgroup
            && Guid.TryParse(parts[2], out Guid setting) && setting == UsbSelectiveSuspendSetting;
    }

    internal static uint ReadPowerPlanIndex(Guid scheme, bool ac)
    {
        Guid subgroup = UsbSettingsSubgroup, setting = UsbSelectiveSuspendSetting;
        uint value;
        uint rc = ac
            ? PowerReadACValueIndex(IntPtr.Zero, ref scheme, ref subgroup, ref setting, out value)
            : PowerReadDCValueIndex(IntPtr.Zero, ref scheme, ref subgroup, ref setting, out value);
        if (rc != 0) throw new InvalidOperationException($"PowerRead{(ac ? "AC" : "DC")}ValueIndex failed (0x{rc:X8}), scheme={scheme:D}.");
        return value;
    }

    internal static uint ReadPowerPlanDefaultIndex(Guid scheme, bool ac)
    {
        // Defaults depend on the personality, not on an arbitrary custom plan GUID.
        Guid none = new("fea3413e-7e05-4911-9a71-700331f1c294");
        Guid personalitySetting = new("245d8541-3943-4422-b025-13a784f679b7");
        byte[] buffer = new byte[16];
        uint size = 16;
        uint rc = PowerReadACValue(IntPtr.Zero, ref scheme, ref none, ref personalitySetting, out _, buffer, ref size);
        if (rc != 0 || size != 16) throw new InvalidOperationException($"Cannot read power scheme personality (0x{rc:X8}), scheme={scheme:D}.");
        Guid personality = new(buffer), subgroup = UsbSettingsSubgroup, setting = UsbSelectiveSuspendSetting;
        uint value;
        rc = ac
            ? PowerReadACDefaultIndex(IntPtr.Zero, ref personality, ref subgroup, ref setting, out value)
            : PowerReadDCDefaultIndex(IntPtr.Zero, ref personality, ref subgroup, ref setting, out value);
        if (rc != 0) throw new InvalidOperationException($"Cannot read power setting default (0x{rc:X8}), scheme={scheme:D}.");
        return value;
    }

    internal static void RestorePowerPlanIndex(Guid scheme, bool ac, uint value)
    {
        QaRun.DemandWrite("RestorePowerPlanIndex");
        PowerPlanRestoreService.Restore(new WindowsPowerPlanBackend(), scheme, ac, value);
    }

    private sealed class WindowsPowerPlanBackend : IPowerPlanBackend
    {
        public Guid GetActiveScheme()
            => TryGetActiveScheme(out Guid scheme) ? scheme : Guid.Empty;
        public uint ReadIndex(Guid scheme, bool ac) => ReadPowerPlanIndex(scheme, ac);
        public void WriteIndex(Guid scheme, bool ac, uint value)
        {
            QaRun.DemandWrite("PowerPlan.WriteIndex");
            Guid subgroup = UsbSettingsSubgroup, setting = UsbSelectiveSuspendSetting;
            uint rc = ac
                ? PowerWriteACValueIndex(IntPtr.Zero, ref scheme, ref subgroup, ref setting, value)
                : PowerWriteDCValueIndex(IntPtr.Zero, ref scheme, ref subgroup, ref setting, value);
            if (rc != 0) throw new InvalidOperationException($"PowerWrite{(ac ? "AC" : "DC")}ValueIndex failed (0x{rc:X8}), scheme={scheme:D}.");
        }
        public void Activate(Guid scheme)
        {
            QaRun.DemandWrite("PowerPlan.Activate");
            uint rc = PowerSetActiveScheme(IntPtr.Zero, ref scheme);
            if (rc != 0) throw new InvalidOperationException($"PowerSetActiveScheme failed (0x{rc:X8}).");
        }
    }

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

    public static ApplyResult ApplyControllerAndHubs(string controllerInstanceId, bool enabled)
    {
        QaRun.DemandWrite("ApplyControllerAndHubs");
        if (string.IsNullOrWhiteSpace(controllerInstanceId))
        {
            throw new InvalidOperationException("USB controller instance id is empty.");
        }

        List<string> targets = [controllerInstanceId, .. EnumerateRootHubs(controllerInstanceId)];
        List<string> notes = [];
        List<string> errors = [];
        int pnpSucceeded = 0;
        int wmiExposed = 0;
        int wmiSucceeded = 0;
        foreach (string target in targets)
        {
            try
            {
                SetDeviceParameterDword(target, SelectiveSuspendEnabledName, enabled ? 1 : 0);
                SetDeviceParameterDword(target, EnhancedPowerManagementEnabledName, enabled ? 1 : 0);
            }
            catch (Exception ex)
            {
                errors.Add($"{target}: registry: {ex.Message}");
            }

            DevicePowerPolicy.WriteStatus pnpStatus = DevicePowerPolicy.ApplyInstancePnPCapabilities(
                target,
                allowTurnOff: enabled,
                out _,
                out string? pnpError);
            if (pnpStatus == DevicePowerPolicy.WriteStatus.Applied)
            {
                pnpSucceeded++;
            }
            else if (pnpStatus == DevicePowerPolicy.WriteStatus.NotExposed)
            {
                notes.Add($"{target}: PnPCapabilities class key is not exposed");
            }
            else
            {
                errors.Add($"{target}: PnPCapabilities: {pnpError ?? "write failed"}");
            }

            DevicePowerPolicy.WriteStatus wmiStatus = DevicePowerPolicy.SetDevicePowerEnable(
                target,
                allowTurnOff: enabled,
                out string? wmiError);
            if (wmiStatus == DevicePowerPolicy.WriteStatus.Applied)
            {
                wmiExposed++;
                wmiSucceeded++;
            }
            else if (wmiStatus is DevicePowerPolicy.WriteStatus.NotExposed or DevicePowerPolicy.WriteStatus.NotSupported)
            {
                notes.Add($"{target}: MSPower_DeviceEnable status={wmiStatus} {wmiError}");
            }
            else
            {
                wmiExposed++;
                errors.Add($"{target}: MSPower_DeviceEnable: {wmiError ?? "write failed"}");
            }
        }

        return new ApplyResult(targets.Count, pnpSucceeded, wmiExposed, wmiSucceeded, notes, errors);
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
        QaRun.DemandWrite("ActivateCurrentPowerScheme");
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
        QaRun.DemandWrite("SetPowerPlanEnabled");
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
        QaRun.DemandWrite("SetDeviceParameterDword");
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
    private static extern uint PowerReadACValueIndex(IntPtr root, ref Guid scheme, ref Guid subgroup, ref Guid setting, out uint value);
    [DllImport("powrprof.dll")]
    private static extern uint PowerReadDCValueIndex(IntPtr root, ref Guid scheme, ref Guid subgroup, ref Guid setting, out uint value);
    [DllImport("powrprof.dll")]
    private static extern uint PowerReadACDefaultIndex(IntPtr root, ref Guid personality, ref Guid subgroup, ref Guid setting, out uint value);
    [DllImport("powrprof.dll")]
    private static extern uint PowerReadDCDefaultIndex(IntPtr root, ref Guid personality, ref Guid subgroup, ref Guid setting, out uint value);
    [DllImport("powrprof.dll")]
    private static extern uint PowerReadACValue(IntPtr root, ref Guid scheme, ref Guid subgroup, ref Guid setting, out uint type, [Out] byte[] buffer, ref uint size);

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
