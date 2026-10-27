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
    internal enum WriteStatus
    {
        Applied,
        NotExposed,
        NotSupported,
        QueryFailed,
        Failed,
        ReadbackUnavailable,
        ReadbackMismatch
    }

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
        QaRun.DemandWrite("ApplyPnPCapabilities");
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
        => ApplyInstancePnPCapabilities(instanceId, allowTurnOff, out pnpCaps, out _) == WriteStatus.Applied;

    public static WriteStatus ApplyInstancePnPCapabilities(
        string instanceId,
        bool allowTurnOff,
        out int? pnpCaps,
        out string? error)
    {
        pnpCaps = null;
        error = null;
        string? classKey = TryGetClassKeyPath(instanceId);
        if (string.IsNullOrWhiteSpace(classKey))
        {
            return WriteStatus.NotExposed;
        }

        try
        {
            pnpCaps = ApplyPnPCapabilities(classKey, allowTurnOff);
            return WriteStatus.Applied;
        }
        catch (Exception ex)
        {
            error = ex.Message;
            return WriteStatus.Failed;
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
    /// Must query full instances (SELECT *) , partial projections make Put() a no-op.
    /// </summary>
    public static bool TrySetDevicePowerEnable(string instanceId, bool allowTurnOff)
        => SetDevicePowerEnable(instanceId, allowTurnOff, out _) == WriteStatus.Applied;

    public static WriteStatus SetDevicePowerEnable(string instanceId, bool allowTurnOff, out string? error)
    {
        QaRun.DemandWrite("SetDevicePowerEnable");
        return SetDevicePowerEnableCore(instanceId, allowTurnOff, new WmiPowerBackend(), out error);
    }

    internal interface IPowerEntry
    {
        void Write(bool enabled);
        bool? ReadBack();
    }

    internal interface IPowerBackend
    {
        IEnumerable<IPowerEntry> Query(string instanceId);
    }

    internal static WriteStatus SetDevicePowerEnableCore(string instanceId, bool allowTurnOff, IPowerBackend backend, out string? error)
    {
        error = null;
        if (string.IsNullOrWhiteSpace(instanceId))
        {
            return WriteStatus.NotExposed;
        }

        int matched = 0, written = 0, verified = 0;
        string stage = "query";
        try
        {
            foreach (IPowerEntry entry in backend.Query(instanceId))
            {
                matched++;
                stage = "write";
                entry.Write(allowTurnOff);
                written++;
                stage = "readback";
                bool? actual = entry.ReadBack();
                if (!actual.HasValue || actual.Value != allowTurnOff)
                {
                    error = $"stage={stage} matched={matched} written={written} verified={verified} requested={allowTurnOff} actual={(actual?.ToString() ?? "unavailable")}";
                    return actual.HasValue ? WriteStatus.ReadbackMismatch : WriteStatus.ReadbackUnavailable;
                }
                verified++;
                stage = "query";
            }
        }
        catch (Exception ex)
        {
            string code = ex is ManagementException wmi ? $" wmiCode={wmi.ErrorCode}" : string.Empty;
            error = $"stage={stage} matched={matched} written={written} verified={verified} hresult=0x{ex.HResult:X8}{code} detail={ex.Message}";
            if (stage == "readback") return WriteStatus.ReadbackUnavailable;
            if (ex is ManagementException { ErrorCode: ManagementStatus.NotSupported }
                || ex.HResult == unchecked((int)0x8004100C)) return WriteStatus.NotSupported;
            if (stage == "query" && matched == 0
                && ex is ManagementException { ErrorCode: ManagementStatus.InvalidClass or ManagementStatus.InvalidNamespace })
                return WriteStatus.NotExposed;
            return stage == "query" ? WriteStatus.QueryFailed : WriteStatus.Failed;
        }

        return verified > 0 ? WriteStatus.Applied : WriteStatus.NotExposed;
    }

    internal static bool IsWriteFailure(WriteStatus status)
        => status is WriteStatus.QueryFailed or WriteStatus.Failed or WriteStatus.ReadbackUnavailable or WriteStatus.ReadbackMismatch;

    private sealed class WmiPowerBackend : IPowerBackend
    {
        public IEnumerable<IPowerEntry> Query(string instanceId)
        {
            using ManagementObjectSearcher searcher = new("root\\WMI", "SELECT * FROM MSPower_DeviceEnable");
            using ManagementObjectCollection instances = searcher.Get();
            foreach (ManagementObject mo in instances)
            {
                using (mo)
                {
                    if (InstanceNameMatches(mo["InstanceName"] as string, instanceId)) yield return new WmiPowerEntry(mo);
                }
            }
        }
    }

    private sealed class WmiPowerEntry(ManagementObject instance) : IPowerEntry
    {
        public void Write(bool enabled)
        {
            QaRun.DemandWrite("SetDevicePowerEnable.WMI.Put");
            instance["Enable"] = enabled;
            instance.Put();
        }

        public bool? ReadBack()
        {
            // Refresh from WMI: the local value assigned before Put is not a readback.
            instance.Get();
            object? value = instance["Enable"];
            return value is bool b ? b : bool.TryParse(value?.ToString(), out bool parsed) ? parsed : null;
        }
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
