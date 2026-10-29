using Microsoft.Win32;
using System.Runtime.InteropServices;

namespace DeviceTweakerCS;

internal sealed record DriverRegistryState(uint? Value, string Evidence)
{
    internal bool? Enabled => Value is uint value ? value != 0 : null;
}

internal sealed record DriverServiceState(string Name, bool? Exists, string State);

internal sealed record DriverEnvironmentSnapshot(
    DriverRegistryState Blocklist,
    DriverRegistryState HvciConfigured,
    bool? HvciRunning,
    DriverServiceState Defender,
    DriverServiceState Faceit,
    DriverServiceState VanguardDriver,
    DriverServiceState VanguardService)
{
    internal bool AntiCheatDetected => Faceit.Exists == true || VanguardDriver.Exists == true || VanguardService.Exists == true;
    internal bool PotentialRestriction => AntiCheatDetected || Blocklist.Enabled == true || HvciRunning == true;
}

// Read-only preflight. Presence/configuration is evidence, not a driver-load verdict.
internal static class DriverEnvironment
{
    internal const string BlocklistKey = @"SYSTEM\CurrentControlSet\Control\CI\Config";
    internal const string HvciKey = @"SYSTEM\CurrentControlSet\Control\DeviceGuard\Scenarios\HypervisorEnforcedCodeIntegrity";

    internal static DriverEnvironmentSnapshot Capture(bool? hvciRunning) => new(
        ReadDword(BlocklistKey, "VulnerableDriverBlocklistEnable"),
        ReadDword(HvciKey, "Enabled"), hvciRunning,
        ReadService("WinDefend"), ReadService("FACEIT"), ReadService("vgk"), ReadService("vgc"));

    private static DriverRegistryState ReadDword(string path, string name)
    {
        try
        {
            using RegistryKey root = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64);
            using RegistryKey? key = root.OpenSubKey(path, writable: false);
            object? value = key?.GetValue(name);
            return value is int number
                ? new(unchecked((uint)number), $"HKLM\\{path}\\{name}")
                : new(null, value is null ? "missing" : "invalid value type");
        }
        catch (Exception ex) when (ex is UnauthorizedAccessException or System.Security.SecurityException or IOException)
        {
            return new(null, $"unavailable: {ex.GetType().Name}");
        }
    }

    internal static DriverServiceState ReadService(string name)
    {
        IntPtr manager = OpenSCManager(null, null, 0x0001); // SC_MANAGER_CONNECT
        if (manager == IntPtr.Zero) return new(name, null, $"unavailable ({Marshal.GetLastWin32Error()})");
        try
        {
            IntPtr service = OpenService(manager, name, 0x0004); // SERVICE_QUERY_STATUS only
            if (service == IntPtr.Zero)
            {
                int error = Marshal.GetLastWin32Error();
                return error == 1060 ? new(name, false, "not found") : new(name, null, $"unavailable ({error})");
            }
            try
            {
                if (!QueryServiceStatusEx(service, 0, out ServiceStatus status, Marshal.SizeOf<ServiceStatus>(), out _))
                    return new(name, true, $"state unavailable ({Marshal.GetLastWin32Error()})");
                string state = status.CurrentState switch
                {
                    1 => "stopped", 2 => "starting", 3 => "stopping", 4 => "running",
                    5 => "continuing", 6 => "pausing", 7 => "paused", _ => $"state {status.CurrentState}"
                };
                return new(name, true, state);
            }
            finally { CloseServiceHandle(service); }
        }
        finally { CloseServiceHandle(manager); }
    }

    internal static string BuildConfirmation(DriverEnvironmentSnapshot snapshot)
    {
        List<string> lines =
        [
            "Driver preflight (read-only):",
            $"Vulnerable Driver Blocklist (registry): {RegistryText(snapshot.Blocklist)}",
            $"Memory Integrity / HVCI: runtime {BoolText(snapshot.HvciRunning)}, registry {RegistryText(snapshot.HvciConfigured)}",
            $"Microsoft Defender service: {ServiceText(snapshot.Defender)}",
            $"FACEIT service: {ServiceText(snapshot.Faceit)}",
            $"Vanguard: vgk {ServiceText(snapshot.VanguardDriver)}, vgc {ServiceText(snapshot.VanguardService)}",
        ];
        if (snapshot.AntiCheatDetected)
            lines.Add("Anti-cheat components were found. Compatibility with DTIMOD/KDU is not confirmed; a conflict is possible, not detected.");
        lines.Add("These checks do not test driver loading. A missing registry value does not mean the protection is disabled.");
        return string.Join("\n", lines);
    }

    internal static string BuildCompactConfirmation(DriverEnvironmentSnapshot snapshot)
    {
        List<string> notes = [];
        bool faceit = snapshot.Faceit.Exists == true;
        bool vanguard = snapshot.VanguardDriver.Exists == true || snapshot.VanguardService.Exists == true;
        if (faceit || vanguard)
        {
            notes.Add(faceit && vanguard
                ? "FACEIT and Vanguard detected.\nThey may conflict with DTIMOD.sys."
                : faceit
                    ? "FACEIT detected.\nIt may conflict with DTIMOD.sys."
                    : "Vanguard detected.\nIt may conflict with DTIMOD.sys.");
        }
        if (snapshot.Blocklist.Enabled == true || snapshot.HvciRunning == true)
            notes.Add("Windows settings may prevent DTIMOD.sys from loading.");
        if (snapshot.HvciRunning is null || snapshot.Blocklist.Evidence.StartsWith("unavailable", StringComparison.Ordinal)
            || snapshot.Blocklist.Evidence == "invalid value type" || snapshot.HvciConfigured.Evidence.StartsWith("unavailable", StringComparison.Ordinal)
            || snapshot.Faceit.Exists is null || snapshot.VanguardDriver.Exists is null || snapshot.VanguardService.Exists is null)
            notes.Add("Some system checks are unavailable.");
        const string question = "Include USB IMOD in auto-optimization?";
        string concern = notes.Count == 0 ? string.Empty : "\n\n" + string.Join("\n", notes);
        return question + concern + "\n\n"
            + "SKIP continues auto-optimization without USB IMOD.";
    }

    private static string RegistryText(DriverRegistryState state) => state.Value is uint value
        ? value == 0 ? "0 (configured off)" : $"{value} (configured on)"
        : state.Evidence == "missing" ? "not explicitly configured" : "unavailable";
    private static string BoolText(bool? value) => value switch { true => "enabled", false => "disabled", _ => "unavailable" };
    private static string ServiceText(DriverServiceState service) => service.Exists switch
    {
        false => "not found", null => "unavailable", _ => $"found ({service.State})"
    };

    [StructLayout(LayoutKind.Sequential)]
    private struct ServiceStatus
    {
        public uint ServiceType, CurrentState, ControlsAccepted, Win32ExitCode, ServiceSpecificExitCode;
        public uint CheckPoint, WaitHint, ProcessId, ServiceFlags;
    }
    [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern IntPtr OpenSCManager(string? machine, string? database, uint access);
    [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern IntPtr OpenService(IntPtr manager, string name, uint access);
    [DllImport("advapi32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool QueryServiceStatusEx(IntPtr service, int level, out ServiceStatus status, int bytes, out int needed);
    [DllImport("advapi32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CloseServiceHandle(IntPtr handle);
}
