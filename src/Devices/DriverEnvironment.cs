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
    internal const string PurposeNotice =
        "DTIMOD provides hardware access for USB IMOD, NIC ITR and PCI register reads.";
    internal const string AntiCheatNotice =
        "Easy Anti-Cheat, Riot Vanguard, FACEIT Anti-Cheat and other protection software may block the driver or disconnect you from a game. Compatibility is not guaranteed.";
    internal const string LifetimeNotice =
        "Closing DEVICE TWEAKER does not unload the driver. Restart Windows to unload it. Saved profiles can load it again at sign-in, wake or device reconnection.";

    internal const string BriefCompatibilityNotice =
        "Anti-cheat software may block DTIMOD or disconnect you from a game.";
    internal const string BriefLifetimeNotice =
        "Closing the app does not unload DTIMOD. Saved profiles can load it automatically.";

    internal static IReadOnlyList<string> CheckSummary(DriverEnvironmentSnapshot snapshot)
    {
        List<string> lines = [snapshot.HvciRunning switch
        {
            true => "Memory integrity (HVCI) is enabled.",
            false => "Memory integrity (HVCI) is not running.",
            _ => "Memory integrity (HVCI) could not be read.",
        }];
        lines.Add(snapshot.Blocklist.Enabled switch
        {
            true => "Driver blocklist is configured on.",
            false => "Driver blocklist registry value is set to off.",
            _ => snapshot.Blocklist.Evidence == "missing"
                ? "Driver blocklist has no explicit registry setting."
                : "Driver blocklist registry setting could not be read.",
        });
        if (snapshot.Faceit.Exists == true) lines.Add("A FACEIT service is installed.");
        if (snapshot.VanguardDriver.Exists == true || snapshot.VanguardService.Exists == true)
            lines.Add("A Riot Vanguard service is installed.");
        if (snapshot.Faceit.Exists is null || snapshot.VanguardDriver.Exists is null || snapshot.VanguardService.Exists is null)
            lines.Add("Some service information could not be read.");
        return lines;
    }

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
            lines.Add("Anti-cheat compatibility is not tested by this check.");
        lines.Add("These checks do not test driver loading. A missing registry value does not mean the protection is disabled.");
        return string.Join("\n", lines);
    }

    internal static string BuildCompactConfirmation(DriverEnvironmentSnapshot snapshot, bool includeCompatibilityNotice = true)
    {
        List<string> notes = [];
        if (snapshot.Blocklist.Enabled == true)
            notes.Add("Driver blocklist is configured on.");
        if (snapshot.HvciRunning == true)
            notes.Add("Memory integrity (HVCI) is enabled.");
        if (snapshot.HvciRunning is null || snapshot.Blocklist.Evidence.StartsWith("unavailable", StringComparison.Ordinal)
            || snapshot.Blocklist.Evidence == "invalid value type" || snapshot.HvciConfigured.Evidence.StartsWith("unavailable", StringComparison.Ordinal)
            || snapshot.Faceit.Exists is null || snapshot.VanguardDriver.Exists is null || snapshot.VanguardService.Exists is null)
            notes.Add("Some system checks are unavailable.");
        const string question = "Include USB IMOD in auto-optimization?";
        string concern = notes.Count == 0 ? string.Empty : "\n\n" + string.Join("\n", notes);
        return question + (includeCompatibilityNotice
            ? "\n\n" + BriefCompatibilityNotice + "\n\n" + BriefLifetimeNotice : string.Empty) + concern;
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
