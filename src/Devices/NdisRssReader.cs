using System.Globalization;
using System.Management;

namespace DeviceTweakerCS;

// Read-only NetAdapter CIM provider. Incomplete/unsupported reads use the
// existing PowerShell path; never infer runtime RSS settings from the registry.
internal static class NdisRssReader
{
    internal sealed record Adapter(string Guid, string PnpId, string Name, string Description);

    internal static Adapter? MatchAdapter(IEnumerable<Adapter> adapters, string? guid, string pnpId)
    {
        Adapter[] candidates = adapters.ToArray();
        Adapter? match = System.Guid.TryParse(guid, out Guid expected)
            ? candidates.FirstOrDefault(a => System.Guid.TryParse(a.Guid, out Guid actual) && actual == expected) : null;
        return match ?? (string.IsNullOrWhiteSpace(pnpId) ? null : candidates.FirstOrDefault(a => string.Equals(a.PnpId, pnpId, StringComparison.OrdinalIgnoreCase)));
    }

    internal static bool TryRead(string? guid, string pnpId, CancellationToken cancellation,
        out NdisRssRuntimeState? state, out string? fallbackReason)
    {
        state = null; fallbackReason = null;
        try
        {
            cancellation.ThrowIfCancellationRequested();
            using ManagementObjectSearcher adapterQuery = Query(
                "SELECT InterfaceGuid, PnPDeviceID, Name, InterfaceDescription FROM MSFT_NetAdapter");
            using ManagementObjectCollection adapterRows = adapterQuery.Get();
            List<Adapter> adapters = [];
            foreach (ManagementBaseObject row in adapterRows)
            {
                using (row)
                {
                    cancellation.ThrowIfCancellationRequested();
                    adapters.Add(new(Text(row["InterfaceGuid"]), Text(row["PnPDeviceID"]),
                        Text(row["Name"]), Text(row["InterfaceDescription"])));
                }
            }
            Adapter? adapter = MatchAdapter(adapters, guid, pnpId);
            if (adapter is null || string.IsNullOrWhiteSpace(adapter.Name))
            { fallbackReason = "adapter-not-found"; return false; }
            string name = Escape(adapter.Name);
            using ManagementObjectSearcher rssQuery = Query(
                "SELECT Source, Enabled, BaseProcessorGroup, BaseProcessorNumber, MaxProcessorGroup, MaxProcessorNumber, "
                + "MaxProcessors, NumberOfReceiveQueues, Profile FROM MSFT_NetAdapterRssSettingData WHERE Name='" + name + "'");
            using ManagementObjectCollection rssRows = rssQuery.Get();
            NdisRssRuntimeState? candidate = null;
            foreach (ManagementBaseObject row in rssRows)
            {
                using (row)
                {
                    cancellation.ThrowIfCancellationRequested();
                    // Source 2 is device state; persistent settings are not live readback.
                    if (Number(row["Source"]) != 2 || candidate is not null)
                    { fallbackReason = "non-runtime-or-ambiguous-rss"; return false; }
                    if (row["Enabled"] is not bool enabled) { fallbackReason = "missing-enabled"; return false; }
                    candidate = new(true, true, enabled, Number(row["BaseProcessorGroup"]), Number(row["BaseProcessorNumber"]),
                        Number(row["MaxProcessorGroup"]), Number(row["MaxProcessorNumber"]), Number(row["MaxProcessors"]),
                        Number(row["NumberOfReceiveQueues"]), adapter.Name, adapter.Description, Text(row["Profile"]), string.Empty);
                }
            }
            if (candidate is not null) { state = candidate; return true; }
            fallbackReason = "rss-not-found";
        }
        catch (Exception error) when (error is ManagementException or System.Runtime.InteropServices.COMException
            or UnauthorizedAccessException or System.Security.SecurityException or InvalidCastException or FormatException)
        { fallbackReason = error.GetType().Name; }
        return false;
    }

    private static ManagementObjectSearcher Query(string query) => new(
        new ManagementScope(@"\\.\root\StandardCimv2"), new ObjectQuery(query),
        new System.Management.EnumerationOptions { Timeout = TimeSpan.FromSeconds(2), Rewindable = false });
    internal static string Escape(string value) => value.Replace("\\", "\\\\").Replace("'", "\\'");
    private static string Text(object? value) => Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
    private static int? Number(object? value) => value is null ? null
        : int.TryParse(Text(value), NumberStyles.Integer, CultureInfo.InvariantCulture, out int number) ? number : null;
}
