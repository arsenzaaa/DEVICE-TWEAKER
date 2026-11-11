namespace DeviceTweakerCS;

internal sealed record PciInterruptReading(string InstanceId, InterruptModeResult Result);

// Materialize all reads while the caller still owns its driver session. Results
// contain device identities and data, never controls or references to the Form.
internal static class PciInterruptBatch
{
    internal static PciInterruptReading[] Read(IReadOnlyList<string> instances,
        Func<string, InterruptModeResult> query, Action<string, Exception> reportError)
    {
        var readings = new PciInterruptReading[instances.Count];
        for (int index = 0; index < instances.Count; index++)
        {
            string instance = instances[index];
            InterruptModeResult result;
            try { result = query(instance); }
            catch (Exception error)
            {
                reportError(instance, error);
                result = new(InterruptMode.Unavailable, error.Message, DateTimeOffset.Now);
            }
            readings[index] = new(instance, result);
        }
        return readings;
    }

    internal static PciInterruptReading[] Unavailable(IReadOnlyList<string> instances, string detail)
        => instances.Select(instance => new PciInterruptReading(instance,
            new(InterruptMode.Unavailable, detail, DateTimeOffset.Now))).ToArray();
}
