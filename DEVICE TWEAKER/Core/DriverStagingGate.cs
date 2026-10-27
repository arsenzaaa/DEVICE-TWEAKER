namespace DeviceTweakerCS;

internal static class DriverStagingGate
{
    private static readonly object Sync = new();
    internal static bool Run(Func<bool> stage)
    {
        lock (Sync) return stage();
    }
}
