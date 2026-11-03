namespace DeviceTweakerCS;

internal sealed record DeviceDiscoveryResult(IReadOnlyList<DeviceInfo>? Devices, string? Error);

// Owns one provider task, including after a timeout. A stuck WMI provider must
// not create more worker threads on each retry, or publish a late device list.
internal sealed class DeviceDiscovery
{
    private readonly object _sync = new();
    private Task<IReadOnlyList<DeviceInfo>>? _active;

    internal async Task<DeviceDiscoveryResult> ReadAsync(
        Func<Task<IReadOnlyList<DeviceInfo>>> provider, TimeSpan timeout, Action<Exception> lateError)
    {
        using CancellationTokenSource deadline = new();
        try { return await ReadAsync(provider, Task.Delay(timeout, deadline.Token), lateError); }
        finally { deadline.Cancel(); }
    }

    internal async Task<DeviceDiscoveryResult> ReadAsync(
        Func<Task<IReadOnlyList<DeviceInfo>>> provider, Task deadline, Action<Exception> lateError)
    {
        Task<IReadOnlyList<DeviceInfo>> pending;
        lock (_sync)
        {
            if (_active is { IsCompleted: false })
                return new(null, "The previous device scan is still running. The current list was preserved.");
            try { pending = _active = provider(); }
            catch (Exception error) { return new(null, error.Message); }
        }
        try
        {
            BackgroundReadOutcome outcome = await BackgroundReadbacks.WaitCoreAsync(pending, deadline, () => { }, lateError);
            return outcome == BackgroundReadOutcome.Completed
                ? new(await pending, null)
                : new(null, "Device discovery timed out. The current list was preserved. Retry after the provider finishes.");
        }
        catch (Exception error) { return new(null, error.Message); }
    }
}
