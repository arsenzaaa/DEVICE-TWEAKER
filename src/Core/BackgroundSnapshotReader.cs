namespace DeviceTweakerCS;

internal sealed record SnapshotReadResult<T>(T? Snapshot, string? Error) where T : class;

// One provider owns the slot until it really finishes, even after the UI stops
// waiting. The worker returns data only; late completion never publishes itself.
internal sealed class BackgroundSnapshotReader<T> where T : class
{
    private readonly object _sync = new();
    private Task<T>? _pending;

    internal async Task<SnapshotReadResult<T>> ReadAsync(Func<T> read, TimeSpan timeout, Action<Exception> lateError)
    {
        using CancellationTokenSource deadline = new();
        try { return await ReadAsync(read, Task.Delay(timeout, deadline.Token), lateError); }
        finally { deadline.Cancel(); }
    }

    internal async Task<SnapshotReadResult<T>> ReadAsync(Func<T> read, Task deadline, Action<Exception> lateError)
    {
        Task<T> work;
        lock (_sync)
        {
            if (_pending is { IsCompleted: false })
                return new(null, "The previous device state read is still running.");
            work = _pending = Task.Run(read);
        }
        try
        {
            BackgroundReadOutcome outcome = await BackgroundReadbacks.WaitCoreAsync(work, deadline, () => { }, lateError);
            return outcome == BackgroundReadOutcome.Completed
                ? new(await work, null)
                : new(null, "Device state read timed out. The current device list was preserved.");
        }
        catch (Exception error) { return new(null, error.Message); }
    }
}
