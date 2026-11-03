namespace DeviceTweakerCS;

internal enum BackgroundReadOutcome { Completed, TimedOut }

// A provider stuck inside Windows cannot always be cancelled. Stop waiting,
// invalidate publication and observe its eventual error without pumping messages.
internal static class BackgroundReadbacks
{
    internal static async Task<BackgroundReadOutcome> WaitAsync(
        Task work, TimeSpan timeout, Action invalidate, Action<Exception> lateError)
    {
        using CancellationTokenSource deadline = new();
        try
        {
            return await WaitCoreAsync(work, Task.Delay(timeout, deadline.Token), invalidate, lateError);
        }
        finally { deadline.Cancel(); }
    }

    internal static async Task<BackgroundReadOutcome> WaitCoreAsync(
        Task work, Task deadline, Action invalidate, Action<Exception> lateError)
    {
        if (await Task.WhenAny(work, deadline) == work || work.IsCompleted)
        {
            await work;
            return BackgroundReadOutcome.Completed;
        }
        await deadline;
        invalidate();
        _ = ObserveLateFailureAsync(work, lateError);
        return BackgroundReadOutcome.TimedOut;
    }

    private static async Task ObserveLateFailureAsync(Task work, Action<Exception> lateError)
    {
        try { await work.ConfigureAwait(false); }
        catch (Exception error) { lateError(error); }
    }
}
