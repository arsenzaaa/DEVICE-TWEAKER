using Xunit;

namespace DeviceTweakerCS.Tests;

public sealed class BackgroundSnapshotReaderTests
{
    [Fact]
    public async Task BlockingProviderRunsOffTheCallingThreadAndReturnsItsData()
    {
        using ManualResetEventSlim release = new();
        using ManualResetEventSlim started = new();
        TaskCompletionSource<SnapshotReadResult<string>> completed = new(TaskCreationOptions.RunContinuationsAsynchronously);
        Thread caller = new(async () =>
        {
            try
            {
                int callerId = Environment.CurrentManagedThreadId, workerId = 0;
                var read = new BackgroundSnapshotReader<string>().ReadAsync(() =>
                {
                    workerId = Environment.CurrentManagedThreadId;
                    started.Set();
                    if (!release.Wait(TimeSpan.FromSeconds(5))) throw new TimeoutException();
                    return "complete snapshot";
                }, new TaskCompletionSource().Task, _ => { });
                // Keep this dedicated caller alive until the provider starts;
                // a pool continuation must not reuse its ID and make this flaky.
                Assert.True(started.Wait(TimeSpan.FromSeconds(5)));
                Assert.NotEqual(callerId, workerId);
                Assert.False(read.IsCompleted);
                release.Set();
                completed.SetResult(await read);
            }
            catch (Exception error) { release.Set(); completed.TrySetException(error); }
        }) { IsBackground = true };
        caller.Start();
        var result = await completed.Task.WaitAsync(TimeSpan.FromSeconds(10));
        Assert.Equal("complete snapshot", result.Snapshot);
        Assert.Null(result.Error);
    }

    [Fact]
    public async Task TimeoutKeepsSlotAndDoesNotReturnLateDataThenAllowsRetry()
    {
        BackgroundSnapshotReader<string> reader = new();
        using ManualResetEventSlim release = new();
        TaskCompletionSource started = new(TaskCreationOptions.RunContinuationsAsynchronously);
        TaskCompletionSource deadline = new(TaskCreationOptions.RunContinuationsAsynchronously);
        Task<SnapshotReadResult<string>> read = reader.ReadAsync(() =>
        {
            started.SetResult();
            if (!release.Wait(TimeSpan.FromSeconds(5))) throw new TimeoutException();
            return "stale";
        }, deadline.Task, _ => { });
        try
        {
            await started.Task.WaitAsync(TimeSpan.FromSeconds(5));
            deadline.SetResult();
            var expired = await read;
            Assert.Null(expired.Snapshot);
            Assert.Contains("timed out", expired.Error);
            bool invoked = false;
            var busy = await reader.ReadAsync(() => { invoked = true; return "unexpected"; }, Task.CompletedTask, _ => { });
            Assert.False(invoked);
            Assert.Contains("still running", busy.Error);
        }
        finally { release.Set(); }
        // Wait for the original worker to release the slot without spawning more.
        SnapshotReadResult<string> retry;
        using CancellationTokenSource retryLimit = new(TimeSpan.FromSeconds(5));
        do { await Task.Delay(1, retryLimit.Token); retry = await reader.ReadAsync(() => "fresh", new TaskCompletionSource().Task, _ => { }); }
        while (retry.Snapshot is null);
        Assert.Equal("fresh", retry.Snapshot);
    }

    [Fact]
    public async Task FailureIsNotAnEmptySuccessfulSnapshotAndCanBeRetried()
    {
        BackgroundSnapshotReader<string> reader = new();
        var failed = await reader.ReadAsync(() => throw new IOException("provider failed"), new TaskCompletionSource().Task, _ => { });
        Assert.Null(failed.Snapshot);
        Assert.Equal("provider failed", failed.Error);
        Assert.Equal("recovered", (await reader.ReadAsync(() => "recovered", new TaskCompletionSource().Task, _ => { })).Snapshot);
    }

    [Fact]
    public async Task LateFaultIsObservedAfterTimeout()
    {
        BackgroundSnapshotReader<string> reader = new();
        using ManualResetEventSlim release = new();
        TaskCompletionSource started = new(TaskCreationOptions.RunContinuationsAsynchronously);
        TaskCompletionSource deadline = new(TaskCreationOptions.RunContinuationsAsynchronously);
        TaskCompletionSource<Exception> observed = new(TaskCreationOptions.RunContinuationsAsynchronously);
        var read = reader.ReadAsync(() =>
        {
            started.SetResult();
            if (!release.Wait(TimeSpan.FromSeconds(5))) throw new TimeoutException();
            throw new IOException("late fault");
        }, deadline.Task, error => observed.TrySetResult(error));
        try
        {
            await started.Task.WaitAsync(TimeSpan.FromSeconds(5));
            deadline.SetResult();
            Assert.Null((await read).Snapshot);
        }
        finally { release.Set(); }
        Assert.Equal("late fault", (await observed.Task.WaitAsync(TimeSpan.FromSeconds(5))).Message);
    }
}
