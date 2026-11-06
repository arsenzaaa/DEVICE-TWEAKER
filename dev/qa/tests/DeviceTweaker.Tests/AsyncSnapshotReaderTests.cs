using Xunit;

namespace DeviceTweakerCS.Tests;

public sealed class AsyncSnapshotReaderTests
{
    [Fact]
    public async Task AwaitedProviderKeepsItsSlotAfterDeadlineAndDoesNotReturnLateData()
    {
        BackgroundSnapshotReader<string> reader = new();
        TaskCompletionSource started = new(TaskCreationOptions.RunContinuationsAsynchronously);
        TaskCompletionSource<string> finish = new(TaskCreationOptions.RunContinuationsAsynchronously);
        TaskCompletionSource deadline = new(TaskCreationOptions.RunContinuationsAsynchronously);
        Task<SnapshotReadResult<string>> read = reader.ReadTaskAsync(async () =>
        { started.SetResult(); return await finish.Task; }, deadline.Task, _ => { });
        await started.Task.WaitAsync(TimeSpan.FromSeconds(5));
        deadline.SetResult();
        Assert.Null((await read).Snapshot);
        bool duplicate = false;
        Assert.Contains("still running", (await reader.ReadTaskAsync(() =>
        { duplicate = true; return Task.FromResult("wrong"); }, Task.CompletedTask, _ => { })).Error);
        Assert.False(duplicate);
        finish.SetResult("late");
        using CancellationTokenSource limit = new(TimeSpan.FromSeconds(5));
        SnapshotReadResult<string> retry;
        do { await Task.Delay(1, limit.Token); retry = await reader.ReadTaskAsync(() => Task.FromResult("fresh"), new TaskCompletionSource().Task, _ => { }); }
        while (retry.Snapshot is null);
        Assert.Equal("fresh", retry.Snapshot);
    }
    [Fact]
    public async Task FaultAfterAwaitIsReportedAndRetryRemainsPossible()
    {
        BackgroundSnapshotReader<string> reader = new();
        var failed = await reader.ReadTaskAsync(async () =>
        { await Task.Yield(); throw new IOException("async provider failed"); }, new TaskCompletionSource().Task, _ => { });
        Assert.Null(failed.Snapshot);
        Assert.Equal("async provider failed", failed.Error);
        Assert.Equal("recovered", (await reader.ReadTaskAsync(() => Task.FromResult("recovered"), new TaskCompletionSource().Task, _ => { })).Snapshot);
    }
}
