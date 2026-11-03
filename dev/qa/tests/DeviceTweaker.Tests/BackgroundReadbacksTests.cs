using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public sealed class BackgroundReadbacksTests
{
    [Fact]
    public async Task CompletedReadDoesNotInvalidateAndAnActualProviderFailurePropagates()
    {
        TaskCompletionSource deadline = new();
        int invalidations = 0;
        Assert.Equal(BackgroundReadOutcome.Completed, await BackgroundReadbacks.WaitCoreAsync(
            Task.CompletedTask, deadline.Task, () => invalidations++, _ => throw new Exception("Unexpected late failure")));
        await Assert.ThrowsAsync<IOException>(() => BackgroundReadbacks.WaitCoreAsync(
            Task.FromException(new IOException("provider failed")), deadline.Task, () => invalidations++, _ => { }));
        Assert.Equal(0, invalidations);
    }

    [Fact]
    public async Task TimeoutInvalidatesBeforeReturningAndObservesAnEventualFailure()
    {
        TaskCompletionSource work = new(TaskCreationOptions.RunContinuationsAsynchronously);
        TaskCompletionSource<Exception> observed = new(TaskCreationOptions.RunContinuationsAsynchronously);
        int invalidations = 0;
        Assert.Equal(BackgroundReadOutcome.TimedOut, await BackgroundReadbacks.WaitCoreAsync(
            work.Task, Task.CompletedTask, () => invalidations++, error => observed.TrySetResult(error)));
        Assert.Equal(1, invalidations);
        work.SetException(new IOException("late failure"));
        Assert.Equal("late failure", (await observed.Task.WaitAsync(TimeSpan.FromSeconds(2))).Message);
    }

    [Fact]
    public async Task TimedOutReadCannotPublishIntoTheNextGeneration()
    {
        int generation = 1, published = 0;
        int requestedGeneration = generation;
        TaskCompletionSource release = new(TaskCreationOptions.RunContinuationsAsynchronously);
        async Task ReadAndPublish()
        {
            await release.Task;
            if (requestedGeneration == generation) published++;
        }
        Task work = ReadAndPublish();
        Assert.Equal(BackgroundReadOutcome.TimedOut, await BackgroundReadbacks.WaitCoreAsync(
            work, Task.CompletedTask, () => generation++, _ => { }));
        release.SetResult();
        await work;
        Assert.Equal(0, published);
    }
}
