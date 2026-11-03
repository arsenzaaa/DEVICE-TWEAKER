using Xunit;

namespace DeviceTweakerCS.Tests;

public sealed class DeviceOperationCoordinatorTests
{
    [Fact]
    public async Task LeaseCoversTheActualCompletionOfAsyncWork()
    {
        DeviceOperationCoordinator coordinator = new();
        TaskCompletionSource finish = new();
        async Task WriteAsync()
        {
            using IDisposable lease = coordinator.TryEnter()!;
            await finish.Task;
        }
        Task operation = WriteAsync();
        Assert.True(coordinator.IsBusy);
        Assert.Null(coordinator.TryEnter());
        finish.SetResult();
        await operation;
        Assert.False(coordinator.IsBusy);
        using IDisposable? retry = coordinator.TryEnter();
        Assert.NotNull(retry);
    }

    [Fact]
    public async Task FailureReleasesTheLease()
    {
        DeviceOperationCoordinator coordinator = new();
        await Assert.ThrowsAsync<IOException>(async () =>
        {
            using IDisposable lease = coordinator.TryEnter()!;
            await Task.Yield();
            throw new IOException("write failed");
        });
        Assert.False(coordinator.IsBusy);
    }

    [Fact]
    public void OldLeaseCannotReleaseANewerOperation()
    {
        DeviceOperationCoordinator coordinator = new();
        IDisposable first = coordinator.TryEnter()!;
        first.Dispose();
        using IDisposable second = coordinator.TryEnter()!;
        first.Dispose();
        Assert.True(coordinator.IsBusy);
        Assert.Null(coordinator.TryEnter());
    }
}
