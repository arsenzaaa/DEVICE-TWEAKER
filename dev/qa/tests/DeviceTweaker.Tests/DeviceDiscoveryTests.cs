using Xunit;

namespace DeviceTweakerCS.Tests;

public sealed class DeviceDiscoveryTests
{
    [Fact]
    public async Task TimedOutProviderKeepsItsSlotUntilItActuallyFinishes()
    {
        DeviceDiscovery reader = new();
        TaskCompletionSource<IReadOnlyList<DeviceInfo>> pending = new();
        TaskCompletionSource deadline = new();
        Task<DeviceDiscoveryResult> first = reader.ReadAsync(() => pending.Task, deadline.Task, _ => { });
        deadline.SetResult();
        Assert.Null((await first).Devices);
        bool secondCalled = false;
        DeviceDiscoveryResult busy = await reader.ReadAsync(() =>
        {
            secondCalled = true;
            return Task.FromResult<IReadOnlyList<DeviceInfo>>([]);
        }, Task.CompletedTask, _ => { });
        Assert.False(secondCalled);
        Assert.Contains("still running", busy.Error);
        pending.SetResult([]);
        DeviceDiscoveryResult retry = await reader.ReadAsync(() => pending.Task, Task.CompletedTask, _ => { });
        Assert.NotNull(retry.Devices);
        Assert.Empty(retry.Devices);
        Assert.Null(retry.Error);
    }

    [Fact]
    public async Task ProviderFailureIsNotReportedAsAnEmptySuccessfulScan()
    {
        DeviceDiscoveryResult result = await new DeviceDiscovery().ReadAsync(
            () => Task.FromException<IReadOnlyList<DeviceInfo>>(new IOException("provider unavailable")),
            new TaskCompletionSource().Task, _ => { });
        Assert.Null(result.Devices);
        Assert.Equal("provider unavailable", result.Error);
    }
}
