using System.Management;
using System.Runtime.InteropServices;
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

    [Fact]
    public void PhysicalDiskWmiFailuresDegradeWithoutThrowing()
    {
        bool wmiOk = WmiInterop.TryGetPhysicalDisksCore(
            () => throw new ManagementException("Invalid class"),
            out List<WmiPhysicalDisk> wmiDisks,
            out string? wmiError);

        Assert.False(wmiOk);
        Assert.Empty(wmiDisks);
        Assert.Equal("Invalid class", wmiError);

        bool comOk = WmiInterop.TryGetPhysicalDisksCore(
            () => throw new COMException("Invalid property", unchecked((int)0x80041031)),
            out List<WmiPhysicalDisk> comDisks,
            out string? comError);

        Assert.False(comOk);
        Assert.Empty(comDisks);
        Assert.Contains("Invalid property", comError);
    }

    [Fact]
    public void StorageFallbackKeepsNvmeAndSataControllersWhenStorageWmiIsUnavailable()
    {
        Assert.Equal("NVMe", MainForm.GetStorageTagForDevice("Standard NVM Express Controller", [], storageInventoryAvailable: false));
        Assert.Equal("SATA", MainForm.GetStorageTagForDevice("Standard SATA AHCI Controller", [], storageInventoryAvailable: false));
        Assert.Equal(string.Empty, MainForm.GetStorageTagForDevice("Microsoft Storage Spaces Controller", [], storageInventoryAvailable: false));

        Assert.Equal("NVMe", MainForm.GetStorageTagForDevice("Standard NVM Express Controller", [], storageInventoryAvailable: true));
        Assert.Equal(string.Empty, MainForm.GetStorageTagForDevice("Standard SATA AHCI Controller", [], storageInventoryAvailable: true));
    }
}
