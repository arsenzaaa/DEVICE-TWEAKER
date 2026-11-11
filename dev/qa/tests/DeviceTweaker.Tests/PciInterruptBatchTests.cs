using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public sealed class PciInterruptBatchTests
{
    [Fact]
    public void ReadsAreMaterializedInDeviceOrderIncludingRepeatedIdentities()
    {
        List<string> calls = new();
        bool sessionOpen = true;
        var readings = PciInterruptBatch.Read(["pci-a", "pci-b", "pci-a"], instance =>
        {
            Assert.True(sessionOpen);
            calls.Add(instance);
            return new(InterruptMode.Msi, instance, DateTimeOffset.Now);
        }, (_, _) => throw new InvalidOperationException());
        sessionOpen = false;
        Assert.Equal(new[] { "pci-a", "pci-b", "pci-a" }, calls);
        Assert.Equal(calls, readings.Select(reading => reading.InstanceId));
        Assert.All(readings, reading => Assert.Equal(InterruptMode.Msi, reading.Result.Mode));
    }

    [Fact]
    public void OneFailedDeviceDoesNotEraseOtherResultsOrTheirIdentities()
    {
        List<string> errors = new();
        var readings = PciInterruptBatch.Read(["good-a", "failed", "good-b"], instance =>
            instance == "failed" ? throw new IOException("device removed")
                : new(InterruptMode.MsiX, instance, DateTimeOffset.Now),
            (instance, error) => errors.Add(instance + ":" + error.Message));
        Assert.Equal(new[] { "failed:device removed" }, errors);
        Assert.Equal(InterruptMode.MsiX, readings[0].Result.Mode);
        Assert.Equal("failed", readings[1].InstanceId);
        Assert.Equal(InterruptMode.Unavailable, readings[1].Result.Mode);
        Assert.Equal("device removed", readings[1].Result.Detail);
        Assert.Equal(InterruptMode.MsiX, readings[2].Result.Mode);
    }

    [Fact]
    public void PreparationFailureProducesAnUnavailableResultForEachRequestedDevice()
    {
        var readings = PciInterruptBatch.Unavailable(["a", "b"], "load failed");
        Assert.Equal(new[] { "a", "b" }, readings.Select(reading => reading.InstanceId));
        Assert.All(readings, reading => {
            Assert.Equal(InterruptMode.Unavailable, reading.Result.Mode);
            Assert.Equal("load failed", reading.Result.Detail);
        });
    }
}
