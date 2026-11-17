namespace DeviceTweaker.Tests;

using DeviceTweakerCS;
using Xunit;

public class NdisRssLayoutTests
{
    private static List<CpuLpInfo> AdjacentSmt(int cores)
        => Enumerable.Range(0, cores * 2).Select(lp => new CpuLpInfo(0, lp, lp / 2, 0, 0, 0)).ToList();

    private static List<CpuLpInfo> NoSmt(int cores)
        => Enumerable.Range(0, cores).Select(lp => new CpuLpInfo(0, lp, lp, 0, 0, 0)).ToList();

    [Fact]
    public void Select_WithSmt_UsesOneProcessorPerPhysicalCore()
    {
        NdisRssLayout.Layout layout = NdisRssLayout.Select(AdjacentSmt(16), 32, requestedBase: 12, queues: 4);

        Assert.Equal([12, 14, 16, 18], layout.Processors);
        Assert.Equal(12, layout.BaseProcessor);
        Assert.Equal(18, layout.MaxProcessor);
    }

    [Fact]
    public void Select_WithoutSmt_UsesConsecutiveProcessors()
    {
        NdisRssLayout.Layout layout = NdisRssLayout.Select(NoSmt(8), 8, requestedBase: 2, queues: 4);

        Assert.Equal([2, 3, 4, 5], layout.Processors);
        Assert.Equal(5, layout.MaxProcessor);
    }

    [Fact]
    public void Select_SnapsSmtSiblingBaseToItsCore()
    {
        NdisRssLayout.Layout layout = NdisRssLayout.Select(AdjacentSmt(8), 16, requestedBase: 13, queues: 2);

        Assert.Equal([12, 14], layout.Processors);
    }

    [Fact]
    public void Select_ShiftsBaseLeftWhenQueuesDoNotFit()
    {
        NdisRssLayout.Layout layout = NdisRssLayout.Select(AdjacentSmt(8), 16, requestedBase: 14, queues: 4);

        Assert.Equal([8, 10, 12, 14], layout.Processors);
    }

    [Fact]
    public void Select_HybridCpu_CrossesFromSmtPCoresToSingleThreadECores()
    {
        List<CpuLpInfo> lps = [.. AdjacentSmt(8), .. Enumerable.Range(16, 16).Select(lp => new CpuLpInfo(0, lp, lp - 8, 0, 0, 0))];

        NdisRssLayout.Layout layout = NdisRssLayout.Select(lps, 32, requestedBase: 14, queues: 3);

        Assert.Equal([14, 16, 17], layout.Processors);
    }

    [Fact]
    public void Select_NonAdjacentSiblings_UsesPrimaryThreads()
    {
        List<CpuLpInfo> lps = Enumerable.Range(0, 16).Select(lp => new CpuLpInfo(0, lp, lp % 8, 0, 0, 0)).ToList();

        NdisRssLayout.Layout layout = NdisRssLayout.Select(lps, 16, requestedBase: 2, queues: 4);

        Assert.Equal([2, 3, 4, 5], layout.Processors);
    }

    [Fact]
    public void Select_ClampsProcessorsToPhysicalCoreCount()
    {
        NdisRssLayout.Layout layout = NdisRssLayout.Select(AdjacentSmt(4), 8, requestedBase: 0, queues: 8);

        Assert.Equal([0, 2, 4, 6], layout.Processors);
        Assert.Equal(4, NdisRssLayout.GetMaxQueues(AdjacentSmt(4), 8));
        Assert.Equal(8, layout.QueueCount);
    }

    [Fact]
    public void Select_WithoutTopology_FallsBackToConsecutiveProcessors()
    {
        NdisRssLayout.Layout layout = NdisRssLayout.Select(null, 8, requestedBase: 3, queues: 2);

        Assert.Equal([3, 4], layout.Processors);
    }

    [Fact]
    public void AllowedQueues_UseDriverEnumIndependentlyOfProcessorCount()
    {
        PciInterruptCapabilities msix = new(PciInterruptSupport.Msi | PciInterruptSupport.MsiX, 5);

        Assert.Equal([1, 2, 4], NdisRssLayout.GetAllowedQueueCounts([4, 1, 2], msix, maxQueues: 8));
        Assert.Equal([1, 2, 4], NdisRssLayout.GetAllowedQueueCounts([1, 2, 4], msix, maxQueues: 2));
    }

    [Theory]
    [InlineData(1)]
    [InlineData(8)]
    public void AllowedQueues_WithoutMsiX_RetainDriverAdvertisedReceiveQueues(int physicalCores)
    {
        // Intel 82579/I217/I218/I219 support two receive queues on one CPU core.
        // Queue count and the interrupt-processing CPU count are separate limits.
        PciInterruptCapabilities msiOnly = new(PciInterruptSupport.LineBased | PciInterruptSupport.Msi, 1);
        Assert.Equal([1, 2], NdisRssLayout.GetAllowedQueueCounts([1, 2], msiOnly, physicalCores));
    }

    [Fact]
    public void AllowedQueues_WithoutMsiX_WithoutDriverOptions_AreSingleQueue()
    {
        PciInterruptCapabilities msiOnly = new(PciInterruptSupport.LineBased | PciInterruptSupport.Msi, 1);

        Assert.Equal([1], NdisRssLayout.GetAllowedQueueCounts(null, msiOnly, maxQueues: 8));
    }

    [Fact]
    public void Select_I219_TwoQueuesUseOnePhysicalCore()
    {
        var adapter = NicItrService.TryGetNicItrProfile(@"PCI\VEN_8086&DEV_15B8");
        Assert.NotNull(adapter);
        NdisRssLayout.Layout layout = NdisRssLayout.Select(AdjacentSmt(8), 16, 9, 2, adapter.RssProcessorLimit);
        Assert.Equal(2, layout.QueueCount);
        Assert.Equal([8], layout.Processors);
        Assert.Equal(8, layout.BaseProcessor);
        Assert.Equal(8, layout.MaxProcessor);
    }

    [Theory]
    [InlineData("PCI\\VEN_8086&DEV_15B8", 1)]
    [InlineData("PCI\\VEN_8086&DEV_153A", 1)]
    [InlineData("PCI\\VEN_8086&DEV_125C", null)]
    [InlineData("PCI\\VEN_10EC&DEV_8125", null)]
    public void ProcessorLimit_IsDeviceSpecific(string instance, int? expected)
    {
        Assert.Equal(expected, NicItrService.TryGetNicItrProfile(instance)?.RssProcessorLimit);
    }

    [Fact]
    public void Select_GenericSharedInterrupt_DoesNotImposeI219CpuLimit()
    {
        PciInterruptCapabilities msiOnly = new(PciInterruptSupport.Msi, 1);
        Assert.Equal([1, 2], NdisRssLayout.GetAllowedQueueCounts([1, 2], msiOnly, 8));
        var layout = NdisRssLayout.Select(AdjacentSmt(8), 16, 8, 2);
        Assert.Equal(2, layout.QueueCount);
        Assert.Equal([8, 10], layout.Processors);
    }

    [Fact]
    public void Select_WithSmt_DoesNotDoubleQueuesForSiblingThreads()
    {
        var single = NdisRssLayout.Select(AdjacentSmt(8), 16, 8, 1);
        var pair = NdisRssLayout.Select(AdjacentSmt(8), 16, 8, 2);
        Assert.Equal(1, single.QueueCount);
        Assert.Equal([8], single.Processors);
        Assert.Equal(2, pair.QueueCount);
        Assert.Equal([8, 10], pair.Processors);
    }

    [Fact]
    public void AllowedQueues_LimitedByMsiXMessages()
    {
        PciInterruptCapabilities msix = new(PciInterruptSupport.MsiX, 3);

        Assert.Equal([1, 2, 3], NdisRssLayout.GetAllowedQueueCounts(null, msix, maxQueues: 8));
    }

    [Fact]
    public void AllowedQueues_UnknownCapabilitiesKeepDriverValues()
    {
        Assert.Equal([1, 2, 4], NdisRssLayout.GetAllowedQueueCounts([1, 2, 4], null, maxQueues: 8));
        Assert.Equal([1, 2, 3], NdisRssLayout.GetAllowedQueueCounts([], null, maxQueues: 3));
    }

    [Theory]
    [InlineData(2, 3, 4)]
    [InlineData(4, 3, 2)]
    [InlineData(4, 5, 4)]
    [InlineData(1, 0, 1)]
    [InlineData(2, 2, 2)]
    public void Snap_FollowsStepDirection(int previous, int requested, int expected)
    {
        Assert.Equal(expected, NdisRssLayout.Snap([1, 2, 4], previous, requested));
    }

    [Theory]
    [InlineData(3, 2)]
    [InlineData(8, 4)]
    [InlineData(0, 1)]
    public void Floor_PicksLargestAllowedNotAbove(int value, int expected)
    {
        Assert.Equal(expected, NdisRssLayout.Floor([1, 2, 4], value));
    }

    [Fact]
    public void Select_IgnoresOtherProcessorGroups()
    {
        List<CpuLpInfo> lps = [.. NoSmt(4), new CpuLpInfo(1, 0, 100, 0, 0, 0)];

        Assert.Equal(4, NdisRssLayout.GetMaxQueues(lps, 64));
    }
}
