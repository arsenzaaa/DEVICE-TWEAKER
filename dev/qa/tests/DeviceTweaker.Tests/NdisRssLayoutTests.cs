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
    public void Select_ClampsQueuesToPhysicalCoreCount()
    {
        NdisRssLayout.Layout layout = NdisRssLayout.Select(AdjacentSmt(4), 8, requestedBase: 0, queues: 8);

        Assert.Equal([0, 2, 4, 6], layout.Processors);
        Assert.Equal(4, NdisRssLayout.GetMaxQueues(AdjacentSmt(4), 8));
    }

    [Fact]
    public void Select_WithoutTopology_FallsBackToConsecutiveProcessors()
    {
        NdisRssLayout.Layout layout = NdisRssLayout.Select(null, 8, requestedBase: 3, queues: 2);

        Assert.Equal([3, 4], layout.Processors);
    }

    [Fact]
    public void AllowedQueues_UseDriverEnumWithinCoreLimit()
    {
        PciInterruptCapabilities msix = new(PciInterruptSupport.Msi | PciInterruptSupport.MsiX, 5);

        Assert.Equal([1, 2, 4], NdisRssLayout.GetAllowedQueueCounts([4, 1, 2], msix, maxQueues: 8));
        Assert.Equal([1, 2], NdisRssLayout.GetAllowedQueueCounts([1, 2, 4], msix, maxQueues: 2));
    }

    [Fact]
    public void AllowedQueues_WithoutMsiX_AreSingleQueue()
    {
        PciInterruptCapabilities msiOnly = new(PciInterruptSupport.LineBased | PciInterruptSupport.Msi, 1);

        Assert.Equal([1], NdisRssLayout.GetAllowedQueueCounts([1, 2], msiOnly, maxQueues: 8));
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
