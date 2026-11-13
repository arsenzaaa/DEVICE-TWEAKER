using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public sealed class RawPollingEstimateCacheTests
{
    private static RawPollingState Filled(double interval = 1d)
    {
        RawPollingState state = new() { Role = "Mouse", InstanceId = "synthetic" };
        for (int i = 0; i < 256; i++) state.RecordInterval(interval);
        return state;
    }

    [Fact]
    public void UnchangedBuffersReuseEstimatesWithoutAllocatingSampleArrays()
    {
        RawPollingState state = Filled();
        for (int i = 0; i < 32; i++) { state.TryEstimateLive(out _); state.TryEstimateStable(out _); }
        long before = GC.GetAllocatedBytesForCurrentThread();
        for (int i = 0; i < 100; i++) { state.TryEstimateLive(out _); state.TryEstimateStable(out _); }
        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;
        Assert.True(allocated < 512, $"Unchanged buffers allocated {allocated} bytes.");
    }

    [Fact]
    public void ReplacementAtCapacityInvalidatesBothEstimates()
    {
        RawPollingState state = Filled();
        Assert.True(state.TryEstimateLive(out double live)); Assert.Equal(1000, live);
        Assert.True(state.TryEstimateStable(out double stable)); Assert.Equal(1000, stable);
        for (int i = 0; i < 256; i++) state.RecordInterval(.125d);
        Assert.True(state.TryEstimateLive(out live)); Assert.Equal(8000, live);
        Assert.True(state.TryEstimateStable(out stable)); Assert.Equal(8000, stable);
    }

    [Fact]
    public void PauseClearsCachedResultsAndSampleThresholdsStillApply()
    {
        RawPollingState state = Filled();
        state.TryEstimateLive(out _); state.TryEstimateStable(out _);
        state.RecordInterval(501);
        Assert.False(state.TryEstimateLive(out _)); Assert.False(state.TryEstimateStable(out _));
        for (int i = 0; i < 11; i++) state.RecordInterval(1);
        Assert.False(state.TryEstimateLive(out _));
        state.RecordInterval(1);
        Assert.True(state.TryEstimateLive(out _)); Assert.False(state.TryEstimateStable(out _));
        for (int i = 12; i < 32; i++) state.RecordInterval(1);
        Assert.True(state.TryEstimateStable(out _));
    }

    [Fact]
    public void CachedReadingsKeepConfirmationOnEveryTimerTick()
    {
        RawPollingState state = Filled();
        for (int tick = 1; tick <= 8; tick++)
        {
            Assert.True(state.TryEstimateLive(out double live));
            Assert.Equal(tick == 4, RawPollingService.TryConfirmLive(state, live, "1K"));
            Assert.True(state.TryEstimateStable(out double stable));
            Assert.Equal(tick == 2, RawPollingService.TryConfirmStable(state, stable, "1K"));
        }
        Assert.Equal(8, state.CandidateCount); Assert.Equal(8, state.LiveCandidateCount);
    }

    [Fact]
    public void MixedInputMatchesUncachedAlgorithmIncludingBoundariesAndPauses()
    {
        RawPollingState state = Filled();
        Queue<double> stable = new(Enumerable.Repeat(1d, 256));
        Queue<double> live = new(Enumerable.Repeat(1d, 128));
        double[] intervals = [.01, .05, .125, .25, 1, 8, 12.5, 12.5001, 20, 200, 200.001, 500, 501, double.NaN];
        Random random = new(714);
        for (int i = 0; i < 1000; i++)
        {
            double interval = intervals[random.Next(intervals.Length)];
            state.RecordInterval(interval);
            if (interval > 500) { live.Clear(); stable.Clear(); }
            else if (interval >= .05 && interval <= 200)
            {
                live.Enqueue(interval); if (live.Count > 128) live.Dequeue();
                if (interval <= 12.5) { stable.Enqueue(interval); if (stable.Count > 256) stable.Dequeue(); }
            }
            for (int read = 0; read < 2; read++)
            {
                Assert.Equal(RawPollingService.TryEstimateLivePollingHz(live, out double expectedLive), state.TryEstimateLive(out double actualLive));
                Assert.Equal(expectedLive, actualLive);
                Assert.Equal(RawPollingService.TryEstimateRawPollingHz(stable, out double expectedStable), state.TryEstimateStable(out double actualStable));
                Assert.Equal(expectedStable, actualStable);
            }
        }
    }
}
