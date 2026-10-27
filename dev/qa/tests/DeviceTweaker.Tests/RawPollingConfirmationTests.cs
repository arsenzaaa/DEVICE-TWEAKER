using DeviceTweakerCS;
using Xunit;
namespace DeviceTweaker.Tests;

public sealed class RawPollingConfirmationTests
{
    private static RawPollingState State() => new() { Role = "Mouse", InstanceId = "memory-device" };
    [Fact] public void MouseDowngradeNeedsRepeatedStableAndLiveConfirmation()
    {
        var s = State();
        Assert.False(RawPollingService.TryConfirmStable(s, 8000, "8K"));
        Assert.True(RawPollingService.TryConfirmStable(s, 8000, "8K"));
        for (int i = 0; i < 8; i++) Assert.False(RawPollingService.TryConfirmStable(s, 1000, "1K"));
        Assert.Equal("8K", s.LastDisplayedTag);
        for (int i = 0; i < 3; i++) Assert.False(RawPollingService.TryConfirmLive(s, 1000, "1K"));
        Assert.True(RawPollingService.TryConfirmLive(s, 1000, "1K"));
        Assert.True(RawPollingService.TryConfirmStable(s, 1000, "1K"));
        Assert.Equal("1K", s.LastDisplayedTag);
        Assert.False(RawPollingService.TryConfirmStable(s, 1000, "1K"));
    }
    [Fact] public void ChangingCandidateResetsConfirmationInsteadOfCombiningDifferentRates()
    {
        var s = State();
        Assert.False(RawPollingService.TryConfirmStable(s, 1000, "1K"));
        Assert.False(RawPollingService.TryConfirmStable(s, 8000, "8K"));
        Assert.False(RawPollingService.TryConfirmStable(s, 1000, "1K"));
        Assert.True(RawPollingService.TryConfirmStable(s, 1000, "1K"));
    }
    [Fact] public void LowLiveActivityCannotInventALowerStableKeyboardRate()
    {
        var s = new RawPollingState { Role = "Keyboard", InstanceId = "memory-keyboard" };
        RawPollingService.TryConfirmStable(s, 8000, "8K"); RawPollingService.TryConfirmStable(s, 8000, "8K");
        for (int i = 0; i < 4; i++) RawPollingService.TryConfirmLive(s, 10, "10Hz");
        for (int i = 0; i < 10; i++) Assert.False(RawPollingService.TryConfirmStable(s, 1000, "1K"));
        Assert.Equal("8K", s.LastDisplayedTag);
    }
}
