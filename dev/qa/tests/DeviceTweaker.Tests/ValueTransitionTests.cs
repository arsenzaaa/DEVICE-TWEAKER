using Xunit;

namespace DeviceTweakerCS.Tests;

public sealed class ValueTransitionTests
{
    [Fact]
    public void ScrollRetargetsFromItsCurrentPositionWithoutOvershooting()
    {
        ValueTransition motion = new(160);
        motion.MoveTo(100, 0, true);
        float visible = motion.Position(40);
        Assert.InRange(visible, 1, 99);
        motion.MoveTo(-20, 40, true);
        Assert.Equal(visible, motion.Position(40));
        float previous = visible;
        foreach (int now in new[] { 41, 70, 100, 150, 200 })
        {
            float value = motion.Position(now);
            Assert.InRange(value, -20, previous);
            previous = value;
        }
        Assert.Equal(-20, previous);
        Assert.False(motion.IsRunning);
    }

    [Fact]
    public void RepeatingDestinationDoesNotPostponeCompletion()
    {
        ValueTransition motion = new(160);
        motion.MoveTo(100, 0, true);
        motion.MoveTo(100, 70, true);
        Assert.Equal(100, motion.Position(160));
        Assert.False(motion.IsRunning);
    }

    [Fact]
    public void ImmediateRangeClampCancelsTheOldDestination()
    {
        ValueTransition motion = new(160);
        motion.MoveTo(1000, 0, true);
        motion.MoveTo(30, 40, false);
        Assert.Equal(30, motion.Position(40));
        Assert.Equal(30, motion.Position(1000));
        Assert.False(motion.IsRunning);
    }

    [Fact]
    public void ADelayedFrameCatchesUpWithoutReplayingQueuedTicks()
    {
        ValueTransition motion = new(100);
        motion.MoveTo(1, 0, true);
        Assert.Equal(1, motion.Position(3000));
        Assert.False(motion.IsRunning);
    }
}
