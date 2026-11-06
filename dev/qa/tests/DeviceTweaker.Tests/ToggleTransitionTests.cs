using Xunit;

namespace DeviceTweakerCS.Tests;

public sealed class ToggleTransitionTests
{
    [Fact]
    public void MotionIsBoundedMonotonicAndFinishesAtItsDeadline()
    {
        ToggleTransition motion = new();
        motion.SetTarget(true, 1000, animate: true);
        float previous = 0;
        foreach (int elapsed in new[] { 0, 7, 31, 60, 92, 119, 120 })
        {
            float position = motion.Position(1000 + elapsed);
            Assert.InRange(position, previous, 1);
            previous = position;
        }
        Assert.Equal(1, previous);
        Assert.False(motion.IsRunning);
    }

    [Fact]
    public void RapidReversalStartsFromTheVisiblePosition()
    {
        ToggleTransition motion = new();
        motion.SetTarget(true, 0, animate: true);
        float visible = motion.Position(40);
        motion.SetTarget(false, 40, animate: true);
        Assert.Equal(visible, motion.Position(40));
        Assert.InRange(motion.Position(80), 0, visible);
        Assert.Equal(0, motion.Position(160));
        Assert.False(motion.IsRunning);
    }

    [Fact]
    public void ReducedMotionCompletesImmediatelyEvenMidTransition()
    {
        ToggleTransition motion = new();
        motion.SetTarget(true, 0, animate: true);
        motion.SetTarget(false, 40, animate: false);
        Assert.Equal(0, motion.Position(40));
        Assert.False(motion.IsRunning);
    }

    [Fact]
    public void ADelayedPaintDoesNotNeedMissedTimerTicksToCatchUp()
    {
        ToggleTransition motion = new();
        motion.SetTarget(true, 0, animate: true);
        Assert.Equal(1, motion.Position(900));
        Assert.False(motion.IsRunning);
    }

    [Fact]
    public void RepeatedTargetDoesNotRestartTheDeadline()
    {
        ToggleTransition motion = new();
        motion.SetTarget(true, 0, animate: true);
        motion.SetTarget(true, 50, animate: true);
        Assert.Equal(1, motion.Position(120));
        Assert.False(motion.IsRunning);
    }
}
