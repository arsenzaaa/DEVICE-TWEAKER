using Xunit;

namespace DeviceTweakerCS.Tests;

public sealed class CardColumnsTests
{
    [Fact]
    public void WideCardKeepsAffinityAndSettingsSideBySide()
    {
        CardColumns layout = CardColumns.Measure(1400, 1f, 600, 460);
        Assert.False(layout.Stacked);
        Assert.True(layout.SettingsLeft >= 16 + layout.AffinityWidth + 20);
    }

    [Fact]
    public void LongDualCcdListUsesTheFullWidthBeforeAddingHorizontalScroll()
    {
        CardColumns layout = CardColumns.Measure(1000, 1f, 800, 460);
        Assert.True(layout.Stacked);
        Assert.True(layout.AffinityWidth >= 800);
        Assert.Equal(16, layout.SettingsLeft);
    }

    [Fact]
    public void SupportedScaleAndWidthCombinationsKeepBothSectionsInsideTheCard()
    {
        foreach (float scale in new[] { 1f, 1.25f, 1.5f })
        foreach (int width in new[] { 960, 1148, 1560 })
        foreach (int affinity in new[] { 400, 650, 900 })
        {
            int S(int value) => (int)Math.Round(value * scale);
            CardColumns layout = CardColumns.Measure(S(width), scale, S(affinity), S(460));
            Assert.InRange(S(16) + layout.AffinityWidth, 1, S(width) - S(24));
            Assert.InRange(layout.SettingsLeft + layout.SettingsWidth, 1, S(width) - S(24));
            if (!layout.Stacked) Assert.True(layout.SettingsLeft >= S(16) + layout.AffinityWidth + S(20));
        }
    }
}
