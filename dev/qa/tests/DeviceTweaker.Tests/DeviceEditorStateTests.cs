using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public sealed class DeviceEditorStateTests
{
    [Fact]
    public void RefreshPreservesEditedFieldWhileAcceptingExternalChangesToUntouchedFields()
    {
        var baseline = new DeviceEditorState { MsiIndex = 0, LimitText = "Unlimited", AffinityMask = 4, PowerSaving = true };
        var draft = baseline with { LimitText = "8", MsiIndex = 1 };
        var latest = baseline with { AffinityMask = 16, PowerSaving = false };
        var merged = DeviceEditorState.MergeDraft(latest, baseline, draft);
        Assert.Equal("8", merged.LimitText);
        Assert.Equal(1, merged.MsiIndex);
        Assert.Equal(16UL, merged.AffinityMask);
        Assert.False(merged.PowerSaving);
    }

    [Fact]
    public void AnUneditedDraftNeverRestoresOldHardwareValues()
    {
        var baseline = new DeviceEditorState { ImodText = "0xC8", RssQueues = 2, NicItrText = "0x0" };
        var latest = baseline with { ImodText = "0x0", RssQueues = 4, NicItrText = "0x28" };
        Assert.Equal(latest, DeviceEditorState.MergeDraft(latest, baseline, baseline));
    }

    [Fact]
    public void NicAndRssDraftsSurviveIndependentlyOfApplyDirtyBadge()
    {
        var baseline = new DeviceEditorState { RssQueues = 2, NicItrText = "0x28" };
        var draft = baseline with { RssQueues = 4, NicItrText = "0x0" };
        var latest = baseline with { NicItrText = "0x40", PriorityIndex = 2 };
        var merged = DeviceEditorState.MergeDraft(latest, baseline, draft);
        Assert.Equal(4m, merged.RssQueues);
        Assert.Equal("0x0", merged.NicItrText);
        Assert.Equal(2, merged.PriorityIndex);
    }
}
