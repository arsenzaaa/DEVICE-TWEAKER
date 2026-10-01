using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public class NicItrTests
{
    [Theory]
    [InlineData(@"PCI\VEN_8086&DEV_15F3&SUBSYS_00000000", "Intel I225/I226 (EITR)", "8086")]
    [InlineData(@"PCI\VEN_8086&DEV_15F2&SUBSYS_00000000", "Intel I225/I226 (EITR)", "8086")]
    [InlineData(@"PCI\VEN_8086&DEV_125B&SUBSYS_00000000", "Intel I225/I226 (EITR)", "8086")]
    [InlineData(@"PCI\VEN_8086&DEV_15F7&SUBSYS_00000000", "Intel I225/I226 (EITR)", "8086")]
    [InlineData(@"PCI\VEN_8086&DEV_15F8&SUBSYS_00000000", "Intel I225/I226 (EITR)", "8086")]
    [InlineData(@"PCI\VEN_8086&DEV_1539&SUBSYS_00000000", "Intel I210/I211 (EITR)", "8086")]
    [InlineData(@"PCI\VEN_8086&DEV_1521&SUBSYS_00000000", "Intel I350 (EITR)", "8086")]
    [InlineData(@"PCI\VEN_10EC&DEV_8125&SUBSYS_00000000", "Realtek RTL8125/8126", "10EC")]
    [InlineData(@"PCI\VEN_10EC&DEV_8126&SUBSYS_00000000", "Realtek RTL8125/8126", "10EC")]
    [InlineData(@"PCI\VEN_10EC&DEV_8168&SUBSYS_00000000", "Realtek RTL8111/8168", "10EC")]
    [InlineData(@"PCI\VEN_10EC&DEV_2600&SUBSYS_00000000", "Killer E2500/E2600", "10EC")]
    public void TryGetNicItrProfile_RecognizesSupportedDevices(string instanceId, string expectedFamily, string expectedVendor)
    {
        var profile = MainForm.TryGetNicItrProfile(instanceId);
        Assert.NotNull(profile);
        Assert.Equal(expectedFamily, profile.FamilyName);
        Assert.Equal(expectedVendor, profile.VendorId, ignoreCase: true);
        Assert.True(profile.BaseOffset > 0);
        Assert.True(profile.MaxQueues >= 1);
        Assert.True(profile.ReadWidth is 16 or 32);
    }

    [Theory]
    [InlineData(@"PCI\VEN_9999&DEV_9999&SUBSYS_00000000")]
    [InlineData(@"USB\VID_046D&PID_C08B")]
    [InlineData("")]
    public void TryGetNicItrProfile_ReturnsNullForUnsupportedDevices(string instanceId)
    {
        var profile = MainForm.TryGetNicItrProfile(instanceId);
        Assert.Null(profile);
    }
}
