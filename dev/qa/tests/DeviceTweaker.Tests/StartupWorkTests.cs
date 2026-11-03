using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public class StartupWorkTests
{
    [Fact]
    public void InFlightProfileOwnsItsUsbAndNicValues()
    {
        MainForm.ImodConfig source = new()
        {
            GlobalInterval = 200, GlobalHcsparamsOffset = 8, GlobalRtsoff = 24,
            HasScript = true, StartupReadError = "unreadable sentinel",
        };
        source.Overrides.Add(new MainForm.ImodConfigEntry
        {
            Hwid = "USB-A", Enabled = false, Interval = 100,
            HcsparamsOffset = 4, Rtsoff = 32, AdaptiveRoleBinding = true,
            Intervals = [1, 2], RoleIntervals = new(StringComparer.OrdinalIgnoreCase) { ["Mouse"] = 3 },
        });
        source.NicItrEntries.Add(new MainForm.NicItrConfigEntry
        {
            Hwid = "NIC-A", FamilyName = "family", BaseOffset = 8, Stride = 4,
            Queues = 2, Width = 4, Mask = 0xFFFF, OrBits = 0x8000, Values = [10, 20],
        });

        var work = MainForm.SnapshotImodConfig(source);
        source.GlobalInterval = 400;
        source.Overrides[0].Intervals![0] = 999;
        source.Overrides[0].RoleIntervals!["Mouse"] = 999;
        source.NicItrEntries[0].Values[0] = 999;
        source.Overrides.Clear();
        source.NicItrEntries.Clear();

        Assert.Equal(200U, work.GlobalInterval);
        Assert.Equal(8U, work.GlobalHcsparamsOffset);
        Assert.Equal(24U, work.GlobalRtsoff);
        Assert.True(work.HasScript);
        Assert.Equal("unreadable sentinel", work.StartupReadError);
        var usb = Assert.Single(work.Overrides);
        Assert.Equal("USB-A", usb.Hwid);
        Assert.False(usb.Enabled);
        Assert.True(usb.AdaptiveRoleBinding);
        Assert.Equal(100U, usb.Interval);
        Assert.Equal(4U, usb.HcsparamsOffset);
        Assert.Equal(32U, usb.Rtsoff);
        Assert.Equal(new uint[] { 1, 2 }, usb.Intervals);
        Assert.Equal(3U, usb.RoleIntervals!["mouse"]);
        var nic = Assert.Single(work.NicItrEntries);
        Assert.Equal("NIC-A", nic.Hwid);
        Assert.Equal("family", nic.FamilyName);
        Assert.Equal(8U, nic.BaseOffset);
        Assert.Equal(4U, nic.Stride);
        Assert.Equal(2, nic.Queues);
        Assert.Equal(4, nic.Width);
        Assert.Equal(0xFFFFUL, nic.Mask);
        Assert.Equal(0x8000UL, nic.OrBits);
        Assert.Equal(new ulong[] { 10, 20 }, nic.Values);
    }

    [Fact]
    public void SnapshotRetainsAbsentOverridesWithoutInventingSettings()
    {
        MainForm.ImodConfig source = new();
        source.Overrides.Add(new MainForm.ImodConfigEntry { Hwid = "USB-A" });
        var work = MainForm.SnapshotImodConfig(source);
        Assert.Null(work.Overrides[0].Intervals);
        Assert.Null(work.Overrides[0].RoleIntervals);
        Assert.Null(work.Overrides[0].Interval);
        Assert.Null(work.StartupReadError);
        Assert.False(work.HasScript);
    }
}
