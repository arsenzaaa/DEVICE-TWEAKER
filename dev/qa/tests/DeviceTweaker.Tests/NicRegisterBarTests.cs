using System.Text.Json;
using Xunit;
using Range = DeviceTweakerImodRuntime.PciMemoryRange;

namespace DeviceTweakerCS.Tests;

// The NIC ITR register window comes from the first memory BAR in PCI config space,
// never from the largest Windows range. Models only: no driver, MMIO or hardware access.
public sealed class NicRegisterBarTests
{
    private sealed record Capture(string Name, string InstanceId, string ExpectedMode, string ConfigBase64, string ResourcesBase64);

    private static byte[] Config(ushort vendor, ushort device, params (int Bar, uint Value)[] bars)
    {
        byte[] config = new byte[256];
        BitConverter.GetBytes(vendor).CopyTo(config, 0);
        BitConverter.GetBytes(device).CopyTo(config, 2);
        foreach ((int bar, uint value) in bars) BitConverter.GetBytes(value).CopyTo(config, 0x10 + (bar * 4));
        return config;
    }

    // MSI-X capability at 0x50: table and PBA in one BAR (BIR in the low three bits of the offsets).
    private static byte[] WithMsiX(byte[] config, int bir, uint tableOffset, uint pbaOffset, int entries)
    {
        BitConverter.GetBytes((ushort)0x0010).CopyTo(config, 6);
        config[0x34] = 0x50;
        config[0x50] = 0x11;
        config[0x51] = 0;
        BitConverter.GetBytes((ushort)(entries - 1)).CopyTo(config, 0x52);
        BitConverter.GetBytes(tableOffset | (uint)bir).CopyTo(config, 0x54);
        BitConverter.GetBytes(pbaOffset | (uint)bir).CopyTo(config, 0x58);
        return config;
    }

    private static bool Select(byte[] config, string instanceId, Range[] ranges, uint offset, uint span,
        out ulong baseAddress, out ulong length, out string detail, out string error)
        => DeviceTweakerImodRuntime.TrySelectNicRegisterRange(config, instanceId, ranges, offset, span,
            out baseAddress, out length, out detail, out error);

    [Fact]
    public void QueryPciIoctlMatchesTheDriverAndTheAppReader()
    {
        Assert.Equal((0x8010u << 16) | (1u << 14) | (0x814u << 2), DeviceTweakerImodRuntime.IoctlImodQueryPci);
    }

    [Fact]
    public void Rtl8125CaptureKeepsTheRegisterBarThatWasAlreadyUsed()
    {
        Capture[] captures = JsonSerializer.Deserialize<Capture[]>(File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "WindowsResourceCaptures.json")))!;
        Capture nic = captures.Single(c => c.Name == "NIC");
        Range[] ranges = [new(0x856F0000, 0x10000), new(0x856EC000, 0x4000)];

        Assert.True(Select(Convert.FromBase64String(nic.ConfigBase64), nic.InstanceId, ranges, 0x0A00, 0x1C,
            out ulong baseAddress, out ulong length, out string detail, out string error), error);

        Assert.Equal((0x856F0000UL, 0x10000UL), (baseAddress, length));
        Assert.Contains("selected=BAR2 0x856F0000+0x10000", detail);
        Assert.Contains("(same)", detail);
        Assert.Contains("BAR4 mem64", detail);
    }

    [Fact]
    public void Rtl8168UsesTheSmallRegisterBarInsteadOfTheLargerMsiXBar()
    {
        // RTL8111/8168: BAR0 I/O, BAR2 64-bit 4 KB registers, BAR4 64-bit prefetchable 16 KB MSI-X.
        byte[] config = WithMsiX(Config(0x10EC, 0x8168, (0, 0xE001), (2, 0xF7100004), (4, 0xF000000C)),
            bir: 4, tableOffset: 0x0, pbaOffset: 0x800, entries: 4);
        Range[] ranges = [new(0xF7100000, 0x1000), new(0xF0000000, 0x4000)];

        Assert.True(Select(config, @"PCI\VEN_10EC&DEV_8168&SUBSYS_00000000&REV_15\4&1&00E4", ranges, 0xE2, 2,
            out ulong baseAddress, out ulong length, out string detail, out string error), error);

        Assert.Equal((0xF7100000UL, 0x1000UL), (baseAddress, length));
        Assert.Contains("largest=0xF0000000+0x4000 (differs; not used)", detail);
        Assert.Contains("msix=[table BAR4", detail);
        Assert.Contains("BAR0 io 0xE000", detail);
    }

    [Fact]
    public void ConventionalPciRtl8169UsesMemoryBarOne()
    {
        byte[] config = Config(0x10EC, 0x8169, (0, 0xD001), (1, 0xFE000000));
        Range[] ranges = [new(0xFE000000, 0x100)];

        Assert.True(Select(config, @"PCI\VEN_10EC&DEV_8169&SUBSYS_0\0", ranges, 0xE2, 2,
            out ulong baseAddress, out _, out _, out string error), error);
        Assert.Equal(0xFE000000UL, baseAddress);
    }

    [Fact]
    public void IntelUsesBarZeroAndIgnoresALargerFlashOrMsiXBar()
    {
        // 82576-style: BAR0 128 KB registers, BAR1 flash 4 MB, BAR3 16 KB MSI-X.
        byte[] config = WithMsiX(Config(0x8086, 0x10C9, (0, 0xFB000000), (1, 0xFA000000), (2, 0xE001), (3, 0xFB100000)),
            bir: 3, tableOffset: 0x0, pbaOffset: 0x2000, entries: 10);
        Range[] ranges = [new(0xFB000000, 0x20000), new(0xFA000000, 0x400000), new(0xFB100000, 0x4000)];

        Assert.True(Select(config, @"PCI\VEN_8086&DEV_10C9\0", ranges, 0x1680, 4 * 9 + 4,
            out ulong baseAddress, out ulong length, out string detail, out string error), error);
        Assert.Equal((0xFB000000UL, 0x20000UL), (baseAddress, length));
        Assert.Contains("(differs; not used)", detail);
    }

    [Fact]
    public void BarMissingFromWindowsRangesIsRefusedInsteadOfGuessed()
    {
        byte[] config = Config(0x10EC, 0x8168, (2, 0xF7100004));
        Range[] ranges = [new(0xF0000000, 0x4000)];

        Assert.False(Select(config, @"PCI\VEN_10EC&DEV_8168\0", ranges, 0xE2, 2, out ulong baseAddress, out _, out string detail, out string error));
        Assert.Equal(0UL, baseAddress);
        Assert.Contains("not among the memory ranges Windows allocated", error);
        Assert.Contains("ranges=[0xF0000000+0x4000]", detail);
    }

    [Fact]
    public void RegistersInsideTheMsiXTableAreRefused()
    {
        byte[] config = WithMsiX(Config(0x10EC, 0x8168, (2, 0xF7100004)), bir: 2, tableOffset: 0x0, pbaOffset: 0x800, entries: 16);
        Range[] ranges = [new(0xF7100000, 0x1000)];

        Assert.False(Select(config, @"PCI\VEN_10EC&DEV_8168\0", ranges, 0xE2, 2, out _, out _, out _, out string error));
        Assert.Contains("overlap the MSI-X table", error);
    }

    [Fact]
    public void RegistersPastTheBarEndAreRefused()
    {
        byte[] config = Config(0x10EC, 0x8125, (2, 0xF7100004));
        Range[] ranges = [new(0xF7100000, 0x100)];

        Assert.False(Select(config, @"PCI\VEN_10EC&DEV_8125\0", ranges, 0x0A00, 0x1C, out _, out _, out _, out string error));
        Assert.Contains("exceed BAR2", error);
    }

    [Fact]
    public void ConfigurationOfAnotherDeviceIsRefused()
    {
        byte[] config = Config(0x8086, 0x15F3, (0, 0xFB000000));
        Range[] ranges = [new(0xFB000000, 0x100000)];

        Assert.False(Select(config, @"PCI\VEN_10EC&DEV_8168\0", ranges, 0xE2, 2, out _, out _, out _, out string error));
        Assert.Contains("does not belong", error);
    }

    [Fact]
    public void DeviceWithOnlyIoBarsIsRefused()
    {
        byte[] config = Config(0x10EC, 0x8168, (0, 0xE001));

        Assert.False(Select(config, @"PCI\VEN_10EC&DEV_8168\0", [new(0xF0000000, 0x1000)], 0xE2, 2, out _, out _, out _, out string error));
        Assert.Contains("no memory BAR", error);
    }

    [Fact]
    public void CyclicCapabilityListIsRefused()
    {
        byte[] config = Config(0x10EC, 0x8168, (2, 0xF7100004));
        BitConverter.GetBytes((ushort)0x0010).CopyTo(config, 6);
        config[0x34] = 0x50;
        config[0x50] = 0x01;
        config[0x51] = 0x50;

        Assert.False(Select(config, @"PCI\VEN_10EC&DEV_8168\0", [new(0xF7100000, 0x1000)], 0xE2, 2, out _, out _, out _, out string error));
        Assert.Contains("capability list", error);
    }

    [Fact]
    public void RegisterSpanCoversEveryQueue()
    {
        NicItrProfile rtl8125 = NicItrService.NicItrProfiles.Single(p => p.FamilyName == "Realtek RTL8125/8126");
        NicItrProfile rtl8168 = NicItrService.NicItrProfiles.Single(p => p.FamilyName == "Realtek RTL8111/8168");

        Assert.Equal(3u * 8u + 4u, NicItrService.GetNicItrRegisterSpan(rtl8125));
        Assert.Equal(2u, NicItrService.GetNicItrRegisterSpan(rtl8168));
    }
}

// Device IDs follow Realtek's INF (rt640x64.inf); a saved entry with an outdated layout is not applied.
public sealed class NicItrProfileTableTests
{
    private static MainForm.NicItrConfigEntry Saved(string hwid, NicItrProfile layout) => new()
    {
        Hwid = hwid, FamilyName = layout.FamilyName, BaseOffset = layout.BaseOffset, Stride = layout.Stride,
        Queues = layout.MaxQueues, Width = layout.ReadWidth, Mask = layout.ReadMask, OrBits = layout.WriteOrBits, Values = [0x10],
    };

    [Theory]
    [InlineData("8162", 0x00E2u, 16)] // RTL8118AS / RTL8168H
    [InlineData("3000", 0x0A00u, 32)] // Killer E3000, RTL8125 family
    [InlineData("8125", 0x0A00u, 32)]
    [InlineData("2600", 0x00E2u, 16)]
    public void RealtekIdsUseTheRegisterMapOfTheirFamily(string device, uint offset, int width)
    {
        NicItrProfile profile = NicItrService.TryGetNicItrProfile($@"PCI\VEN_10EC&DEV_{device}&SUBSYS_00000000\0")!;
        Assert.Equal((offset, width), (profile.BaseOffset, profile.ReadWidth));
    }

    [Fact]
    public void EveryDeviceIdBelongsToOneProfile()
    {
        var ids = NicItrService.NicItrProfiles.SelectMany(p => p.DeviceIds.Select(id => p.VendorId + ":" + id)).ToList();
        Assert.Equal(ids.Count, ids.Distinct(StringComparer.OrdinalIgnoreCase).Count());
    }

    [Fact]
    public void EntrySavedWithTheOldKillerE3000LayoutIsRemoved()
    {
        NicItrProfile rtl8168 = NicItrService.NicItrProfiles.Single(p => p.FamilyName == "Realtek RTL8111/8168");
        NicItrProfile rtl8125 = NicItrService.NicItrProfiles.Single(p => p.FamilyName == "Realtek RTL8125/8126");
        MainForm.ImodConfig config = new();
        config.NicItrEntries.Add(Saved("VEN_10EC&DEV_3000", rtl8168));
        config.NicItrEntries.Add(Saved("VEN_10EC&DEV_8125", rtl8125));
        config.NicItrEntries.Add(Saved("legacy-key-without-pci-id", rtl8168));

        List<string> removed = MainForm.RemoveStaleNicItrEntries(config);

        string line = Assert.Single(removed);
        Assert.Contains("hwid=VEN_10EC&DEV_3000", line);
        Assert.Contains("reason=layout-changed saved=[offset=0xE2", line);
        Assert.Contains("current=[offset=0xA00", line);
        Assert.Equal(["VEN_10EC&DEV_8125", "legacy-key-without-pci-id"], config.NicItrEntries.Select(e => e.Hwid));
    }

    [Fact]
    public void EntryForAnIdWithoutProfileIsRemoved()
    {
        NicItrProfile rtl8168 = NicItrService.NicItrProfiles.Single(p => p.FamilyName == "Realtek RTL8111/8168");
        Assert.Contains("reason=no-current-profile", MainForm.DescribeStaleNicItrEntry(Saved("VEN_10EC&DEV_FFFF", rtl8168)));
        Assert.Null(MainForm.DescribeStaleNicItrEntry(Saved("VEN_10EC&DEV_8168", rtl8168)));
    }
}

public sealed class MsiDisableGuardTests
{
    // DEVPKEY_PciDevice_InterruptSupport bits: 1 = line-based, 2 = MSI, 4 = MSI-X.
    [Theory]
    [InlineData(4)]
    [InlineData(2)]
    [InlineData(6)]
    public void MessageOnlyDeviceCannotBeSwitchedToLineBased(int support)
    {
        Assert.False(MainForm.TryValidateMsiDisable(new PciInterruptCapabilities((PciInterruptSupport)support, 8), out string error));
        Assert.Contains("no line-based", error);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(3)]
    [InlineData(7)]
    [InlineData(0)]
    public void DevicesWithLineBasedOrNoInterruptsAreNotRestricted(int support)
    {
        Assert.True(MainForm.TryValidateMsiDisable(new PciInterruptCapabilities((PciInterruptSupport)support, null), out string error), error);
    }

    [Fact]
    public void UnknownCapabilityIsNotRestricted()
    {
        Assert.True(MainForm.TryValidateMsiDisable(null, out _));
    }
}
