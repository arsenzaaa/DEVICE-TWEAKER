namespace DeviceTweakerCS.Tests;

using Xunit;

public class CcdSelectionTests
{
    [Fact]
    public void TrySelectNonVCacheCcd_SelectsUniqueSmallerL3Ccd()
    {
        Dictionary<int, long> sizes = new()
        {
            [0] = 96L * 1024 * 1024,
            [1] = 32L * 1024 * 1024,
        };

        bool selected = MainForm.TrySelectNonVCacheCcd([0, 1], sizes, out int ccd);

        Assert.True(selected);
        Assert.Equal(1, ccd);
    }

    [Fact]
    public void TrySelectNonVCacheCcd_SelectsFirstCcdWhenItHasSmallerL3()
    {
        Dictionary<int, long> sizes = new()
        {
            [0] = 32L * 1024 * 1024,
            [1] = 96L * 1024 * 1024,
        };

        Assert.True(MainForm.TrySelectNonVCacheCcd([0, 1], sizes, out int ccd));
        Assert.Equal(0, ccd);
    }

    [Fact]
    public void TrySelectNonVCacheCcd_RejectsIncompleteL3Data()
    {
        Dictionary<int, long> sizes = new()
        {
            [0] = 32L * 1024 * 1024,
        };

        Assert.False(MainForm.TrySelectNonVCacheCcd([0, 1], sizes, out _));
    }

    [Theory]
    [InlineData(32, 32)]
    [InlineData(96, 96)]
    [InlineData(64, 48)]
    public void TrySelectNonVCacheCcd_RejectsSymmetricOrSmallDifferences(int firstMb, int secondMb)
    {
        Dictionary<int, long> sizes = new()
        {
            [0] = firstMb * 1024L * 1024L,
            [1] = secondMb * 1024L * 1024L,
        };

        Assert.False(MainForm.TrySelectNonVCacheCcd([0, 1], sizes, out _));
    }

    [Fact]
    public void BuildCcdL3CacheSizeMap_SumsDistinctCcxCachesPerCcd()
    {
        List<CpuLpInfo> lps =
        [
            new(0, 0, 0, 0, 0, 0, L3CacheSizeBytes: 16L * 1024 * 1024),
            new(0, 1, 1, 0, 0, 0, L3CacheSizeBytes: 16L * 1024 * 1024),
            new(0, 2, 2, 1, 0, 0, L3CacheSizeBytes: 16L * 1024 * 1024),
            new(0, 3, 3, 1, 0, 0, L3CacheSizeBytes: 16L * 1024 * 1024),
            new(0, 4, 4, 2, 0, 0, L3CacheSizeBytes: 96L * 1024 * 1024),
            new(0, 5, 5, 2, 0, 0, L3CacheSizeBytes: 96L * 1024 * 1024),
        ];
        CpuTopology topology = new(lps);
        Dictionary<int, int> ccdMap = new()
        {
            [0] = 0, [1] = 0, [2] = 0, [3] = 0,
            [4] = 1, [5] = 1,
        };

        Dictionary<int, long> sizes = MainForm.BuildCcdL3CacheSizeMap(topology, ccdMap);

        Assert.Equal(32L * 1024 * 1024, sizes[0]);
        Assert.Equal(96L * 1024 * 1024, sizes[1]);
    }

    [Fact]
    public void IrqInfo_FlagsCombinedIdenticalDeviceMatches()
    {
        DeviceIrqInfo info = new();
        info.AddIrq(1000, @"PCI\VEN_1234&DEV_5678\A");
        info.AddIrq(1001, @"PCI\VEN_1234&DEV_5678\B");

        Assert.True(info.IsAmbiguous);
        Assert.Equal(2, info.MatchedDeviceCount);
    }

    [Theory]
    [InlineData("Unlimited")]
    [InlineData("0")]
    [InlineData("")]
    public void MsiLimitParser_AcceptsUnlimitedAliases(string input)
    {
        Assert.True(MainForm.TryParseMsiLimitInput(input, out bool unlocked, out int value));
        Assert.True(unlocked);
        Assert.Equal(0, value);
    }

    [Theory]
    [InlineData("1")]
    [InlineData("2")]
    [InlineData("4")]
    [InlineData("8")]
    [InlineData("16")]
    public void MsiLimitValidation_AcceptsLegalMsiOnlyValues(string input)
    {
        PciInterruptCapabilities capabilities = new(
            PciInterruptSupport.LineBased | PciInterruptSupport.Msi,
            16);

        Assert.True(MainForm.TryValidateMsiLimitInput(input, capabilities, out bool unlocked, out _, out string error), error);
        Assert.False(unlocked);
    }

    [Theory]
    [InlineData("3")]
    [InlineData("15")]
    public void MsiLimitValidation_RejectsNonPowerOfTwoMsiOnlyValues(string input)
    {
        PciInterruptCapabilities capabilities = new(PciInterruptSupport.Msi, 16);

        Assert.False(MainForm.TryValidateMsiLimitInput(input, capabilities, out _, out _, out string error));
        Assert.Contains("MSI-only", error, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void MsiLimitValidation_UsesMsiXHardwareMaximum()
    {
        PciInterruptCapabilities capabilities = new(PciInterruptSupport.MsiX, 9);

        Assert.True(MainForm.TryValidateMsiLimitInput("9", capabilities, out _, out _, out _));
        Assert.False(MainForm.TryValidateMsiLimitInput("10", capabilities, out _, out _, out string error));
        Assert.Contains("hardware maximum", error, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void MsiLimitValidation_AllowsUnlimitedEvenWithoutMessageSupport()
    {
        PciInterruptCapabilities capabilities = new(PciInterruptSupport.LineBased, null);

        Assert.True(MainForm.TryValidateMsiLimitInput("0", capabilities, out bool unlocked, out _, out _));
        Assert.True(unlocked);
        Assert.False(MainForm.TryValidateMsiLimitInput("1", capabilities, out _, out _, out _));
    }
}
