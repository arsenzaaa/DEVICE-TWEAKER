using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public class UsbChipPathTests
{
    [Theory]
    // AMD CPU-direct (CHIP 0)
    [InlineData(@"PCI\VEN_1022&DEV_15B6&SUBSYS_00000000", "CpuDirect", "CHIP 0")]
    [InlineData(@"PCI\VEN_1022&DEV_15B7&SUBSYS_00000000", "CpuDirect", "CHIP 0")]
    [InlineData(@"PCI\VEN_1022&DEV_1587&SUBSYS_00000000", "CpuDirect", "CHIP 0")]
    [InlineData(@"PCI\VEN_1022&DEV_15E2&SUBSYS_00000000", "CpuDirect", "CHIP 0")]
    [InlineData(@"PCI\VEN_1022&DEV_1502&SUBSYS_00000000", "CpuDirect", "CHIP 0")]
    [InlineData(@"PCI\VEN_1022&DEV_149C&SUBSYS_00000000", "CpuDirect", "CHIP 0")]
    // AMD Chipset (CHIP 1)
    [InlineData(@"PCI\VEN_1022&DEV_43FC&SUBSYS_00000000", "Chipset", "CHIP 1")]
    [InlineData(@"PCI\VEN_1022&DEV_43FE&SUBSYS_00000000", "Chipset", "CHIP 1")]
    [InlineData(@"PCI\VEN_1022&DEV_43F7&SUBSYS_00000000", "Chipset", "CHIP 1")]
    // Intel CPU-direct & Thunderbolt (CHIP 0)
    [InlineData(@"PCI\VEN_8086&DEV_7F35&SUBSYS_00000000", "CpuDirect", "CHIP 0")]
    [InlineData(@"PCI\VEN_8086&DEV_AE78&SUBSYS_00000000", "CpuDirect", "CHIP 0")]
    [InlineData(@"PCI\VEN_8086&DEV_5782&SUBSYS_00000000", "Thunderbolt", "CHIP 0")]
    // Intel PCH Chipset (CHIP 1)
    [InlineData(@"PCI\VEN_8086&DEV_7F6E&SUBSYS_00000000", "Chipset", "CHIP 1")]
    [InlineData(@"PCI\VEN_8086&DEV_777D&SUBSYS_00000000", "Chipset", "CHIP 1")]
    [InlineData(@"PCI\VEN_8086&DEV_7A60&SUBSYS_00000000", "Chipset", "CHIP 1")]
    // Addon (CHIP 1+)
    [InlineData(@"PCI\VEN_1B21&DEV_4242&SUBSYS_00000000", "Addon", "CHIP 1")]
    [InlineData(@"PCI\VEN_1B21&DEV_3242&SUBSYS_00000000", "Addon", "CHIP 1")]
    public void Classify_RecognizesKnownControllers(string instanceId, string expectedOrigin, string expectedTagPrefix)
    {
        UsbChipPathInfo info = UsbChipPath.Classify(instanceId);
        Assert.Equal(expectedOrigin, info.Origin.ToString());
        Assert.StartsWith(expectedTagPrefix, info.CompactTag);
    }

    [Theory]
    [InlineData("")]
    [InlineData("NOT_A_PCI_DEVICE")]
    [InlineData(@"USB\VID_046D&PID_C08B")]
    public void Classify_HandlesInvalidOrNonPciDevices(string instanceId)
    {
        UsbChipPathInfo info = UsbChipPath.Classify(instanceId);
        Assert.Equal("Unknown", info.Origin.ToString());
    }
}
