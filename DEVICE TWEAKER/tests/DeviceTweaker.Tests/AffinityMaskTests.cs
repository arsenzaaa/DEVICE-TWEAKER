using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public class AffinityMaskTests
{
    [Fact]
    public void ReadAffinityMaskValue_Parses8ByteBinary()
    {
        ulong expected = 0x00000000_0000000F;
        byte[] bytes = BitConverter.GetBytes(expected);
        ulong actual = MainForm.ReadAffinityMaskValue(bytes);
        Assert.Equal(expected, actual);
    }

    [Fact]
    public void ReadAffinityMaskValue_Parses4ByteBinary()
    {
        uint expected32 = 0x000000FF;
        byte[] bytes = BitConverter.GetBytes(expected32);
        ulong actual = MainForm.ReadAffinityMaskValue(bytes);
        Assert.Equal(expected32, actual);
    }

    [Fact]
    public void ReadAffinityMaskValue_ParsesIntegers()
    {
        Assert.Equal(15UL, MainForm.ReadAffinityMaskValue(15));
        Assert.Equal(255UL, MainForm.ReadAffinityMaskValue(255U));
        Assert.Equal(0x1000UL, MainForm.ReadAffinityMaskValue(0x1000L));
        Assert.Equal(0x2000UL, MainForm.ReadAffinityMaskValue(0x2000UL));
    }

    [Fact]
    public void ReadAffinityMaskValue_ReturnsZeroForNullOrInvalid()
    {
        Assert.Equal(0UL, MainForm.ReadAffinityMaskValue(null));
        Assert.Equal(0UL, MainForm.ReadAffinityMaskValue("not a number"));
        Assert.Equal(0UL, MainForm.ReadAffinityMaskValue(new byte[] { 1, 2 }));
    }
}
