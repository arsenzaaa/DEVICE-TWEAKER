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

    [Theory]
    [InlineData(new byte[] { 0x0C }, 0x0CUL)]
    [InlineData(new byte[] { 0x50 }, 0x50UL)]
    [InlineData(new byte[] { 0x00, 0x03 }, 0x0300UL)]
    [InlineData(new byte[] { 0x00, 0x30 }, 0x3000UL)]
    [InlineData(new byte[] { 0x00, 0x0C }, 0x0C00UL)]
    [InlineData(new byte[] { 0x01, 0x02 }, 0x0201UL)]
    [InlineData(new byte[] { 0x01, 0x02, 0x03 }, 0x030201UL)]
    [InlineData(new byte[] { 0x01, 0x02, 0x03, 0x04, 0x05 }, 0x0504030201UL)]
    [InlineData(new byte[] { 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x80, 0xFF }, 0x80000000_00000001UL)]
    public void ReadAffinityMaskValue_ParsesShortAndMultiByteBinaryLittleEndian(byte[] bytes, ulong expected)
    {
        Assert.Equal(expected, MainForm.ReadAffinityMaskValue(bytes));
        Assert.Equal(expected, MainForm.TryReadAffinityMaskValue(bytes));
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
        Assert.Equal(0UL, MainForm.ReadAffinityMaskValue(Array.Empty<byte>()));
        Assert.Null(MainForm.TryReadAffinityMaskValue(null));
        Assert.Null(MainForm.TryReadAffinityMaskValue("not a number"));
        Assert.Null(MainForm.TryReadAffinityMaskValue(Array.Empty<byte>()));
        Assert.Equal(0UL, MainForm.TryReadAffinityMaskValue(new byte[] { 0x00 }));
    }
}
