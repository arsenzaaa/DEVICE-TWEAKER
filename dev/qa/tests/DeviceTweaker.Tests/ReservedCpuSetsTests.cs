using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public class ReservedCpuSetsTests
{
    [Fact]
    public void EncodeAndDecode_RoundTripsAccurately()
    {
        bool[] bits = new bool[16];
        bits[2] = true;
        bits[3] = true;
        bits[5] = true;
        bits[10] = true;

        byte[] encoded = MainForm.EncodeReservedCpuSetsBytes(bits);
        Assert.Equal(8, encoded.Length);
        Assert.Equal(new byte[] { 0x2C, 0x04, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 }, encoded);

        List<int> decoded = MainForm.DecodeReservedCpuSetsRawIds(encoded);
        Assert.Equal([2, 3, 5, 10], decoded);
    }

    [Fact]
    public void Encode_MultiGroup_PadsToMultipleOfEightBytes()
    {
        bool[] bits = new bool[70];
        bits[1] = true;
        bits[64] = true;

        byte[] encoded = MainForm.EncodeReservedCpuSetsBytes(bits);
        Assert.Equal(16, encoded.Length);
        Assert.Equal(0x02, encoded[0]);
        Assert.Equal(0x01, encoded[8]);
        Assert.Equal([1, 64], MainForm.DecodeReservedCpuSetsRawIds(encoded));
    }

    [Fact]
    public void Decode_LegacyShortMask_StillReadsCleanly()
    {
        List<int> decoded = MainForm.DecodeReservedCpuSetsRawIds([0x2C, 0x04]);
        Assert.Equal([2, 3, 5, 10], decoded);
    }

    [Fact]
    public void Encode_EmptyOrAllFalse_ReturnsEmptyBytes()
    {
        byte[] empty = MainForm.EncodeReservedCpuSetsBytes([]);
        Assert.Empty(empty);

        byte[] allFalse = MainForm.EncodeReservedCpuSetsBytes(new bool[8]);
        Assert.Empty(allFalse);
    }

    [Fact]
    public void Decode_Empty_ReturnsEmptyList()
    {
        List<int> decoded = MainForm.DecodeReservedCpuSetsRawIds([]);
        Assert.Empty(decoded);
    }
}
