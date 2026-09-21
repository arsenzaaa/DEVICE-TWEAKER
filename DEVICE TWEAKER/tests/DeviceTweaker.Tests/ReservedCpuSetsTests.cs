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
        Assert.Equal(2, encoded.Length);

        List<int> decoded = MainForm.DecodeReservedCpuSetsRawIds(encoded);
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
