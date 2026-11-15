using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public sealed class ImodReadbackTests
{
    private sealed class Registers(int failureIndex = -1) : IDeviceRegisters
    {
        internal readonly List<ulong> Reads = [];
        public bool TryRead(ulong address, uint size, out ulong value, out string error)
        {
            Assert.Equal(4u, size);
            Reads.Add(address);
            int index = (int)((address - 0x1024) / 0x20);
            value = 0xABCD0000u | (uint)(index + 1);
            error = "modeled read failure";
            return index != failureIndex;
        }
        public bool TryWrite(ulong address, uint size, ulong value, out string error)
            => throw new InvalidOperationException("Readback must not write");
    }

    [Fact]
    public void SuccessfulReadKeepsIndicesAndDiscardsCounterBits()
    {
        Registers registers = new();
        Assert.True(ImodReadbackService.TryRead(registers, 0x1000, 3, out var values, out var error));
        Assert.Null(error);
        Assert.Equal(new[] { 1u, 2u, 3u }, values);
        Assert.Equal(new ulong[] { 0x1024, 0x1044, 0x1064 }, registers.Reads);
    }

    [Fact]
    public void MissingMiddleRegisterDoesNotRelabelTheNextInterrupter()
    {
        Registers registers = new(failureIndex: 1);
        Assert.False(ImodReadbackService.TryRead(registers, 0x1000, 3, out var values, out var error));
        Assert.Empty(values);
        Assert.Equal(new ulong[] { 0x1024, 0x1044 }, registers.Reads);
        Assert.Contains("interrupter=1 address=0x1044", error);
    }

    [Theory]
    [InlineData(0x1000UL, 0u)]
    [InlineData(0x1000UL, 2049u)]
    [InlineData(ulong.MaxValue - 0x24, 1u)]
    public void InvalidRangesNeverRead(ulong address, uint count)
    {
        Registers registers = new();
        Assert.False(ImodReadbackService.TryRead(registers, address, count, out var values, out _));
        Assert.Empty(values);
        Assert.Empty(registers.Reads);
    }
}
