using Xunit;
namespace DeviceTweaker.Tests;

public sealed class ImodRegisterTests
{
    private sealed class Bus(string fault) : IDeviceRegisters
    {
        internal readonly List<string> Calls = [];
        internal uint Value = 0xABCD1234;
        private int reads;
        public bool TryRead(ulong address, uint size, out ulong value, out string error)
        {
            Calls.Add($"read:{address:X}:{size}"); reads++; value = Value; error = "read failure";
            if (fault == "initial-read" && reads == 1 || fault == "readback" && reads == 2) return false;
            if (fault == "mismatch" && reads == 2) value ^= 1;
            // Hardware can advance the moderation counter (upper 16 bits).
            if (fault == "counter" && reads == 2) value ^= 0xFFFF0000;
            error = null!; return true;
        }
        public bool TryWrite(ulong address, uint size, ulong value, out string error)
        {
            Calls.Add($"write:{address:X}:{size}"); error = "write failure";
            if (fault == "write") return false;
            Value = (uint)value; error = null!; return true;
        }
    }
    [Theory] [InlineData(0u)] [InlineData(8u)] [InlineData(0xFFFFu)] [InlineData(0x12340008u)]
    public void IntervalWritePreservesCounterAndVerifiesLowBits(uint value)
    {
        Bus bus = new("none"); Assert.True(DeviceTweakerImodRuntime.TryWriteImodInterval(bus, 0x1020, value, out var error));
        Assert.Null(error); Assert.Equal(0xABCD0000u | (value & 0xFFFF), bus.Value);
        Assert.Equal(new[] { "read:1020:4", "write:1020:4", "read:1020:4" }, bus.Calls);
    }
    [Theory] [InlineData("initial-read", 1)] [InlineData("write", 2)] [InlineData("readback", 3)] [InlineData("mismatch", 3)]
    public void FailureIsReportedAndStopsFurtherIo(string fault, int calls)
    {
        Bus bus = new(fault); Assert.False(DeviceTweakerImodRuntime.TryWriteImodInterval(bus, 0x1020, 8, out var error));
        Assert.NotNull(error); Assert.Equal(calls, bus.Calls.Count);
    }
    [Fact] public void ChangingHardwareCounterDoesNotFailIntervalVerification()
    {
        Bus bus = new("counter"); Assert.True(DeviceTweakerImodRuntime.TryWriteImodInterval(bus, 0x1020, 8, out _));
    }
}
