using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public sealed class NicItrProgrammingTests
{
    private static NicItrProfile I225 => MainForm.TryGetNicItrProfile(@"PCI\VEN_8086&DEV_15F3")!;

    // Regression baseline for the current I225 layout, independent of the profile under test.
    [Fact]
    public void I225_WritesFiveExactAddressesAndVerifiesEachRegister()
    {
        MemoryRegisters bus = new();
        Assert.True(MainForm.TryProgramNicItrRegisters(I225, 0x100000, [0xFFFF], bus, out string? error), error);
        Assert.Equal(new ulong[] { 0x101680, 0x101684, 0x101688, 0x10168C, 0x101690 }, bus.Writes.Select(w => w.Address));
        Assert.All(bus.Writes, w => { Assert.Equal(32, w.Width); Assert.Equal(0x80007FFCUL, w.Value); });
        Assert.Equal(new[] { "W0", "R0", "W1", "R1", "W2", "R2", "W3", "R3", "W4", "R4" }, bus.Trace);
    }

    [Fact]
    public void PerQueueInputAndMasking_ArePreserved()
    {
        MemoryRegisters bus = new();
        Assert.True(MainForm.TryProgramNicItrRegisters(I225, 0, [4, 8, 12, 16, 20], bus, out _));
        Assert.Equal(new ulong[] { 0x80000004, 0x80000008, 0x8000000C, 0x80000010, 0x80000014 }, bus.Writes.Select(w => w.Value));
    }

    [Fact]
    public void Realtek16BitProfile_UsesOne16BitRegister()
    {
        var profile = MainForm.TryGetNicItrProfile(@"PCI\VEN_10EC&DEV_8168")!;
        MemoryRegisters bus = new();
        Assert.True(MainForm.TryProgramNicItrRegisters(profile, 0x2000, [0x12345], bus, out _));
        Assert.Equal((0x20E2UL, 16, 0x2345UL), Assert.Single(bus.Writes));
        Assert.Equal(new[] { "W0", "R0" }, bus.Trace);
    }

    [Theory]
    [InlineData("write", 0, 1, 0)]
    [InlineData("write", 2, 3, 2)]
    [InlineData("read", 0, 1, 1)]
    [InlineData("read", 2, 3, 3)]
    [InlineData("mismatch", 0, 1, 1)]
    [InlineData("mismatch", 2, 3, 3)]
    public void Failure_StopsBeforeLaterQueues(string failure, int index, int writes, int reads)
    {
        MemoryRegisters bus = new() { Failure = failure, FailureIndex = index };
        Assert.False(MainForm.TryProgramNicItrRegisters(I225, 0, [8], bus, out string? error));
        Assert.NotNull(error);
        Assert.Equal(writes, bus.Writes.Count);
        Assert.Equal(reads, bus.ReadCount);
        if (failure == "mismatch") Assert.Contains("expected", error);
        if (failure == "read") Assert.Contains("write completed but readback failed", error);
    }

    [Fact]
    public void EmptyInput_PerformsNoIo()
    {
        MemoryRegisters bus = new();
        Assert.False(MainForm.TryProgramNicItrRegisters(I225, 0, [], bus, out _));
        Assert.Empty(bus.Trace);
    }

    private sealed class MemoryRegisters : INicItrRegisters
    {
        public string? Failure { get; init; }
        public int FailureIndex { get; init; }
        public List<(ulong Address, int Width, ulong Value)> Writes { get; } = [];
        public List<string> Trace { get; } = [];
        public int ReadCount { get; private set; }
        private readonly Dictionary<ulong, ulong> _memory = [];

        public bool TryWrite(ulong address, int width, ulong value, out string? error)
        {
            int index = Writes.Count;
            Trace.Add($"W{index}");
            Writes.Add((address, width, value));
            error = Failure == "write" && index == FailureIndex ? "injected write failure" : null;
            if (error is not null) return false;
            _memory[address] = value;
            return true;
        }

        public bool TryRead(ulong address, int width, out ulong value, out string? error)
        {
            int index = ReadCount++;
            Trace.Add($"R{index}");
            value = _memory[address];
            if (Failure == "mismatch" && index == FailureIndex) value ^= 4;
            error = Failure == "read" && index == FailureIndex ? "injected read failure" : null;
            return error is null;
        }
    }
}
