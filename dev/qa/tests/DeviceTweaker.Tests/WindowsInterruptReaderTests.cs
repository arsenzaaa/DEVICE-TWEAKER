using System.Buffers.Binary;
using System.Text.Json;
using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public sealed class WindowsInterruptReaderTests
{
    private static byte[] Resource(params (bool Message, ushort Count)[] irqs)
    {
        byte[] data = new byte[20 + irqs.Length * 20];
        BinaryPrimitives.WriteUInt32LittleEndian(data, 1);
        BinaryPrimitives.WriteUInt32LittleEndian(data.AsSpan(16), (uint)irqs.Length);
        for (int i = 0; i < irqs.Length; i++)
        {
            int offset = 20 + i * 20; data[offset] = 2;
            BinaryPrimitives.WriteUInt16LittleEndian(data.AsSpan(offset + 2), irqs[i].Message ? (ushort)3 : (ushort)0);
            BinaryPrimitives.WriteUInt16LittleEndian(data.AsSpan(offset + 6), irqs[i].Count);
        }
        return data;
    }

    [Theory]
    [InlineData(1, 1, 3, "Msi")]
    [InlineData(1, 1, 5, "MsiX")]
    [InlineData(1, 1, 7, "Message")]
    [InlineData(1, 8, 7, "Msi")]
    [InlineData(9, 1, 7, "MsiX")]
    [InlineData(32, 1, 7, "MsiX")]
    [InlineData(1, 3, 7, "Unavailable")]
    [InlineData(1, 0, 7, "Unavailable")]
    [InlineData(1, 1, 1, "Unavailable")]
    [InlineData(2, 2, 7, "Unavailable")]
    [InlineData(1, 1, 15, "Unavailable")]
    public void AllocationShapeAndCapabilitiesDetermineResult(int descriptors, int count, uint support, string expected)
    {
        byte[] raw = Resource(Enumerable.Repeat((true, (ushort)count), descriptors).ToArray());
        var result = WindowsInterruptReader.Decode(raw, raw, support);
        Assert.Equal(expected, result.Mode.ToString());
        if (result.Mode != InterruptMode.Unavailable) Assert.Contains("PCI Enable bits were not read", result.Detail);
    }

    [Fact] public void MissingCapabilitiesDoesNotInventMsiX()
        => Assert.Equal(InterruptMode.Message, WindowsInterruptReader.Decode(Resource((true, 1)), Resource((true, 1)), null).Mode);
    [Fact] public void EdgeTriggeredLineIsNotTreatedAsAMessage()
    {
        byte[] data = Resource((false, 0)); data[22] = 1;
        Assert.Equal(InterruptMode.Line, WindowsInterruptReader.Decode(data, data, 7).Mode);
    }
    [Fact] public void EmptyValidListMeansNoAllocatedIrq()
        => Assert.Equal(InterruptMode.None, WindowsInterruptReader.Decode(Resource(), Resource(), 7).Mode);
    [Fact] public void ConflictingListsFail()
        => Assert.Equal(InterruptMode.Unavailable, WindowsInterruptReader.Decode(Resource((true, 1)), Resource((false, 0)), 7).Mode);
    [Fact] public void MixedLineAndMessageFails()
        => Assert.Equal(InterruptMode.Unavailable, WindowsInterruptReader.Decode(Resource((true, 1), (false, 0)), Resource((true, 1), (false, 0)), 7).Mode);
    [Theory] [InlineData(0)] [InlineData(3)] [InlineData(10)] [InlineData(21)] [InlineData(39)]
    public void TruncatedResourcesFail(int length)
    {
        byte[] raw = Resource((true, 1)); Array.Resize(ref raw, length);
        Assert.Equal(InterruptMode.Unavailable, WindowsInterruptReader.Decode(raw, Resource((true, 1)), 7).Mode);
    }

    private sealed class Source(params WindowsIrqSnapshot[] snapshots) : IWindowsIrqSource
    {
        internal int Calls;
        public WindowsIrqSnapshot Read(string _) => snapshots[Math.Min(Calls++, snapshots.Length - 1)];
    }
    private static WindowsIrqSnapshot Snapshot(string pdo = @"\Device\Test", bool started = true)
        => new(pdo, started, Resource((true, 1)), Resource((true, 1)), 3);

    [Fact] public void ChangingPdoRetriesAndAcceptsOnlyStableSample()
    {
        Source source = new(Snapshot(), Snapshot(@"\Device\Replacement"), Snapshot(), Snapshot());
        Assert.Equal(InterruptMode.Msi, WindowsInterruptReader.ReadStable(@"PCI\VEN_8086", source).Mode);
        Assert.Equal(4, source.Calls);
    }
    [Fact] public void RepeatedChangesFailInsteadOfReturningStaleMode()
    {
        Source source = new(Snapshot(), Snapshot(@"\Device\Other"), Snapshot(), Snapshot(@"\Device\Other"));
        Assert.Equal(InterruptMode.Unavailable, WindowsInterruptReader.ReadStable(@"PCI\VEN_8086", source).Mode);
    }
    [Fact] public void StoppedDeviceDoesNotClaimNoIrqOrActiveMode()
        => Assert.Equal(InterruptMode.Unavailable, WindowsInterruptReader.ReadStable(@"PCI\VEN_8086", new Source(Snapshot(started: false))).Mode);
    [Fact] public void NonPciDoesNotAccessBackend()
    {
        Source source = new(Snapshot());
        Assert.Equal(InterruptMode.NotPci, WindowsInterruptReader.ReadStable(@"USB\Test", source).Mode);
        Assert.Equal(0, source.Calls);
    }
    [Fact] public void AutonomousQaNeverReadsHardwareOrRegistry()
        => Assert.Equal(InterruptMode.NotRead, WindowsInterruptReader.Read(@"PCI\VEN_8086").Mode);

    [Fact] public void CapturedRealWindowsResourcesDecodeWithoutDriverCalls()
    {
        using JsonDocument doc = JsonDocument.Parse(File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "WindowsIrqCaptures.json")));
        var devices = doc.RootElement.GetProperty("devices").EnumerateArray().ToDictionary(x => x.GetProperty("data").GetProperty("InstanceId").GetString()!);
        int count = 0;
        foreach (var map in doc.RootElement.GetProperty("resourceMaps").EnumerateArray())
        {
            var device = devices[map.GetProperty("instanceId").GetString()!];
            uint support = device.GetProperty("data").GetProperty("SupportMask").GetUInt32();
            var result = WindowsInterruptReader.Decode(Convert.FromBase64String(map.GetProperty("rawBase64").GetString()!), Convert.FromBase64String(map.GetProperty("translatedBase64").GetString()!), support);
            Assert.Equal(support == 3 ? InterruptMode.Msi : InterruptMode.MsiX, result.Mode); count++;
        }
        Assert.Equal(6, count);
    }
}
