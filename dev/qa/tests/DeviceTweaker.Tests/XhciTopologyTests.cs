using Xunit;

namespace DeviceTweaker.Tests;

// Shared xHCI readers of the app and the logon script, run against fake controller memory.
public sealed class XhciTopologyTests
{
    private const ulong Bar = 0x1000_0000;
    private const ulong BarLength = 0x10000;

    [Fact]
    public void EndpointTargetWinsOverSlotTargetAndHubsAreSkipped()
    {
        FakeXhci xhci = new();
        // Mouse: slot says I1, its interrupt IN ring says I3.
        xhci.Device(slot: 1, rootPort: 2, address: 5, slotInterrupter: 1, endpointInterrupter: 3);
        // No endpoint ring: the slot target is used.
        xhci.Device(slot: 2, rootPort: 4, address: 6, slotInterrupter: 2, endpointInterrupter: null);
        xhci.Device(slot: 3, rootPort: 1, address: 7, slotInterrupter: 4, endpointInterrupter: null, hub: true);

        Assert.True(DeviceTweakerImodRuntime.TryReadXhciInterrupterTopology(
            xhci, Bar, BarLength, 4, out uint maxIntrs, out var topology, out string detail), detail);

        Assert.Equal(8u, maxIntrs);
        Assert.Equal(new uint[] { 2, 4 }, topology.ByRootPort.Keys.Order());
        Assert.Equal([3u], topology.ByRootPort[2]);
        Assert.Equal([2u], topology.ByRootPort[4]);
        Assert.Equal([3u], topology.ByDeviceAddress[5]);
        Assert.Equal(1, topology.EndpointTargetCount);
        Assert.Equal(1, topology.SlotTargetCount);
        Assert.Contains("rootPortMap=[2:I3, 4:I2]", detail);
    }

    [Fact]
    public void RuntimeLayoutStaysInsideTheBar()
    {
        FakeXhci xhci = new();
        Assert.True(DeviceTweakerImodRuntime.TryResolveXhciRuntimeLayout(
            xhci, Bar, BarLength, 4, 0x18, out uint maxIntrs, out ulong runtime, out _, out _, out string error), error);
        Assert.Equal(8u, maxIntrs);
        Assert.Equal(Bar + 0x2000, runtime);

        xhci.Dwords[Bar + 0x18] = 0xFFE0;
        Assert.False(DeviceTweakerImodRuntime.TryResolveXhciRuntimeLayout(
            xhci, Bar, BarLength, 4, 0x18, out _, out _, out _, out _, out error));
        Assert.Contains("RTSOFF=0xFFE0", error);
        Assert.Empty(xhci.Writes);
    }

    [Theory]
    [InlineData(0x0000_0080u, "unexpected xHCI version")]
    [InlineData(0x0110_0010u, "invalid xHCI CAPLENGTH")]
    public void ForeignOrBrokenHeaderIsRejected(uint header, string message)
    {
        FakeXhci xhci = new();
        xhci.Dwords[Bar] = header;
        Assert.False(DeviceTweakerImodRuntime.TryResolveXhciRuntimeLayout(
            xhci, Bar, BarLength, 4, 0x18, out _, out _, out _, out _, out string error));
        Assert.Contains(message, error);
    }

    [Fact]
    public void RolePriorityThenNameDecidesSharedInterrupters()
    {
        Dictionary<string, uint> intervals = new(StringComparer.OrdinalIgnoreCase) { ["Audio"] = 0xFA0, ["Mouse"] = 0, ["Zeta"] = 1, ["Beta"] = 2 };
        Assert.True(DeviceTweakerImodRuntime.TrySelectAdaptiveRoleInterval(new(StringComparer.OrdinalIgnoreCase) { "Audio", "Mouse" }, intervals, out string role, out uint value));
        Assert.Equal(("Mouse", 0u), (role, value));
        Assert.True(DeviceTweakerImodRuntime.TrySelectAdaptiveRoleInterval(new(StringComparer.OrdinalIgnoreCase) { "Zeta", "Beta" }, intervals, out role, out value));
        Assert.Equal(("Beta", 2u), (role, value));
        Assert.False(DeviceTweakerImodRuntime.TrySelectAdaptiveRoleInterval(new(StringComparer.OrdinalIgnoreCase) { "Webcam" }, intervals, out _, out _));
    }

    private sealed class FakeXhci : IDeviceRegisters
    {
        private const ulong Dcbaa = 0x5000_0000;
        public Dictionary<ulong, uint> Dwords { get; } = [];
        public List<ulong> Writes { get; } = [];

        public FakeXhci()
        {
            Dwords[Bar] = 0x0110_0080;          // HCIVERSION 1.10, CAPLENGTH 0x80
            Dwords[Bar + 0x04] = (8u << 8) | 4; // MaxIntrs 8, MaxSlots 4
            Dwords[Bar + 0x10] = 0;             // 32-byte contexts
            Dwords[Bar + 0x18] = 0x2000;        // RTSOFF
            Dwords[Bar + 0x80 + 0x30] = unchecked((uint)Dcbaa);
        }

        public void Device(uint slot, uint rootPort, uint address, uint slotInterrupter, uint? endpointInterrupter, bool hub = false)
        {
            ulong context = 0x6000_0000 + (slot * 0x1000UL);
            Dwords[Dcbaa + (slot * 8)] = unchecked((uint)context);
            uint entries = endpointInterrupter.HasValue ? 3u : 1u;
            Dwords[context] = (entries << 27) | (hub ? 1u << 26 : 0);
            Dwords[context + 0x04] = rootPort << 16;
            Dwords[context + 0x08] = slotInterrupter << 22;
            Dwords[context + 0x0C] = (3u << 27) | address;
            if (endpointInterrupter is uint target)
            {
                ulong endpoint = context + (3 * 32);
                ulong ring = 0x7000_0000 + (slot * 0x1000UL);
                Dwords[endpoint] = 1;                         // running
                Dwords[endpoint + 0x04] = 7u << 3;            // interrupt IN
                Dwords[endpoint + 0x08] = unchecked((uint)ring) | 1; // dequeue pointer + DCS
                Dwords[ring + 0x08] = target << 22;
                Dwords[ring + 0x0C] = 1u << 10;               // normal TRB
            }
        }

        public bool TryRead(ulong address, uint size, out ulong value, out string error)
        {
            value = Dwords.TryGetValue(address, out uint dword) ? dword : 0;
            error = null!;
            return size == 4;
        }

        public bool TryWrite(ulong address, uint size, ulong value, out string error)
        {
            Writes.Add(address);
            error = "read-only fake";
            return false;
        }
    }
}
