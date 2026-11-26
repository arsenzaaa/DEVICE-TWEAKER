using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public sealed class ImodRoleMapTests
{
    private const string Controller = @"PCI\VEN_8086&DEV_7AE0&SUBSYS_00000000&REV_11\3&11583659&0&A0";
    private const string HostPath = @"\\?\PCI#VEN_8086&DEV_7AE0&SUBSYS_00000000&REV_11#3&11583659&0&A0#{3abf6f2d-71c4-462a-8a92-1e6861e6af27}";
    private const string OtherHostPath = @"\\?\PCI#VEN_1022&DEV_15B6&SUBSYS_00000000&REV_00#4&2b0d2a4&0&0341#{3abf6f2d-71c4-462a-8a92-1e6861e6af27}";

    private static DeviceTweakerImodRuntime.XhciInterrupterTopology Topology()
    {
        DeviceTweakerImodRuntime.XhciInterrupterTopology topology = new();
        topology.ByDeviceAddress[5] = [1];
        topology.ByRootPort[3] = [2];
        topology.ByRootPort[4] = [3];
        return topology;
    }

    private static UsbEndpointInfo Endpoint(string cls, string sub, string proto, int address, string topologyPath, string host = HostPath)
        => new() { HostControllerPath = host, InterfaceClass = cls, InterfaceSubClass = sub, InterfaceProtocol = proto,
            DeviceAddress = address, TopologyPath = topologyPath };

    [Fact]
    public void EndpointsMapByDeviceAddressThenRootPort()
    {
        UsbEndpointInfo[] endpoints =
        [
            Endpoint("03", "01", "02", address: 5, "Port1"),          // mouse, address match
            Endpoint("03", "01", "01", address: 9, "Port3"),          // keyboard, no address, root port 3
            Endpoint("01", "02", "00", address: 0, "Port4/Port2"),    // audio, first root port 4
            Endpoint("03", "00", "00", address: 7, "Port3"),          // generic HID: gamepad
            Endpoint("03", "01", "02", address: 5, "Port1", OtherHostPath),
        ];

        Dictionary<uint, HashSet<string>> roles = ImodRoleMap.MapEndpoints(endpoints, Topology(), Controller,
            new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Mouse", "Keyboard", "Audio" }, out ImodRoleMap.EndpointMapCounts counts);

        Assert.Equal(new uint[] { 1, 2, 3 }, roles.Keys.Order().ToArray());
        Assert.Equal(new[] { "Mouse" }, roles[1]);
        Assert.Equal(new[] { "Keyboard" }, roles[2]);
        Assert.Equal(new[] { "Audio" }, roles[3]);
        Assert.Equal(new ImodRoleMap.EndpointMapCounts(AddressMatches: 1, RootPortMatches: 2, Classified: 3, FilteredGamepads: 1), counts);
    }

    [Fact]
    public void GenericHidCountsAsGamepadOnlyWhenTheControllerReportsOne()
    {
        UsbEndpointInfo[] endpoints = [Endpoint("03", "00", "00", address: 5, "Port1")];

        Dictionary<uint, HashSet<string>> roles = ImodRoleMap.MapEndpoints(endpoints, Topology(), Controller,
            new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Gamepad" }, out ImodRoleMap.EndpointMapCounts counts);

        Assert.Equal(new[] { "Gamepad" }, roles[1]);
        Assert.Equal(0, counts.FilteredGamepads);
    }

    [Fact]
    public void UnmatchedEndpointsProduceNoMap()
    {
        UsbEndpointInfo[] endpoints = [Endpoint("03", "01", "02", address: 8, "Port9"), Endpoint("08", "06", "50", address: 5, "Port1")];

        Dictionary<uint, HashSet<string>> roles = ImodRoleMap.MapEndpoints(endpoints, Topology(), Controller,
            new HashSet<string>(), out ImodRoleMap.EndpointMapCounts counts);

        Assert.Empty(roles);
        Assert.Equal(new ImodRoleMap.EndpointMapCounts(0, 0, 1, 0), counts);
    }

    [Theory]
    [InlineData("Port3/Port1", true, 3u)]
    [InlineData("Port0/Port2", true, 2u)]
    [InlineData(@"Hub\Port12", true, 12u)]
    [InlineData("USB(1).P7", true, 7u)]
    [InlineData("Port", false, 0u)]
    [InlineData("", false, 0u)]
    public void RootPortIsTakenFromTheTopologyPath(string path, bool found, uint expected)
    {
        Assert.Equal(found, ImodRoleMap.TryGetUsbRootPort(path, out uint rootPort));
        Assert.Equal(expected, rootPort);
    }

    [Theory]
    [InlineData("03", "01", "02", "Mouse")]
    [InlineData("0x03", "0x01", "0x01", "Keyboard")]
    [InlineData("HID", "00", "00", "Gamepad")]
    [InlineData("03", "01", "00", "")]
    [InlineData("Audio", "", "", "Audio")]
    [InlineData("0E", "01", "00", "Webcam")]
    [InlineData("08", "06", "50", "")]
    public void InterfaceDescriptorsClassifyRoles(string cls, string sub, string proto, string expected)
    {
        bool classified = ImodRoleMap.TryClassifyUsbRole(Endpoint(cls, sub, proto, 1, "Port1"), out string role);

        Assert.Equal(expected.Length > 0, classified);
        Assert.Equal(expected, role);
    }

    [Theory]
    [InlineData("Мышь 8K", "Mouse")]
    [InlineData("Клавиатура", "Keyboard")]
    [InlineData("Headphones", "Audio")]
    [InlineData("Геймпад", "Gamepad")]
    [InlineData("USB Camera", "Webcam")]
    public void RussianAndEnglishLabelsShareRoles(string label, string expected)
    {
        Assert.True(ImodRoleMap.TryGetRoleBase(label, out string role));
        Assert.Equal(expected, role);
    }
}

public sealed class ImodMapTextTests
{
    private const string Controller = @"PCI\VEN_8086&DEV_1234\A";
    private static readonly ImodRoleSnapshot Snapshot = new([new(Controller, "Mouse 1K, Keyboard", "")], new Dictionary<string, string>());

    [Fact]
    public void ComposeShowsMappedAndUnmappedDevices()
    {
        Dictionary<uint, HashSet<string>> roles = new() { [1] = new(StringComparer.OrdinalIgnoreCase) { "Mouse" } };

        (string mapText, string detail) = ImodMapText.Compose([0x0, 0xC8], roles, Snapshot, Controller, "exact");

        string devices = "devices: Mouse 1K -> intr1  0xC8   50 us\r\n"
            + "         Keyboard -> intr?  mixed";
        string interrupters = "interrupters 0-1: intr0  0x0   0 ns  | intr1  0xC8  50 us";
        Assert.Equal($"{devices}\r\n\r\n{interrupters}", mapText);
        Assert.Equal($"{devices}\r\n\r\n{interrupters}\r\n"
            + "detail: I0=0x0 (0 ns) | I1 Mouse 1K=0xC8 (50 us)\r\n"
            + "raw: I0:0x0 | I1:0xC8\r\n"
            + "time: I0:0 ns | I1:50 us\r\n"
            + "map: exact", detail);
    }

    [Fact]
    public void WithoutARoleMapEveryDeviceIsUnknown()
    {
        (string mapText, string detail) = ImodMapText.Compose([0xC8, 0xC8], null, Snapshot, Controller, string.Empty);

        Assert.StartsWith("devices: Mouse 1K -> intr?  0xC8\r\n         Keyboard -> intr?  0xC8\r\n\r\n", mapText);
        Assert.DoesNotContain("map:", detail);
    }

    [Fact]
    public void InterrupterLinesShowEightAndCountTheRest()
    {
        uint[] values = Enumerable.Range(0, 9).Select(i => (uint)(i * 4)).ToArray();

        string[] lines = ImodMapText.VisibleInterrupterLines(values).Split("\r\n");

        Assert.Equal(new[] { "interrupters 0-1:", "interrupters 2-3:", "interrupters 4-5:", "interrupters 6-7:" },
            lines.Select(line => line[..line.IndexOf(':')] + ":"));
        Assert.EndsWith(" | +1 more", lines[^1]);
        Assert.Equal("interrupters: -", ImodMapText.VisibleInterrupterLines([]));
    }

    [Theory]
    [InlineData(0x0u, "0 ns")]
    [InlineData(0x3u, "750 ns")]
    [InlineData(0x4u, "1 us")]
    [InlineData(0x5u, "1.25 us")]
    [InlineData(0xFA0u, "1000 us")]
    [InlineData(0x1_00C8u, "50 us")]
    public void UsbTimeUses250NanosecondUnitsOfTheLow16Bits(uint raw, string expected)
    {
        Assert.Equal(expected, ImodMapText.UsbTime(raw));
    }

    [Fact]
    public void ValueListsSummarizeEqualValues()
    {
        Assert.Equal("0xC8 x3", ImodMapText.ValueList([0xC8, 0xC8, 0xC8]));
        Assert.Equal("varies by interrupter", ImodMapText.ValueList([0xC8, 0x0]));
        Assert.Equal("-", ImodMapText.ValueList([]));
        Assert.Equal("0x00C8,0x0000", ImodMapText.ValueListForLog([0x1_00C8, 0x0]));
    }
}
