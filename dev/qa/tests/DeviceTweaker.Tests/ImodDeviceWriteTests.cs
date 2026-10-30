using System.Buffers.Binary;
using System.Text;
using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

// IOCTL_IMOD_WRITE_DEVICE request built by the shared runtime must match IMOD\Driver\imod_driver.h.
public sealed class ImodDeviceWriteTests
{
    private const string Pdo = @"\Device\NTPNP_PCI0012";
    private const string Instance = @"PCI\VEN_8086&DEV_7AE0&SUBSYS_7D361462&REV_11\3&11583659&0&A0";

    [Fact]
    public void RequestMatchesTheDriverStructLayout()
    {
        byte[] request = DeviceTweakerImodRuntime.BuildDeviceWriteRequest(Pdo, Instance, 0xF7000024, 4, 0x80000028);

        Assert.NotNull(request);
        Assert.Equal(1048, request.Length);
        Assert.Equal(1u, BinaryPrimitives.ReadUInt32LittleEndian(request.AsSpan(0)));
        Assert.Equal(4u, BinaryPrimitives.ReadUInt32LittleEndian(request.AsSpan(4)));
        Assert.Equal(0xF7000024UL, BinaryPrimitives.ReadUInt64LittleEndian(request.AsSpan(8)));
        Assert.Equal(0x80000028UL, BinaryPrimitives.ReadUInt64LittleEndian(request.AsSpan(16)));
        Assert.Equal(Pdo, Encoding.Unicode.GetString(request, 24, 512).TrimEnd('\0'));
        Assert.Equal(Instance, Encoding.Unicode.GetString(request, 536, 512).TrimEnd('\0'));
    }

    [Fact]
    public void IoctlUsesIndexFiveWithWriteAccess()
    {
        Assert.Equal((0x8010u << 16) | (2u << 14) | (0x815u << 2), DeviceTweakerImodRuntime.IoctlImodWriteDevice);
    }

    [Theory]
    [InlineData("NTPNP_PCI0012", Instance, 4)]
    [InlineData(Pdo, @"USB\VID_046D&PID_C547\1", 4)]
    [InlineData(Pdo, Instance, 3)]
    [InlineData(Pdo, Instance, 8)]
    public void InvalidRequestsAreNotBuilt(string pdo, string instance, uint size)
    {
        Assert.Null(DeviceTweakerImodRuntime.BuildDeviceWriteRequest(pdo, instance, 0x1000, size, 0));
    }

    [Fact]
    public void NamesMustFitTheDriverBuffers()
    {
        Assert.Null(DeviceTweakerImodRuntime.BuildDeviceWriteRequest(@"\Device\" + new string('A', 248), Instance, 0x1000, 4, 0));
        Assert.NotNull(DeviceTweakerImodRuntime.BuildDeviceWriteRequest(@"\Device\" + new string('A', 247), Instance, 0x1000, 4, 0));
    }

    [Fact]
    public void DriverRefusalsAreExplained()
    {
        Assert.Contains(@"0xF7000024 is outside the memory of PCI\X", DeviceTweakerImodRuntime.DescribeWriteFailure(5, 0xF7000024, @"PCI\X"));
        // ERROR_INVALID_FUNCTION: the resident DTIMOD predates device-checked writes.
        Assert.Equal(DriverRename.RestartMessage, DeviceTweakerImodRuntime.DescribeWriteFailure(1, 0, @"PCI\X"));
        Assert.StartsWith("failed to write physical memory via driver: ", DeviceTweakerImodRuntime.DescribeWriteFailure(87, 0, @"PCI\X"));
    }

    [Fact]
    public void OlderDriverProbeReportsNothingWithoutADriver()
    {
        Assert.False(DeviceTweakerImodRuntime.IsOlderDriverLoaded(new IntPtr(-1)));
    }
}
