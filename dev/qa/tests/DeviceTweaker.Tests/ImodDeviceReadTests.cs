using Xunit;

namespace DeviceTweakerCS.Tests;

public sealed class ImodDeviceReadTests
{
    private static byte[] Reply(uint contextSize = 32)
    {
        byte[] bytes = new byte[40];
        uint[] words = [1, 2, 8, contextSize, 1, 1, 0, 3, 5, 6];
        for (int i = 0; i < words.Length; i++) BitConverter.GetBytes(words[i]).CopyTo(bytes, i * 4);
        return bytes;
    }

    [Fact]
    public void NewQueriesRequireReadAccessAndUseDistinctFunctionNumbers()
    {
        const uint device = 0x8010;
        Assert.Equal((device << 16) | (1u << 14) | (0x816u << 2), DeviceTweakerImodRuntime.IoctlImodReadDevice);
        Assert.Equal((device << 16) | (1u << 14) | (0x817u << 2), DeviceTweakerImodRuntime.IoctlImodQueryXhci);
        Assert.NotEqual(DeviceTweakerImodRuntime.IoctlImodWriteDevice, DeviceTweakerImodRuntime.IoctlImodReadDevice);
    }

    [Theory]
    [InlineData(32u)]
    [InlineData(64u)]
    public void TopologyReplyPreservesAddressPortAndEndpointTarget(uint contextSize)
    {
        byte[] bytes = Reply(contextSize);
        Assert.True(DeviceTweakerImodRuntime.TryParseNativeTopology(bytes, bytes.Length,
            out uint maxIntrs, out var topology, out string detail));
        Assert.Equal(8u, maxIntrs);
        Assert.Equal(2u, topology.MaxSlots);
        Assert.Equal(contextSize, topology.ContextSize);
        Assert.Equal(new[] { 6u }, topology.ByRootPort[3]);
        Assert.Equal(new[] { 6u }, topology.ByDeviceAddress[5]);
        Assert.Equal(1, topology.EndpointTargetCount);
        Assert.Equal(0, topology.SlotTargetCount);
        Assert.Contains("endpointTargets=1, slotTargets=0", detail);
    }

    [Fact]
    public void SlotFallbackIsReportedWithoutAnEndpointTarget()
    {
        byte[] bytes = Reply();
        BitConverter.GetBytes(0u).CopyTo(bytes, 20);
        BitConverter.GetBytes(1u).CopyTo(bytes, 24);
        BitConverter.GetBytes(2u).CopyTo(bytes, 36);
        Assert.True(DeviceTweakerImodRuntime.TryParseNativeTopology(bytes, bytes.Length, out _, out var topology, out _));
        Assert.Equal(new[] { 2u }, topology.ByRootPort[3]);
        Assert.Equal(1, topology.SlotTargetCount);
    }

    [Theory]
    [InlineData(0, 99u)]
    [InlineData(4, 256u)]
    [InlineData(8, 0u)]
    [InlineData(8, 2048u)]
    [InlineData(12, 16u)]
    [InlineData(16, uint.MaxValue)]
    [InlineData(20, 3u)]
    [InlineData(24, 2u)]
    [InlineData(28, 0u)]
    [InlineData(32, 256u)]
    [InlineData(36, 8u)]
    public void MalformedTopologyIsRejectedWithoutPublishingPartialMaps(int offset, uint value)
    {
        byte[] bytes = Reply();
        BitConverter.GetBytes(value).CopyTo(bytes, offset);
        Assert.False(DeviceTweakerImodRuntime.TryParseNativeTopology(bytes, bytes.Length,
            out uint maxIntrs, out var topology, out _));
        Assert.Equal(0u, maxIntrs);
        Assert.Empty(topology.ByRootPort);
        Assert.Empty(topology.ByDeviceAddress);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(27)]
    [InlineData(39)]
    [InlineData(41)]
    public void TruncatedOrOversizedReplyIsRejected(int returned)
    {
        Assert.False(DeviceTweakerImodRuntime.TryParseNativeTopology(Reply(), returned, out _, out _, out _));
    }

    [Fact]
    public void EmptyTopologyDoesNotBecomeASuccessfulMap()
    {
        byte[] bytes = Reply()[..28];
        BitConverter.GetBytes(0u).CopyTo(bytes, 16);
        Assert.False(DeviceTweakerImodRuntime.TryParseNativeTopology(bytes, bytes.Length, out _, out var topology, out string error));
        Assert.Empty(topology.ByRootPort);
        Assert.Contains("no active xHCI", error);
    }
}
