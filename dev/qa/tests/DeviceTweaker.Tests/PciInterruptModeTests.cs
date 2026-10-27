using System.Buffers.Binary;
using System.Text;
using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public sealed class PciInterruptModeTests
{
    private const string Id = @"PCI\VEN_8086&DEV_7AE0&SUBSYS_00000000\EXACT_INSTANCE";
    private static void U32(byte[] bytes, int offset, uint value) => BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(offset), value);
    private static void U16(byte[] bytes, int offset, ushort value) => BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(offset), value);

    private static byte[] Reply(bool msi = false, bool msiX = false, bool line = false, int messages = 1)
    {
        int count = (line ? 1 : 0) + messages;
        byte[] data = new byte[PciInterruptMode.ReplyHeaderSize + 20 + count * 20];
        U32(data, 0, 1); U32(data, 12, 256); U32(data, 16, (uint)(data.Length - PciInterruptMode.ReplyHeaderSize));
        U16(data, 20, 0x8086); U16(data, 22, 0x7AE0); U16(data, 26, 0x10);
        data[20 + 0x34] = 0x40;
        data[20 + 0x40] = 5; data[20 + 0x41] = 0x50;
        U16(data, 20 + 0x42, (ushort)(0x80 | (msi ? 1 : 0))); // 64-bit capable is not Enable.
        data[20 + 0x50] = 0x11;
        U16(data, 20 + 0x52, (ushort)(7 | (msiX ? 0x8000 : 0))); // Table size is not Enable.
        int resource = PciInterruptMode.ReplyHeaderSize;
        U32(data, resource, 1); U32(data, resource + 16, (uint)count);
        for (int i = 0; i < count; i++)
        {
            data[resource + 20 + i * 20] = 2;
            U16(data, resource + 22 + i * 20, (ushort)(line && i == 0 ? 0 : 2));
        }
        return data;
    }

    [Fact]
    public void SingleInterruptDistinguishesMsiAndMsiXByEnableBits()
    {
        Assert.Equal(InterruptMode.Msi, PciInterruptMode.DecodeReply(Id, Reply(msi: true)).Mode);
        Assert.Equal(InterruptMode.MsiX, PciInterruptMode.DecodeReply(Id, Reply(msiX: true)).Mode);
    }

    [Fact]
    public void SupportedCapabilitiesDoNotBecomeTheCurrentMode()
        => Assert.Equal(InterruptMode.Line, PciInterruptMode.DecodeReply(Id, Reply(line: true, messages: 0)).Mode);

    [Theory]
    [InlineData(1)] [InlineData(8)] [InlineData(32)]
    public void NumberOfDescriptorsDoesNotChooseMsiX(int descriptors)
        => Assert.Equal(InterruptMode.Msi, PciInterruptMode.DecodeReply(Id, Reply(msi: true, messages: descriptors)).Mode);

    [Fact]
    public void MsiXFunctionMaskDoesNotChangeTheEnabledMode()
    {
        byte[] data = Reply(msiX: true); U16(data, 20 + 0x52, 0xC007);
        Assert.Equal(InterruptMode.MsiX, PciInterruptMode.DecodeReply(Id, data).Mode);
    }

    [Theory]
    [InlineData(true, true, false, 1)]
    [InlineData(true, false, true, 0)]
    [InlineData(false, true, true, 0)]
    [InlineData(false, false, false, 1)]
    [InlineData(true, false, true, 1)]
    public void ConflictingEvidenceIsNeverReportedAsConfirmed(bool msi, bool msiX, bool line, int messages)
        => Assert.Equal(InterruptMode.Unavailable, PciInterruptMode.DecodeReply(Id, Reply(msi, msiX, line, messages)).Mode);

    [Fact]
    public void EnabledPciBitWithoutAllocatedInterruptsIsNoIrq()
        => Assert.Equal(InterruptMode.None, PciInterruptMode.DecodeReply(Id, Reply(msiX: true, messages: 0)).Mode);

    [Theory]
    [InlineData(false, "Message")] [InlineData(true, "Line")]
    public void MissingPciReadPreservesOnlyWhatAllocationProves(bool line, string expected)
    {
        byte[] data = Reply(line: line, messages: line ? 0 : 1); U32(data, 4, 0xC00000BB); U32(data, 12, 0);
        var result = PciInterruptMode.DecodeReply(Id, data);
        Assert.Equal(expected, result.Mode.ToString()); Assert.Contains("0xC00000BB", result.Detail);
    }

    [Fact]
    public void FailedResourceReadDoesNotPromotePciStateToActiveMode()
    {
        byte[] data = Reply(msiX: true); U32(data, 8, 0xC0000022);
        var result = PciInterruptMode.DecodeReply(Id, data);
        Assert.Equal(InterruptMode.Unavailable, result.Mode); Assert.Contains("0xC0000022", result.Detail);
    }

    [Theory]
    [InlineData(0)] [InlineData(0xFFFF)] [InlineData(0x10DE)]
    public void WrongOrRemovedDeviceIsRejected(int vendor)
    {
        byte[] data = Reply(msiX: true); U16(data, 20, (ushort)vendor);
        Assert.Equal(InterruptMode.Unavailable, PciInterruptMode.DecodeReply(Id, data).Mode);
    }

    [Theory]
    [InlineData(0x40)] [InlineData(0x20)] [InlineData(0xFF)] [InlineData(0x43)]
    public void CapabilityCyclesAndBadPointersAreRejected(int next)
    {
        byte[] data = Reply(msiX: true); data[20 + 0x51] = (byte)next;
        Assert.Equal(InterruptMode.Unavailable, PciInterruptMode.DecodeReply(Id, data).Mode);
    }

    [Fact]
    public void DuplicateCapabilityIsRejected()
    {
        byte[] data = Reply(msi: true); data[20 + 0x50] = 5;
        Assert.Equal(InterruptMode.Unavailable, PciInterruptMode.DecodeReply(Id, data).Mode);
    }

    [Theory]
    [InlineData(0xFF)] [InlineData(0)]
    public void UnreadableCapabilityIsRejected(int id)
    {
        byte[] data = Reply(msi: true); data[20 + 0x50] = (byte)id;
        Assert.Equal(InterruptMode.Unavailable, PciInterruptMode.DecodeReply(Id, data).Mode);
    }

    [Fact]
    public void BridgeHeaderIsSupportedAndCardBusHeaderIsExplicitlyRejected()
    {
        byte[] data = Reply(msiX: true); data[20 + 0x0E] = 0x81;
        Assert.Equal(InterruptMode.MsiX, PciInterruptMode.DecodeReply(Id, data).Mode);
        data[20 + 0x0E] = 2;
        Assert.Equal(InterruptMode.Unavailable, PciInterruptMode.DecodeReply(Id, data).Mode);
    }

    [Fact]
    public void TruncatedAndOversizedProtocolLengthsAreRejected()
    {
        byte[] data = Reply(msi: true);
        Assert.Equal(InterruptMode.Unavailable, PciInterruptMode.DecodeReply(Id, data.AsSpan(0, 100)).Mode);
        Assert.Equal(InterruptMode.Unavailable, PciInterruptMode.DecodeReply(Id, data.AsSpan(0, data.Length - 1)).Mode);
        U32(data, 16, uint.MaxValue);
        Assert.Equal(InterruptMode.Unavailable, PciInterruptMode.DecodeReply(Id, data).Mode);
        U32(data, 16, 0); U32(data, 12, 257);
        Assert.Equal(InterruptMode.Unavailable, PciInterruptMode.DecodeReply(Id, data).Mode);
    }

    [Fact]
    public void ProtocolVersionIsChecked()
    {
        byte[] data = Reply(msi: true); U32(data, 0, 2);
        Assert.Equal(InterruptMode.Unavailable, PciInterruptMode.DecodeReply(Id, data).Mode);
    }

    [Fact]
    public void ResourceCountsAndVariableDataAreBoundsChecked()
    {
        byte[] data = Reply(msi: true); int r = PciInterruptMode.ReplyHeaderSize;
        U32(data, r + 16, uint.MaxValue);
        Assert.Equal(InterruptMode.Unavailable, PciInterruptMode.DecodeReply(Id, data).Mode);
        U32(data, r + 16, 1); data[r + 20] = 5; U32(data, r + 24, 40);
        Assert.Equal(InterruptMode.Unavailable, PciInterruptMode.DecodeReply(Id, data).Mode);
    }

    [Fact]
    public void DeviceSpecificDataAndMultipleFullDescriptorsAreParsed()
    {
        byte[] resources = new byte[4 + 16 + 20 + 4 + 16 + 20]; U32(resources, 0, 2);
        U32(resources, 16, 1); resources[20] = 5; U32(resources, 24, 4);
        U32(resources, 44 + 12, 1); resources[44 + 16] = 2; U16(resources, 44 + 18, 2);
        Assert.True(PciInterruptMode.TryReadAllocatedInterrupts(resources, out bool line, out bool message, out string error), error);
        Assert.False(line); Assert.True(message);
    }

    [Fact]
    public void RawAndTranslatedListsMustAgreeAndConsumeAllBytes()
    {
        byte[] reply = Reply(msi: true);
        byte[] raw = reply[PciInterruptMode.ReplyHeaderSize..];
        byte[] pair = raw.Concat(raw).ToArray();
        Assert.True(PciInterruptMode.TryReadAllocatedInterrupts(pair, out _, out bool message, out _));
        Assert.True(message);
        U16(pair, raw.Length + 22, 0);
        Assert.False(PciInterruptMode.TryReadAllocatedInterrupts(pair, out _, out _, out string disagreement));
        Assert.Contains("disagree", disagreement);
        Assert.False(PciInterruptMode.TryReadAllocatedInterrupts(raw.Concat(raw).Concat(raw).ToArray(), out _, out _, out _));
        Assert.False(PciInterruptMode.TryReadAllocatedInterrupts(raw.Concat(raw[..^1]).ToArray(), out _, out _, out _));
        Assert.False(PciInterruptMode.TryReadAllocatedInterrupts(raw.Concat(new byte[] { 1, 2, 3 }).ToArray(), out _, out _, out _));
    }

    [Fact]
    public void RequestIncludesCompleteIdentityAndTerminatedUnicodeNames()
    {
        Assert.True(PciInterruptReader.TryBuildRequest(@"\Device\NTPNP_PCI0001", Id, out byte[] request));
        Assert.Equal(1028, request.Length); Assert.Equal(1u, BinaryPrimitives.ReadUInt32LittleEndian(request));
        Assert.Equal(Id, Encoding.Unicode.GetString(request, 516, 512).TrimEnd('\0'));
        Assert.False(PciInterruptReader.TryBuildRequest(@"\Device\X", Id + '\0', out _));
        Assert.False(PciInterruptReader.TryBuildRequest(@"\Device\X", "USB\\VID_1234", out _));
        Assert.False(PciInterruptReader.TryBuildRequest("C:\\X", Id, out _));
    }

    [Fact]
    public void QaNeverOpensOrLoadsDriverForModeQuery()
    {
        Assert.True(QaRun.IsReadOnly);
        Assert.Equal(InterruptMode.NotRead, PciInterruptReader.Query(new IntPtr(-1), Id).Mode);
        Assert.Equal(InterruptMode.NotRead, PciInterruptReader.QueryLoaded(Id).Mode);
    }

    [Fact]
    public void NonPciCapabilitiesNeverComeFromAPciParent()
    {
        Assert.False(NativeCfgMgr32.TryGetPciInterruptCapabilities(@"HID\VID_1234&PID_5678\X", out var capabilities));
        Assert.Null(capabilities);
    }
}
