using System.Text.Json;
using System.Buffers.Binary;
using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public sealed class WindowsResourceCaptureTests
{
    private sealed record Capture(string Name, string InstanceId, string ExpectedMode, string ConfigBase64, string ResourcesBase64);
    [Theory]
    [InlineData(0)] [InlineData(1)] [InlineData(2)] [InlineData(3)] [InlineData(4)] [InlineData(5)]
    public void ActualWindowsRawAndTranslatedListsAreBothDecoded(int index)
    {
        Capture[] captures = JsonSerializer.Deserialize<Capture[]>(File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "WindowsResourceCaptures.json")))!;
        Assert.Equal(6, captures.Length);
        Capture capture = captures[index];
        byte[] config = Convert.FromBase64String(capture.ConfigBase64), resources = Convert.FromBase64String(capture.ResourcesBase64);
        byte[] reply = new byte[PciInterruptMode.ReplyHeaderSize + resources.Length];
        BinaryPrimitives.WriteUInt32LittleEndian(reply, 1);
        BinaryPrimitives.WriteUInt32LittleEndian(reply.AsSpan(12), (uint)config.Length);
        BinaryPrimitives.WriteUInt32LittleEndian(reply.AsSpan(16), (uint)resources.Length);
        config.CopyTo(reply, 20); resources.CopyTo(reply, PciInterruptMode.ReplyHeaderSize);
        InterruptModeResult result = PciInterruptMode.DecodeReply(capture.InstanceId, reply);
        Assert.True(result.Mode.ToString() == capture.ExpectedMode, $"{capture.Name}: {result.Mode} / {result.Detail}");
    }
}
