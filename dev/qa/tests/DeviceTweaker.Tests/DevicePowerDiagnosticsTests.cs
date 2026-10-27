using System.Runtime.InteropServices;
using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public sealed class DevicePowerDiagnosticsTests
{
    private sealed class Entry : DevicePowerPolicy.IPowerEntry
    {
        internal readonly List<string> Calls = [];
        internal bool? Actual;
        internal Exception? WriteError, ReadError;
        public void Write(bool enabled) { Calls.Add($"write:{enabled}"); if (WriteError is not null) throw WriteError; }
        public bool? ReadBack() { Calls.Add("readback"); if (ReadError is not null) throw ReadError; return Actual; }
    }

    private sealed class Backend(params Entry[] entries) : DevicePowerPolicy.IPowerBackend
    {
        internal Exception? QueryError;
        internal bool Disposed;
        public IEnumerable<DevicePowerPolicy.IPowerEntry> Query(string id)
        {
            try
            {
                if (QueryError is not null) throw QueryError;
                foreach (Entry entry in entries) yield return entry;
            }
            finally { Disposed = true; }
        }
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void AppliedRequiresFreshMatchingReadbackForEveryEntry(bool requested)
    {
        Entry first = new() { Actual = requested }, second = new() { Actual = requested };
        Backend backend = new(first, second);
        Assert.Equal(DevicePowerPolicy.WriteStatus.Applied,
            DevicePowerPolicy.SetDevicePowerEnableCore("TEST", requested, backend, out string? detail));
        Assert.Null(detail);
        Assert.Equal(new[] { $"write:{requested}", "readback" }, first.Calls);
        Assert.Equal(first.Calls, second.Calls);
        Assert.True(backend.Disposed);
    }

    [Theory]
    [InlineData("query", "QueryFailed")]
    [InlineData("write", "Failed")]
    [InlineData("readback", "ReadbackUnavailable")]
    public void FailureKeepsItsStageAndCompletedWriteCount(string stage, string expected)
    {
        Entry entry = new() { Actual = false };
        Backend backend = new(entry);
        IOException failure = new("probe denial");
        if (stage == "query") backend.QueryError = failure;
        if (stage == "write") entry.WriteError = failure;
        if (stage == "readback") entry.ReadError = failure;
        var actual = DevicePowerPolicy.SetDevicePowerEnableCore("TEST", false, backend, out string? detail);
        Assert.Equal(expected, actual.ToString());
        Assert.True(DevicePowerPolicy.IsWriteFailure(actual));
        Assert.Contains($"stage={stage}", detail);
        Assert.Contains(stage == "readback" ? "written=1" : "written=0", detail);
        Assert.True(backend.Disposed);
    }

    [Theory]
    [InlineData("query")]
    [InlineData("write")]
    public void UnsupportedWmiIsNotAnAccessDeniedOrGenericWriteError(string stage)
    {
        Entry entry = new();
        Backend backend = new(entry);
        COMException unsupported = new("not supported", unchecked((int)0x8004100C));
        if (stage == "query") backend.QueryError = unsupported; else entry.WriteError = unsupported;
        var status = DevicePowerPolicy.SetDevicePowerEnableCore("TEST", false, backend, out string? detail);
        Assert.Equal(DevicePowerPolicy.WriteStatus.NotSupported, status);
        Assert.False(DevicePowerPolicy.IsWriteFailure(status));
        Assert.Contains($"stage={stage}", detail);
        Assert.DoesNotContain("readback", entry.Calls);
    }

    [Fact]
    public void NoMatchingInstanceIsNotExposed()
    {
        Assert.Equal(DevicePowerPolicy.WriteStatus.NotExposed,
            DevicePowerPolicy.SetDevicePowerEnableCore("TEST", false, new Backend(), out _));
    }

    [Theory]
    [InlineData(null, "ReadbackUnavailable")]
    [InlineData(true, "ReadbackMismatch")]
    public void AcceptedPutDoesNotProveTheRequestedValue(bool? actual, string expected)
    {
        Entry entry = new() { Actual = actual };
        Backend backend = new(entry);
        var result = DevicePowerPolicy.SetDevicePowerEnableCore("TEST", false, backend, out string? detail);
        Assert.Equal(expected, result.ToString());
        Assert.Contains("written=1 verified=0", detail);
        Assert.Equal(new[] { "write:False", "readback" }, entry.Calls);
        Assert.True(backend.Disposed);
    }

    [Fact]
    public void ASecondInstanceFailureDoesNotEraseTheFirstVerifiedWrite()
    {
        Backend backend = new(new Entry { Actual = false }, new Entry { WriteError = new UnauthorizedAccessException("denied") });
        Assert.Equal(DevicePowerPolicy.WriteStatus.Failed,
            DevicePowerPolicy.SetDevicePowerEnableCore("TEST", false, backend, out string? detail));
        Assert.Contains("matched=2 written=1 verified=1", detail);
    }
}
