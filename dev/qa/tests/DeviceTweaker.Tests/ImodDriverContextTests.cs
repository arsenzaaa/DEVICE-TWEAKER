using System.Reflection;
using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public sealed class ImodDriverContextTests
{
    [Fact]
    public void DiagnosticVersionMatchesTheEmbeddedStartupScript()
    {
        string script = StartupScriptTests.ReadResource("DeviceTweakerCS.Startup.ApplyIMOD.ps1");
        var version = System.Text.RegularExpressions.Regex.Match(script, @"\$imodScriptVersion\s*=\s*(\d+)");
        Assert.True(version.Success);
        Assert.Equal(ImodDriverContext.StartupScriptVersion.ToString(System.Globalization.CultureInfo.InvariantCulture), version.Groups[1].Value);
    }

    [Fact]
    public void NativeDeviceOpenBoundaryAlsoRefusesReadOnlyMode()
    {
        Assert.False(ImodDriverNative.TryOpenImodDriverDevice(out var handle, out var error));
        Assert.Equal(new IntPtr(-1), handle);
        Assert.Contains("QA read-only", error);
        Assert.False(ImodDriverNative.TryOpenImodDriverDeviceOnce(out handle, out error));
        Assert.Equal(new IntPtr(-1), handle);
        Assert.Contains("QA read-only", error);
    }

    private static ImodDriverContext Create(Action<string> log) => (ImodDriverContext)Activator.CreateInstance(
        typeof(ImodDriverContext), BindingFlags.Instance | BindingFlags.NonPublic, null,
        new object?[] { "unused-driver-path", log }, null)!;

    [Fact]
    public void DisposingAnUnownedSessionIsIdempotent()
    {
        int notifications = 0;
        var session = Create(_ => notifications++);
        session.Dispose();
        session.Dispose();
        Assert.Equal(1, notifications);
        Assert.Equal(new IntPtr(-1), session.DriverHandle);
        Assert.False(session.ServiceCreated);
        Assert.False(session.ServiceStartedByContext);
    }

    [Fact]
    public void AFailingLogObserverDoesNotMakeDisposalRepeatable()
    {
        int notifications = 0;
        var session = Create(_ => { notifications++; throw new IOException("observer failure"); });
        Assert.Throws<IOException>(session.Dispose);
        session.Dispose();
        Assert.Equal(1, notifications);
        Assert.Equal(new IntPtr(-1), session.DriverHandle);
    }

    [Fact]
    public void ReadOnlyInitializationStopsBeforeCallingTheSuppliedLoader()
    {
        bool called = false;
        bool Load(string path, Action<string>? log, out string? error)
        { called = true; throw new InvalidOperationException("Must not reach a backend in QA"); }
        Assert.False(ImodDriverContext.TryInitialize("nonexistent.sys", null, Load, out var session, out var error));
        Assert.Null(session);
        Assert.Contains("QA read-only", error);
        Assert.False(called);
    }
}
