using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public sealed class LegacyServiceCleanupTests
{
    [Theory] [InlineData("StopServiceIfNeeded")] [InlineData("CleanupDriverArtifacts")]
    public void NativeCleanupRefusesImmutableQaModeBeforeCallingWindows(string method)
    {
        const System.Reflection.BindingFlags flags = System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance;
        Type type = typeof(MainForm).GetNestedType("ImodDriverContext", System.Reflection.BindingFlags.NonPublic)!;
        object context = Activator.CreateInstance(type, flags, null, new object?[] { "unused-driver-path", null }, null)!;
        object?[] args = method == "StopServiceIfNeeded" ? [] : ["unused-driver-path"];
        var error = Assert.Throws<System.Reflection.TargetInvocationException>(() => type.GetMethod(method, flags)!.Invoke(context, args));
        Assert.Contains("QA read-only", error.InnerException!.Message);
    }
    [Theory]
    [InlineData(false, false, false, "close,release")]
    [InlineData(false, false, true, "close,release")]
    [InlineData(true, false, false, "close,stop,release")]
    [InlineData(false, true, false, "close,stop,release")]
    [InlineData(true, false, true, "close,stop,cleanup,release")]
    [InlineData(false, true, true, "close,stop,cleanup,release")]
    [InlineData(true, true, true, "close,stop,cleanup,release")]
    public void CleanupUsesActualOwnershipAndClosesHandleFirst(bool created, bool started, bool initialized, string expected)
    {
        List<string> events = new();
        LegacyServiceCleanup.Run(created, started, initialized, () => events.Add("close"),
            () => events.Add("stop"), () => events.Add("cleanup"), () => events.Add("release"));
        Assert.Equal(expected, string.Join(',', events));
    }
    [Theory] [InlineData("close")] [InlineData("stop")] [InlineData("cleanup")]
    public void FailureStillReleasesGateAndIsNotReportedAsSuccess(string stage)
    {
        bool released = false;
        void Step(string current) { if (stage == current) throw new IOException(current); }
        Assert.Throws<IOException>(() => LegacyServiceCleanup.Run(true, true, true,
            () => Step("close"), () => Step("stop"), () => Step("cleanup"), () => released = true));
        Assert.True(released);
    }
}
