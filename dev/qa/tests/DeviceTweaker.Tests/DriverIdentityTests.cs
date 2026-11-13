using System.Reflection;
using System.Security.Cryptography;
using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public sealed class DriverIdentityTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LegacyDeviceConflictIsCheckedBeforeLoading(bool loaded)
    {
        List<string> requests = [];
        bool Available(string path) { requests.Add(path); return loaded; }
        if (loaded)
            Assert.Equal(DriverIdentity.RestartMessage,
                Assert.Throws<InvalidOperationException>(() => DriverIdentity.CheckLegacyDevice(Available)).Message);
        else
            DriverIdentity.CheckLegacyDevice(Available);
        Assert.Equal(new[] { DriverIdentity.LegacyDevicePath }, requests);
    }

    [Fact]
    public void NativePreflightCannotBypassImmutableQaMode()
        => Assert.Throws<QaWriteBlockedException>(() => DriverIdentity.CheckLegacyDevice());

    [Fact]
    public void DeviceProbeFailureIsNotReportedAsAbsence()
        => Assert.Throws<IOException>(() => DriverIdentity.CheckLegacyDevice(_ => throw new IOException("access failed")));

    [Fact]
    public void ApplicationStartupAndBundledDriverUseTheSameIdentity()
    {
        const BindingFlags flags = BindingFlags.NonPublic | BindingFlags.Static;
        Assert.Equal(DriverIdentity.ServiceName, typeof(ImodDriverNative).GetField("ImodDriverServiceName", flags)!.GetRawConstantValue());
        Assert.Equal(DriverIdentity.DevicePath, typeof(ImodDriverNative).GetField("ImodDriverDevicePath", flags)!.GetRawConstantValue());
        string script = StartupScriptTests.ReadResource("DeviceTweakerCS.Startup.ApplyIMOD.ps1");
        Assert.Contains("-drvn " + DriverIdentity.ServiceName + " -drvr " + DriverIdentity.ServiceName, script);
        Assert.Contains("$imodScriptVersion = " + ImodDriverContext.StartupScriptVersion, script);
        Assert.DoesNotContain("-drvn DTIMODMapped", script);
        using Stream stream = typeof(MainForm).Assembly.GetManifestResourceStream("DeviceTweakerCS.IMOD.DTIMOD.sys")!;
        using MemoryStream memory = new();
        stream.CopyTo(memory);
        byte[] bytes = memory.ToArray();
        foreach (string name in new[] { @"\Device\DeviceTweakerImod2", @"\DosDevices\DeviceTweakerImod2" })
            Assert.True(bytes.AsSpan().IndexOf(System.Text.Encoding.Unicode.GetBytes(name)) >= 0, name);
        Assert.True(bytes.AsSpan().IndexOf(System.Text.Encoding.Unicode.GetBytes(@"\Device\DTIMOD")) < 0);
        Assert.Equal("F2BD92D5455F64CEB764CC4EFA4EA2D7F83DB6F7FDF8EE5FFA7247D733EA59FD", Convert.ToHexString(SHA256.HashData(bytes)));
    }

    [Fact]
    public void ReadOnlyPciQueryReturnsBeforeTouchingAnInvalidNativeHandle()
    {
        Assert.Equal(InterruptMode.NotRead, PciInterruptReader.Query(new IntPtr(-1), @"PCI\VEN_8086&DEV_7AE0").Mode);
        Assert.Equal(InterruptMode.NotRead, PciInterruptReader.QueryLoaded(@"PCI\VEN_8086&DEV_7AE0").Mode);
    }
}
