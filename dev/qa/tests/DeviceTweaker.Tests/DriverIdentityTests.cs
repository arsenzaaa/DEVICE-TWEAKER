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
    public void StartupAndApplicationKeepTheBundledDeviceInterfaceAndSeparateServiceNames()
    {
        const BindingFlags flags = BindingFlags.NonPublic | BindingFlags.Static;
        Assert.Equal("DeviceTweakerImod", DriverIdentity.ServiceName);
        Assert.Equal("DEVICE TWEAKER IMOD", DeviceTweakerImodRuntime.ImodDriverServiceDisplayName);
        Assert.Equal(DriverIdentity.DevicePath, typeof(ImodDriverNative).GetField("ImodDriverDevicePath", flags)!.GetRawConstantValue());
        string script = StartupScriptTests.ReadResource("DeviceTweakerCS.Startup.ApplyIMOD.ps1");
        Assert.Contains("-drvn DeviceTweakerImod2 -drvr DeviceTweakerImod2", script);
        Assert.Contains("$imodScriptVersion = " + ImodDriverContext.StartupScriptVersion, script);
        Assert.DoesNotContain("-drvn DTIMODMapped", script);
        using Stream stream = typeof(MainForm).Assembly.GetManifestResourceStream("DeviceTweakerCS.IMOD.DTIMOD.sys")!;
        using MemoryStream memory = new();
        stream.CopyTo(memory);
        byte[] bytes = memory.ToArray();
        foreach (string name in new[] { @"\Device\DeviceTweakerImod2", @"\DosDevices\DeviceTweakerImod2" })
            Assert.True(bytes.AsSpan().IndexOf(System.Text.Encoding.Unicode.GetBytes(name)) >= 0, name);
        Assert.True(bytes.AsSpan().IndexOf(System.Text.Encoding.Unicode.GetBytes(@"\Device\DTIMOD")) < 0);
        Assert.Equal("B4998F6ADFAFE44E1F41CE7C42BB5ADB2C634ED02E1B9E76A98499E64DB4EB52", Convert.ToHexString(SHA256.HashData(bytes)));
    }

    [Fact]
    public void ReadOnlyPciQueryReturnsBeforeTouchingAnInvalidNativeHandle()
    {
        Assert.Equal(InterruptMode.NotRead, PciInterruptReader.Query(new IntPtr(-1), @"PCI\VEN_8086&DEV_7AE0").Mode);
        Assert.Equal(InterruptMode.NotRead, PciInterruptReader.QueryLoaded(@"PCI\VEN_8086&DEV_7AE0").Mode);
    }
}
