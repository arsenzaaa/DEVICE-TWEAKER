using System.Reflection;
using System.Security.Cryptography;
using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public sealed class LegacyCompatibilityTests
{
    private sealed class Backend : IDriverRenameBackend
    {
        public bool NewDriverLoaded;
        public bool OldDriverLoaded = true;
        public bool DeviceAvailable(string path) => path == @"\\.\DTIMOD" ? NewDriverLoaded : OldDriverLoaded;
        public LegacyDriverService ReadLegacyService() => throw new Exception("Compatibility mode must not inspect or migrate services");
        public void DeleteStoppedLegacyService() => throw new Exception("Compatibility mode must not delete services");
    }

    [Fact] public void NewDriverAlreadyLoadedStopsBeforeAnyServiceAction()
    {
        Assert.Throws<InvalidOperationException>(() => DriverRename.Prepare(new Backend { NewDriverLoaded = true }, false));
    }
    [Fact] public void OldDeviceCanBeUsedWithoutServiceMigration()
    {
        Backend backend = new();
        DriverRename.Prepare(backend, false);
        Assert.Contains("migration disabled", DriverRename.Complete(backend, false));
    }
    [Fact] public void MissingOldDeviceIsNotReportedReady()
    {
        Assert.Throws<InvalidOperationException>(() => DriverRename.Complete(new Backend { OldDriverLoaded = false }, false));
    }
    [Fact] public void ReadOnlyPciQueryReturnsBeforeTouchingAnInvalidNativeHandle()
    {
        Assert.Equal(InterruptMode.NotRead, PciInterruptReader.Query(new IntPtr(-1), @"PCI\VEN_8086&DEV_7AE0").Mode);
        Assert.Equal(InterruptMode.NotRead, PciInterruptReader.QueryLoaded(@"PCI\VEN_8086&DEV_7AE0").Mode);
    }
    [Fact] public void BundledSysMatchesThePinnedPayload()
    {
        using Stream stream = typeof(MainForm).Assembly.GetManifestResourceStream("DeviceTweakerCS.IMOD.DTIMOD.sys")!;
        Assert.Equal("DBBF0081B5D2E931B5347C16CA457BDBC3988740C07C0224119749B6D26CA3A8", Convert.ToHexString(SHA256.HashData(stream)));
    }
    [Fact] public void StartupAndApplicationAddressTheSameOldDevice()
    {
        const BindingFlags flags = BindingFlags.NonPublic | BindingFlags.Static;
        Assert.Equal("DeviceTweakerImod2", typeof(MainForm).GetField("ImodDriverServiceName", flags)!.GetRawConstantValue());
        Assert.Equal(@"\\.\DeviceTweakerImod2", typeof(MainForm).GetField("ImodDriverDevicePath", flags)!.GetRawConstantValue());
        string script = StartupScriptTests.ReadResource("DeviceTweakerCS.Startup.ApplyIMOD.ps1");
        Assert.Contains("-drvn DeviceTweakerImod2 -drvr DeviceTweakerImod2", script);
        Assert.Contains("$imodScriptVersion = 31", script);
        Assert.DoesNotContain("-drvn DTIMODMapped", script);
    }
}
