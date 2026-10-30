using System.Reflection;
using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public sealed class DriverRenameTests
{
    [Fact] public void BundledDriverContainsThePinnedKernelObjects()
    {
        using Stream stream = typeof(MainForm).Assembly.GetManifestResourceStream("DeviceTweakerCS.IMOD.DTIMOD.sys")!;
        using MemoryStream memory = new(); stream.CopyTo(memory); byte[] bytes = memory.ToArray();
        foreach (string name in new[] { @"\Device\DeviceTweakerImod2", @"\DosDevices\DeviceTweakerImod2" })
            Assert.True(bytes.AsSpan().IndexOf(System.Text.Encoding.Unicode.GetBytes(name)) >= 0, name);
        Assert.True(bytes.AsSpan().IndexOf(System.Text.Encoding.Unicode.GetBytes(@"\Device\DTIMOD")) < 0);
    }
    private sealed class Backend : IDriverRenameBackend
    {
        public bool Current = true, Legacy;
        public LegacyDriverService Service = new() { Exists = true, Stopped = true, Owned = true };
        public int Reads, Deletes;
        public bool FailDelete;
        public bool DeviceAvailable(string path) => path == DriverRename.DevicePath ? Current : Legacy;
        public LegacyDriverService ReadLegacyService() { Reads++; return Service; }
        public void DeleteStoppedLegacyService() { if (FailDelete) throw new IOException("SCM denied delete"); Deletes++; }
    }

    [Fact] public void ApplicationAndStartupUseTheSameIdentity()
    {
        const BindingFlags flags = BindingFlags.NonPublic | BindingFlags.Static;
        Assert.Equal("DeviceTweakerImod2", typeof(MainForm).GetField("ImodDriverServiceName", flags)!.GetRawConstantValue());
        Assert.Equal(@"\\.\DeviceTweakerImod2", typeof(MainForm).GetField("ImodDriverDevicePath", flags)!.GetRawConstantValue());
        string script = StartupScriptTests.ReadResource("DeviceTweakerCS.Startup.ApplyIMOD.ps1");
        Assert.Contains("-drvn DeviceTweakerImod2 -drvr DeviceTweakerImod2", script);
        Assert.DoesNotContain("-drvn DTIMODMapped", script);
    }

    [Fact] public void ReadOnlyBlocksBeforeAnyBackendAccess()
    {
        Assert.Contains("QA read-only", Assert.Throws<InvalidOperationException>(() => DriverRename.Prepare(null!, true)).Message);
        Assert.Contains("QA read-only", Assert.Throws<InvalidOperationException>(() => DriverRename.Complete(null!, true)).Message);
    }

    [Fact] public void NativeBackendCannotBypassImmutableQaMode()
    {
        IDriverRenameBackend native = DriverRename.Native(Path.Combine(Path.GetTempPath(), "DTIMOD.sys"));
        Assert.Contains("QA read-only", Assert.Throws<InvalidOperationException>(() => native.DeviceAvailable(DriverRename.DevicePath)).Message);
        Assert.Contains("QA read-only", Assert.Throws<InvalidOperationException>(() => native.ReadLegacyService()).Message);
        Assert.Contains("QA read-only", Assert.Throws<InvalidOperationException>(() => native.DeleteStoppedLegacyService()).Message);
    }

    [Fact] public void LoadedOldDriverRequiresRestartWithoutRemovingService()
    {
        Backend b = new() { Legacy = true };
        Assert.Equal(DriverRename.RestartMessage, Assert.Throws<InvalidOperationException>(() => DriverRename.Prepare(b, false)).Message);
        Assert.Equal(0, b.Reads); Assert.Equal(0, b.Deletes);
    }

    [Fact] public void AvailableDeviceRetainsOtherServices()
    {
        Backend b = new(); DriverRename.Prepare(b, false); DriverRename.Complete(b, false);
        Assert.Equal(0, b.Deletes); Assert.Equal(0, b.Reads);
    }

    [Fact] public void UnrelatedRunningServiceIsNotTouched()
    {
        Backend b = new() { Service = new() { Exists = true, Owned = true, Stopped = false } };
        DriverRename.Prepare(b, false);
        Assert.Equal(0, b.Reads); Assert.Equal(0, b.Deletes);
    }

    [Fact] public void MissingNewDevicePreservesLegacyService()
    {
        Backend b = new() { Current = false };
        Assert.Throws<InvalidOperationException>(() => DriverRename.Complete(b, false));
        Assert.Equal(0, b.Reads); Assert.Equal(0, b.Deletes);
    }

    [Theory]
    [InlineData(true, true, false)] [InlineData(true, false, true)] [InlineData(false, true, true)]
    public void UnownedRunningOrMissingLegacyServiceIsNotDeleted(bool exists, bool stopped, bool owned)
    {
        Backend b = new() { Service = new() { Exists = exists, Stopped = stopped, Owned = owned } };
        DriverRename.Complete(b, false); Assert.Equal(0, b.Deletes);
    }

    [Fact] public void CompletionDoesNotInvokeServiceDeletion()
    {
        Backend b = new() { FailDelete = true }; DriverRename.Complete(b, false); Assert.Equal(0, b.Reads); Assert.Equal(0, b.Deletes);
    }

}
