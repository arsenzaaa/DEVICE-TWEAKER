using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.ExceptionServices;
using System.Windows.Forms;
using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public sealed class QaWriteGuardTests
{
    [Fact]
    public void ProcessLock_SurvivesEnvironmentChangesAndHasNoReleaseSetter()
    {
        Environment.SetEnvironmentVariable("DEVICE_TWEAKER_QA_READ_ONLY", "0");
        try
        {
            Assert.True(QaRun.IsReadOnly);
            MainForm form = FormWithoutStartup();
            var property = typeof(MainForm).GetProperty("_testAutoDryRun", BindingFlags.Instance | BindingFlags.NonPublic)!;
            if (property.CanWrite) property.SetValue(form, false);
            Assert.Equal(true, property.GetValue(form));
            Assert.Throws<QaWriteBlockedException>(() => DevicePowerPolicy.ApplyPnPCapabilities("", false));
        }
        finally { Environment.SetEnvironmentVariable("DEVICE_TWEAKER_QA_READ_ONLY", "1"); }
    }

    [Fact]
    public void FileWriter_RejectsAnExistingFileOutsideOutput_WithoutChangingIt()
    {
        string dir = Path.Combine(Path.GetTempPath(), "dt-qa-sentinel-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        string path = Path.Combine(dir, "sentinel.txt");
        File.WriteAllText(path, "unchanged");
        try
        {
            AssertBlocked("WriteAllTextAtomic", path, "changed", System.Text.Encoding.UTF8);
            AssertBlocked("WriteAllBytesAtomic", path, new byte[] { 0 });
            Assert.Equal("unchanged", File.ReadAllText(path));
            Assert.Single(Directory.GetFiles(dir));
        }
        finally { Directory.Delete(dir, recursive: true); }
    }

    [Fact]
    public void ArtifactWriter_CanWriteOnlyInsideRunDirectory()
    {
        string path = Path.Combine(QaRun.OutputDirectory, "artifact.txt");
        Method("WriteAllTextAtomic").Invoke(null, [path, "qa artifact", System.Text.Encoding.UTF8]);
        Assert.Equal("qa artifact", File.ReadAllText(path));
        File.Delete(path);
        Assert.Throws<QaWriteBlockedException>(() => QaRun.DemandArtifactPath(QaRun.OutputDirectory + "-sibling/file.txt"));
    }

    [Fact]
    public void RegistryPowerAndWmiWriters_BlockBeforeLookingUpDevices()
    {
        Assert.Throws<QaWriteBlockedException>(() => DevicePowerPolicy.ApplyPnPCapabilities("", false));
        Assert.Throws<QaWriteBlockedException>(() => DevicePowerPolicy.SetDevicePowerEnable("", false, out _));
        Assert.Throws<QaWriteBlockedException>(() => UsbChipPath.SetSelectiveSuspendEnabled("", false));
        Assert.Throws<QaWriteBlockedException>(() => UsbSelectiveSuspendPolicy.SetDeviceParameterDword("", "", 0));
        Assert.Throws<QaWriteBlockedException>(() => UsbSelectiveSuspendPolicy.ApplyControllerAndHubs("", false));
        Assert.Throws<QaWriteBlockedException>(() => UsbSelectiveSuspendPolicy.SetPowerPlanEnabled(false));
        Assert.Throws<QaWriteBlockedException>(() => UsbSelectiveSuspendPolicy.ActivateCurrentPowerScheme());
        Assert.Throws<QaWriteBlockedException>(() => UsbSelectiveSuspendPolicy.RestorePowerPlanIndex(Guid.Empty, true, 0));
        Assert.Throws<QaWriteBlockedException>(() => new WindowsRegistryValues(null!).Set("", 0, Microsoft.Win32.RegistryValueKind.DWord));
        Assert.Throws<QaWriteBlockedException>(() => new WindowsRegistryValues(null!).Delete(""));
    }

    [Fact]
    public void DeviceApplyAndReset_RejectRealDevicesBeforeReadingUiOrRegistry()
    {
        DeviceBlock block = (DeviceBlock)RuntimeHelpers.GetUninitializedObject(typeof(DeviceBlock));
        typeof(DeviceBlock).GetProperty("Device")!.SetValue(block, new DeviceInfo
        {
            Name = "QA sentinel", InstanceId = "", Class = "Display", RegBase = "",
            Kind = DeviceKind.GPU, IsTestDevice = false
        });
        AssertBlocked("SaveBlockSettings", block, false, null);
        AssertBlocked("ResetBlockSettings", block, null);
        AssertBlocked("WriteNdisIrqPolicy", null, null);
        AssertBlocked("ClearNdisIrqPolicy", null, null);
    }

    [Fact]
    public void RestoreAndStartupWriters_BlockBeforeReadingBackupOrDereferencingArguments()
    {
        AssertBlocked("RestoreDeviceTweakerBackup", "missing-backup.json");
        AssertBlocked("RestoreRegistryValue", (object?)null);
        AssertBlocked("RestoreFileBackup", (object?)null);
        AssertBlocked("DeleteDeviceTweakerBackups", (object?)null);
        AssertBlocked("PruneDeviceTweakerBackups", "", 0);
        AssertBlocked("WriteImodScript", null, "");
        AssertBlocked("UpgradeManagedImodStartupForRename");
        AssertBlocked("DeleteImodStartupRegistration", (object?)null);
        AssertBlocked("RemoveImodPersistenceFiles", (object?)null);
        AssertBlocked("SetRawMouseThrottle", false, 0);
        AssertBlocked("SetReservedCpuSets", (object?)null);
        AssertBlocked("SetNdisBaseCore", "", 0, null, null);
        AssertBlocked("SetNdisRssQueues", "", 0, null, null);
        AssertBlocked("SetNdisRssExtraValues", "", null, null, null);
        AssertBlocked("WriteNicDriverModeration", "", null, "OFF", null);
        AssertBlocked("ClearNdisBaseCore", "", null, null);
        AssertBlocked("ClearNdisRssQueues", "", null, null);
        AssertBlocked("ClearNdisRssExtraValues", "", null, null);
        AssertBlocked("TryDisableVulnerableDriverBlocklist", (object?)null);
    }

    [Fact]
    public void DriverStagingAndInitialization_AreSkippedBeforeAnyDriverAccess()
    {
        object?[] args = [false, null, null];
        Assert.Equal(false, Method("EnsureImodDriverOnDisk").Invoke(FormWithoutStartup(), args));
        Assert.Equal(string.Empty, args[1]);
        Assert.Contains("QA read-only", (string)args[2]!);
        Type ctx = typeof(MainForm).GetNestedType("ImodDriverContext", BindingFlags.NonPublic)!;
        object?[] initializeArgs = ["does-not-exist.sys", null, null, null];
        Assert.Equal(false, ctx.GetMethod("TryInitialize")!.Invoke(null, initializeArgs));
        Assert.Null(initializeArgs[2]);
        Assert.Contains("QA read-only", (string)initializeArgs[3]!);
        AssertBlocked("TryReadPhysicalMemory", null, 0UL, 4U, null, null);
        AssertBlocked("TryWritePhysicalMemory", null, 0UL, 4U, 0UL, null);
        AssertBlocked("RunImodKduMap", "", "", null, null);
    }

    [Fact]
    public void RealNicItrSave_IsBlockedAtTheUiEntryBeforeCreatingBackup()
    {
        RunSta(() =>
        {
            using TextBox input = new() { Text = "8" };
            using Label status = new();
            DeviceBlock block = (DeviceBlock)RuntimeHelpers.GetUninitializedObject(typeof(DeviceBlock));
            typeof(DeviceBlock).GetProperty("Device")!.SetValue(block, new DeviceInfo
            {
                Name = "QA sentinel", InstanceId = @"PCI\VEN_8086&DEV_15F3", Class = "Net",
                RegBase = "", Kind = DeviceKind.NET_NDIS, IsTestDevice = false
            });
            typeof(DeviceBlock).GetProperty("NicItrBox")!.SetValue(block, input);
            typeof(DeviceBlock).GetProperty("NicItrStatusLabel")!.SetValue(block, status);
            int before = QaRun.BlockedCount;
            Method("SaveNicItrPersistenceFromBlock").Invoke(FormWithoutStartup(), [block]);
            Assert.Equal(before + 1, QaRun.BlockedCount);
            Assert.Equal("NIC ITR SAVE", QaRun.BlockedOperations[^1].Operation);
            Assert.DoesNotContain("saved for startup", status.Text);
        });
    }

    private static MethodInfo Method(string name) => typeof(MainForm).GetMethod(name,
        BindingFlags.Instance | BindingFlags.Static | BindingFlags.NonPublic)!;

    private static MainForm FormWithoutStartup()
    {
        // No MainForm constructor, WMI scan, window, or OnShown startup.
        MainForm form = (MainForm)RuntimeHelpers.GetUninitializedObject(typeof(MainForm));
        return form;
    }

    private static void AssertBlocked(string name, params object?[] args)
    {
        MethodInfo method = Method(name);
        var exception = Assert.Throws<TargetInvocationException>(() =>
            method.Invoke(method.IsStatic ? null : FormWithoutStartup(), args));
        Assert.IsType<QaWriteBlockedException>(exception.InnerException);
    }

    private static void RunSta(Action action)
    {
        Exception? failure = null;
        Thread thread = new(() => { try { action(); } catch (Exception ex) { failure = ex; } });
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        Assert.True(thread.Join(TimeSpan.FromSeconds(15)), "STA check timed out");
        if (failure is not null) ExceptionDispatchInfo.Capture(failure).Throw();
    }
}
