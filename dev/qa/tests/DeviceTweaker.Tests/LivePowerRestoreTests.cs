using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public sealed class LivePowerRestoreTests
{
    private const string Id = @"PCI\USB";
    private const string Class = @"SYSTEM\CurrentControlSet\Control\Class\{test}\0000";
    private static RegistryValueBackup Value(string path, string name, bool exists, string? data = null, string kind = "DWord")
        => new() { Hive = "HKLM", Path = path, Name = name, Exists = exists, Kind = kind, Data = data };
    private static OperationReport Run(IEnumerable<LivePowerRestoreTarget> targets,
        Func<string, bool, LivePowerWriteResult> write, Action? activate = null)
        => LivePowerRestoreService.Restore(targets, write, activate ?? (() => { }), _ => { });

    [Fact]
    public void RestoredPersistenceIsPassedToWriterAndLiveReadbackFailureMakesParentPartial()
    {
        var backup = new DeviceTweakerBackup { RegistryValues = [Value(Class, "PnPCapabilities", true, "280")] };
        bool? requested = null;
        bool? enabled = LivePowerRestoreService.ReadRestoredEnabled(backup, Id, Class, usb: true);
        var child = Run([new(Id, enabled)], (_, value) => { requested = value; return new(DevicePowerPolicy.WriteStatus.ReadbackMismatch, "still enabled"); });
        var parent = new OperationReport();
        parent.AddSuccess("Registry");
        parent.MergeIssues(child);
        Assert.False(requested);
        Assert.False(parent.Succeeded);
        Assert.False(parent.NoChangesMade);
        Assert.Contains("ReadbackMismatch", Assert.Single(parent.Issues.Where(i => i.Severity == OperationIssueSeverity.Error)).TechnicalDetails);
    }

    [Theory]
    [InlineData(false, false, true)]
    [InlineData(true, false, false)]
    public void UsbParametersDetermineTargetWhenClassValueWasAbsent(bool selectiveExists, bool enhancedExists, bool expected)
    {
        string path = UsbSelectiveSuspendPolicy.DeviceParametersPath(Id);
        var backup = new DeviceTweakerBackup { RegistryValues = [Value(Class, "PnPCapabilities", false),
            Value(path, "SelectiveSuspendEnabled", selectiveExists, "0"), Value(path, "EnhancedPowerManagementEnabled", enhancedExists, "1")] };
        Assert.Equal(expected, LivePowerRestoreService.ReadRestoredEnabled(backup, Id, Class, usb: true));
    }

    [Fact]
    public void MissingOrFailedCaptureDoesNotGuessAnEnabledValue()
    {
        var backup = new DeviceTweakerBackup();
        Assert.Null(LivePowerRestoreService.ReadRestoredEnabled(backup, Id, Class, usb: false));
        backup.RegistryValues.Add(Value(Class, "PnPCapabilities", false, "denied", "ReadError"));
        Assert.Null(LivePowerRestoreService.ReadRestoredEnabled(backup, Id, Class, usb: false));
        var report = Run([new(Id, null)], (_, _) => throw new Exception("must not write"));
        Assert.True(report.Succeeded);
        Assert.True(Assert.Single(report.Warnings).IsSkipped);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void UnsupportedLiveSettingIsReportedAsSkipped(bool unsupported)
    {
        var status = unsupported ? DevicePowerPolicy.WriteStatus.NotSupported : DevicePowerPolicy.WriteStatus.NotExposed;
        var report = Run([new(Id, false)], (_, _) => new(status, "unsupported"));
        Assert.True(report.Succeeded);
        Assert.True(Assert.Single(report.Warnings).IsSkipped);
    }

    [Fact]
    public void FailureForOneDeviceDoesNotSkipOthersAndActivationErrorIsPreserved()
    {
        var writes = new List<string>();
        var report = Run([new("first", false), new("second", true)], (id, _) =>
        {
            writes.Add(id);
            if (id == "first") throw new IOException("WMI failed");
            return new(DevicePowerPolicy.WriteStatus.Applied, null);
        }, () => throw new IOException("activation failed"));
        Assert.Equal(new[] { "first", "second" }, writes);
        Assert.Equal(2, report.Errors.Count);
    }

    [Fact]
    public void DuplicateDevicesWriteOnceAndConflictingTargetsNeverWrite()
    {
        int writes = 0;
        var report = Run([new(Id, false), new(Id.ToLowerInvariant(), false)], (_, _) => { writes++; return new(DevicePowerPolicy.WriteStatus.Applied, null); });
        Assert.True(report.Succeeded);
        Assert.Equal(1, writes);
        report = Run([new(Id, false), new(Id, true)], (_, _) => throw new Exception("must not write"));
        Assert.False(report.Succeeded);
    }

    [Fact]
    public void ChildNoChangesMadeDoesNotHideParentChanges()
    {
        var child = new OperationReport(); child.MarkNoChangesMade(); child.AddError("WMI", "failed");
        var parent = new OperationReport(); parent.AddSuccess("Registry"); parent.MergeIssues(child);
        Assert.False(parent.Succeeded);
        Assert.False(parent.NoChangesMade);
    }
}
