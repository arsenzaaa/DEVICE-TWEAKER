namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    internal sealed record BackupDeviceTarget(string InstanceId, string RegBase, DeviceKind Kind, bool Wifi);
    private sealed record BackupCreationResult(string? Path, Exception? Error);

    private BackupDeviceTarget[] CaptureBackupTargets() => _blocks
        .Where(block => !block.Device.IsTestDevice)
        .Select(block => new BackupDeviceTarget(block.Device.InstanceId, block.Device.RegBase, block.Kind, block.Device.Wifi))
        .ToArray();

    private Task<bool> CreateDeviceTweakerBackup(string reason, bool showDialog,
        BackupLocation location = BackupLocation.Local, string? protectedPath = null)
    {
        DemandRealWrite("CREATE BACKUP");
        if (!_deviceDiscoveryCompleted)
            return Task.FromResult(PublishBackupResult(new(null, new InvalidOperationException(
                "The initial device scan has not completed. Refresh devices before creating a backup.")), reason, showDialog));
        BackupDeviceTarget[] targets = CaptureBackupTargets();
        return CreateBackupAndPublishAsync(reason, showDialog, location, protectedPath, targets);
    }

    private async Task<bool> CreateBackupAndPublishAsync(string reason, bool showDialog,
        BackupLocation location, string? protectedPath, BackupDeviceTarget[] targets)
    {
        BackupCreationResult result = await Task.Run(() => CreateDeviceTweakerBackupCore(reason, location, protectedPath, targets));
        return PublishBackupResult(result, reason, showDialog);
    }

    private bool PublishBackupResult(BackupCreationResult result, string reason, bool showDialog)
    {
        if (result.Error is null && result.Path is { } path)
        {
            _lastBackupPath = path;
            if (showDialog) ShowBackupFileResult(path, restored: false);
            return true;
        }
        WriteLog($"BACKUP: failed reason={reason}: {result.Error?.Message}");
        if (showDialog)
        {
            OperationReport report = new();
            report.MarkNoChangesMade();
            report.AddError("BACKUP", "The backup file could not be created.", result.Error?.ToString());
            ShowOperationResult(report, string.Empty, "No backup was created.", operationName: "BACKUP");
        }
        return false;
    }

    private Task<bool> EnsureOriginalDeviceTweakerBackup(BackupLocation location)
    {
        DemandRealWrite("CREATE ORIGINAL BACKUP");
        BackupDeviceTarget[] targets = CaptureBackupTargets();
        return Task.Run(() => EnsureOriginalDeviceTweakerBackupCore(location, targets));
    }
}
