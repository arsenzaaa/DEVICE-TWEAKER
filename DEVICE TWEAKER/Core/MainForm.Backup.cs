using Microsoft.Win32;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private const int DeviceTweakerBackupVersion = 1;
    private const string BackupFolderName = "Backups";
    private const string BackupFilePrefix = "DeviceTweakerBackup_";
    private string? _lastBackupPath;
    private const string OriginalBackupFileName = "DeviceTweakerBackup_ORIGINAL.json";

    private enum BackupLocation
    {
        Local,
        Roaming,
    }

    private enum AutoBackupChoice
    {
        Skip,
        Local,
        Roaming,
        Cancel,
    }

    private enum RestoreChoice
    {
        Cancel,
        SafeReset,
        RestoreLatest,
        RestoreBackup,
        DeleteBackup,
    }

    private sealed class BackupSnapshotInfo
    {
        public required string Path { get; init; }
        public required string Location { get; init; }
        public required DateTime LastWriteUtc { get; init; }
        public string Reason { get; init; } = string.Empty;
        public DateTime? CreatedAt { get; init; }
        public bool IsOriginal { get; init; }

        public override string ToString()
        {
            DateTime stamp = CreatedAt ?? LastWriteUtc.ToLocalTime();
            string reason = IsOriginal
                ? "ORIGINAL STATE"
                : string.IsNullOrWhiteSpace(Reason) ? "backup" : Reason;
            return $"{stamp:yyyy-MM-dd HH:mm:ss} [{Location}] {UiLanguage.Text(reason)}";
        }
    }

    private sealed class DeviceTweakerBackup
    {
        public int Version { get; set; } = DeviceTweakerBackupVersion;
        public DateTime CreatedAt { get; set; }
        public string Reason { get; set; } = string.Empty;
        public List<RegistryValueBackup> RegistryValues { get; set; } = [];
        public FileBackup? ImodScript { get; set; }
    }

    private sealed class RegistryValueBackup
    {
        public string Hive { get; set; } = string.Empty;
        public string Path { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public bool Exists { get; set; }
        public string Kind { get; set; } = string.Empty;
        public string? Data { get; set; }
    }

    private sealed class FileBackup
    {
        public string Path { get; set; } = string.Empty;
        public bool Exists { get; set; }
        public string? Text { get; set; }
    }

    private sealed record BackupValidationResult(
        int ValueCount,
        int ExistingCount,
        int MissingCount,
        int ReadErrorCount,
        string ImodScriptState,
        long Bytes,
        string Sha256);

    private string GetBackupDirectory()
    {
        return GetBackupDirectory(BackupLocation.Local);
    }

    private string GetBackupDirectory(BackupLocation location)
    {
        if (location == BackupLocation.Roaming)
        {
            string root = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            if (string.IsNullOrWhiteSpace(root))
            {
                root = AppContext.BaseDirectory;
            }

            return Path.Combine(root, "DEVICE TWEAKER", BackupFolderName);
        }

        return Path.Combine(GetScriptRoot(), BackupFolderName);
    }

    private bool CreateDeviceTweakerBackup(string reason, bool showDialog)
    {
        return CreateDeviceTweakerBackup(reason, showDialog, BackupLocation.Local);
    }

    private bool CreateDeviceTweakerBackup(string reason, bool showDialog, BackupLocation location)
    {
        try
        {
            if (_blocks.Count == 0)
            {
                RefreshBlocks();
            }

            if (reason.StartsWith("pre-", StringComparison.OrdinalIgnoreCase)
                && !EnsureOriginalDeviceTweakerBackup(location))
            {
                throw new InvalidOperationException("The original-state backup could not be created or validated.");
            }

            DeviceTweakerBackup backup = CaptureDeviceTweakerBackup(reason);
            string directory = GetBackupDirectory(location);
            Directory.CreateDirectory(directory);

            string safeReason = MakeBackupFileReason(reason);
            string stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff", CultureInfo.InvariantCulture);
            string path = Path.Combine(directory, $"{BackupFilePrefix}{stamp}_{safeReason}.json");
            if (File.Exists(path))
            {
                path = Path.Combine(directory, $"{BackupFilePrefix}{stamp}_{safeReason}_{Guid.NewGuid().ToString("N")[..8]}.json");
            }
            JsonSerializerOptions options = new() { WriteIndented = true };
            ValidateCapturedBackupForWrite(backup);
            WriteAllTextAtomic(path, JsonSerializer.Serialize(backup, options), Encoding.UTF8);
            BackupValidationResult validation = ValidateWrittenBackup(path, backup);
            WriteLog(
                $"BACKUP.VALIDATION: status=passed schema={backup.Version} values={validation.ValueCount} " +
                $"existing={validation.ExistingCount} missing={validation.MissingCount} readErrors={validation.ReadErrorCount} " +
                $"imodScript={validation.ImodScriptState} bytes={validation.Bytes} sha256={validation.Sha256}");
            WriteLog($"BACKUP: saved location={location} path={path} values={backup.RegistryValues.Count} reason={reason} atomic=true validated=true");
            PruneDeviceTweakerBackups(directory, keepLatest: 10);
            _lastBackupPath = path;

            if (showDialog)
            {
                ShowThemedInfo($"Backup saved.\n{path}");
            }

            return true;
        }
        catch (Exception ex)
        {
            WriteLog($"BACKUP: failed reason={reason}: {ex.Message}");
            if (showDialog)
            {
                OperationReport report = new();
                report.MarkNoChangesMade();
                report.AddError("BACKUP", "The backup file could not be created.", ex.ToString());
                ShowOperationResult(
                    report,
                    string.Empty,
                    "No backup was created.",
                    operationName: "BACKUP");
            }

            return false;
        }
    }

    private bool EnsureOriginalDeviceTweakerBackup(BackupLocation preferredLocation)
    {
        foreach (string directory in EnumerateBackupDirectories())
        {
            string existingPath = Path.Combine(directory, OriginalBackupFileName);
            if (!File.Exists(existingPath))
            {
                continue;
            }

            try
            {
                string json = File.ReadAllText(existingPath, Encoding.UTF8);
                DeviceTweakerBackup? existing = JsonSerializer.Deserialize<DeviceTweakerBackup>(json);
                if (existing is null)
                {
                    WriteLog($"BACKUP.ORIGINAL.ERROR: invalid snapshot path={existingPath}");
                    return false;
                }

                ValidateDeviceTweakerBackup(existing);

                BackupValidationResult validation = ValidateWrittenBackup(existingPath, existing);
                WriteLog(
                    $"BACKUP.ORIGINAL.VALIDATION: status=passed schema={existing.Version} values={validation.ValueCount} " +
                    $"existing={validation.ExistingCount} missing={validation.MissingCount} readErrors={validation.ReadErrorCount} " +
                    $"imodScript={validation.ImodScriptState} bytes={validation.Bytes} sha256={validation.Sha256}");
                WriteLog($"BACKUP.ORIGINAL: existing path={existingPath} created={existing.CreatedAt:O} validated=true");
                return true;
            }
            catch (Exception ex)
            {
                WriteLog($"BACKUP.ORIGINAL.ERROR: validation failed path={existingPath} error=\"{FlattenLogText(ex.ToString())}\"");
                return false;
            }
        }

        try
        {
            string directory = GetBackupDirectory(preferredLocation);
            Directory.CreateDirectory(directory);
            string path = Path.Combine(directory, OriginalBackupFileName);

            DeviceTweakerBackup? original = null;
            BackupSnapshotInfo? oldestKnown = GetBackupSnapshots()
                .Where(snapshot => !snapshot.IsOriginal)
                .OrderBy(snapshot => snapshot.CreatedAt ?? snapshot.LastWriteUtc.ToLocalTime())
                .FirstOrDefault();
            if (oldestKnown is not null)
            {
                string oldestJson = File.ReadAllText(oldestKnown.Path, Encoding.UTF8);
                original = JsonSerializer.Deserialize<DeviceTweakerBackup>(oldestJson);
                if (original is null)
                {
                    throw new InvalidOperationException($"The oldest known backup is invalid: {oldestKnown.Path}");
                }

                ValidateDeviceTweakerBackup(original);

                WriteLog($"BACKUP.ORIGINAL: importing oldest known snapshot path={oldestKnown.Path} created={original.CreatedAt:O}");
            }

            original ??= CaptureDeviceTweakerBackup("original-state");
            ValidateCapturedBackupForWrite(original);
            string json = JsonSerializer.Serialize(original, new JsonSerializerOptions { WriteIndented = true });
            WriteAllTextAtomic(path, json, Encoding.UTF8);
            BackupValidationResult validation = ValidateWrittenBackup(path, original);
            WriteLog(
                $"BACKUP.ORIGINAL.VALIDATION: status=passed schema={original.Version} values={validation.ValueCount} " +
                $"existing={validation.ExistingCount} missing={validation.MissingCount} readErrors={validation.ReadErrorCount} " +
                $"imodScript={validation.ImodScriptState} bytes={validation.Bytes} sha256={validation.Sha256}");
            WriteLog($"BACKUP.ORIGINAL: created location={preferredLocation} path={path} values={original.RegistryValues.Count} atomic=true validated=true");
            return true;
        }
        catch (Exception ex)
        {
            WriteLog($"BACKUP.ORIGINAL.ERROR: creation failed location={preferredLocation} error=\"{FlattenLogText(ex.ToString())}\"");
            return false;
        }
    }

    private DeviceTweakerBackup CaptureDeviceTweakerBackup(string reason)
    {
        DeviceTweakerBackup backup = new()
        {
            CreatedAt = DateTime.Now,
            Reason = reason,
        };

        HashSet<string> seen = new(StringComparer.OrdinalIgnoreCase);
        void AddValues(RegistryHive hive, string path, params string[] names)
        {
            foreach (string name in names)
            {
                string id = $"{hive}|{path}|{name}";
                if (!seen.Add(id))
                {
                    continue;
                }

                backup.RegistryValues.Add(CaptureRegistryValue(hive, path, name));
            }
        }

        foreach (DeviceBlock block in _blocks)
        {
            if (block.Device.IsTestDevice)
            {
                continue;
            }

            string intBase = block.Device.RegBase + @"\Device Parameters\Interrupt Management";
            AddValues(
                RegistryHive.LocalMachine,
                intBase + @"\MessageSignaledInterruptProperties",
                "MSISupported",
                "MessageNumberLimit");

            AddValues(
                RegistryHive.LocalMachine,
                intBase + @"\Affinity Policy",
                "DevicePriority",
                "DevicePolicy",
                "AssignmentSetOverride");

            if (block.Kind == DeviceKind.USB)
            {
                foreach (string instanceId in UsbSelectiveSuspendPolicy.EnumerateBackupInstanceIds(block.Device.InstanceId))
                {
                    AddValues(
                        RegistryHive.LocalMachine,
                        UsbSelectiveSuspendPolicy.DeviceParametersPath(instanceId),
                        UsbSelectiveSuspendPolicy.SelectiveSuspendEnabledName,
                        UsbSelectiveSuspendPolicy.EnhancedPowerManagementEnabledName);

                    string? usbClassKey = DevicePowerPolicy.TryGetClassKeyPath(instanceId);
                    if (!string.IsNullOrWhiteSpace(usbClassKey))
                    {
                        AddValues(
                            RegistryHive.LocalMachine,
                            usbClassKey,
                            DevicePowerPolicy.PnPCapabilitiesName);
                    }
                }
            }

            if (block.Kind is DeviceKind.NET_NDIS or DeviceKind.NET_CX)
            {
                string? classKey = GetClassKeyForDevice(block.Device.InstanceId);
                if (!string.IsNullOrWhiteSpace(classKey))
                {
                    if (block.Kind == DeviceKind.NET_NDIS)
                    {
                        AddValues(
                            RegistryHive.LocalMachine,
                            classKey,
                            "*RssBaseProcNumber",
                            "*NumRssQueues",
                            "*RssBaseProcGroup",
                            "*MaxRssProcessors",
                            "*RSSMaxProcGroup",
                            "*RssMaxProcNumber",
                            "*NumaNodeId");
                    }

                    AddValues(
                        RegistryHive.LocalMachine,
                        classKey,
                        DevicePowerPolicy.PnPCapabilitiesName);
                }
            }
        }

        AddValues(
            RegistryHive.CurrentUser,
            @"Control Panel\Mouse",
            RawMouseThrottleValueName);

        AddValues(
            RegistryHive.CurrentUser,
            ImodStartupRunKeyPath,
            ImodStartupRunValueName);

        AddValues(
            RegistryHive.LocalMachine,
            @"SYSTEM\CurrentControlSet\Control\Session Manager\Kernel",
            "ReservedCpuSets");

        if (UsbSelectiveSuspendPolicy.TryGetActiveScheme(out Guid usbSsScheme))
        {
            AddValues(
                RegistryHive.LocalMachine,
                UsbSelectiveSuspendPolicy.PowerPlanSettingPath(usbSsScheme),
                "ACSettingIndex",
                "DCSettingIndex");
        }

        string startupScript = GetImodStartupPath();
        backup.ImodScript = new FileBackup
        {
            Path = startupScript,
            Exists = File.Exists(startupScript),
            Text = File.Exists(startupScript) ? File.ReadAllText(startupScript, Encoding.UTF8) : null,
        };

        return backup;
    }

    private bool ValidateImodBackupPersistenceContract(out string error)
    {
        try
        {
            DeviceTweakerBackup backup = CaptureDeviceTweakerBackup("qa-imod-persistence");
            ValidateDeviceTweakerBackup(backup);
            RegistryValueBackup? runValue = backup.RegistryValues.FirstOrDefault(value =>
                string.Equals(value.Hive, "HKCU", StringComparison.OrdinalIgnoreCase)
                && string.Equals(value.Path, ImodStartupRunKeyPath, StringComparison.OrdinalIgnoreCase)
                && string.Equals(value.Name, ImodStartupRunValueName, StringComparison.OrdinalIgnoreCase));
            if (runValue is null)
            {
                error = "HKCU Run value is not captured";
                return false;
            }

            if (backup.ImodScript is null || !IsManagedImodScriptPath(backup.ImodScript.Path))
            {
                error = "managed IMOD script state is not captured";
                return false;
            }

            error = string.Empty;
            return true;
        }
        catch (Exception ex)
        {
            error = FlattenLogText(ex.ToString());
            return false;
        }
    }

    private RegistryValueBackup CaptureRegistryValue(RegistryHive hive, string path, string name)
    {
        RegistryValueBackup backup = new()
        {
            Hive = hive == RegistryHive.CurrentUser ? "HKCU" : "HKLM",
            Path = path,
            Name = name,
            Exists = false,
        };

        try
        {
            using RegistryKey baseKey = RegistryKey.OpenBaseKey(hive, RegistryView.Registry64);
            using RegistryKey? key = baseKey.OpenSubKey(path);
            object? value = key?.GetValue(name, null, RegistryValueOptions.DoNotExpandEnvironmentNames);
            if (key is null || value is null)
            {
                return backup;
            }

            RegistryValueKind kind = key.GetValueKind(name);
            backup.Exists = true;
            backup.Kind = kind.ToString();
            backup.Data = EncodeRegistryValue(value, kind);
        }
        catch (Exception ex)
        {
            backup.Kind = "ReadError";
            backup.Data = ex.Message;
            backup.Exists = true; // must not look like "absent" - restore deletes Exists=false
            WriteLog($"BACKUP.REG: read failed {backup.Hive}\\{path}\\{name}: {ex.Message}");
        }

        return backup;
    }

    private static string EncodeRegistryValue(object value, RegistryValueKind kind)
    {
        return kind switch
        {
            RegistryValueKind.Binary => Convert.ToBase64String(value as byte[] ?? []),
            RegistryValueKind.MultiString => JsonSerializer.Serialize(value as string[] ?? []),
            RegistryValueKind.DWord or RegistryValueKind.QWord => Convert.ToInt64(value, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture),
            _ => value.ToString() ?? string.Empty,
        };
    }

    private void RestoreLatestDeviceTweakerBackup()
    {
        try
        {
            List<BackupSnapshotInfo> backups = GetBackupSnapshots();
            RestoreChoice choice = ShowRestoreChoiceDialog(backups, out string? path);
            if (choice == RestoreChoice.Cancel)
            {
                WriteLog("BACKUP.RESTORE: canceled by user");
                return;
            }

            if (choice == RestoreChoice.SafeReset)
            {
                RunSafeResetFromRestore();
                return;
            }

            if (choice == RestoreChoice.DeleteBackup)
            {
                int deleted = string.IsNullOrWhiteSpace(path)
                    ? 0
                    : DeleteDeviceTweakerBackups(backups.Where(b => string.Equals(b.Path, path, StringComparison.OrdinalIgnoreCase)).ToList());
                WriteLog($"BACKUP.RESTORE: deleted selected backup count={deleted} path={path}");
                ShowThemedInfo($"Backup files deleted: {deleted}");
                return;
            }

            if (choice == RestoreChoice.RestoreLatest)
            {
                path = backups.FirstOrDefault()?.Path;
            }

            if (string.IsNullOrWhiteSpace(path))
            {
                WriteLog("BACKUP.RESTORE: restore-backup requested but no backup found");
                ShowThemedInfo("No backup files were found in the EXE folder or APPDATA.");
                return;
            }

            if (!CreateDeviceTweakerBackup("pre-restore", showDialog: false))
            {
                throw new InvalidOperationException("A rollback backup could not be created; restore was cancelled.");
            }

            WriteLog($"BACKUP.RESTORE: restore-backup requested choice={choice} path={path}");
            RestoreDeviceTweakerBackup(path);
            BeginDevicesBusyWork("Refreshing devices...", 4);
            try
            {
                RefreshBlocks();
            }
            finally
            {
                EndDevicesBusy();
            }

            ShowThemedInfo($"Backup restored.\n{path}\n\nPlease reboot your PC to finish applying restored settings.");
        }
        catch (Exception ex)
        {
            WriteLog($"BACKUP.RESTORE: failed: {FlattenLogText(ex.ToString())}");
            OperationReport report = new();
            report.AddError(
                "BACKUP RESTORE",
                "Restore did not finish. Any completed rollback steps were preserved.",
                ex.ToString());
            ShowOperationResult(
                report,
                "Some restore steps may already have run.",
                "Restore did not finish. Open DETAILS before rebooting.",
                operationName: "BACKUP RESTORE");
        }
    }

    private void RunSafeResetFromRestore()
    {
        WriteLog("BACKUP.RESTORE: safe-reset requested");
        OperationReport report = new();
        if (!_testAutoDryRun && !CreateDeviceTweakerBackup("pre-reset-tweaks", showDialog: false))
        {
            report.MarkNoChangesMade();
            report.AddError("Automatic backup", "backup could not be created; no settings were changed");
            ShowOperationResult(
                report,
                string.Empty,
                "RESET WINDOWS DEFAULT was cancelled because the rollback backup failed.",
                operationName: "RESET WINDOWS DEFAULT");
            return;
        }

        BeginDevicesBusyWork(
            _testAutoDryRun ? "Previewing RESET WINDOWS DEFAULT..." : "Running RESET WINDOWS DEFAULT...",
            Math.Max(4, _blocks.Count + 4));
        try
        {
            SafeResetTweaks(report);
        }
        finally
        {
            EndDevicesBusy();
        }
    }

    private List<BackupSnapshotInfo> GetBackupSnapshots()
    {
        List<BackupSnapshotInfo> backups = [];
        foreach (string directory in EnumerateBackupDirectories())
        {
            try
            {
                if (!Directory.Exists(directory))
                {
                    WriteLog($"BACKUP.SCAN: directory={directory} exists=false count=0");
                    continue;
                }

                int before = backups.Count;
                foreach (string path in Directory.EnumerateFiles(directory, $"{BackupFilePrefix}*.json"))
                {
                    BackupSnapshotInfo? info = CreateBackupSnapshotInfo(path, directory);
                    if (info is not null)
                    {
                        backups.Add(info);
                    }
                }

                WriteLog($"BACKUP.SCAN: directory={directory} exists=true count={backups.Count - before}");
            }
            catch (Exception ex)
            {
                WriteLog($"BACKUP.SCAN.WARN: directory={directory} failed={ex.Message}");
            }
        }

        return backups
            .OrderByDescending(info => info.CreatedAt ?? info.LastWriteUtc.ToLocalTime())
            .ThenByDescending(info => info.LastWriteUtc)
            .ToList();
    }

    private IEnumerable<string> EnumerateBackupDirectories()
    {
        HashSet<string> seen = new(StringComparer.OrdinalIgnoreCase);
        foreach (BackupLocation location in new[] { BackupLocation.Local, BackupLocation.Roaming })
        {
            string directory = GetBackupDirectory(location);
            if (seen.Add(directory))
            {
                yield return directory;
            }
        }
    }

    private BackupSnapshotInfo? CreateBackupSnapshotInfo(string path, string directory)
    {
        try
        {
            FileInfo file = new(path);
            DeviceTweakerBackup? backup = null;
            try
            {
                string json = File.ReadAllText(path, Encoding.UTF8);
                backup = JsonSerializer.Deserialize<DeviceTweakerBackup>(json);
                if (backup is null)
                {
                    throw new InvalidDataException("Backup JSON is empty.");
                }

                ValidateDeviceTweakerBackup(backup);
            }
            catch (Exception ex)
            {
                WriteLog($"BACKUP.SCAN.WARN: invalid backup excluded path={path} error=\"{FlattenLogText(ex.ToString())}\"");
                return null;
            }

            string location = string.Equals(directory, GetBackupDirectory(BackupLocation.Roaming), StringComparison.OrdinalIgnoreCase)
                ? "APPDATA"
                : "EXE";

            return new BackupSnapshotInfo
            {
                Path = path,
                Location = location,
                LastWriteUtc = file.LastWriteTimeUtc,
                CreatedAt = backup?.CreatedAt,
                Reason = backup?.Reason ?? string.Empty,
                IsOriginal = string.Equals(file.Name, OriginalBackupFileName, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(backup?.Reason, "original-state", StringComparison.OrdinalIgnoreCase),
            };
        }
        catch
        {
            return null;
        }
    }

    private int DeleteDeviceTweakerBackups(IReadOnlyList<BackupSnapshotInfo> backups)
    {
        int deleted = 0;
        foreach (BackupSnapshotInfo backup in backups)
        {
            if (backup.IsOriginal)
            {
                WriteLog($"BACKUP.DELETE: protected original snapshot path={backup.Path}");
                continue;
            }

            try
            {
                if (File.Exists(backup.Path))
                {
                    File.Delete(backup.Path);
                    deleted++;
                }
            }
            catch (Exception ex)
            {
                WriteLog($"BACKUP.DELETE: failed path={backup.Path}: {ex.Message}");
            }
        }

        return deleted;
    }

    private void PruneDeviceTweakerBackups(string directory, int keepLatest)
    {
        try
        {
            if (!Directory.Exists(directory))
            {
                return;
            }

            List<FileInfo> files = Directory.EnumerateFiles(directory, $"{BackupFilePrefix}*.json")
                .Where(file => !string.Equals(Path.GetFileName(file), OriginalBackupFileName, StringComparison.OrdinalIgnoreCase))
                .Select(path => new FileInfo(path))
                .OrderByDescending(file => file.LastWriteTimeUtc)
                .ToList();

            foreach (FileInfo file in files.Skip(Math.Max(1, keepLatest)))
            {
                try
                {
                    file.Delete();
                    WriteLog($"BACKUP.PRUNE: deleted old backup path={file.FullName}");
                }
                catch (Exception ex)
                {
                    WriteLog($"BACKUP.PRUNE.WARN: failed path={file.FullName}: {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            WriteLog($"BACKUP.PRUNE.WARN: failed directory={directory}: {ex.Message}");
        }
    }

    private AutoBackupChoice PromptBackupLocationForAuto()
    {
        AutoBackupChoice choice = ShowAutoBackupChoiceDialog();
        WriteLog($"BACKUP.PROMPT.AUTO: choice={choice}");
        return choice;
    }

    private void RestoreDeviceTweakerBackup(string path)
    {
        string json = File.ReadAllText(path, Encoding.UTF8);
        DeviceTweakerBackup? backup = JsonSerializer.Deserialize<DeviceTweakerBackup>(json);
        if (backup is null)
        {
            throw new InvalidOperationException("Backup file is empty or invalid.");
        }

        ValidateDeviceTweakerBackup(backup);

        List<RegistryValueBackup> rollbackValues = [];
        FileBackup? rollbackScript = null;
        if (backup.ImodScript is not null)
        {
            string startupPath = GetImodStartupPath();
            rollbackScript = new FileBackup
            {
                Path = startupPath,
                Exists = File.Exists(startupPath),
                Text = File.Exists(startupPath) ? File.ReadAllText(startupPath, Encoding.UTF8) : null,
            };
        }

        try
        {
            foreach (RegistryValueBackup value in backup.RegistryValues)
            {
                if (string.Equals(value.Kind, "ReadError", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (!TryGetBackupHive(value.Hive, out RegistryHive hive))
                {
                    throw new InvalidOperationException($"Unsupported registry hive: {value.Hive}.");
                }

                RegistryValueBackup rollback = CaptureRegistryValue(hive, value.Path, value.Name);
                if (string.Equals(rollback.Kind, "ReadError", StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException(
                        $"Cannot capture rollback value: {value.Hive}\\{value.Path}\\{value.Name}. {rollback.Data}");
                }

                rollbackValues.Add(rollback);
                RestoreRegistryValue(value);
            }

            if (backup.ImodScript is not null)
            {
                RestoreFileBackup(backup.ImodScript);
                InvalidateImodCache();
            }
        }
        catch (Exception restoreError)
        {
            WriteLog($"BACKUP.RESTORE.ROLLBACK: starting after error=\"{FlattenLogText(restoreError.ToString())}\"");
            List<string> rollbackErrors = [];
            foreach (RegistryValueBackup rollback in rollbackValues.AsEnumerable().Reverse())
            {
                try
                {
                    RestoreRegistryValue(rollback);
                }
                catch (Exception ex)
                {
                    rollbackErrors.Add($"{rollback.Hive}\\{rollback.Path}\\{rollback.Name}: {ex.Message}");
                }
            }

            if (rollbackScript is not null)
            {
                try
                {
                    RestoreFileBackup(rollbackScript);
                    InvalidateImodCache();
                }
                catch (Exception ex)
                {
                    rollbackErrors.Add($"IMOD script: {ex.Message}");
                }
            }

            if (rollbackErrors.Count > 0)
            {
                throw new InvalidOperationException(
                    $"Backup restore failed and rollback was incomplete. Restore error: {restoreError.Message}. "
                    + $"Rollback errors: {string.Join(" | ", rollbackErrors)}",
                    restoreError);
            }

            WriteLog("BACKUP.RESTORE.ROLLBACK: completed");
            throw;
        }

        WriteLog($"BACKUP.RESTORE: restored path={path} values={backup.RegistryValues.Count} reason={backup.Reason}");
        SyncLivePowerManagementAfterRestore();
    }

    private void ValidateDeviceTweakerBackup(DeviceTweakerBackup backup)
    {
        if (backup.Version != DeviceTweakerBackupVersion)
        {
            throw new InvalidOperationException($"Unsupported backup version: {backup.Version}.");
        }

        if (backup.CreatedAt == default)
        {
            throw new InvalidOperationException("Backup creation time is missing.");
        }

        if (backup.RegistryValues is null)
        {
            throw new InvalidOperationException("Backup registry data is missing.");
        }

        HashSet<string> restoreTargets = new(StringComparer.OrdinalIgnoreCase);
        foreach (RegistryValueBackup value in backup.RegistryValues)
        {
            if (value is null || !IsManagedBackupRegistryValue(value))
            {
                throw new InvalidOperationException("Backup contains an unmanaged or empty registry target.");
            }

            string target = $"{value.Hive.Trim()}|{value.Path.Trim().Trim('\\')}|{value.Name.Trim()}";
            if (!restoreTargets.Add(target))
            {
                throw new InvalidOperationException(
                    $"Backup contains a duplicate registry target: {value.Hive}\\{value.Path}\\{value.Name}");
            }

            ValidateRegistryValueBackup(value);
        }

        if (backup.ImodScript is not null
            && !IsManagedImodScriptPath(backup.ImodScript.Path))
        {
            throw new InvalidOperationException("Backup contains an unmanaged IMOD script path.");
        }
    }

    private void ValidateCapturedBackupForWrite(DeviceTweakerBackup backup)
    {
        ValidateDeviceTweakerBackup(backup);
        List<RegistryValueBackup> readErrors = backup.RegistryValues
            .Where(value => string.Equals(value.Kind, "ReadError", StringComparison.OrdinalIgnoreCase))
            .ToList();
        if (readErrors.Count == 0)
        {
            return;
        }

        string targets = string.Join(
            " | ",
            readErrors.Take(5).Select(value => $"{value.Hive}\\{value.Path}\\{value.Name}"));
        throw new InvalidOperationException(
            $"Backup capture is incomplete: {readErrors.Count} registry value(s) could not be read. {targets}");
    }

    private BackupValidationResult ValidateWrittenBackup(string path, DeviceTweakerBackup expected)
    {
        FileInfo file = new(path);
        if (!file.Exists || file.Length <= 0)
        {
            throw new InvalidOperationException($"The written backup is missing or empty: {path}");
        }

        string json = File.ReadAllText(path, Encoding.UTF8);
        DeviceTweakerBackup? actual = JsonSerializer.Deserialize<DeviceTweakerBackup>(json);
        if (actual is null)
        {
            throw new InvalidOperationException($"The written backup could not be read back: {path}");
        }

        ValidateDeviceTweakerBackup(actual);
        string expectedPayload = JsonSerializer.Serialize(expected);
        string actualPayload = JsonSerializer.Serialize(actual);
        if (!string.Equals(actualPayload, expectedPayload, StringComparison.Ordinal))
        {
            throw new InvalidOperationException($"The written backup does not match the captured snapshot: {path}");
        }

        int existing = actual.RegistryValues.Count(value => value.Exists);
        int readErrors = actual.RegistryValues.Count(value =>
            string.Equals(value.Kind, "ReadError", StringComparison.OrdinalIgnoreCase));
        using FileStream stream = file.Open(FileMode.Open, FileAccess.Read, FileShare.Read);
        string sha256 = Convert.ToHexString(SHA256.HashData(stream));
        string imodScriptState = actual.ImodScript switch
        {
            null => "not-captured",
            { Exists: true } => "present",
            _ => "absent",
        };

        return new BackupValidationResult(
            actual.RegistryValues.Count,
            existing,
            actual.RegistryValues.Count - existing,
            readErrors,
            imodScriptState,
            file.Length,
            sha256);
    }

    private static void ValidateRegistryValueBackup(RegistryValueBackup value)
    {
        if (!value.Exists || string.Equals(value.Kind, "ReadError", StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        if (!TryParseRegistryValueKind(value.Kind, out RegistryValueKind kind))
        {
            throw new InvalidOperationException(
                $"Backup contains an unsupported registry value kind: {value.Hive}\\{value.Path}\\{value.Name} kind={value.Kind}");
        }

        try
        {
            _ = DecodeRegistryValue(value.Data, kind);
        }
        catch (Exception ex) when (ex is FormatException or JsonException or OverflowException)
        {
            throw new InvalidOperationException(
                $"Backup contains invalid registry data: {value.Hive}\\{value.Path}\\{value.Name} kind={value.Kind}",
                ex);
        }
    }

    private static bool IsManagedBackupRegistryValue(RegistryValueBackup value)
    {
        string hive = value.Hive.Trim();
        string path = value.Path.Trim().Trim('\\');
        string name = value.Name.Trim();

        if (path.Length == 0
            || name.Length == 0
            || path.Contains("..", StringComparison.Ordinal)
            || path.Contains('\0')
            || name.Contains('\\')
            || name.Contains('\0'))
        {
            return false;
        }

        if (string.Equals(hive, "HKCU", StringComparison.OrdinalIgnoreCase))
        {
            bool isMouseSetting = string.Equals(path, @"Control Panel\Mouse", StringComparison.OrdinalIgnoreCase)
                && string.Equals(name, RawMouseThrottleValueName, StringComparison.OrdinalIgnoreCase);
            bool isImodStartup = string.Equals(path, ImodStartupRunKeyPath, StringComparison.OrdinalIgnoreCase)
                && string.Equals(name, ImodStartupRunValueName, StringComparison.OrdinalIgnoreCase);
            return isMouseSetting || isImodStartup;
        }

        if (!string.Equals(hive, "HKLM", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (string.Equals(
                path,
                @"SYSTEM\CurrentControlSet\Control\Session Manager\Kernel",
                StringComparison.OrdinalIgnoreCase))
        {
            return string.Equals(name, "ReservedCpuSets", StringComparison.OrdinalIgnoreCase);
        }

        const string powerSchemesRoot = @"SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes\";
        string usbSsSuffix =
            $@"\{UsbSelectiveSuspendPolicy.UsbSettingsSubgroup:D}\{UsbSelectiveSuspendPolicy.UsbSelectiveSuspendSetting:D}";
        if (path.StartsWith(powerSchemesRoot, StringComparison.OrdinalIgnoreCase)
            && path.EndsWith(usbSsSuffix, StringComparison.OrdinalIgnoreCase))
        {
            return name is "ACSettingIndex" or "DCSettingIndex";
        }

        const string enumRoot = @"SYSTEM\CurrentControlSet\Enum\";
        if (path.StartsWith(enumRoot, StringComparison.OrdinalIgnoreCase)
            && path.EndsWith(@"\Device Parameters", StringComparison.OrdinalIgnoreCase)
            && !path.Contains(@"\Interrupt Management", StringComparison.OrdinalIgnoreCase))
        {
            return name is "SelectiveSuspendEnabled" or "EnhancedPowerManagementEnabled";
        }

        if (path.StartsWith(enumRoot, StringComparison.OrdinalIgnoreCase)
            && path.EndsWith(
                @"\Device Parameters\Interrupt Management\MessageSignaledInterruptProperties",
                StringComparison.OrdinalIgnoreCase))
        {
            return name is "MSISupported" or "MessageNumberLimit";
        }

        if (path.StartsWith(enumRoot, StringComparison.OrdinalIgnoreCase)
            && path.EndsWith(
                @"\Device Parameters\Interrupt Management\Affinity Policy",
                StringComparison.OrdinalIgnoreCase))
        {
            return name is "DevicePriority" or "DevicePolicy" or "AssignmentSetOverride";
        }

        const string classRoot = @"SYSTEM\CurrentControlSet\Control\Class\";
        if (path.StartsWith(classRoot, StringComparison.OrdinalIgnoreCase))
        {
            return name is "*RssBaseProcNumber"
                or "*NumRssQueues"
                or "*RssBaseProcGroup"
                or "*MaxRssProcessors"
                or "*RSSMaxProcGroup"
                or "*RssMaxProcNumber"
                or "*NumaNodeId"
                or "PnPCapabilities";
        }

        return false;
    }

    private void RestoreRegistryValue(RegistryValueBackup value)
    {
        if (!TryGetBackupHive(value.Hive, out RegistryHive hive))
        {
            throw new InvalidOperationException($"Unsupported registry hive for restore: {value.Hive}.");
        }

        using RegistryKey baseKey = RegistryKey.OpenBaseKey(hive, RegistryView.Registry64);
        using RegistryKey? key = baseKey.CreateSubKey(value.Path, writable: true);
        if (key is null)
        {
            throw new InvalidOperationException(
                $"Failed to open registry key for restore: {value.Hive}\\{value.Path}");
        }

        if (!value.Exists)
        {
            key.DeleteValue(value.Name, throwOnMissingValue: false);
            WriteLog($"BACKUP.RESTORE.REG: deleted {value.Hive}\\{value.Path}\\{value.Name}");
            return;
        }

        if (string.Equals(value.Kind, "ReadError", StringComparison.OrdinalIgnoreCase))
        {
            WriteLog($"BACKUP.RESTORE.REG: skipped {value.Hive}\\{value.Path}\\{value.Name} kind=ReadError");
            return;
        }

        if (!TryParseRegistryValueKind(value.Kind, out RegistryValueKind kind))
        {
            throw new InvalidOperationException(
                $"Unsupported registry value kind for restore: {value.Hive}\\{value.Path}\\{value.Name} kind={value.Kind}");
        }

        object data = DecodeRegistryValue(value.Data, kind);
        key.SetValue(value.Name, data, kind);
        WriteLog($"BACKUP.RESTORE.REG: set {value.Hive}\\{value.Path}\\{value.Name} kind={kind}");
    }

    private static bool TryGetBackupHive(string text, out RegistryHive hive)
    {
        if (string.Equals(text, "HKCU", StringComparison.OrdinalIgnoreCase))
        {
            hive = RegistryHive.CurrentUser;
            return true;
        }

        if (string.Equals(text, "HKLM", StringComparison.OrdinalIgnoreCase))
        {
            hive = RegistryHive.LocalMachine;
            return true;
        }

        hive = RegistryHive.LocalMachine;
        return false;
    }

    private static bool TryParseRegistryValueKind(string text, out RegistryValueKind kind)
    {
        return Enum.TryParse(text, ignoreCase: true, out kind)
            && kind is RegistryValueKind.String
                or RegistryValueKind.ExpandString
                or RegistryValueKind.Binary
                or RegistryValueKind.DWord
                or RegistryValueKind.MultiString
                or RegistryValueKind.QWord;
    }

    private static object DecodeRegistryValue(string? data, RegistryValueKind kind)
    {
        string text = data ?? string.Empty;
        return kind switch
        {
            RegistryValueKind.Binary => Convert.FromBase64String(text),
            RegistryValueKind.MultiString => JsonSerializer.Deserialize<string[]>(text) ?? [],
            RegistryValueKind.DWord => int.Parse(text, CultureInfo.InvariantCulture),
            RegistryValueKind.QWord => long.Parse(text, CultureInfo.InvariantCulture),
            _ => text,
        };
    }

    private void RestoreFileBackup(FileBackup backup)
    {
        if (string.IsNullOrWhiteSpace(backup.Path))
        {
            return;
        }

        if (!IsManagedImodScriptPath(backup.Path))
        {
            throw new InvalidOperationException("Refusing to restore an unmanaged file path.");
        }

        string expectedPath = Path.GetFullPath(GetImodStartupPath());

        if (!backup.Exists)
        {
            DeleteImodStartupRegistration();
            File.Delete(expectedPath);
            File.Delete(GetLegacyImodStartupPath());
            return;
        }

        string? directory = Path.GetDirectoryName(expectedPath);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        WriteAllTextAtomic(expectedPath, backup.Text ?? string.Empty, new UTF8Encoding(false));
        if (!EnsureImodStartupRegistration(out string? registrationError))
        {
            throw new InvalidOperationException($"IMOD startup registration failed: {registrationError}");
        }

        File.Delete(GetLegacyImodStartupPath());
    }

    private static string MakeBackupFileReason(string reason)
    {
        if (string.IsNullOrWhiteSpace(reason))
        {
            return "manual";
        }

        StringBuilder sb = new(reason.Length);
        foreach (char ch in reason.Trim())
        {
            if (char.IsLetterOrDigit(ch) || ch is '-' or '_')
            {
                sb.Append(ch);
            }
        }

        return sb.Length > 0 ? sb.ToString() : "manual";
    }
}
