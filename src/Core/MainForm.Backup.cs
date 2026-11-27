using Microsoft.Win32;
using System.Text;

namespace DeviceTweakerCS;

// Which settings a backup holds, the restore dialog flow and write authorization.
// BackupStore owns the files and BackupRegistry reads and writes the values.
public sealed partial class MainForm
{
    private const string BackupFolderName = "Backups";
    private string? _lastBackupPath;

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

    private BackupStore CreateBackupStore()
        => new(GetBackupDirectory(BackupLocation.Local), GetBackupDirectory(BackupLocation.Roaming),
            ValidateDeviceTweakerBackup, WriteLog);

    private BackupRegistry CreateBackupRegistry() => new(WriteLog);

    private string GetBackupDirectory()
    {
        return GetBackupDirectory(BackupLocation.Local);
    }

    private string GetBackupDirectory(BackupLocation location)
    {
        if (QaRun.IsReadOnly)
        {
            return Path.Combine(QaRun.OutputDirectory, "Backups", location.ToString());
        }

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

    private BackupCreationResult CreateDeviceTweakerBackupCore(string reason, BackupLocation location,
        string? protectedPath, IReadOnlyList<BackupDeviceTarget> targets)
    {
        DemandRealWrite("CREATE BACKUP worker");
        try
        {
            if (reason.StartsWith("pre-", StringComparison.OrdinalIgnoreCase)
                && !EnsureOriginalDeviceTweakerBackupCore(location, targets))
            {
                throw new InvalidOperationException("The original-state backup could not be created or validated.");
            }

            DeviceTweakerBackup backup = CaptureDeviceTweakerBackup(reason, targets);
            string directory = GetBackupDirectory(location);
            string path = CreateBackupStore().Save(directory, location.ToString(), reason, backup);
            PruneDeviceTweakerBackups(directory, keepLatest: 10, protectedPaths: new[] { path, protectedPath ?? path });
            return new(path, null);
        }
        catch (Exception ex) { return new(null, ex); }
    }

    private bool EnsureOriginalDeviceTweakerBackupCore(BackupLocation preferredLocation, IReadOnlyList<BackupDeviceTarget> targets)
    {
        DemandRealWrite("CREATE ORIGINAL BACKUP");
        return CreateBackupStore().EnsureOriginal(GetBackupDirectory(preferredLocation), preferredLocation.ToString(),
            () => CaptureDeviceTweakerBackup("original-state", targets));
    }

    private DeviceTweakerBackup CaptureDeviceTweakerBackup(string reason, IReadOnlyList<BackupDeviceTarget> targets)
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

        foreach (BackupDeviceTarget block in targets)
        {
            string intBase = block.RegBase + @"\Device Parameters\Interrupt Management";
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
                foreach (string instanceId in UsbSelectiveSuspendPolicy.EnumerateBackupInstanceIds(block.InstanceId))
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
                string? classKey = GetClassKeyForDevice(block.InstanceId);
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
                            "*RSSProfile",
                            "*NumaNodeId");
                    }

                    AddValues(
                        RegistryHive.LocalMachine,
                        classKey,
                        [DevicePowerPolicy.PnPCapabilitiesName, .. NicDriverModeration.KeywordNames]);
                }
            }
        }

        AddValues(
            RegistryHive.CurrentUser,
            @"Control Panel\Mouse",
            RawMouseThrottleValueName);

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

        backup.StartupProfile = CaptureStartupProfileForBackup();
        return backup;
    }

    private bool ValidateImodBackupPersistenceContract(out string error)
    {
        try
        {
            DeviceTweakerBackup backup = CaptureDeviceTweakerBackup("qa-imod-persistence", CaptureBackupTargets());
            ValidateDeviceTweakerBackup(backup);
            if (backup.StartupProfile is null
                || !IsManagedImodScriptPath(backup.StartupProfile.Path))
            {
                error = "startup profile state is not captured";
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
        => CreateBackupRegistry().Capture(hive, path, name);

    private async Task RestoreLatestDeviceTweakerBackup()
    {
        bool settingsMayHaveChanged = false;
        string? recoveryBackupPath = null;
        try
        {
            List<BackupSnapshotInfo> backups = await Task.Run(GetBackupSnapshots);
            RestoreChoice choice = ShowRestoreChoiceDialog(backups, out string? path, selected =>
            {
                if (TryBlockSandboxHardwareWrite("BACKUP DELETE")) return 0;
                int deleted = DeleteDeviceTweakerBackups([selected]);
                WriteLog($"BACKUP.RESTORE: deleted selected backup count={deleted} path={selected.Path}");
                return deleted;
            });
            if (choice == RestoreChoice.Cancel)
            {
                WriteLog("BACKUP.RESTORE: canceled by user");
                return;
            }

            if (choice == RestoreChoice.SafeReset)
            {
                await RunSafeResetFromRestore();
                return;
            }

            if (TryBlockSandboxHardwareWrite("BACKUP RESTORE / DELETE")) return;

            if (choice == RestoreChoice.RestoreLatest)
            {
                path = backups.FirstOrDefault()?.Path;
            }

            if (string.IsNullOrWhiteSpace(path))
            {
                WriteLog("BACKUP.RESTORE: restore-backup requested but no backup found");
                ShowThemedInfo("No backups found next to the EXE or in AppData.");
                return;
            }

            WriteLog($"BACKUP.RESTORE: restore-backup requested choice={choice} path={path}");
            DeviceTweakerBackup selectedBackup = await Task.Run(() => ReadDeviceTweakerBackup(path));
            if (selectedBackup.StartupProfile?.Exists == true && !ConfirmDriverUse()) return;
            BackupDeviceTarget[] targets = CaptureBackupTargets();
            BackupCreationResult? rollback = null;
            OperationReport report;
            try
            {
                report = await Task.Run(() =>
                {
                    BackupRestoreWorkflow workflow = new(
                        _ => selectedBackup,
                        selected =>
                        {
                            rollback = CreateDeviceTweakerBackupCore("pre-restore", BackupLocation.Local, selected, targets);
                            if (rollback.Error is not null)
                                WriteLog($"BACKUP: failed reason=pre-restore: {rollback.Error.Message}");
                            return rollback.Error is null && rollback.Path is not null;
                        },
                        backup =>
                        {
                            settingsMayHaveChanged = true;
                            return RestoreDeviceTweakerBackup(backup, path, targets);
                        });
                    return workflow.Restore(path);
                });
            }
            finally
            {
                if (rollback?.Path is { } rollbackPath) _lastBackupPath = recoveryBackupPath = rollbackPath;
                if (settingsMayHaveChanged) InvalidateImodCache();
            }
            report.SetBackupPath(_lastBackupPath);
            _pendingReservedCpuSets = null;
            _reservedCpuSetsDirty = false;
            _rawMouseThrottlePending = false;
            BeginDevicesBusyWork("Refreshing devices...", 4);
            OperationReport refreshReport;
            try
            {
                refreshReport = await RefreshBlocksAsync(preserveEdits: false);
                if (IsDisposed || Disposing) return;
            }
            finally
            {
                EndDevicesBusy();
            }

            ShowCompletedBackupRestore(report, path, refreshReport);
        }
        catch (Exception ex)
        {
            WriteLog($"BACKUP.RESTORE: failed: {FlattenLogText(ex.ToString())}");
            OperationReport report = new();
            if (!settingsMayHaveChanged) report.MarkNoChangesMade();
            if (recoveryBackupPath is not null) report.SetBackupPath(recoveryBackupPath);
            report.AddError(
                "BACKUP RESTORE",
                settingsMayHaveChanged
                    ? "Restore did not finish. Any completed rollback steps were preserved."
                    : "Restore was cancelled before settings were changed.",
                ex.ToString());
            ShowOperationResult(
                report,
                "Some restore steps may already have run.",
                "Restore did not finish. Open DETAILS before rebooting.",
                operationName: "BACKUP RESTORE");
        }
    }

    private void ShowCompletedBackupRestore(OperationReport report, string path, OperationReport? refreshReport = null)
    {
        // Discovery and display run after the saved settings have been restored.
        // Failure here can leave the displayed values stale, not the restore partial.
        if (refreshReport is not null)
            foreach (OperationIssue issue in refreshReport.Issues)
                report.AddWarning(issue.Component, issue.UserMessage, issue.TechnicalDetails);

        if (report.Succeeded && report.Warnings.Count == 0)
            ShowBackupFileResult(path, restored: true);
        else
            ShowOperationResult(report,
                string.Empty,
                "Saved settings were restored.",
                operationName: "BACKUP RESTORE");
    }

    private async Task RunSafeResetFromRestore()
    {
        WriteLog("BACKUP.RESTORE: safe-reset requested");
        OperationReport report = new();
        if (!_testAutoDryRun && !await CreateDeviceTweakerBackup("pre-reset-tweaks", showDialog: false))
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
            await SafeResetTweaks(report);
        }
        finally
        {
            EndDevicesBusy();
        }
    }

    private List<BackupSnapshotInfo> GetBackupSnapshots() => CreateBackupStore().GetSnapshots();

    private IEnumerable<string> EnumerateBackupDirectories() => CreateBackupStore().Directories();

    private int DeleteDeviceTweakerBackups(IReadOnlyList<BackupSnapshotInfo> backups)
    {
        DemandRealWrite("DeleteDeviceTweakerBackups");
        return CreateBackupStore().Delete(backups);
    }

    private void PruneDeviceTweakerBackups(string directory, int keepLatest, IEnumerable<string>? protectedPaths = null)
    {
        DemandRealWrite("PruneDeviceTweakerBackups");
        CreateBackupStore().Prune(directory, keepLatest, protectedPaths);
    }

    private AutoBackupChoice PromptBackupLocationForAuto()
    {
        AutoBackupChoice choice = ShowAutoBackupChoiceDialog();
        WriteLog($"BACKUP.PROMPT.AUTO: choice={choice}");
        return choice;
    }

    private DeviceTweakerBackup ReadDeviceTweakerBackup(string path)
    {
        DeviceTweakerBackup backup = CreateBackupStore().Read(path);
        NormalizeStartupProfileBackup(backup);
        return backup;
    }

    private OperationReport RestoreDeviceTweakerBackup(DeviceTweakerBackup backup, string path, IReadOnlyList<BackupDeviceTarget> targets)
    {
        DemandRealWrite("RestoreDeviceTweakerBackup");
        if (backup.StartupProfile?.Exists == true && !_driverUseConfirmed)
            throw new InvalidOperationException("Driver profile restoration was not confirmed.");
        using DriverLoadGate gate = DriverLoadGate.Enter();
        DriverLoadGate.AssertLoaderIdle();

        BackupRegistry registry = CreateBackupRegistry();
        BackupRestoreService service = new(
            ValidateDeviceTweakerBackup,
            registry.Capture,
            RestoreRegistryValue,
            CaptureStartupProfileFile,
            RestoreStartupProfile,
            WriteLog,
            PrepareRollbackStartupProfile);
        service.Restore(backup);

        WriteLog($"BACKUP.RESTORE: restored path={path} values={backup.RegistryValues.Count} reason={backup.Reason}");
        OperationReport report = CreateBackupRestoreReport(backup);
        report.MergeIssues(SyncLivePowerManagementAfterRestore(backup, targets));
        return report;
    }

    internal static OperationReport CreateBackupRestoreReport(DeviceTweakerBackup backup)
    {
        OperationReport report = new();
        report.AddSuccess("Saved settings", "Registry and startup profile restored and verified.");
        foreach (string skipped in backup.SkippedNicItrEntries)
            report.AddWarning("NIC ITR startup profile",
                "An outdated NIC ITR startup setting was skipped. Save NIC ITR again for this adapter.", skipped);
        foreach (RegistryValueBackup skipped in backup.RegistryValues.Where(value =>
            string.Equals(value.Kind, "ReadError", StringComparison.OrdinalIgnoreCase)))
            report.AddSkipped("Saved setting", "This setting was unreadable when the backup was created. Its current value was kept.",
                $"{skipped.Hive}\\{skipped.Path}\\{skipped.Name}: {skipped.Data}");
        return report;
    }

    private void ValidateDeviceTweakerBackup(DeviceTweakerBackup backup)
        => BackupFormat.ValidateDocument(backup, IsManagedImodScriptPath);

    private void RestoreRegistryValue(RegistryValueBackup value)
    {
        DemandRealWrite("RestoreRegistryValue");
        CreateBackupRegistry().Restore(value);
    }

    private static FileBackup CaptureStartupProfileFile()
    {
        string path = GetStartupProfilePath();
        bool exists = File.Exists(path);
        return new FileBackup { Path = path, Exists = exists, Text = exists ? File.ReadAllText(path, Encoding.UTF8) : null };
    }

    private static FileBackup CaptureStartupProfileForBackup()
    {
        string path = GetStartupProfilePath();
        static FileBackup Read(string file) => new() { Path = file, Exists = true, Text = File.ReadAllText(file, Encoding.UTF8) };
        return StartupConfigReader.Read(path, new[] { GetLegacyAppDataImodScriptPath(), GetLegacyImodStartupPath() }, Read, Read)
            .RequireValueOrCreate(() => new FileBackup { Path = path, Exists = false });
    }

    // Brings a backup to the current startup layout before restore: a script saved
    // by an older version becomes profile.json, and the profile is rebuilt with the
    // hashes and paths of this build so the restored file matches what is installed.
    private void NormalizeStartupProfileBackup(DeviceTweakerBackup backup)
    {
        backup.SkippedNicItrEntries.Clear();
        backup.RegistryValues.RemoveAll(value =>
            string.Equals(value.Hive, "HKCU", StringComparison.OrdinalIgnoreCase)
            && string.Equals(value.Path.Trim('\\'), ImodStartupRunKeyPath, StringComparison.OrdinalIgnoreCase)
            && string.Equals(value.Name, ImodStartupRunValueName, StringComparison.OrdinalIgnoreCase));

        FileBackup? source = backup.StartupProfile ?? backup.ImodScript;
        backup.ImodScript = null;
        if (source is null)
        {
            backup.StartupProfile = null;
            return;
        }

        if (!source.Exists)
        {
            backup.StartupProfile = new FileBackup { Path = GetStartupProfilePath(), Exists = false };
            return;
        }

        string text = source.Text ?? throw new InvalidDataException("The backup is missing the saved startup profile contents.");
        bool legacyScript = !string.Equals(Path.GetFullPath(source.Path), Path.GetFullPath(GetStartupProfilePath()), StringComparison.OrdinalIgnoreCase);
        ImodConfig config = legacyScript
            ? ParseImodScriptText(text)
            : ConvertStartupProfile(StartupProfile.FromJson(text));
        foreach (string removed in RemoveStaleNicItrEntries(config))
        {
            backup.SkippedNicItrEntries.Add(removed);
            WriteLog($"BACKUP.RESTORE.NIC.SKIPPED: {removed} action=not-restored; set NIC ITR again for this adapter");
        }
        backup.StartupProfile = new FileBackup
        {
            Path = GetStartupProfilePath(),
            Exists = true,
            Text = BuildStartupProfile(config, GetImodDriverSystemPath()).ToJson(),
        };
    }

    // Rollback must use the same supported layout and current payload hashes as installation.
    // Readback itself stays raw so verification cannot hide an incorrectly written profile.
    private FileBackup PrepareRollbackStartupProfile(FileBackup profile)
    {
        DeviceTweakerBackup rollback = new() { StartupProfile = profile };
        NormalizeStartupProfileBackup(rollback);
        return rollback.StartupProfile!;
    }

    private void RestoreStartupProfile(FileBackup backup)
    {
        DemandRealWrite("RestoreStartupProfile");
        if (!string.Equals(Path.GetFullPath(backup.Path), Path.GetFullPath(GetStartupProfilePath()), StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException("Refusing to restore an unmanaged file path.");
        }

        if (!backup.Exists || backup.Text is null)
        {
            RemoveStartupBundle(report: null);
            return;
        }

        InstallStartupBundle(ConvertStartupProfile(StartupProfile.FromJson(backup.Text)));
    }
}
