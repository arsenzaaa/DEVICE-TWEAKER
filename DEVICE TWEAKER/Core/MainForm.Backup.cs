using Microsoft.Win32;
using System.Globalization;
using System.Text;
using System.Text.Json;

namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private const int DeviceTweakerBackupVersion = 1;
    private const string BackupFolderName = "Backups";
    private const string BackupFilePrefix = "DeviceTweakerBackup_";

    private enum BackupLocation
    {
        Local,
        Roaming,
    }

    private enum RestoreChoice
    {
        Cancel,
        ResetDefault,
        RestoreBackup,
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

            DeviceTweakerBackup backup = CaptureDeviceTweakerBackup(reason);
            string directory = GetBackupDirectory(location);
            Directory.CreateDirectory(directory);

            string safeReason = MakeBackupFileReason(reason);
            string stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture);
            string path = Path.Combine(directory, $"{BackupFilePrefix}{stamp}_{safeReason}.json");
            JsonSerializerOptions options = new() { WriteIndented = true };
            File.WriteAllText(path, JsonSerializer.Serialize(backup, options), Encoding.UTF8);
            WriteLog($"BACKUP: saved location={location} path={path} values={backup.RegistryValues.Count} reason={reason}");

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
                ShowThemedInfo($"Backup failed.\n{ex.Message}");
            }

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

            if (block.Kind == DeviceKind.NET_NDIS)
            {
                string? classKey = GetClassKeyForDevice(block.Device.InstanceId);
                if (!string.IsNullOrWhiteSpace(classKey))
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

        string startupScript = GetImodStartupPath();
        backup.ImodScript = new FileBackup
        {
            Path = startupScript,
            Exists = File.Exists(startupScript),
            Text = File.Exists(startupScript) ? File.ReadAllText(startupScript, Encoding.UTF8) : null,
        };

        return backup;
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
            string? path = GetLatestBackupPath();
            RestoreChoice choice = ShowRestoreChoiceDialog(path);
            if (choice == RestoreChoice.Cancel)
            {
                WriteLog("BACKUP.RESTORE: canceled by user");
                return;
            }

            if (choice == RestoreChoice.ResetDefault)
            {
                WriteLog("BACKUP.RESTORE: reset-default requested");
                ResetAllTweaks();
                RefreshBlocks();
                return;
            }

            if (string.IsNullOrWhiteSpace(path))
            {
                WriteLog("BACKUP.RESTORE: restore-backup requested but no backup found");
                ShowThemedInfo("No backup files were found in the EXE folder or APPDATA.");
                return;
            }

            WriteLog($"BACKUP.RESTORE: restore-backup requested path={path}");
            RestoreDeviceTweakerBackup(path);
            ShowThemedInfo($"Backup restored.\n{path}\n\nPlease reboot your PC to finish applying restored settings.");
            RefreshBlocks();
        }
        catch (Exception ex)
        {
            WriteLog($"BACKUP.RESTORE: failed: {ex.Message}");
            ShowThemedInfo($"Backup restore failed.\n{ex.Message}");
        }
    }

    private string? GetLatestBackupPath()
    {
        return EnumerateBackupDirectories()
            .Where(Directory.Exists)
            .SelectMany(directory => Directory.EnumerateFiles(directory, $"{BackupFilePrefix}*.json"))
            .Select(path => new FileInfo(path))
            .OrderByDescending(file => file.LastWriteTimeUtc)
            .FirstOrDefault()
            ?.FullName;
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

    private BackupLocation PromptBackupLocationForAuto()
    {
        bool local = ShowThemedConfirm(
            "Where should DEVICE TWEAKER save the pre-auto backup?\n\nEXE FOLDER = portable backup next to the app.\nAPPDATA = user profile backup that survives app folder cleanup.",
            "AUTO BACKUP",
            "EXE FOLDER",
            "APPDATA");
        BackupLocation location = local ? BackupLocation.Local : BackupLocation.Roaming;
        WriteLog($"BACKUP.PROMPT.AUTO: location={location}");
        return location;
    }

    private void RestoreDeviceTweakerBackup(string path)
    {
        string json = File.ReadAllText(path, Encoding.UTF8);
        DeviceTweakerBackup? backup = JsonSerializer.Deserialize<DeviceTweakerBackup>(json);
        if (backup is null)
        {
            throw new InvalidOperationException("Backup file is empty or invalid.");
        }

        foreach (RegistryValueBackup value in backup.RegistryValues)
        {
            RestoreRegistryValue(value);
        }

        if (backup.ImodScript is not null)
        {
            RestoreFileBackup(backup.ImodScript);
            InvalidateImodCache();
        }

        WriteLog($"BACKUP.RESTORE: restored path={path} values={backup.RegistryValues.Count} reason={backup.Reason}");
    }

    private void RestoreRegistryValue(RegistryValueBackup value)
    {
        if (!TryGetBackupHive(value.Hive, out RegistryHive hive))
        {
            return;
        }

        using RegistryKey baseKey = RegistryKey.OpenBaseKey(hive, RegistryView.Registry64);
        using RegistryKey? key = baseKey.CreateSubKey(value.Path, writable: true);
        if (key is null)
        {
            WriteLog($"BACKUP.RESTORE.REG: failed to open {value.Hive}\\{value.Path}");
            return;
        }

        if (!value.Exists)
        {
            key.DeleteValue(value.Name, throwOnMissingValue: false);
            WriteLog($"BACKUP.RESTORE.REG: deleted {value.Hive}\\{value.Path}\\{value.Name}");
            return;
        }

        if (!TryParseRegistryValueKind(value.Kind, out RegistryValueKind kind))
        {
            WriteLog($"BACKUP.RESTORE.REG: skipped {value.Hive}\\{value.Path}\\{value.Name} kind={value.Kind}");
            return;
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

    private static void RestoreFileBackup(FileBackup backup)
    {
        if (string.IsNullOrWhiteSpace(backup.Path))
        {
            return;
        }

        if (!backup.Exists)
        {
            if (File.Exists(backup.Path))
            {
                File.Delete(backup.Path);
            }
            return;
        }

        string? directory = Path.GetDirectoryName(backup.Path);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        File.WriteAllText(backup.Path, backup.Text ?? string.Empty, Encoding.UTF8);
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
