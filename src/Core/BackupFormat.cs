using Microsoft.Win32;
using System.Globalization;
using System.Text.Json;

namespace DeviceTweakerCS;

internal static class BackupFormat
{
    internal static string EncodeRegistryValue(object value, RegistryValueKind kind)
    {
        return kind switch
        {
            RegistryValueKind.Binary => Convert.ToBase64String(value as byte[] ?? []),
            RegistryValueKind.MultiString => JsonSerializer.Serialize(value as string[] ?? []),
            RegistryValueKind.DWord or RegistryValueKind.QWord => Convert.ToInt64(value, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture),
            _ => value.ToString() ?? string.Empty,
        };
    }

    internal static void ValidateRegistryValueBackup(RegistryValueBackup value)
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
            if (UsbSelectiveSuspendPolicy.TryParsePowerPlanTarget(value.Hive, value.Path, value.Name, out _, out _)
                && (kind != RegistryValueKind.DWord || !uint.TryParse(value.Data, out uint index) || index > 1))
                throw new FormatException("USB selective suspend power index must be a DWORD 0 or 1.");
        }
        catch (Exception ex) when (ex is FormatException or JsonException or OverflowException)
        {
            throw new InvalidOperationException(
                $"Backup contains invalid registry data: {value.Hive}\\{value.Path}\\{value.Name} kind={value.Kind}",
                ex);
        }
    }

    internal static bool IsManagedBackupRegistryValue(RegistryValueBackup value)
    {
        if (string.IsNullOrWhiteSpace(value.Hive) || string.IsNullOrWhiteSpace(value.Path) || string.IsNullOrWhiteSpace(value.Name)) return false;
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
                && string.Equals(name, "RawMouseThrottleDuration", StringComparison.OrdinalIgnoreCase);
            bool isImodStartup = string.Equals(path, @"Software\Microsoft\Windows\CurrentVersion\Run", StringComparison.OrdinalIgnoreCase)
                && string.Equals(name, "DEVICE TWEAKER IMOD", StringComparison.OrdinalIgnoreCase);
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

        if (UsbSelectiveSuspendPolicy.TryParsePowerPlanTarget(hive, path, name, out _, out _))
            return true;

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
                or "PnPCapabilities"
                or "*InterruptModeration"
                or "ITR";
        }

        return false;
    }

    internal static bool TryGetBackupHive(string text, out RegistryHive hive)
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

    internal static bool TryParseRegistryValueKind(string text, out RegistryValueKind kind)
    {
        return Enum.TryParse(text, ignoreCase: true, out kind)
            && kind is RegistryValueKind.String
                or RegistryValueKind.ExpandString
                or RegistryValueKind.Binary
                or RegistryValueKind.DWord
                or RegistryValueKind.MultiString
                or RegistryValueKind.QWord;
    }

    internal static object DecodeRegistryValue(string? data, RegistryValueKind kind)
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
    internal static void ValidateDocument(DeviceTweakerBackup backup, Func<string, bool> isManagedScriptPath)
    {
        if (backup.Version != 1)
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

        foreach (FileBackup? file in new[] { backup.StartupProfile, backup.ImodScript })
        {
            if (file is not null && !isManagedScriptPath(file.Path))
            {
                throw new InvalidOperationException("Backup contains an unmanaged startup file path.");
            }
        }
    }

}
