using Microsoft.Win32;
using System.Globalization;

namespace DeviceTweakerCS;

// Reads and writes one backed-up value: a registry value in the 64-bit view, or a USB
// selective-suspend power-plan index, which Windows exposes only through the Power API.
internal sealed class BackupRegistry(Action<string> log)
{
    public RegistryValueBackup Capture(RegistryHive hive, string path, string name)
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
            if (UsbSelectiveSuspendPolicy.TryParsePowerPlanTarget(backup.Hive, path, name, out Guid scheme, out bool ac))
            {
                backup.Exists = true;
                backup.Kind = RegistryValueKind.DWord.ToString();
                backup.Data = UsbSelectiveSuspendPolicy.ReadPowerPlanIndex(scheme, ac).ToString(CultureInfo.InvariantCulture);
                return backup;
            }
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
            backup.Data = BackupFormat.EncodeRegistryValue(value, kind);
        }
        catch (Exception ex)
        {
            backup.Kind = "ReadError";
            backup.Data = ex.Message;
            backup.Exists = true; // must not look like "absent" - restore deletes Exists=false
            log($"BACKUP.REG: read failed {backup.Hive}\\{path}\\{name}: {ex.Message}");
        }

        return backup;
    }

    public RegistryValueBackup Capture(RegistryValueBackup value)
    {
        if (!BackupFormat.TryGetBackupHive(value.Hive, out RegistryHive hive))
            throw new InvalidOperationException($"Unsupported registry hive: {value.Hive}.");
        return Capture(hive, value.Path, value.Name);
    }

    public void Restore(RegistryValueBackup value)
    {
        QaRun.DemandWrite("BackupRegistry.Restore");
        if (!BackupFormat.TryGetBackupHive(value.Hive, out RegistryHive hive))
        {
            throw new InvalidOperationException($"Unsupported registry hive for restore: {value.Hive}.");
        }

        if (UsbSelectiveSuspendPolicy.TryParsePowerPlanTarget(value.Hive, value.Path, value.Name, out Guid scheme, out bool ac))
        {
            uint index = value.Exists
                ? uint.Parse(value.Data ?? string.Empty, CultureInfo.InvariantCulture)
                : UsbSelectiveSuspendPolicy.ReadPowerPlanDefaultIndex(scheme, ac);
            UsbSelectiveSuspendPolicy.RestorePowerPlanIndex(scheme, ac, index);
            log($"BACKUP.RESTORE.POWER: scheme={scheme:D} source={(ac ? "AC" : "DC")} index={index} readback=verified");
            return;
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
            log($"BACKUP.RESTORE.REG: deleted {value.Hive}\\{value.Path}\\{value.Name}");
            return;
        }

        if (string.Equals(value.Kind, "ReadError", StringComparison.OrdinalIgnoreCase))
        {
            log($"BACKUP.RESTORE.REG: skipped {value.Hive}\\{value.Path}\\{value.Name} kind=ReadError");
            return;
        }

        if (!BackupFormat.TryParseRegistryValueKind(value.Kind, out RegistryValueKind kind))
        {
            throw new InvalidOperationException(
                $"Unsupported registry value kind for restore: {value.Hive}\\{value.Path}\\{value.Name} kind={value.Kind}");
        }

        object data = BackupFormat.DecodeRegistryValue(value.Data, kind);
        key.SetValue(value.Name, data, kind);
        log($"BACKUP.RESTORE.REG: set {value.Hive}\\{value.Path}\\{value.Name} kind={kind}");
    }
}
