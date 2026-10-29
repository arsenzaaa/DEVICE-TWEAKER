namespace DeviceTweakerCS;

internal sealed class DeviceTweakerBackup
{
    public int Version { get; set; } = 1;
    public DateTime CreatedAt { get; set; }
    public string Reason { get; set; } = string.Empty;
    public List<RegistryValueBackup> RegistryValues { get; set; } = [];
    public FileBackup? ImodScript { get; set; }
}

internal sealed class RegistryValueBackup
{
    public string Hive { get; set; } = string.Empty;
    public string Path { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public bool Exists { get; set; }
    public string Kind { get; set; } = string.Empty;
    public string? Data { get; set; }
}

internal sealed class FileBackup
{
    public string Path { get; set; } = string.Empty;
    public bool Exists { get; set; }
    public string? Text { get; set; }
}

// The same transaction runs against Windows in the app and memory in developer QA.
// UI prompts, backup selection and post-restore refresh belong to MainForm.
internal sealed class BackupRestoreService(
    Action<DeviceTweakerBackup> validate,
    Func<RegistryValueBackup, RegistryValueBackup> captureValue,
    Action<RegistryValueBackup> restoreValue,
    Func<FileBackup> captureScript,
    Action<FileBackup> restoreScript,
    Action<string> log)
{
    internal void Restore(DeviceTweakerBackup backup)
    {
        validate(backup);
        var targets = backup.RegistryValues
            .Where(v => !string.Equals(v.Kind, "ReadError", StringComparison.OrdinalIgnoreCase)).ToArray();
        // Capture every rollback target before the first write. A failed read cannot
        // leave earlier targets already changed.
        var original = targets.Select(captureValue).ToArray();
        foreach (var value in original)
            if (string.Equals(value.Kind, "ReadError", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException($"Cannot capture rollback value: {value.Hive}\\{value.Path}\\{value.Name}. {value.Data}");
        FileBackup? script = backup.ImodScript is null ? null : captureScript();
        var attempted = new List<RegistryValueBackup>();
        bool scriptAttempted = false;
        try
        {
            for (int i = 0; i < targets.Length; i++)
            {
                attempted.Add(original[i]); // A failing backend may have written partially.
                restoreValue(targets[i]);
                Verify(targets[i], captureValue(targets[i]));
            }
            if (backup.ImodScript is not null)
            {
                scriptAttempted = true;
                restoreScript(backup.ImodScript);
                VerifyScript(backup.ImodScript, captureScript());
            }
        }
        catch (Exception error)
        {
            log($"BACKUP.RESTORE.ROLLBACK: starting after error=\"{error.Message.Replace('\r', ' ').Replace('\n', ' ')}\"");
            var errors = new List<string>();
            // Script operations can also alter the Run value. Restore it first,
            // then restore registry targets so the original snapshot wins.
            if (scriptAttempted && script is not null)
            {
                try { restoreScript(script); VerifyScript(script, captureScript()); }
                catch (Exception ex) { errors.Add($"IMOD script: {ex.Message}"); }
            }
            foreach (var value in attempted.AsEnumerable().Reverse())
            {
                try { restoreValue(value); Verify(value, captureValue(value)); }
                catch (Exception ex) { errors.Add($"{value.Hive}\\{value.Path}\\{value.Name}: {ex.Message}"); }
            }
            if (errors.Count > 0)
                throw new InvalidOperationException($"Backup restore failed and rollback was incomplete. Restore error: {error.Message}. Rollback errors: {string.Join(" | ", errors)}", error);
            log("BACKUP.RESTORE.ROLLBACK: completed");
            throw;
        }
    }

    private static void Verify(RegistryValueBackup expected, RegistryValueBackup actual)
    {
        // Old backups can lack a power index. The Windows backend restores the
        // scheme personality's default; it cannot delete a Power API setting.
        if (!expected.Exists && UsbSelectiveSuspendPolicy.TryParsePowerPlanTarget(expected.Hive, expected.Path, expected.Name, out _, out _)) return;
        if (expected.Exists != actual.Exists || (expected.Exists &&
            (!string.Equals(expected.Kind, actual.Kind, StringComparison.OrdinalIgnoreCase) || expected.Data != actual.Data)))
            throw new InvalidOperationException($"Restore readback mismatch: {expected.Hive}\\{expected.Path}\\{expected.Name}.");
    }

    private static void VerifyScript(FileBackup expected, FileBackup actual)
    {
        if (expected.Exists != actual.Exists || (expected.Exists && expected.Text != actual.Text))
            throw new InvalidOperationException("Restore readback mismatch: IMOD script.");
    }
}
