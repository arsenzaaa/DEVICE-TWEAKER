using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace DeviceTweakerCS;

internal sealed class BackupSnapshotInfo
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
        string reason = BackupReasonLabel.Get(IsOriginal ? "original-state" : Reason);
        return $"{stamp:yyyy-MM-dd HH:mm:ss} [{Location}] {reason}";
    }
}

// Backup files in the EXE folder and in %APPDATA%: writing a snapshot and proving the
// file reads back identically, the protected original-state snapshot, listing,
// reading, deleting and rotation. Which values a snapshot holds, and restoring them,
// belong to the caller. A read-only QA process may change files only in its run directory.
internal sealed class BackupStore(
    string localDirectory,
    string roamingDirectory,
    Action<DeviceTweakerBackup> validate,
    Action<string> log)
{
    internal const string FilePrefix = "DeviceTweakerBackup_";
    internal const string OriginalFileName = BackupFileRetention.OriginalFileName;

    internal sealed record ValidationResult(
        int ValueCount,
        int ExistingCount,
        int MissingCount,
        int ReadErrorCount,
        string ImodScriptState,
        long Bytes,
        string Sha256);

    public IEnumerable<string> Directories()
    {
        HashSet<string> seen = new(StringComparer.OrdinalIgnoreCase);
        foreach (string directory in new[] { localDirectory, roamingDirectory })
        {
            if (seen.Add(directory))
            {
                yield return directory;
            }
        }
    }

    // Writes a new snapshot; the caller prunes the directory afterwards.
    public string Save(string directory, string location, string reason, DeviceTweakerBackup backup)
    {
        Directory.CreateDirectory(directory);

        string safeReason = MakeFileReason(reason);
        string stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff", CultureInfo.InvariantCulture);
        string path = Path.Combine(directory, $"{FilePrefix}{stamp}_{safeReason}.json");
        if (File.Exists(path))
        {
            path = Path.Combine(directory, $"{FilePrefix}{stamp}_{safeReason}_{Guid.NewGuid().ToString("N")[..8]}.json");
        }
        JsonSerializerOptions options = new() { WriteIndented = true };
        ValidateForWrite(backup);
        AtomicFile.WriteAllText(path, JsonSerializer.Serialize(backup, options), Encoding.UTF8);
        ValidationResult validation = ValidateWritten(path, backup);
        log(
            $"BACKUP.VALIDATION: status=passed schema={backup.Version} values={validation.ValueCount} " +
            $"existing={validation.ExistingCount} missing={validation.MissingCount} readErrors={validation.ReadErrorCount} " +
            $"imodScript={validation.ImodScriptState} bytes={validation.Bytes} sha256={validation.Sha256}");
        log($"BACKUP: saved location={location} path={path} values={backup.RegistryValues.Count} reason={reason} atomic=true validated=true");
        return path;
    }

    // An existing original-state snapshot must validate. Without one, the oldest known
    // snapshot becomes the original, or a fresh capture when there is none.
    public bool EnsureOriginal(string preferredDirectory, string location, Func<DeviceTweakerBackup> captureCurrent)
    {
        foreach (string directory in Directories())
        {
            string existingPath = Path.Combine(directory, OriginalFileName);
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
                    log($"BACKUP.ORIGINAL.ERROR: invalid snapshot path={existingPath}");
                    return false;
                }

                validate(existing);

                ValidationResult validation = ValidateWritten(existingPath, existing);
                log(
                    $"BACKUP.ORIGINAL.VALIDATION: status=passed schema={existing.Version} values={validation.ValueCount} " +
                    $"existing={validation.ExistingCount} missing={validation.MissingCount} readErrors={validation.ReadErrorCount} " +
                    $"imodScript={validation.ImodScriptState} bytes={validation.Bytes} sha256={validation.Sha256}");
                log($"BACKUP.ORIGINAL: existing path={existingPath} created={existing.CreatedAt:O} validated=true");
                return true;
            }
            catch (Exception ex)
            {
                log($"BACKUP.ORIGINAL.ERROR: validation failed path={existingPath} error=\"{FlattenLogText(ex.ToString())}\"");
                return false;
            }
        }

        try
        {
            Directory.CreateDirectory(preferredDirectory);
            string path = Path.Combine(preferredDirectory, OriginalFileName);

            DeviceTweakerBackup? original = null;
            BackupSnapshotInfo? oldestKnown = GetSnapshots()
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

                validate(original);

                log($"BACKUP.ORIGINAL: importing oldest known snapshot path={oldestKnown.Path} created={original.CreatedAt:O}");
            }

            original ??= captureCurrent();
            ValidateForWrite(original);
            string json = JsonSerializer.Serialize(original, new JsonSerializerOptions { WriteIndented = true });
            AtomicFile.WriteAllText(path, json, Encoding.UTF8);
            ValidationResult validation = ValidateWritten(path, original);
            log(
                $"BACKUP.ORIGINAL.VALIDATION: status=passed schema={original.Version} values={validation.ValueCount} " +
                $"existing={validation.ExistingCount} missing={validation.MissingCount} readErrors={validation.ReadErrorCount} " +
                $"imodScript={validation.ImodScriptState} bytes={validation.Bytes} sha256={validation.Sha256}");
            log($"BACKUP.ORIGINAL: created location={location} path={path} values={original.RegistryValues.Count} atomic=true validated=true");
            return true;
        }
        catch (Exception ex)
        {
            log($"BACKUP.ORIGINAL.ERROR: creation failed location={location} error=\"{FlattenLogText(ex.ToString())}\"");
            return false;
        }
    }

    public List<BackupSnapshotInfo> GetSnapshots()
    {
        List<BackupSnapshotInfo> backups = [];
        foreach (string directory in Directories())
        {
            try
            {
                if (!Directory.Exists(directory))
                {
                    log($"BACKUP.SCAN: directory={directory} exists=false count=0");
                    continue;
                }

                int before = backups.Count;
                foreach (string path in Directory.EnumerateFiles(directory, $"{FilePrefix}*.json"))
                {
                    BackupSnapshotInfo? info = CreateSnapshotInfo(path, directory);
                    if (info is not null)
                    {
                        backups.Add(info);
                    }
                }

                log($"BACKUP.SCAN: directory={directory} exists=true count={backups.Count - before}");
            }
            catch (Exception ex)
            {
                log($"BACKUP.SCAN.WARN: directory={directory} failed={ex.Message}");
            }
        }

        return backups
            .OrderByDescending(info => info.CreatedAt ?? info.LastWriteUtc.ToLocalTime())
            .ThenByDescending(info => info.LastWriteUtc)
            .ToList();
    }

    public DeviceTweakerBackup Read(string path)
    {
        string json = File.ReadAllText(path, Encoding.UTF8);
        DeviceTweakerBackup? backup = JsonSerializer.Deserialize<DeviceTweakerBackup>(json);
        if (backup is null)
        {
            throw new InvalidOperationException("Backup file is empty or invalid.");
        }

        validate(backup);
        return backup;
    }

    // The original-state snapshot is never deleted.
    public int Delete(IReadOnlyList<BackupSnapshotInfo> backups)
    {
        int deleted = 0;
        foreach (BackupSnapshotInfo backup in backups)
        {
            if (backup.IsOriginal)
            {
                log($"BACKUP.DELETE: protected original snapshot path={backup.Path}");
                continue;
            }

            QaRun.DemandArtifactPath(backup.Path);
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
                log($"BACKUP.DELETE: failed path={backup.Path}: {ex.Message}");
            }
        }

        return deleted;
    }

    public void Prune(string directory, int keepLatest, IEnumerable<string>? protectedPaths = null)
    {
        try
        {
            if (!Directory.Exists(directory))
            {
                return;
            }

            BackupFileEntry[] files = Directory.EnumerateFiles(directory, $"{FilePrefix}*.json")
                .Select(path => new FileInfo(path))
                .Select(file => new BackupFileEntry(file.FullName, file.LastWriteTimeUtc))
                .ToArray();

            foreach (string path in BackupFileRetention.SelectForRemoval(files, keepLatest, protectedPaths))
            {
                QaRun.DemandArtifactPath(path);
                try
                {
                    File.Delete(path);
                    log($"BACKUP.PRUNE: deleted old backup path={path}");
                }
                catch (Exception ex)
                {
                    log($"BACKUP.PRUNE.WARN: failed path={path}: {ex.Message}");
                }
            }
        }
        catch (QaWriteBlockedException)
        {
            throw;
        }
        catch (Exception ex)
        {
            log($"BACKUP.PRUNE.WARN: failed directory={directory}: {ex.Message}");
        }
    }

    // A snapshot with unreadable values could not restore them, so it is not written.
    internal void ValidateForWrite(DeviceTweakerBackup backup)
    {
        validate(backup);
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

    internal ValidationResult ValidateWritten(string path, DeviceTweakerBackup expected)
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

        validate(actual);
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
        string imodScriptState = actual.StartupProfile switch
        {
            null => "not-captured",
            { Exists: true } => "present",
            _ => "absent",
        };

        return new ValidationResult(
            actual.RegistryValues.Count,
            existing,
            actual.RegistryValues.Count - existing,
            readErrors,
            imodScriptState,
            file.Length,
            sha256);
    }

    internal static string MakeFileReason(string reason)
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

    private BackupSnapshotInfo? CreateSnapshotInfo(string path, string directory)
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

                validate(backup);
            }
            catch (Exception ex)
            {
                log($"BACKUP.SCAN.WARN: invalid backup excluded path={path} error=\"{FlattenLogText(ex.ToString())}\"");
                return null;
            }

            string location = string.Equals(directory, roamingDirectory, StringComparison.OrdinalIgnoreCase)
                ? "APPDATA"
                : "EXE";

            return new BackupSnapshotInfo
            {
                Path = path,
                Location = location,
                LastWriteUtc = file.LastWriteTimeUtc,
                CreatedAt = backup?.CreatedAt,
                Reason = backup?.Reason ?? string.Empty,
                IsOriginal = string.Equals(file.Name, OriginalFileName, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(backup?.Reason, "original-state", StringComparison.OrdinalIgnoreCase),
            };
        }
        catch
        {
            return null;
        }
    }

    private static string FlattenLogText(string? value) => LogText.Flatten(value);
}
