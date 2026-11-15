using System.Text.Json;
using Microsoft.Win32;
using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

// Backup files go to EXE/APPDATA-like folders inside the QA run directory, the only
// place a read-only test process may change. Validation is the production rule set.
public sealed class BackupStoreTests : IDisposable
{
    private readonly string _root = Path.Combine(QaRun.OutputDirectory, "backup-store-" + Guid.NewGuid().ToString("N"));
    private readonly List<string> _log = [];
    private int _captures;

    private string Local => Path.Combine(_root, "EXE", "Backups");
    private string Roaming => Path.Combine(_root, "AppData", "DEVICE TWEAKER", "Backups");

    private BackupStore Store() => new(Local, Roaming, backup => BackupFormat.ValidateDocument(backup, _ => false), _log.Add);

    public void Dispose()
    {
        if (Directory.Exists(_root)) Directory.Delete(_root, recursive: true);
    }

    private static DeviceTweakerBackup Snapshot(string reason, string data = "AQ==", DateTime? created = null) => new()
    {
        CreatedAt = created ?? new DateTime(2026, 11, 15, 12, 0, 0),
        Reason = reason,
        RegistryValues = [new() { Hive = "HKLM", Path = @"SYSTEM\CurrentControlSet\Control\Session Manager\Kernel",
            Name = "ReservedCpuSets", Exists = true, Kind = "Binary", Data = data }],
    };

    private DeviceTweakerBackup Capture()
    {
        _captures++;
        return Snapshot("original-state", "Ag==");
    }

    private static string WriteFile(string directory, string name, DeviceTweakerBackup backup, int ageDays = 0)
    {
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, name);
        File.WriteAllText(path, JsonSerializer.Serialize(backup));
        File.SetLastWriteTimeUtc(path, DateTime.UtcNow.AddDays(-ageDays));
        return path;
    }

    [Fact]
    public void SaveWritesAValidatedSnapshotNamedAfterTheReason()
    {
        DeviceTweakerBackup backup = Snapshot("pre-apply");

        string path = Store().Save(Local, "Local", "pre-apply", backup);

        Assert.Equal(Local, Path.GetDirectoryName(path));
        Assert.Matches(@"^DeviceTweakerBackup_\d{8}_\d{6}_\d{3}_pre-apply\.json$", Path.GetFileName(path));
        Assert.Equal(JsonSerializer.Serialize(backup), JsonSerializer.Serialize(Store().Read(path)));
        Assert.Contains(_log, line => line.StartsWith("BACKUP.VALIDATION: status=passed schema=1 values=1 existing=1 missing=0 readErrors=0 imodScript=not-captured", StringComparison.Ordinal));
        Assert.Contains(_log, line => line.StartsWith($"BACKUP: saved location=Local path={path} values=1 reason=pre-apply", StringComparison.Ordinal));
        Assert.Empty(Directory.GetFiles(Local, "*.tmp"));
    }

    [Fact]
    public void SnapshotWithUnreadableValuesIsNotWritten()
    {
        DeviceTweakerBackup backup = Snapshot("pre-apply");
        backup.RegistryValues[0].Kind = "ReadError";
        backup.RegistryValues[0].Data = "access denied";

        InvalidOperationException error = Assert.Throws<InvalidOperationException>(() => Store().Save(Local, "Local", "pre-apply", backup));

        Assert.StartsWith("Backup capture is incomplete: 1 registry value(s) could not be read.", error.Message);
        Assert.Empty(Directory.GetFiles(Local));
    }

    [Theory]
    [InlineData("pre apply/../x", "preapplyx")]
    [InlineData("  ", "manual")]
    [InlineData("***", "manual")]
    [InlineData("pre-nic_itr-save", "pre-nic_itr-save")]
    public void FileReasonKeepsOnlySafeCharacters(string reason, string expected)
    {
        Assert.Equal(expected, BackupStore.MakeFileReason(reason));
    }

    [Fact]
    public void FirstOriginalIsCapturedOnceAndThenValidated()
    {
        Assert.True(Store().EnsureOriginal(Local, "Local", Capture));
        string original = Path.Combine(Local, BackupStore.OriginalFileName);
        Assert.Equal("Ag==", Store().Read(original).RegistryValues[0].Data);
        Assert.Contains(_log, line => line.StartsWith($"BACKUP.ORIGINAL: created location=Local path={original}", StringComparison.Ordinal));

        Assert.True(Store().EnsureOriginal(Roaming, "Roaming", Capture));

        Assert.Equal(1, _captures);
        Assert.False(Directory.Exists(Roaming));
        Assert.Contains(_log, line => line.StartsWith($"BACKUP.ORIGINAL: existing path={original}", StringComparison.Ordinal));
    }

    [Fact]
    public void OldestKnownSnapshotBecomesTheOriginal()
    {
        WriteFile(Roaming, "DeviceTweakerBackup_old.json", Snapshot("pre-apply", "Bw==", new DateTime(2026, 1, 1)), ageDays: 300);
        WriteFile(Local, "DeviceTweakerBackup_new.json", Snapshot("pre-apply", "CA==", new DateTime(2026, 10, 1)), ageDays: 40);

        Assert.True(Store().EnsureOriginal(Local, "Local", Capture));

        Assert.Equal(0, _captures);
        Assert.Equal("Bw==", Store().Read(Path.Combine(Local, BackupStore.OriginalFileName)).RegistryValues[0].Data);
    }

    [Fact]
    public void CorruptOriginalIsReportedAndKept()
    {
        Directory.CreateDirectory(Local);
        string original = Path.Combine(Local, BackupStore.OriginalFileName);
        File.WriteAllText(original, "{ not json");

        Assert.False(Store().EnsureOriginal(Local, "Local", Capture));

        Assert.Equal("{ not json", File.ReadAllText(original));
        Assert.Equal(0, _captures);
        Assert.Contains(_log, line => line.StartsWith($"BACKUP.ORIGINAL.ERROR: validation failed path={original}", StringComparison.Ordinal));
    }

    [Fact]
    public void SnapshotsFromBothFoldersAreListedNewestFirst()
    {
        string oldLocal = WriteFile(Local, "DeviceTweakerBackup_a.json", Snapshot("pre-apply", created: new DateTime(2026, 3, 1)));
        string newRoaming = WriteFile(Roaming, "DeviceTweakerBackup_b.json", Snapshot("pre-auto", created: new DateTime(2026, 9, 1)));
        string original = WriteFile(Local, BackupStore.OriginalFileName, Snapshot("original-state", created: new DateTime(2026, 1, 1)));
        WriteFile(Local, "DeviceTweakerBackup_broken.json", new DeviceTweakerBackup { Version = 99 });
        File.WriteAllText(Path.Combine(Local, "unrelated.json"), "{}");

        List<BackupSnapshotInfo> snapshots = Store().GetSnapshots();

        Assert.Equal(new[] { newRoaming, oldLocal, original }, snapshots.Select(s => s.Path));
        Assert.Equal(new[] { "APPDATA", "EXE", "EXE" }, snapshots.Select(s => s.Location));
        Assert.Equal(new[] { false, false, true }, snapshots.Select(s => s.IsOriginal));
        Assert.Contains(_log, line => line.StartsWith("BACKUP.SCAN.WARN: invalid backup excluded path=", StringComparison.Ordinal)
            && line.Contains("DeviceTweakerBackup_broken.json", StringComparison.Ordinal));
    }

    [Fact]
    public void MissingFoldersAreReportedAsEmpty()
    {
        Assert.Empty(Store().GetSnapshots());
        Assert.Contains($"BACKUP.SCAN: directory={Local} exists=false count=0", _log);
        Assert.Contains($"BACKUP.SCAN: directory={Roaming} exists=false count=0", _log);
    }

    [Fact]
    public void DeleteNeverRemovesTheOriginal()
    {
        WriteFile(Local, "DeviceTweakerBackup_a.json", Snapshot("pre-apply"));
        WriteFile(Local, BackupStore.OriginalFileName, Snapshot("original-state"));
        List<BackupSnapshotInfo> snapshots = Store().GetSnapshots();

        Assert.Equal(1, Store().Delete(snapshots));

        BackupSnapshotInfo remaining = Assert.Single(Store().GetSnapshots());
        Assert.True(remaining.IsOriginal);
    }

    [Fact]
    public void PruneKeepsTheLatestProtectedAndOriginalFiles()
    {
        List<string> paths = Enumerable.Range(0, 13)
            .Select(i => WriteFile(Local, $"DeviceTweakerBackup_{i:D2}.json", Snapshot("pre-apply"), ageDays: i))
            .ToList();
        string original = WriteFile(Local, BackupStore.OriginalFileName, Snapshot("original-state"), ageDays: 500);

        Store().Prune(Local, keepLatest: 10, protectedPaths: [paths[12]]);

        string[] kept = Directory.GetFiles(Local).Order(StringComparer.OrdinalIgnoreCase).ToArray();
        Assert.Equal(paths.Take(9).Append(paths[12]).Append(original).Order(StringComparer.OrdinalIgnoreCase), kept);
        Assert.Equal(3, _log.Count(line => line.StartsWith("BACKUP.PRUNE: deleted old backup path=", StringComparison.Ordinal)));
    }

    [Fact]
    public void FilesOutsideTheQaRunAreNeverDeleted()
    {
        string outside = Path.Combine(Path.GetTempPath(), "dt-backup-outside-" + Guid.NewGuid().ToString("N"));
        try
        {
            List<string> paths = Enumerable.Range(0, 3)
                .Select(i => WriteFile(outside, $"DeviceTweakerBackup_{i}.json", Snapshot("pre-apply"), ageDays: i))
                .ToList();
            BackupStore store = new(outside, Roaming, backup => BackupFormat.ValidateDocument(backup, _ => false), _log.Add);

            Assert.Throws<QaWriteBlockedException>(() => store.Prune(outside, keepLatest: 1));
            Assert.Throws<QaWriteBlockedException>(() => store.Delete(store.GetSnapshots()));
            Assert.Throws<QaWriteBlockedException>(() => store.Save(outside, "Local", "pre-apply", Snapshot("pre-apply")));

            Assert.All(paths, path => Assert.True(File.Exists(path)));
            Assert.Equal(3, Directory.GetFiles(outside).Length);
        }
        finally
        {
            Directory.Delete(outside, recursive: true);
        }
    }
}

public sealed class BackupRegistryTests
{
    private readonly List<string> _log = [];

    [Fact]
    public void MissingValueIsCapturedAsAbsent()
    {
        RegistryValueBackup value = new BackupRegistry(_log.Add)
            .Capture(RegistryHive.CurrentUser, @"Software\DEVICE TWEAKER QA\" + Guid.NewGuid().ToString("N"), "Missing");

        Assert.Equal(("HKCU", false, string.Empty, (string?)null), (value.Hive, value.Exists, value.Kind, value.Data));
        Assert.Empty(_log);
    }

    [Fact]
    public void ExistingValueIsCapturedWithItsKind()
    {
        using RegistryKey? desktop = Registry.CurrentUser.OpenSubKey(@"Control Panel\Desktop");
        Assert.NotNull(desktop);
        string name = desktop.GetValueNames().First(n => desktop.GetValueKind(n) == RegistryValueKind.String);

        RegistryValueBackup value = new BackupRegistry(_log.Add).Capture(RegistryHive.CurrentUser, @"Control Panel\Desktop", name);

        Assert.True(value.Exists);
        Assert.Equal("String", value.Kind);
        Assert.Equal((string)desktop.GetValue(name, null, RegistryValueOptions.DoNotExpandEnvironmentNames)!, value.Data);
    }

    [Fact]
    public void CaptureByDocumentRejectsAnUnknownHive()
    {
        Assert.Throws<InvalidOperationException>(() => new BackupRegistry(_log.Add)
            .Capture(new RegistryValueBackup { Hive = "HKCR", Path = "x", Name = "y" }));
    }

    [Fact]
    public void RestoreIsBlockedBeforeTouchingTheRegistry()
    {
        string path = @"Software\DEVICE TWEAKER QA\" + Guid.NewGuid().ToString("N");

        Assert.Throws<QaWriteBlockedException>(() => new BackupRegistry(_log.Add).Restore(new RegistryValueBackup
        {
            Hive = "HKCU", Path = path, Name = "Value", Exists = true, Kind = "DWord", Data = "1",
        }));

        Assert.Null(Registry.CurrentUser.OpenSubKey(path));
        Assert.Empty(_log);
    }
}
