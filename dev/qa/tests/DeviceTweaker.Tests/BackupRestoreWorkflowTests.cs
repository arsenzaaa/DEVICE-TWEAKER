using System.Text.Json;
using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public sealed class BackupRestoreWorkflowTests : IDisposable
{
    private readonly string _directory = Path.Combine(Path.GetTempPath(), "DeviceTweakerRestore-" + Guid.NewGuid().ToString("N"));
    public BackupRestoreWorkflowTests() => Directory.CreateDirectory(_directory);
    public void Dispose() => Directory.Delete(_directory, recursive: true);

    private string Write(string name, int age, string data = "AQ==")
    {
        string path = Path.Combine(_directory, name);
        File.WriteAllText(path, JsonSerializer.Serialize(new DeviceTweakerBackup
        {
            CreatedAt = DateTime.Today,
            RegistryValues = [new() { Hive = "HKLM", Path = @"SYSTEM\CurrentControlSet\Control\Session Manager\Kernel",
                Name = "ReservedCpuSets", Exists = true, Kind = "Binary", Data = data }],
        }));
        File.SetLastWriteTimeUtc(path, DateTime.UtcNow.AddDays(-age));
        return path;
    }

    private static DeviceTweakerBackup Read(string path)
    {
        var backup = JsonSerializer.Deserialize<DeviceTweakerBackup>(File.ReadAllText(path))!;
        BackupFormat.ValidateDocument(backup, _ => false);
        return backup;
    }

    [Fact]
    public void OldestSelectedFileSurvivesRollbackRotationAndPreparedValueIsRestored()
    {
        string selected = Write("DeviceTweakerBackup_old.json", 100);
        for (int i = 0; i < 9; i++) Write($"DeviceTweakerBackup_{i}.json", i + 1);
        string? restored = null;
        var workflow = new BackupRestoreWorkflow(Read, protectedPath =>
        {
            string rollback = Write("DeviceTweakerBackup_rollback.json", 0, "AA==");
            var entries = Directory.GetFiles(_directory).Select(path => new BackupFileEntry(path, File.GetLastWriteTimeUtc(path)));
            foreach (string path in BackupFileRetention.SelectForRemoval(entries, 10, [protectedPath, rollback])) File.Delete(path);
            return true;
        }, backup =>
        {
            var service = new BackupRestoreService(_ => { }, value => new RegistryValueBackup
                { Hive = value.Hive, Path = value.Path, Name = value.Name, Exists = restored is not null, Kind = "Binary", Data = restored },
                value => restored = value.Data, () => new(), _ => { }, _ => { });
            service.Restore(backup);
            return new();
        });

        Assert.True(workflow.Restore(selected).Succeeded);
        Assert.True(File.Exists(selected));
        Assert.True(File.Exists(Path.Combine(_directory, "DeviceTweakerBackup_rollback.json")));
        Assert.Equal(10, Directory.GetFiles(_directory).Length);
        Assert.Equal("AQ==", restored);
    }

    [Fact]
    public void PreparedDocumentIsUsedEvenIfSelectedFileChangesDuringRollbackCapture()
    {
        string selected = Write("selected.json", 0);
        DeviceTweakerBackup? applied = null;
        var workflow = new BackupRestoreWorkflow(Read, path => { File.WriteAllText(path, "broken"); return true; },
            backup => { applied = backup; return new(); });
        workflow.Restore(selected);
        Assert.Equal("AQ==", applied!.RegistryValues.Single().Data);
    }

    [Fact]
    public void InvalidDocumentStopsBeforeRollbackCreation()
    {
        string selected = Write("selected.json", 0, "invalid base64");
        bool rollback = false, applied = false;
        var workflow = new BackupRestoreWorkflow(Read, _ => { rollback = true; return true; }, _ => { applied = true; return new(); });
        Assert.Throws<InvalidOperationException>(() => workflow.Restore(selected));
        Assert.False(rollback);
        Assert.False(applied);
    }

    [Fact]
    public void FailedRollbackCreationPreventsRestore()
    {
        string selected = Write("selected.json", 0);
        bool applied = false;
        var workflow = new BackupRestoreWorkflow(Read, _ => false, _ => { applied = true; return new(); });
        Assert.Throws<InvalidOperationException>(() => workflow.Restore(selected));
        Assert.False(applied);
    }

    [Fact]
    public void OriginalAndProtectedFilesSurviveEvenWithFutureTimestampsAndSmallLimit()
    {
        string original = Write(BackupFileRetention.OriginalFileName, 100);
        string selected = Write("selected.json", 50);
        string rollback = Write("rollback.json", 0);
        string future = Write("future.json", -100);
        var entries = Directory.GetFiles(_directory).Select(path => new BackupFileEntry(path, File.GetLastWriteTimeUtc(path)));
        var removed = BackupFileRetention.SelectForRemoval(entries, 1, [selected.ToUpperInvariant(), rollback]);
        Assert.Equal(new[] { future }, removed);
        Assert.DoesNotContain(original, removed);
    }
}
