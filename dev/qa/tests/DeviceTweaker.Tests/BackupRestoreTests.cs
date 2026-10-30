using System.Text.Json;
using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public sealed class BackupRestoreTests
{
    private static RegistryValueBackup Value(string name, string data = "1") => new()
    { Hive = "HKLM", Path = @"SYSTEM\CurrentControlSet\Control\Session Manager\Kernel", Name = name, Exists = true, Kind = "Binary", Data = data };
    private static DeviceTweakerBackup Document(params RegistryValueBackup[] values) => new()
    { CreatedAt = DateTime.Today, RegistryValues = [.. values] };
    private sealed class Memory
    {
        internal readonly Dictionary<string, RegistryValueBackup> Values = new();
        internal readonly List<string> Writes = [];
        internal string? FailData;
        internal bool IgnoreWrite;
        internal bool DenyRollback;
        internal string? ReadErrorName;
        internal FileBackup Script = new() { Path = "managed", Exists = true, Text = "original" };
        internal bool FailScript;
        private static string Key(RegistryValueBackup v) => v.Hive + "|" + v.Path + "|" + v.Name;
        internal RegistryValueBackup Read(RegistryValueBackup v)
        {
            if (v.Name == ReadErrorName) return new() { Hive = v.Hive, Path = v.Path, Name = v.Name, Exists = true, Kind = "ReadError", Data = "denied" };
            return Values.TryGetValue(Key(v), out var saved) ? Clone(saved) : new() { Hive = v.Hive, Path = v.Path, Name = v.Name };
        }
        internal void Seed(RegistryValueBackup v) => Values[Key(v)] = Clone(v);
        internal void Write(RegistryValueBackup v)
        {
            Writes.Add(v.Name + ":" + v.Data);
            if (DenyRollback && v.Data == "AA==") throw new IOException("rollback denied");
            if (!IgnoreWrite) Seed(v);
            if (FailData is not null && v.Data == FailData) throw new IOException("write failed after modification");
        }
        internal BackupRestoreService Service(bool validate = false) => new(
            b => { if (validate) BackupFormat.ValidateDocument(b, p => p == "managed"); }, Read, Write,
            () => new() { Path = Script.Path, Exists = Script.Exists, Text = Script.Text },
            s => { Script = new() { Path = "managed", Exists = s.Exists, Text = s.Text }; if (FailScript && s.Text == "new") throw new IOException("script failed after modification"); }, _ => { });
        private static RegistryValueBackup Clone(RegistryValueBackup v) => JsonSerializer.Deserialize<RegistryValueBackup>(JsonSerializer.Serialize(v))!;
    }

    [Fact] public void RestoreChangesValuesAndDeletesMissingValue()
    {
        Memory m = new(); var a = Value("first", "AA=="); var b = Value("second", "AA=="); m.Seed(a); m.Seed(b);
        b.Exists = false; m.Service().Restore(Document(Value("first", "AQ=="), b));
        Assert.Equal("AQ==", m.Read(a).Data); Assert.False(m.Read(b).Exists);
    }
    [Fact] public void ReadFailureBeforeWritesLeavesAllTargetsUntouched()
    {
        Memory m = new() { ReadErrorName = "second" }; m.Seed(Value("first", "AA=="));
        Assert.Throws<InvalidOperationException>(() => m.Service().Restore(Document(Value("first", "AQ=="), Value("second", "AQ=="))));
        Assert.Empty(m.Writes); Assert.Equal("AA==", m.Read(Value("first")).Data);
    }
    [Theory] [InlineData(0)] [InlineData(1)] [InlineData(2)]
    public void WriteFailureRestoresEveryAttemptedTargetInReverseOrder(int failed)
    {
        Memory m = new(); var targets = Enumerable.Range(0, 3).Select(i => Value($"v{i}", Convert.ToBase64String([(byte)(i + 1)]))).ToArray();
        foreach (var t in targets) m.Seed(Value(t.Name, "AA==")); m.FailData = targets[failed].Data;
        Assert.Throws<IOException>(() => m.Service().Restore(Document(targets)));
        Assert.All(targets, t => Assert.Equal("AA==", m.Read(t).Data));
        Assert.Equal(Enumerable.Range(0, failed + 1).Reverse().Select(i => $"v{i}:AA=="), m.Writes.Skip(failed + 1));
    }
    [Fact] public void SilentWriteFailureIsDetectedByReadback()
    {
        Memory m = new() { IgnoreWrite = true }; m.Seed(Value("first", "AA=="));
        var error = Assert.Throws<InvalidOperationException>(() => m.Service().Restore(Document(Value("first", "AQ=="))));
        Assert.Contains("readback mismatch", error.Message); Assert.Equal("AA==", m.Read(Value("first")).Data);
    }
    [Fact] public void RollbackFailureKeepsBothErrors()
    {
        Memory m = new() { FailData = "AQ==", DenyRollback = true }; m.Seed(Value("first", "AA=="));
        var error = Assert.Throws<InvalidOperationException>(() => m.Service().Restore(Document(Value("first", "AQ=="))));
        Assert.Contains("write failed", error.Message); Assert.Contains("rollback denied", error.Message);
    }
    [Fact] public void ScriptFailureRestoresScriptAndRegistry()
    {
        Memory m = new() { FailScript = true }; m.Seed(Value("first", "AA=="));
        var b = Document(Value("first", "AQ==")); b.StartupProfile = new() { Path = "managed", Exists = true, Text = "new" };
        Assert.Throws<IOException>(() => m.Service().Restore(b));
        Assert.Equal("original", m.Script.Text); Assert.Equal("AA==", m.Read(Value("first")).Data);
    }
    [Theory] [InlineData("duplicate")] [InlineData("version")] [InlineData("data")] [InlineData("target")] [InlineData("script")]
    public void InvalidDocumentIsRejectedBeforeAnyWrite(string fault)
    {
        Memory m = new(); var v = Value("ReservedCpuSets", "AA=="); var b = Document(v);
        switch (fault) { case "duplicate": b.RegistryValues.Add(v); break; case "version": b.Version = 99; break; case "data": v.Data = "invalid base64!"; break; case "target": v.Name = "unmanaged"; break; case "script": b.ImodScript = new() { Path = "unmanaged" }; break; }
        Assert.Throws<InvalidOperationException>(() => m.Service(true).Restore(b)); Assert.Empty(m.Writes);
    }
    [Theory] [InlineData("ACSettingIndex", true)] [InlineData("DCSettingIndex", false)]
    public void PowerTargetsPreserveSavedSchemeAndAcDc(string name, bool ac)
    {
        Guid scheme = new("8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c");
        Assert.True(UsbSelectiveSuspendPolicy.TryParsePowerPlanTarget("HKLM", UsbSelectiveSuspendPolicy.PowerPlanSettingPath(scheme), name, out var parsed, out var source));
        Assert.Equal(scheme, parsed); Assert.Equal(ac, source);
        Assert.False(UsbSelectiveSuspendPolicy.TryParsePowerPlanTarget("HKLM", UsbSelectiveSuspendPolicy.PowerPlanSettingPath(scheme).Replace(scheme.ToString(), "invalid"), name, out _, out _));
    }

    public static IEnumerable<object[]> ExistingBackupFiles()
    {
        string? directory = Environment.GetEnvironmentVariable("DEVICE_TWEAKER_QA_BACKUPS");
        if (directory is not null && Directory.Exists(directory))
            foreach (var path in Directory.GetFiles(directory, "DeviceTweakerBackup_*.json")) yield return [path];
        else yield return ["synthetic"];
    }
    [Theory] [MemberData(nameof(ExistingBackupFiles))]
    public void VersionOneBackupUsesProductionRestoreTransaction(string path)
    {
        DeviceTweakerBackup b = path == "synthetic" ? Document(Value("ReservedCpuSets", "AQ=="))
            : JsonSerializer.Deserialize<DeviceTweakerBackup>(File.ReadAllText(path))!;
        BackupFormat.ValidateDocument(b, _ => true); Memory m = new();
        foreach (var v in b.RegistryValues) { var before = new RegistryValueBackup { Hive = v.Hive, Path = v.Path, Name = v.Name, Exists = false }; m.Seed(before); }
        m.Service().Restore(b);
        foreach (var v in b.RegistryValues.Where(v => v.Kind != "ReadError")) { var restored = m.Read(v); Assert.Equal(v.Exists, restored.Exists); Assert.Equal(v.Data, restored.Data); }
        if (b.StartupProfile is not null) { Assert.Equal(b.StartupProfile.Exists, m.Script.Exists); Assert.Equal(b.StartupProfile.Text, m.Script.Text); }
    }
}
