using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.Json;
using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public sealed class StartupProfileRestoreCompatibilityTests
{
    private static readonly string ProfilePath = Path.Combine(MainForm.GetStartupRoot(), "profile.json");
    private const string DriverPath = @"C:\Windows\DTIMOD.sys";
    private static MainForm Form() => (MainForm)RuntimeHelpers.GetUninitializedObject(typeof(MainForm));
    private static void Normalize(DeviceTweakerBackup backup)
        => typeof(MainForm).GetMethod("NormalizeStartupProfileBackup", BindingFlags.NonPublic | BindingFlags.Instance)!.Invoke(Form(), [backup]);
    private static FileBackup PrepareRollback(FileBackup profile)
        => (FileBackup)typeof(MainForm).GetMethod("PrepareRollbackStartupProfile", BindingFlags.NonPublic | BindingFlags.Instance)!.Invoke(Form(), [profile])!;
    private static FileBackup Copy(FileBackup source) => new() { Path = source.Path, Exists = source.Exists, Text = source.Text };
    private static FileBackup File(MainForm.ImodConfig config)
        => new() { Path = ProfilePath, Exists = true, Text = MainForm.BuildStartupProfile(config, DriverPath).ToJson() };
    private static MainForm.ImodConfig OldConfig(string id)
    {
        bool oldGigabit = id == "3000";
        MainForm.ImodConfig config = new();
        config.NicItrEntries.Add(new MainForm.NicItrConfigEntry
        {
            Hwid = "VEN_10EC&DEV_" + id, BaseOffset = oldGigabit ? 0xE2u : 0xA00u,
            Stride = oldGigabit ? 0u : 8u, Queues = oldGigabit ? 1 : 4,
            Width = oldGigabit ? 16 : 32, Mask = oldGigabit ? 0xFFFFUL : 0x7F7F7F7FUL, Values = [0x10],
        });
        return config;
    }
    // Models installation's real pure transformations. It does not stage files, register tasks or access hardware.
    private static FileBackup Install(FileBackup target)
    {
        if (!target.Exists) return Copy(target);
        StartupProfile profile = StartupProfile.FromJson(target.Text!);
        MainForm.ImodConfig config = MainForm.ConvertStartupProfile(profile);
        MainForm.RemoveStaleNicItrEntries(config);
        return new() { Path = target.Path, Exists = true, Text = MainForm.BuildStartupProfile(config, profile.Driver.Path!).ToJson() };
    }

    [Theory]
    [InlineData("3000")]
    [InlineData("8162")]
    public void OldNicProfileRestoresWithoutFalseReadbackFailure(string id)
    {
        FileBackup original = File(OldConfig(id));
        string originalText = original.Text!;
        DeviceTweakerBackup backup = new() { StartupProfile = Copy(original) };
        Normalize(backup);
        FileBackup stored = File(new());
        BackupRestoreService service = new(_ => { }, _ => throw new InvalidOperationException(),
            _ => throw new InvalidOperationException(), () => Copy(stored), target => stored = Install(target), _ => { });

        service.Restore(backup);

        Assert.Empty(StartupProfile.FromJson(stored.Text!).NicItr);
        Assert.Equal(originalText, original.Text); // Reading a backup must not alter its source.
        Assert.Contains("VEN_10EC&DEV_" + id, Assert.Single(backup.SkippedNicItrEntries));
        OperationReport report = MainForm.CreateBackupRestoreReport(backup);
        Assert.True(report.Succeeded);
        OperationIssue warning = Assert.Single(report.Issues.Where(issue => issue.Severity == OperationIssueSeverity.Warning));
        Assert.Contains("Save NIC ITR again", warning.UserMessage);
        Assert.DoesNotContain("SkippedNicItrEntries", JsonSerializer.Serialize(backup));
    }

    [Fact]
    public void MixedProfilePreservesUsbAndCurrentNicValues()
    {
        MainForm.ImodConfig config = OldConfig("3000");
        config.Overrides.Add(new MainForm.ImodConfigEntry { Hwid = "VEN_8086&DEV_7AE0", Enabled = true, Interval = 0 });
        NicItrProfile current = NicItrService.TryGetNicItrProfile("VEN_10EC&DEV_8125")!;
        config.NicItrEntries.Add(new MainForm.NicItrConfigEntry
        {
            Hwid = "VEN_10EC&DEV_8125", FamilyName = current.FamilyName, BaseOffset = current.BaseOffset,
            Stride = current.Stride, Queues = current.MaxQueues, Width = current.ReadWidth,
            Mask = current.ReadMask, OrBits = current.WriteOrBits, Values = [0x20],
        });
        DeviceTweakerBackup backup = new() { StartupProfile = File(config) };

        Normalize(backup);

        MainForm.ImodConfig restored = MainForm.ConvertStartupProfile(StartupProfile.FromJson(backup.StartupProfile!.Text!));
        Assert.Equal(0u, Assert.Single(restored.Overrides).Interval);
        Assert.Equal("VEN_10EC&DEV_8125", Assert.Single(restored.NicItrEntries).Hwid);
        Assert.Equal(0x20UL, restored.NicItrEntries[0].Values.Single());
        Assert.Equal(backup.StartupProfile.Text, Install(backup.StartupProfile).Text);
        Assert.Single(backup.SkippedNicItrEntries);
    }

    [Fact]
    public void AbsentAndUnspecifiedProfilesStayDistinct()
    {
        DeviceTweakerBackup unspecified = new();
        DeviceTweakerBackup absent = new() { StartupProfile = new() { Path = ProfilePath, Exists = false } };
        Normalize(unspecified);
        Normalize(absent);
        Assert.Null(unspecified.StartupProfile);
        Assert.False(absent.StartupProfile!.Exists);
        Assert.Empty(absent.SkippedNicItrEntries);
        Assert.Empty(MainForm.CreateBackupRestoreReport(absent).Warnings);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FailedInstallRollsBackToPreparedProfileAndReportsRealRollbackFailure(bool denyRollback)
    {
        DeviceTweakerBackup backup = new() { StartupProfile = File(new()) };
        Normalize(backup);
        FileBackup stored = File(OldConfig("3000"));
        string expectedRollback = PrepareRollback(stored).Text!;
        int attempts = 0;
        List<string> log = [];
        BackupRestoreService service = new(_ => { }, _ => throw new InvalidOperationException(),
            _ => throw new InvalidOperationException(), () => Copy(stored), target =>
            {
                attempts++;
                if (attempts == 2 && denyRollback) throw new IOException("rollback write denied");
                stored = Install(target);
                if (attempts == 1) throw new IOException("install failed after partial replacement");
            }, log.Add, PrepareRollback);

        Exception? error = Record.Exception(() => service.Restore(backup));

        Assert.NotNull(error);
        Assert.Equal(2, attempts);
        if (denyRollback)
        {
            Assert.Contains("rollback was incomplete", error.Message);
            Assert.Contains("rollback write denied", error.Message);
        }
        else
        {
            Assert.IsType<IOException>(error);
            Assert.Equal("install failed after partial replacement", error.Message);
            Assert.Equal(expectedRollback, stored.Text);
            Assert.Contains("BACKUP.RESTORE.ROLLBACK: completed", log);
        }
    }

    [Fact]
    public void NewWarningsHaveRussianTranslations()
    {
        Assert.False(MainForm.TryValidateMsiDisable(new PciInterruptCapabilities(PciInterruptSupport.MsiX, 8), out string english));
        var translate = typeof(UiLanguage).GetMethod("TranslateRussian", BindingFlags.NonPublic | BindingFlags.Static)!;
        string russian = (string)translate.Invoke(null, [english])!;
        Assert.Contains("Настройка MSI не изменена", russian);
        string skipped = (string)translate.Invoke(null, ["An outdated NIC ITR startup setting was skipped. Save NIC ITR again for this adapter."])!;
        Assert.Contains("Сохраните NIC ITR", skipped);
    }
}
