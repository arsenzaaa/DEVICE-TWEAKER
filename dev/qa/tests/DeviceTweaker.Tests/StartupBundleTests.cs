using System.Text;
using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

// The bundle runs against a layout inside the QA run directory, the only place a
// read-only test process may write. Folder ACLs, the task and HKCU are fakes.
public sealed class StartupBundleTests : IDisposable
{
    private readonly string _root = Path.Combine(QaRun.OutputDirectory, "startup-bundle-" + Guid.NewGuid().ToString("N"));
    private readonly FakePayload _payload = new();
    private readonly FakeTasks _tasks = new();
    private readonly List<string> _secured = [];
    private readonly List<string> _log = [];
    private int _legacyRunEntryRemovals;

    private StartupBundleLayout Layout => new(
        Path.Combine(_root, "ProgramData", "DEVICE TWEAKER"),
        Path.Combine(_root, "AppData", "DEVICE TWEAKER", "IMOD"),
        Path.Combine(_root, "StartMenu", "Startup"));

    private StartupBundle Bundle(Action<string>? secureFolder = null) => new(
        Layout, _payload, _tasks, secureFolder ?? _secured.Add, _ => _legacyRunEntryRemovals++, _log.Add);

    public void Dispose()
    {
        if (Directory.Exists(_root)) Directory.Delete(_root, recursive: true);
    }

    [Fact]
    public void LayoutKeepsTheShippedFileNames()
    {
        StartupBundleLayout layout = new(@"C:\ProgramData\DEVICE TWEAKER", @"C:\AppData\IMOD", @"C:\Startup");

        Assert.Equal(@"C:\ProgramData\DEVICE TWEAKER\Startup", layout.Root);
        Assert.Equal(@"C:\ProgramData\DEVICE TWEAKER\Startup\ApplyIMOD.ps1", layout.ScriptPath);
        Assert.Equal(@"C:\ProgramData\DEVICE TWEAKER\Startup\DeviceTweakerImodRuntime.cs", layout.RuntimePath);
        Assert.Equal(@"C:\ProgramData\DEVICE TWEAKER\Startup\profile.json", layout.ProfilePath);
        Assert.Equal(@"C:\ProgramData\DEVICE TWEAKER\Startup\Loader", layout.LoaderDirectory);
        Assert.Equal(@"C:\AppData\IMOD\ApplyIMOD.ps1", layout.LegacyAppDataScriptPath);
        Assert.Equal(@"C:\Startup\ApplyIMOD.ps1", layout.LegacyStartupScriptPath);
    }

    [Fact]
    public void CurrentUserLayoutIsUnderProgramData()
    {
        string programData = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);

        Assert.Equal(Path.Combine(programData, "DEVICE TWEAKER", "Startup"), StartupBundleLayout.ForCurrentUser().Root);
        Assert.Equal(MainForm.GetStartupRoot(), StartupBundleLayout.ForCurrentUser().Root);
    }

    [Fact]
    public void InstallWritesPayloadProfileAndRegistersTask()
    {
        Bundle().Install("{\"schemaVersion\":1}");
        StartupBundleLayout layout = Layout;

        Assert.Equal(new[] { layout.ParentDirectory, layout.Root, layout.LoaderDirectory }, _secured);
        byte[] script = File.ReadAllBytes(layout.ScriptPath);
        Assert.Equal(new byte[] { 0xEF, 0xBB, 0xBF }, script.Take(3));
        Assert.Equal(_payload.Script, File.ReadAllText(layout.ScriptPath, Encoding.UTF8));
        byte[] runtime = File.ReadAllBytes(layout.RuntimePath);
        Assert.NotEqual(0xEF, runtime[0]);
        Assert.Equal(_payload.Runtime, Encoding.UTF8.GetString(runtime));
        Assert.Equal("{\"schemaVersion\":1}", File.ReadAllText(layout.ProfilePath));
        Assert.StartsWith(Path.Combine(layout.Root, ".prepare-"), Assert.Single(_payload.LoaderWrites));
        Assert.Empty(Directory.GetDirectories(layout.Root, ".prepare-*"));
        Assert.Equal(new[] { layout.ScriptPath }, _tasks.Registered);
        Assert.Equal(1, _legacyRunEntryRemovals);
        Assert.False(Bundle().NeedsAutomaticRepair(staleNic: false));
    }

    [Fact]
    public void InstallRemovesFilesOfOlderLayouts()
    {
        StartupBundleLayout layout = Layout;
        Directory.CreateDirectory(Path.Combine(layout.LegacyAppDataDirectory, "Loader"));
        Directory.CreateDirectory(layout.LegacyStartupFolder);
        File.WriteAllText(layout.LegacyAppDataScriptPath, "old");
        File.WriteAllText(layout.LegacyAppDataScriptPath + ".pre-DTIMOD-1.bak", "old");
        File.WriteAllText(layout.LegacyStartupScriptPath, "old");
        File.WriteAllText(Path.Combine(layout.LegacyAppDataDirectory, "keep.txt"), "user file");

        Bundle().Install("{}");

        Assert.False(File.Exists(layout.LegacyAppDataScriptPath));
        Assert.False(File.Exists(layout.LegacyAppDataScriptPath + ".pre-DTIMOD-1.bak"));
        Assert.False(File.Exists(layout.LegacyStartupScriptPath));
        Assert.False(Directory.Exists(Path.Combine(layout.LegacyAppDataDirectory, "Loader")));
        Assert.True(File.Exists(Path.Combine(layout.LegacyAppDataDirectory, "keep.txt")));
        Assert.Contains(_log, line => line.StartsWith("IMOD.STARTUP.LEGACY: deleted ", StringComparison.Ordinal));
    }

    [Fact]
    public void LoaderFailureStopsBeforeTheProfileAndTask()
    {
        _payload.LoaderError = "staged payload hash mismatch: kdu.exe";

        IOException error = Assert.Throws<IOException>(() => Bundle().Install("{}"));

        Assert.Equal("The startup loader could not be prepared: staged payload hash mismatch: kdu.exe", error.Message);
        Assert.False(File.Exists(Layout.ProfilePath));
        Assert.Empty(_tasks.Registered);
        Assert.Equal(0, _legacyRunEntryRemovals);
    }

    [Fact]
    public void UnsecurableLoaderFolderIsReportedAsALoaderFailure()
    {
        string loader = Layout.LoaderDirectory;
        void Secure(string path)
        {
            if (path == loader) throw new UnauthorizedAccessException("owned by another account");
            _secured.Add(path);
        }

        IOException error = Assert.Throws<IOException>(() => Bundle(Secure).Install("{}"));

        Assert.Equal("The startup loader could not be prepared: owned by another account", error.Message);
        Assert.Empty(_payload.LoaderWrites);
        Assert.Empty(_tasks.Registered);
    }

    [Fact]
    public void TaskFailureIsReportedAndLeavesLegacyFilesInPlace()
    {
        _tasks.RegisterError = "access denied";
        Directory.CreateDirectory(Layout.LegacyStartupFolder);
        File.WriteAllText(Layout.LegacyStartupScriptPath, "old");

        InvalidOperationException error = Assert.Throws<InvalidOperationException>(() => Bundle().Install("{}"));

        Assert.Equal("The startup task could not be registered: access denied", error.Message);
        Assert.False(File.Exists(Layout.ProfilePath));
        Assert.True(File.Exists(Layout.LegacyStartupScriptPath));
        Assert.Equal(0, _legacyRunEntryRemovals);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TaskRegistrationFailureDoesNotPublishNewSettings(bool existingProfile)
    {
        const string previous = "{\"usb\":\"previous\"}";
        if (existingProfile) Bundle().Install(previous);
        string? previousTask = existingProfile ? StartupTask.BuildXml(Layout.ScriptPath + ".old") : null;
        _tasks.CurrentXml = previousTask;
        string? previousRuntime = existingProfile ? File.ReadAllText(Layout.RuntimePath) : null;
        _payload.Runtime = "// changed payload";
        _payload.Loader = "changed loader";
        string driver = Path.Combine(_root, "DTIMOD.sys");
        Directory.CreateDirectory(_root);
        File.WriteAllText(driver, "previous driver");
        _tasks.RegisterError = "query failed after registration";

        Assert.Throws<InvalidOperationException>(() => Bundle().Install("{\"usb\":\"new\"}", driver,
            () => File.WriteAllText(driver, "replacement driver")));

        Assert.Equal(previousTask, _tasks.CurrentXml);
        Assert.Equal("previous driver", File.ReadAllText(driver));
        Assert.Equal(existingProfile, File.Exists(Layout.ProfilePath));
        Assert.Equal(existingProfile, File.Exists(Layout.RuntimePath));
        Assert.Equal(existingProfile, File.Exists(Path.Combine(Layout.LoaderDirectory, "drv64.dll")));
        if (existingProfile)
        {
            Assert.Equal(previous, File.ReadAllText(Layout.ProfilePath));
            Assert.Equal(previousRuntime, File.ReadAllText(Layout.RuntimePath));
            Assert.Equal("loader", File.ReadAllText(Path.Combine(Layout.LoaderDirectory, "kdu.exe")));
        }
        Assert.Empty(Directory.GetDirectories(Layout.Root, ".prepare-*"));
    }

    [Fact]
    public void ProfileWriteFailurePreservesPreviousSettingsAndLegacyFiles()
    {
        Bundle().Install("previous settings");
        Directory.CreateDirectory(Layout.LegacyStartupFolder);
        File.WriteAllText(Layout.LegacyStartupScriptPath, "legacy settings");
        int removals = _legacyRunEntryRemovals;
        using FileStream lockedProfile = new(Layout.ProfilePath, FileMode.Open, FileAccess.Read, FileShare.Read);

        Exception? error = Record.Exception(() => Bundle().Install("new settings"));
        Assert.True(error is IOException or UnauthorizedAccessException, error?.ToString() ?? "Expected a file replacement error.");

        Assert.Equal("previous settings", File.ReadAllText(Layout.ProfilePath));
        Assert.True(File.Exists(Layout.LegacyStartupScriptPath));
        Assert.Equal(removals, _legacyRunEntryRemovals);
        Assert.Empty(Directory.GetFiles(Layout.Root, ".profile.json.*.tmp"));
    }

    [Fact]
    public void ProfileOnlySaveDoesNotRewriteCurrentPayloadOrTask()
    {
        Bundle().Install("previous settings");
        Bundle().Install("new settings");

        Assert.Equal("new settings", File.ReadAllText(Layout.ProfilePath));
        Assert.Single(_tasks.Registered);
        Assert.Single(_payload.LoaderWrites);
    }

    [Fact]
    public void RepairKeepsADisabledTaskDisabled()
    {
        Bundle().Install("previous settings");
        _tasks.CurrentXml = StartupTask.BuildXml(Layout.ScriptPath, enabled: false);
        Assert.False(Bundle().NeedsAutomaticRepair(staleNic: false));
        Assert.False(Bundle().Install("new settings"));
        Assert.Single(_tasks.Registered);

        _tasks.CurrentXml = _tasks.CurrentXml!.Replace("<LogonTrigger>", "<BootTrigger>", StringComparison.Ordinal)
            .Replace("</LogonTrigger>", "</BootTrigger>", StringComparison.Ordinal);
        Assert.True(Bundle().NeedsAutomaticRepair(staleNic: false));
        Assert.False(Bundle().Install("new settings"));
        Assert.Equal(2, _tasks.Registered.Count);
        Assert.True(StartupTask.IsExplicitlyDisabled(_tasks.CurrentXml));
        Assert.False(Bundle().NeedsAutomaticRepair(staleNic: false));
    }

    [Fact]
    public void FailedModerationRestoreLeavesTheProfileForRetryAndDoesNotRecreateTheTask()
    {
        Bundle().Install("saved settings");
        OperationReport failed = new();
        Bundle().Remove(failed, () => false);

        Assert.False(failed.Succeeded);
        Assert.Null(_tasks.ReadXml());
        Assert.Equal("saved settings", File.ReadAllText(Layout.ProfilePath));
        Assert.False(Bundle().NeedsAutomaticRepair(staleNic: true));

        OperationReport retried = new();
        Bundle().Remove(retried, () => true);
        Assert.True(retried.Succeeded);
        Assert.False(File.Exists(Layout.ProfilePath));
        Assert.Null(_tasks.ReadXml());
    }

    [Fact]
    public void AutomaticRepairChecksPayloadOnlyWhileATaskExists()
    {
        Bundle().Install("{}");
        Assert.False(Bundle().NeedsAutomaticRepair(staleNic: false));

        _payload.Runtime = "// newer runtime";
        Assert.True(Bundle().NeedsAutomaticRepair(staleNic: false));
        _payload.Runtime = FakePayload.InitialRuntime;

        _payload.LoaderCurrent = false;
        Assert.True(Bundle().NeedsAutomaticRepair(staleNic: false));
        _payload.LoaderCurrent = true;

        _tasks.TaskExists = false;
        Assert.False(Bundle().NeedsAutomaticRepair(staleNic: true));
        _tasks.TaskExists = true;

        File.Delete(Layout.ScriptPath);
        Assert.True(Bundle().NeedsAutomaticRepair(staleNic: false));
    }

    [Fact]
    public void UnreadablePayloadIsReportedToTheCaller()
    {
        Bundle().Install("{}");
        _payload.ScriptError = new InvalidDataException("Missing embedded startup resource");

        Assert.Throws<InvalidDataException>(() => Bundle().NeedsAutomaticRepair(staleNic: false));
    }

    [Fact]
    public void TaskDeletionFailureLeavesTheProfileAndPayloadInPlace()
    {
        Bundle().Install("{}");
        _tasks.DeleteError = "task is in use";
        OperationReport report = new();

        Bundle().Remove(report, () => throw new InvalidOperationException("must not modify saved settings"));

        StartupBundleLayout layout = Layout;
        Assert.True(File.Exists(layout.ProfilePath));
        Assert.True(File.Exists(layout.ScriptPath));
        Assert.True(File.Exists(layout.RuntimePath));
        Assert.True(Directory.Exists(layout.LoaderDirectory));
        Assert.Contains("IMOD.STARTUP.TASK: task is in use", _log);
        OperationIssue issue = Assert.Single(report.Issues);
        Assert.Equal(("IMOD.STARTUP", "task is in use"), (issue.Component, issue.TechnicalDetails));
        Assert.False(report.Succeeded);
        Assert.Equal(1, _legacyRunEntryRemovals);
    }

    [Fact]
    public void RemoveWithoutABundleChangesNothing()
    {
        OperationReport report = new();

        Bundle().Remove(report);

        Assert.True(report.Succeeded);
        Assert.False(Directory.Exists(_root));
        Assert.DoesNotContain(_log, line => line.Contains("deleted", StringComparison.Ordinal));
    }

    [Fact]
    public void RealFolderSecuringIsBlockedOutsideTheQaRun()
    {
        string outside = Path.Combine(Path.GetTempPath(), "dt-startup-outside-" + Guid.NewGuid().ToString("N"));

        Assert.Throws<QaWriteBlockedException>(() => StartupBundle.SecureFolder(outside));
        Assert.False(Directory.Exists(outside));
    }

    [Fact]
    public void InstallIntoTheRealLayoutIsBlockedInQa()
    {
        StartupBundle bundle = new(StartupBundleLayout.ForCurrentUser(), _payload, _tasks,
            StartupBundle.SecureFolder, _ => _legacyRunEntryRemovals++, _log.Add);

        Assert.Throws<QaWriteBlockedException>(() => bundle.Install("{}"));
        Assert.Throws<QaWriteBlockedException>(() => bundle.Remove(null));
        Assert.Empty(_tasks.Registered);
    }

    private sealed class FakePayload : IStartupPayload
    {
        internal const string InitialRuntime = "// runtime\r\npublic static class R { }\r\n";
        public string Script
        {
            get => ScriptError is null ? "param()\r\n# script\r\n" : throw ScriptError;
        }
        public string Runtime { get; set; } = InitialRuntime;
        public Exception? ScriptError { get; set; }
        public string? LoaderError { get; set; }
        public string Loader { get; set; } = "loader";
        public bool LoaderCurrent { get; set; } = true;
        public List<string> LoaderWrites { get; } = [];

        public bool TryWriteLoader(string directory, out string? error)
        {
            error = LoaderError;
            if (error is not null) return false;
            LoaderWrites.Add(directory);
            Directory.CreateDirectory(directory);
            File.WriteAllText(Path.Combine(directory, "kdu.exe"), Loader);
            File.WriteAllText(Path.Combine(directory, "drv64.dll"), Loader);
            return true;
        }

        public bool IsLoaderCurrent(string directory)
        {
            string file = Path.Combine(directory, "kdu.exe");
            return LoaderCurrent && File.Exists(file) && File.ReadAllText(file) == Loader;
        }
    }

    private sealed class FakeTasks : IStartupTaskRegistry
    {
        public bool TaskExists { get; set; } = true;
        public string? RegisterError { get; set; }
        public string? DeleteError { get; set; }
        public List<string> Registered { get; } = [];

        public string? CurrentXml { get; set; }
        public string? ReadXml() => TaskExists ? CurrentXml : null;

        public bool Register(string scriptPath, bool enabled, Action<string> log, out string? error)
        {
            error = RegisterError;
            CurrentXml = StartupTask.BuildXml(scriptPath, enabled);
            Registered.Add(scriptPath);
            return error is null;
        }

        public bool Restore(string scriptPath, string? previousXml, Action<string> log, out string? error)
        {
            CurrentXml = previousXml;
            error = null;
            return true;
        }

        public bool Delete(Action<string> log, out string? error)
        {
            error = DeleteError;
            if (error is not null) return false;
            CurrentXml = null;
            return true;
        }
    }
}
