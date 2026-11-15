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
    public void InstallWritesEveryFileThenRegistersTheTask()
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
        Assert.Equal(new[] { layout.LoaderDirectory }, _payload.LoaderWrites);
        Assert.Equal(new[] { layout.ScriptPath }, _tasks.Registered);
        Assert.Equal(1, _legacyRunEntryRemovals);
        Assert.True(Bundle().IsCurrent());
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
        Assert.True(File.Exists(Layout.ProfilePath));
        Assert.True(File.Exists(Layout.LegacyStartupScriptPath));
        Assert.Equal(0, _legacyRunEntryRemovals);
    }

    [Fact]
    public void BundleIsOutdatedWhenAnyPartDiffers()
    {
        Bundle().Install("{}");
        Assert.True(Bundle().IsCurrent());

        _payload.Runtime = "// newer runtime";
        Assert.False(Bundle().IsCurrent());
        _payload.Runtime = FakePayload.InitialRuntime;

        _payload.LoaderCurrent = false;
        Assert.False(Bundle().IsCurrent());
        _payload.LoaderCurrent = true;

        _tasks.TaskExists = false;
        Assert.False(Bundle().IsCurrent());
        _tasks.TaskExists = true;

        File.Delete(Layout.ScriptPath);
        Assert.False(Bundle().IsCurrent());
    }

    [Fact]
    public void UnreadablePayloadMakesTheBundleOutdated()
    {
        Bundle().Install("{}");
        _payload.ScriptError = new InvalidDataException("Missing embedded startup resource");

        Assert.False(Bundle().IsCurrent());
    }

    [Fact]
    public void RemoveDeletesTheBundleAndReportsATaskFailure()
    {
        Bundle().Install("{}");
        _tasks.DeleteError = "task is in use";
        OperationReport report = new();

        Bundle().Remove(report);

        StartupBundleLayout layout = Layout;
        Assert.False(File.Exists(layout.ProfilePath));
        Assert.False(File.Exists(layout.ScriptPath));
        Assert.False(File.Exists(layout.RuntimePath));
        Assert.False(Directory.Exists(layout.LoaderDirectory));
        Assert.Contains("IMOD.STARTUP.TASK: task is in use", _log);
        OperationIssue issue = Assert.Single(report.Issues);
        Assert.Equal(("IMOD.STARTUP", "task is in use"), (issue.Component, issue.TechnicalDetails));
        Assert.False(report.Succeeded);
        Assert.Equal(2, _legacyRunEntryRemovals);
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
        public bool LoaderCurrent { get; set; } = true;
        public List<string> LoaderWrites { get; } = [];

        public bool TryWriteLoader(string directory, out string? error)
        {
            error = LoaderError;
            if (error is not null) return false;
            LoaderWrites.Add(directory);
            Directory.CreateDirectory(directory);
            File.WriteAllText(Path.Combine(directory, "kdu.exe"), "loader");
            return true;
        }

        public bool IsLoaderCurrent(string directory) => LoaderCurrent && File.Exists(Path.Combine(directory, "kdu.exe"));
    }

    private sealed class FakeTasks : IStartupTaskRegistry
    {
        public bool TaskExists { get; set; } = true;
        public string? RegisterError { get; set; }
        public string? DeleteError { get; set; }
        public List<string> Registered { get; } = [];

        public bool Exists() => TaskExists && Registered.Count > 0;

        public bool Register(string scriptPath, Action<string> log, out string? error)
        {
            error = RegisterError;
            if (error is not null) return false;
            Registered.Add(scriptPath);
            return true;
        }

        public bool Delete(Action<string> log, out string? error)
        {
            error = DeleteError;
            return error is null;
        }
    }
}
