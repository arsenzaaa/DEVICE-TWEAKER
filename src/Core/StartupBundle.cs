using System.Security.AccessControl;
using System.Security.Principal;
using System.Text;

namespace DeviceTweakerCS;

// Locations of the saved USB IMOD / NIC ITR startup profile and of the layouts
// used before it: %APPDATA% started from HKCU\...\Run, and the user Startup folder.
internal sealed record StartupBundleLayout(string ParentDirectory, string LegacyAppDataDirectory, string LegacyStartupFolder)
{
    internal const string ScriptFileName = "ApplyIMOD.ps1";
    internal const string RuntimeFileName = "DeviceTweakerImodRuntime.cs";
    internal const string ProfileFileName = "profile.json";

    public string Root => Path.Combine(ParentDirectory, "Startup");
    public string ScriptPath => Path.Combine(Root, ScriptFileName);
    public string RuntimePath => Path.Combine(Root, RuntimeFileName);
    public string ProfilePath => Path.Combine(Root, ProfileFileName);
    public string LoaderDirectory => Path.Combine(Root, "Loader");
    public string LegacyAppDataScriptPath => Path.Combine(LegacyAppDataDirectory, ScriptFileName);
    public string LegacyStartupScriptPath => Path.Combine(LegacyStartupFolder, ScriptFileName);

    internal static StartupBundleLayout ForCurrentUser()
        => new(CurrentParentDirectory(), CurrentLegacyAppDataDirectory(), CurrentLegacyStartupFolder());

    internal static string CurrentParentDirectory()
    {
        string programData = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);
        if (string.IsNullOrWhiteSpace(programData))
        {
            throw new InvalidOperationException("%ProgramData% is not available.");
        }

        return Path.Combine(programData, "DEVICE TWEAKER");
    }

    internal static string CurrentLegacyAppDataDirectory()
        => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "DEVICE TWEAKER", "IMOD");

    internal static string CurrentLegacyStartupFolder()
        => Environment.GetFolderPath(Environment.SpecialFolder.Startup);
}

// The task that runs the startup script as SYSTEM.
internal interface IStartupTaskRegistry
{
    string? ReadXml();
    bool Register(string scriptPath, bool enabled, Action<string> log, out string? error);
    bool Restore(string scriptPath, string? previousXml, Action<string> log, out string? error);
    bool Delete(Action<string> log, out string? error);
}

internal sealed class ScheduledStartupTask : IStartupTaskRegistry
{
    public string? ReadXml() => StartupTask.ReadXml();
    public bool Register(string scriptPath, bool enabled, Action<string> log, out string? error) => StartupTask.Register(scriptPath, enabled, log, out error);
    public bool Restore(string scriptPath, string? previousXml, Action<string> log, out string? error)
        => StartupTask.Restore(scriptPath, previousXml, log, out error);
    public bool Delete(Action<string> log, out string? error) => StartupTask.Delete(log, out error);
}

// Files shipped inside the EXE: the script, its shared C# runtime and the loader pair.
internal interface IStartupPayload
{
    string Script { get; }
    string Runtime { get; }
    bool TryWriteLoader(string directory, out string? error);
    bool IsLoaderCurrent(string directory);
}

// Installs, checks and removes the startup bundle. The caller authorizes the write,
// holds the driver load gate and builds the profile; this class owns the files,
// folder permissions, the scheduled task and cleanup of older layouts.
internal sealed class StartupBundle(
    StartupBundleLayout layout,
    IStartupPayload payload,
    IStartupTaskRegistry tasks,
    Action<string> secureFolder,
    Action<OperationReport?> removeLegacyRunEntry,
    Action<string> log)
{
    public StartupBundleLayout Layout => layout;

    public bool NeedsAutomaticRepair(bool staleNic)
    {
        string? xml = tasks.ReadXml();
        return xml is not null && (staleNic || !PayloadIsCurrent() || !TaskIsCurrent(xml));
    }

    public bool Install(string profileJson, string? driverPath = null, Action? stageDriver = null)
    {
        QaRun.DemandArtifactPath(layout.Root);
        secureFolder(layout.ParentDirectory);
        secureFolder(layout.Root);
        string? previousTask = tasks.ReadXml();
        bool enabled = !StartupTask.IsExplicitlyDisabled(previousTask);
        bool updateTask = !TaskIsCurrent(previousTask);
        bool updatePayload = !PayloadIsCurrent();
        List<(string Path, byte[]? Content)> previous = [];
        bool taskWriteStarted = false;
        string? prepared = null;
        try
        {
            if (updatePayload)
            {
                try { secureFolder(layout.LoaderDirectory); }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    throw new IOException("The startup loader could not be prepared: " + ex.Message, ex);
                }
                prepared = Path.Combine(layout.Root, $".prepare-{Guid.NewGuid():N}");
                Directory.CreateDirectory(prepared);
                if (!payload.TryWriteLoader(prepared, out string? stagingError))
                    throw new IOException("The startup loader could not be prepared: " + stagingError);

                string script = payload.Script;
                string runtime = payload.Runtime;
                // Write the script with a BOM for Windows PowerShell 5.1.
                ReplaceFile(layout.ScriptPath, () => AtomicFile.WriteAllText(layout.ScriptPath, script, new UTF8Encoding(true)));
                ReplaceFile(layout.RuntimePath, () => AtomicFile.WriteAllText(layout.RuntimePath, runtime, new UTF8Encoding(false)));
                foreach (string file in new[] { ImodDriverNative.ImodKduFileName, ImodDriverNative.ImodKduDatabaseFileName })
                {
                    string target = Path.Combine(layout.LoaderDirectory, file);
                    byte[] bytes = File.ReadAllBytes(Path.Combine(prepared, file));
                    ReplaceFile(target, () => AtomicFile.WriteAllBytes(target, bytes));
                }
            }

            if (driverPath is not null && stageDriver is not null)
                ReplaceFile(driverPath, stageDriver);

            if (updateTask)
            {
                taskWriteStarted = true;
                if (!tasks.Register(layout.ScriptPath, enabled, log, out string? taskError))
                    throw new InvalidOperationException("The startup task could not be registered: " + taskError);
            }

            AtomicFile.WriteAllText(layout.ProfilePath, profileJson, new UTF8Encoding(false));
        }
        catch (Exception failure)
        {
            List<string> errors = [];
            for (int i = previous.Count - 1; i >= 0; i--)
            {
                (string path, byte[]? bytes) = previous[i];
                try
                {
                    byte[]? current = ReadOptionalFile(path);
                    if (bytes is null)
                    {
                        if (current is not null) File.Delete(path);
                    }
                    else if (!bytes.AsSpan().SequenceEqual(current))
                        AtomicFile.WriteAllBytes(path, bytes);
                }
                catch (Exception ex) { errors.Add($"{path}: {ex.Message}"); }
            }
            if (taskWriteStarted)
            {
                try
                {
                    if (!tasks.Restore(layout.ScriptPath, previousTask, log, out string? error))
                        errors.Add(error ?? "the previous task could not be restored");
                }
                catch (Exception ex) { errors.Add(ex.Message); }
            }
            if (errors.Count > 0)
                throw new InvalidOperationException("Startup update failed and rollback was incomplete: "
                    + string.Join(" | ", errors), failure);
            throw;
        }
        finally
        {
            if (prepared is not null) DeleteDirectory(prepared, "IMOD.STARTUP.STAGING", report: null);
        }

        RemoveLegacy(report: null);
        if (!enabled) log("IMOD.STARTUP.TASK: saved profile; task remains disabled");
        return enabled;

        void ReplaceFile(string path, Action write)
        {
            QaRun.DemandArtifactPath(path);
            previous.Add((path, ReadOptionalFile(path)));
            write();
        }
    }

    private bool PayloadIsCurrent() => FileTextEquals(layout.ScriptPath, payload.Script)
        && FileTextEquals(layout.RuntimePath, payload.Runtime)
        && payload.IsLoaderCurrent(layout.LoaderDirectory);

    private bool TaskIsCurrent(string? xml)
        => xml is not null && StartupTask.VerifyRegisteredXml(xml, layout.ScriptPath, !StartupTask.IsExplicitlyDisabled(xml)) is null;

    private static byte[]? ReadOptionalFile(string path)
    {
        try { return File.ReadAllBytes(path); }
        catch (FileNotFoundException) { return null; }
        catch (DirectoryNotFoundException) { return null; }
    }

    public void Remove(OperationReport? report, Func<bool>? beforeFilesRemoved = null)
    {
        QaRun.DemandArtifactPath(layout.Root);
        if (!tasks.Delete(log, out string? taskError))
        {
            string error = taskError ?? "startup task was not deleted";
            log($"IMOD.STARTUP.TASK: {error}");
            if (report is null) throw new InvalidOperationException(error);
            report.AddError("IMOD.STARTUP", error);
            return;
        }

        if (beforeFilesRemoved is not null && !beforeFilesRemoved())
        {
            const string error = "Could not restore network moderation. The profile was kept so you can retry.";
            log($"IMOD.STARTUP.REMOVE: {error}");
            if (report is null) throw new InvalidOperationException(error);
            report.AddError("IMOD.STARTUP", error);
            return;
        }
        DeleteFile(layout.ProfilePath, "IMOD.STARTUP", report);
        DeleteFile(layout.ScriptPath, "IMOD.STARTUP", report);
        DeleteFile(layout.RuntimePath, "IMOD.STARTUP", report);
        DeleteDirectory(layout.LoaderDirectory, "IMOD.STARTUP", report);
        RemoveLegacy(report);
    }

    public void RemoveLegacy(OperationReport? report)
    {
        removeLegacyRunEntry(report);
        DeleteFile(layout.LegacyAppDataScriptPath, "IMOD.STARTUP.LEGACY", report);
        DeleteFile(layout.LegacyStartupScriptPath, "IMOD.STARTUP.LEGACY", report);
        DeleteDirectory(Path.Combine(layout.LegacyAppDataDirectory, "Loader"), "IMOD.STARTUP.LEGACY", report);
        if (Directory.Exists(layout.LegacyAppDataDirectory))
        {
            foreach (string previous in Directory.GetFiles(layout.LegacyAppDataDirectory, StartupBundleLayout.ScriptFileName + ".pre-DTIMOD-*.bak"))
            {
                DeleteFile(previous, "IMOD.STARTUP.LEGACY", report);
            }
        }
    }

    // A non-admin user can pre-create folders in %ProgramData% or plant a junction.
    // The folder must be a real directory owned by Administrators that only
    // administrators and SYSTEM can change; users keep read access to review the script.
    internal static void SecureFolder(string path)
    {
        QaRun.DemandArtifactPath(path);
        DirectoryInfo directory = new(path);
        if (directory.Exists && directory.Attributes.HasFlag(FileAttributes.ReparsePoint))
        {
            directory.Delete();
        }

        directory.Refresh();
        if (!directory.Exists)
        {
            directory.Create();
        }

        directory.Refresh();
        if (directory.Attributes.HasFlag(FileAttributes.ReparsePoint))
        {
            throw new IOException($"Refusing to use a reparse point as the startup folder: {path}");
        }

        DirectorySecurity security = new();
        security.SetOwner(new SecurityIdentifier(WellKnownSidType.BuiltinAdministratorsSid, null));
        security.SetAccessRuleProtection(isProtected: true, preserveInheritance: false);
        void Allow(WellKnownSidType sid, FileSystemRights rights)
            => security.AddAccessRule(new FileSystemAccessRule(
                new SecurityIdentifier(sid, null),
                rights,
                InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit,
                PropagationFlags.None,
                AccessControlType.Allow));
        Allow(WellKnownSidType.LocalSystemSid, FileSystemRights.FullControl);
        Allow(WellKnownSidType.BuiltinAdministratorsSid, FileSystemRights.FullControl);
        Allow(WellKnownSidType.BuiltinUsersSid, FileSystemRights.ReadAndExecute);
        try
        {
            directory.SetAccessControl(security);
        }
        catch (UnauthorizedAccessException ex)
        {
            throw new UnauthorizedAccessException(
                $"{path} belongs to another account and cannot be secured. Delete the folder and save again.", ex);
        }
    }

    private static bool FileTextEquals(string path, string expected)
        => File.Exists(path) && string.Equals(File.ReadAllText(path, Encoding.UTF8), expected, StringComparison.Ordinal);

    private void DeleteFile(string path, string label, OperationReport? report)
    {
        QaRun.DemandArtifactPath(path);
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
                log($"{label}: deleted {path}");
            }
        }
        catch (Exception ex)
        {
            log($"{label}: failed to delete {path}: {ex.Message}");
            report?.AddError(label, $"failed to delete {path}: {ex.Message}");
        }
    }

    private void DeleteDirectory(string path, string label, OperationReport? report)
    {
        QaRun.DemandArtifactPath(path);
        try
        {
            if (Directory.Exists(path))
            {
                Directory.Delete(path, recursive: true);
                log($"{label}: deleted {path}");
            }
        }
        catch (Exception ex)
        {
            log($"{label}: failed to delete {path}: {ex.Message}");
            report?.AddError(label, $"failed to delete {path}: {ex.Message}");
        }
    }
}
