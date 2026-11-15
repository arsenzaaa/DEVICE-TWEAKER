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
    bool Exists();
    bool Register(string scriptPath, Action<string> log, out string? error);
    bool Delete(Action<string> log, out string? error);
}

internal sealed class ScheduledStartupTask : IStartupTaskRegistry
{
    public bool Exists() => StartupTask.Exists();
    public bool Register(string scriptPath, Action<string> log, out string? error) => StartupTask.Register(scriptPath, log, out error);
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

    public bool IsCurrent()
    {
        try
        {
            return FileTextEquals(layout.ScriptPath, payload.Script)
                && FileTextEquals(layout.RuntimePath, payload.Runtime)
                && payload.IsLoaderCurrent(layout.LoaderDirectory)
                && tasks.Exists();
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or InvalidDataException)
        {
            return false;
        }
    }

    public void Install(string profileJson)
    {
        secureFolder(layout.ParentDirectory);
        secureFolder(layout.Root);
        // PowerShell 5.1 reads a script without BOM in the ANSI code page.
        AtomicFile.WriteAllText(layout.ScriptPath, payload.Script, new UTF8Encoding(true));
        AtomicFile.WriteAllText(layout.RuntimePath, payload.Runtime, new UTF8Encoding(false));

        string? loaderError;
        try
        {
            secureFolder(layout.LoaderDirectory);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            throw new IOException("The startup loader could not be prepared: " + ex.Message, ex);
        }

        if (!payload.TryWriteLoader(layout.LoaderDirectory, out loaderError))
        {
            throw new IOException("The startup loader could not be prepared: " + loaderError);
        }

        if (!tasks.Register(layout.ScriptPath, log, out string? taskError))
        {
            throw new InvalidOperationException("The startup task could not be registered: " + taskError);
        }

        // Publish new settings only after registration and verification succeed.
        // An existing task must retain its previous profile on registration failure.
        AtomicFile.WriteAllText(layout.ProfilePath, profileJson, new UTF8Encoding(false));
        RemoveLegacy(report: null);
    }

    public void Remove(OperationReport? report)
    {
        if (!tasks.Delete(log, out string? taskError))
        {
            log($"IMOD.STARTUP.TASK: {taskError}");
            report?.AddError("IMOD.STARTUP", taskError ?? "startup task was not deleted");
        }

        DeleteFile(layout.ProfilePath, "IMOD.STARTUP", report);
        DeleteFile(layout.ScriptPath, "IMOD.STARTUP", report);
        DeleteFile(layout.RuntimePath, "IMOD.STARTUP", report);
        DeleteDirectory(layout.LoaderDirectory, "IMOD.STARTUP");
        RemoveLegacy(report);
    }

    public void RemoveLegacy(OperationReport? report)
    {
        removeLegacyRunEntry(report);
        DeleteFile(layout.LegacyAppDataScriptPath, "IMOD.STARTUP.LEGACY", report);
        DeleteFile(layout.LegacyStartupScriptPath, "IMOD.STARTUP.LEGACY", report);
        DeleteDirectory(Path.Combine(layout.LegacyAppDataDirectory, "Loader"), "IMOD.STARTUP.LEGACY");
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

    private void DeleteDirectory(string path, string label)
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
        }
    }
}
