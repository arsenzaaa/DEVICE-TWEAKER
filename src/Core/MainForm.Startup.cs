using System.Security.AccessControl;
using System.Security.Principal;
using System.Text;
using System.Text.Json;

namespace DeviceTweakerCS;

// Saved USB IMOD / NIC ITR profile and the standalone script that applies it.
// The script, its C# runtime and the loader live in %ProgramData%\DEVICE TWEAKER\Startup,
// run from Task Scheduler as SYSTEM and keep working without DEVICE TWEAKER.exe.
public sealed partial class MainForm
{
    private const string StartupScriptResource = "DeviceTweakerCS.Startup.ApplyIMOD.ps1";
    private const string StartupRuntimeResource = "DeviceTweakerCS.Startup.DeviceTweakerImodRuntime.cs";
    private const string StartupRuntimeFileName = "DeviceTweakerImodRuntime.cs";
    private const string StartupProfileFileName = "profile.json";

    private static string GetStartupParentDirectory()
    {
        string programData = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);
        if (string.IsNullOrWhiteSpace(programData))
        {
            throw new InvalidOperationException("%ProgramData% is not available.");
        }

        return Path.Combine(programData, "DEVICE TWEAKER");
    }

    internal static string GetStartupRoot() => Path.Combine(GetStartupParentDirectory(), "Startup");

    private static string GetStartupScriptPath() => Path.Combine(GetStartupRoot(), ImodScriptFileName);

    private static string GetStartupRuntimePath() => Path.Combine(GetStartupRoot(), StartupRuntimeFileName);

    private static string GetStartupProfilePath() => Path.Combine(GetStartupRoot(), StartupProfileFileName);

    private static string GetImodStartupKduDirectory() => Path.Combine(GetStartupRoot(), "Loader");

    // Before this layout the script lived in %APPDATA% and started from HKCU\...\Run;
    // older versions used the Startup folder.
    private static string GetLegacyAppDataImodDirectory()
        => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "DEVICE TWEAKER", "IMOD");

    private static string GetLegacyAppDataImodScriptPath() => Path.Combine(GetLegacyAppDataImodDirectory(), ImodScriptFileName);

    private static string GetLegacyImodStartupPath()
        => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Startup), ImodScriptFileName);

    private static string ReadStartupResource(string name)
    {
        using Stream stream = typeof(MainForm).Assembly.GetManifestResourceStream(name)
            ?? throw new InvalidDataException("Missing embedded startup resource: " + name);
        using StreamReader reader = new(stream, new UTF8Encoding(false), detectEncodingFromByteOrderMarks: true);
        return reader.ReadToEnd();
    }

    // Reads the saved profile without migrating or repairing anything.
    private ImodConfig ReadSavedStartupConfig(out string? source)
    {
        string profilePath = GetStartupProfilePath();
        source = null;
        try
        {
            if (File.Exists(profilePath))
            {
                source = profilePath;
                ImodConfig config = ConvertStartupProfile(StartupProfile.FromJson(File.ReadAllText(profilePath, Encoding.UTF8)));
                config.HasScript = true;
                return config;
            }

            foreach (string legacyPath in new[] { GetLegacyAppDataImodScriptPath(), GetLegacyImodStartupPath() })
            {
                if (File.Exists(legacyPath))
                {
                    source = legacyPath;
                    ImodConfig config = ParseImodScriptFile(legacyPath);
                    config.HasScript = true;
                    return config;
                }
            }
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or InvalidDataException
            or JsonException or FormatException or OverflowException)
        {
            WriteLog($"IMOD.CONFIG: failed to read {source}: {FlattenLogText(ex.Message)}");
        }

        return new ImodConfig { HasScript = false };
    }

    // Loads the profile; a profile saved by an older layout is moved to the
    // current one, and an outdated script / runtime / loader is replaced.
    private ImodConfig LoadStartupConfig()
    {
        ImodConfig config = ReadSavedStartupConfig(out string? source);
        if (!config.HasScript || QaRun.IsReadOnly)
        {
            return config;
        }

        try
        {
            bool legacy = !string.Equals(source, GetStartupProfilePath(), StringComparison.OrdinalIgnoreCase);
            if (legacy || !IsStartupBundleCurrent())
            {
                InstallStartupBundle(config);
                WriteLog($"IMOD.STARTUP.{(legacy ? "MIGRATE" : "UPDATE")}: source=\"{source}\" target=\"{GetStartupRoot()}\" status=success");
            }
        }
        catch (Exception ex)
        {
            WriteLog($"IMOD.STARTUP.UPDATE: source=\"{source}\" status=failed error=\"{FlattenLogText(ex.ToString())}\"");
        }

        return config;
    }

    private bool IsStartupBundleCurrent()
    {
        try
        {
            return FileTextEquals(GetStartupScriptPath(), ReadStartupResource(StartupScriptResource))
                && FileTextEquals(GetStartupRuntimePath(), ReadStartupResource(StartupRuntimeResource))
                && TryVerifyEmbeddedImodResourceFile("DeviceTweakerCS.IMOD.Loader.kdu.exe", ".kdu.exe",
                    Path.Combine(GetImodStartupKduDirectory(), ImodKduFileName), out _)
                && TryVerifyEmbeddedImodResourceFile("DeviceTweakerCS.IMOD.Loader.drv64.dll", ".drv64.dll",
                    Path.Combine(GetImodStartupKduDirectory(), ImodKduDatabaseFileName), out _)
                && StartupTask.Exists();
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or InvalidDataException)
        {
            return false;
        }
    }

    private static bool FileTextEquals(string path, string expected)
        => File.Exists(path) && string.Equals(File.ReadAllText(path, Encoding.UTF8), expected, StringComparison.Ordinal);

    private void SaveStartupProfile(ImodConfig config)
    {
        InstallStartupBundle(config);
        config.HasScript = true;
        _imodConfigCache = config;
        _imodConfigLoaded = true;
    }

    private void InstallStartupBundle(ImodConfig config)
    {
        DemandRealWrite("InstallStartupBundle");
        using DriverLoadGate gate = DriverLoadGate.Enter();
        DriverLoadGate.AssertLoaderIdle();
        if (!EnsureImodDriverOnDisk(persistDriver: true, out string driverPath, out string? driverError))
        {
            throw new IOException("DTIMOD.sys could not be staged: " + driverError);
        }

        SecureStartupFolder(GetStartupParentDirectory());
        SecureStartupFolder(GetStartupRoot());
        // PowerShell 5.1 reads a script without BOM in the ANSI code page.
        WriteAllTextAtomic(GetStartupScriptPath(), ReadStartupResource(StartupScriptResource), new UTF8Encoding(true));
        WriteAllTextAtomic(GetStartupRuntimePath(), ReadStartupResource(StartupRuntimeResource), new UTF8Encoding(false));
        if (!EnsureImodStartupKduPayload(out _, out _, out string? loaderError))
        {
            throw new IOException("The startup loader could not be prepared: " + loaderError);
        }

        StartupProfile profile = BuildStartupProfile(config, driverPath);
        WriteAllTextAtomic(GetStartupProfilePath(), profile.ToJson(), new UTF8Encoding(false));
        if (!StartupTask.Register(GetStartupScriptPath(), WriteLog, out string? taskError))
        {
            throw new InvalidOperationException("The startup task could not be registered: " + taskError);
        }

        RemoveLegacyStartup(report: null);
        WriteLog($"IMOD.STARTUP: saved profile usb={profile.Usb.Controllers.Count} nic={profile.NicItr.Count} root=\"{GetStartupRoot()}\"");
    }

    private void RemoveStartupBundle(OperationReport? report)
    {
        DemandRealWrite("RemoveStartupBundle");
        if (!StartupTask.Delete(WriteLog, out string? taskError))
        {
            WriteLog($"IMOD.STARTUP.TASK: {taskError}");
            report?.AddError("IMOD.STARTUP", taskError ?? "startup task was not deleted");
        }

        DeleteFileIfExists(GetStartupProfilePath(), "IMOD.STARTUP", report);
        DeleteFileIfExists(GetStartupScriptPath(), "IMOD.STARTUP", report);
        DeleteFileIfExists(GetStartupRuntimePath(), "IMOD.STARTUP", report);
        DeleteDirectoryIfExists(GetImodStartupKduDirectory(), "IMOD.STARTUP");
        RemoveLegacyStartup(report);
    }

    private void RemoveLegacyStartup(OperationReport? report)
    {
        DeleteImodStartupRegistration(report);
        DeleteFileIfExists(GetLegacyAppDataImodScriptPath(), "IMOD.STARTUP.LEGACY", report);
        DeleteFileIfExists(GetLegacyImodStartupPath(), "IMOD.STARTUP.LEGACY", report);
        string legacyDirectory = GetLegacyAppDataImodDirectory();
        DeleteDirectoryIfExists(Path.Combine(legacyDirectory, "Loader"), "IMOD.STARTUP.LEGACY");
        if (Directory.Exists(legacyDirectory))
        {
            foreach (string previous in Directory.GetFiles(legacyDirectory, ImodScriptFileName + ".pre-DTIMOD-*.bak"))
            {
                DeleteFileIfExists(previous, "IMOD.STARTUP.LEGACY", report);
            }
        }
    }

    // A non-admin user can pre-create folders in %ProgramData% or plant a junction.
    // The folder must be a real directory owned by Administrators that only
    // administrators and SYSTEM can change; users keep read access to review the script.
    private static void SecureStartupFolder(string path)
    {
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

    internal static StartupProfile BuildStartupProfile(ImodConfig config, string driverPath)
    {
        StartupProfile profile = new()
        {
            CreatedBy = "DEVICE TWEAKER " + AppBuildIdentity.ProductVersion,
            Driver = new StartupFile { Path = driverPath, Sha256 = ReadEmbeddedImodDriverHash() ?? string.Empty },
            Runtime = new StartupFile { File = StartupRuntimeFileName, Sha256 = ComputeTextSha256(ReadStartupResource(StartupRuntimeResource)) },
            Loader = new StartupLoader
            {
                Kdu = new StartupFile
                {
                    File = Path.Combine("Loader", ImodKduFileName),
                    Sha256 = ComputeEmbeddedImodResourceHash("DeviceTweakerCS.IMOD.Loader.kdu.exe", ".kdu.exe"),
                },
                Database = new StartupFile
                {
                    File = Path.Combine("Loader", ImodKduDatabaseFileName),
                    Sha256 = ComputeEmbeddedImodResourceHash("DeviceTweakerCS.IMOD.Loader.drv64.dll", ".drv64.dll"),
                },
            },
            Logging = new StartupLogging { Enabled = ImodStartupScriptLoggingEnabled, Verbose = ImodStartupScriptVerboseLoggingEnabled },
            Usb = new StartupUsb
            {
                Apply = HasCustomUsbImod(config),
                Interval = StartupProfile.Hex(config.GlobalInterval),
                HcsparamsOffset = StartupProfile.Hex(config.GlobalHcsparamsOffset),
                Rtsoff = StartupProfile.Hex(config.GlobalRtsoff),
            },
        };

        foreach (ImodConfigEntry entry in config.Overrides
            .Where(entry => !string.IsNullOrWhiteSpace(entry.Hwid))
            .OrderBy(entry => entry.Hwid, StringComparer.OrdinalIgnoreCase))
        {
            profile.Usb.Controllers.Add(new StartupUsbController
            {
                Match = entry.Hwid,
                Enabled = entry.Enabled,
                Interval = entry.Interval is uint interval ? StartupProfile.Hex(interval) : null,
                Intervals = entry.Intervals is { Count: > 0 } intervals ? intervals.Select(value => StartupProfile.Hex(value)).ToList() : null,
                AdaptiveRoleBinding = entry.AdaptiveRoleBinding,
                RoleIntervals = entry.RoleIntervals is { Count: > 0 } roles ? FormatImodRoleIntervals(roles) : null,
                HcsparamsOffset = entry.HcsparamsOffset is uint hcs ? StartupProfile.Hex(hcs) : null,
                Rtsoff = entry.Rtsoff is uint rts ? StartupProfile.Hex(rts) : null,
            });
        }

        foreach (NicItrConfigEntry entry in config.NicItrEntries
            .Where(entry => !string.IsNullOrWhiteSpace(entry.Hwid) && entry.Values.Count > 0)
            .OrderBy(entry => entry.Hwid, StringComparer.OrdinalIgnoreCase))
        {
            profile.NicItr.Add(new StartupNicItr
            {
                Match = entry.Hwid,
                Family = entry.FamilyName,
                BaseOffset = StartupProfile.Hex(entry.BaseOffset),
                Stride = StartupProfile.Hex(entry.Stride),
                Queues = entry.Queues,
                Width = entry.Width,
                Mask = StartupProfile.Hex(entry.Mask),
                OrBits = StartupProfile.Hex(entry.OrBits),
                Values = entry.Values.Select(StartupProfile.Hex).ToList(),
            });
        }

        return profile;
    }

    internal static ImodConfig ConvertStartupProfile(StartupProfile profile)
    {
        ImodConfig config = new()
        {
            GlobalInterval = checked((uint)StartupProfile.ParseNumber(profile.Usb.Interval)),
            GlobalHcsparamsOffset = checked((uint)StartupProfile.ParseNumber(profile.Usb.HcsparamsOffset)),
            GlobalRtsoff = checked((uint)StartupProfile.ParseNumber(profile.Usb.Rtsoff)),
        };

        static uint? ToUInt(string? text) => text is null ? null : checked((uint)StartupProfile.ParseNumber(text));

        foreach (StartupUsbController controller in profile.Usb.Controllers)
        {
            Dictionary<string, uint>? roles = null;
            if (!string.IsNullOrWhiteSpace(controller.RoleIntervals)
                && !TryParseImodRoleIntervals(controller.RoleIntervals, out roles))
            {
                throw new InvalidDataException($"Invalid roleIntervals for {controller.Match}.");
            }

            config.Overrides.Add(new ImodConfigEntry
            {
                Hwid = controller.Match,
                Enabled = controller.Enabled,
                Interval = ToUInt(controller.Interval),
                Intervals = controller.Intervals?.Select(value => checked((uint)StartupProfile.ParseNumber(value))).ToList(),
                AdaptiveRoleBinding = controller.AdaptiveRoleBinding,
                RoleIntervals = roles,
                HcsparamsOffset = ToUInt(controller.HcsparamsOffset),
                Rtsoff = ToUInt(controller.Rtsoff),
            });
        }

        foreach (StartupNicItr nic in profile.NicItr)
        {
            config.NicItrEntries.Add(new NicItrConfigEntry
            {
                Hwid = nic.Match,
                FamilyName = nic.Family,
                BaseOffset = checked((uint)StartupProfile.ParseNumber(nic.BaseOffset)),
                Stride = checked((uint)StartupProfile.ParseNumber(nic.Stride)),
                Queues = nic.Queues,
                Width = nic.Width,
                Mask = StartupProfile.ParseNumber(nic.Mask),
                OrBits = StartupProfile.ParseNumber(nic.OrBits),
                Values = nic.Values.Select(StartupProfile.ParseNumber).ToList(),
            });
        }

        return config;
    }

    private static string ComputeTextSha256(string text)
        => ComputeSha256(new UTF8Encoding(false).GetBytes(text));
}
