using System.Text;

using static DeviceTweakerCS.ImodDriverNative;
using static DeviceTweakerCS.ImodDriverFiles;

namespace DeviceTweakerCS;

// Saved USB IMOD / NIC ITR profile and the standalone script that applies it.
// The script, its C# runtime and the loader live in %ProgramData%\DEVICE TWEAKER\Startup,
// run from Task Scheduler as SYSTEM and keep working without DEVICE TWEAKER.exe.
// StartupBundle owns those files and the task; this part maps settings to the profile.
public sealed partial class MainForm
{
    private const string StartupScriptResource = "DeviceTweakerCS.Startup.ApplyIMOD.ps1";
    private const string StartupRuntimeResource = "DeviceTweakerCS.Startup.DeviceTweakerImodRuntime.cs";

    internal static string GetStartupRoot() => StartupBundleLayout.ForCurrentUser().Root;

    private static string GetStartupScriptPath() => StartupBundleLayout.ForCurrentUser().ScriptPath;

    private static string GetStartupProfilePath() => StartupBundleLayout.ForCurrentUser().ProfilePath;

    // Before this layout the script lived in %APPDATA% and started from HKCU\...\Run;
    // older versions used the Startup folder.
    private static string GetLegacyAppDataImodScriptPath()
        => Path.Combine(StartupBundleLayout.CurrentLegacyAppDataDirectory(), ImodScriptFileName);

    private static string GetLegacyImodStartupPath()
        => Path.Combine(StartupBundleLayout.CurrentLegacyStartupFolder(), ImodScriptFileName);

    private StartupBundle CreateStartupBundle()
        => new(StartupBundleLayout.ForCurrentUser(), new EmbeddedStartupPayload(), new ScheduledStartupTask(),
            StartupBundle.SecureFolder, DeleteImodStartupRegistration, WriteLog);

    // The script, runtime and loader embedded in this EXE.
    private sealed class EmbeddedStartupPayload : IStartupPayload
    {
        private const string KduResource = "DeviceTweakerCS.IMOD.Loader.kdu.exe";
        private const string DatabaseResource = "DeviceTweakerCS.IMOD.Loader.drv64.dll";

        public string Script => ReadStartupResource(StartupScriptResource);
        public string Runtime => ReadStartupResource(StartupRuntimeResource);

        public bool TryWriteLoader(string directory, out string? error)
        {
            string kduPath = Path.Combine(directory, ImodKduFileName);
            string dbPath = Path.Combine(directory, ImodKduDatabaseFileName);
            return TryWriteEmbeddedImodResource(KduResource, ".kdu.exe", kduPath, out error)
                && TryWriteEmbeddedImodResource(DatabaseResource, ".drv64.dll", dbPath, out error)
                && TryVerifyEmbeddedImodResourceFile(KduResource, ".kdu.exe", kduPath, out error)
                && TryVerifyEmbeddedImodResourceFile(DatabaseResource, ".drv64.dll", dbPath, out error);
        }

        public bool IsLoaderCurrent(string directory)
            => TryVerifyEmbeddedImodResourceFile(KduResource, ".kdu.exe", Path.Combine(directory, ImodKduFileName), out _)
                && TryVerifyEmbeddedImodResourceFile(DatabaseResource, ".drv64.dll", Path.Combine(directory, ImodKduDatabaseFileName), out _);
    }

    private static string ReadStartupResource(string name)
    {
        using Stream stream = typeof(MainForm).Assembly.GetManifestResourceStream(name)
            ?? throw new InvalidDataException("Missing embedded startup resource: " + name);
        using StreamReader reader = new(stream, new UTF8Encoding(false), detectEncodingFromByteOrderMarks: true);
        return reader.ReadToEnd();
    }

    // Reads without migration. An unreadable current file never falls back to legacy settings.
    private StartupConfigReadResult<ImodConfig> ReadSavedStartupConfigResult()
    {
#if DEVELOPER_QA
        DevRecordDataRead("ReadSavedStartupConfigResult");
#endif
        return StartupConfigReader.Read(GetStartupProfilePath(),
            new[] { GetLegacyAppDataImodScriptPath(), GetLegacyImodStartupPath() },
            path =>
            {
                ImodConfig config = ConvertStartupProfile(StartupProfile.FromJson(File.ReadAllText(path, Encoding.UTF8)));
                config.HasScript = true;
                return config;
            },
            path =>
            {
                ImodConfig config = ParseImodScriptFile(path);
                config.HasScript = true;
                return config;
            });
    }

    private ImodConfig ReadSavedStartupConfig(out string? source)
    {
        StartupConfigReadResult<ImodConfig> result = ReadSavedStartupConfigResult();
        source = result.Source;
        return result.RequireValueOrCreate(() => new ImodConfig { HasScript = false });
    }

    // Loads the profile; a profile saved by an older layout is moved to the
    // current one, and an outdated script / runtime / loader is replaced.
    private ImodConfig LoadStartupConfig(bool virtualOnly)
    {
        // Virtual fixtures own their saved values; never inherit the host's profile.
        if (virtualOnly) return new ImodConfig { HasScript = false };
        StartupConfigReadResult<ImodConfig> result = ReadSavedStartupConfigResult();
        string? source = result.Source;
        if (result.State == StartupConfigReadState.Error)
        {
            string diagnostic = $"{source}: {result.Error}";
            WriteLog($"IMOD.CONFIG: read failed; save/apply blocked: {FlattenLogText(diagnostic)}");
            return new ImodConfig { StartupReadError = diagnostic };
        }
        ImodConfig config = result.RequireValueOrCreate(() => new ImodConfig { HasScript = false });
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

    private bool IsStartupBundleCurrent() => CreateStartupBundle().IsCurrent();

    private void InstallStartupBundle(ImodConfig config)
    {
        DemandRealWrite("InstallStartupBundle");
        using DriverLoadGate gate = DriverLoadGate.Enter();
        DriverLoadGate.AssertLoaderIdle();
        if (!EnsureImodDriverOnDisk(persistDriver: true, out string driverPath, out string? driverError))
        {
            throw new IOException("DTIMOD.sys could not be staged: " + driverError);
        }

        StartupProfile profile = BuildStartupProfile(config, driverPath);
        string profileJson = profile.ToJson();
        _ = ConvertStartupProfile(StartupProfile.FromJson(profileJson));

        StartupBundle bundle = CreateStartupBundle();
        bundle.Install(profileJson);
        WriteLog($"IMOD.STARTUP: saved profile usb={profile.Usb.Controllers.Count} nic={profile.NicItr.Count} root=\"{bundle.Layout.Root}\"");
    }

    private void RemoveStartupBundle(OperationReport? report)
    {
        DemandRealWrite("RemoveStartupBundle");
        CreateStartupBundle().Remove(report);
    }

    internal static StartupProfile BuildStartupProfile(ImodConfig config, string driverPath)
    {
        StartupProfile profile = new()
        {
            CreatedBy = "DEVICE TWEAKER " + AppBuildIdentity.ProductVersion,
            Driver = new StartupFile { Path = driverPath, Sha256 = ReadEmbeddedImodDriverHash() ?? string.Empty },
            Runtime = new StartupFile { File = StartupBundleLayout.RuntimeFileName, Sha256 = ComputeTextSha256(ReadStartupResource(StartupRuntimeResource)) },
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
