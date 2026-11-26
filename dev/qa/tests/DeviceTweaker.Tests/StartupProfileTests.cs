using System.Text.Json;
using System.Xml.Linq;
using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public sealed class StartupProfileTests
{
    private const string DriverPath = @"C:\Windows\DTIMOD.sys";

    internal static MainForm.ImodConfig Sample()
    {
        MainForm.ImodConfig config = new() { GlobalInterval = 0xC8, GlobalHcsparamsOffset = 0x4, GlobalRtsoff = 0x18 };
        config.Overrides.Add(new MainForm.ImodConfigEntry
        {
            Hwid = "VEN_8086&DEV_7AE0",
            Enabled = true,
            Interval = 0x0,
            AdaptiveRoleBinding = true,
            RoleIntervals = new(StringComparer.OrdinalIgnoreCase) { ["Mouse"] = 0x0, ["Audio"] = 0xFA0 },
        });
        config.Overrides.Add(new MainForm.ImodConfigEntry
        {
            Hwid = "VEN_1022&DEV_15B6",
            Intervals = [0x0, 0x64, 0xC8],
            HcsparamsOffset = 0x4,
            Rtsoff = 0x18,
        });
        config.NicItrEntries.Add(new MainForm.NicItrConfigEntry
        {
            Hwid = "VEN_8086&DEV_125C",
            FamilyName = "Intel I225/I226 (EITR)",
            BaseOffset = 0x1680,
            Stride = 0x4,
            Queues = 5,
            Width = 32,
            Mask = 0x7FFC,
            OrBits = 0x80000000,
            Values = [0x0, 0x28, 0x28, 0x28, 0x28],
        });
        return config;
    }

    [Fact]
    public void UsbProfileMatchUsesExactInstancesAndRejectsAmbiguousLegacyKeys()
    {
        const string first = @"PCI\VEN_1022&DEV_15B6\A";
        const string second = @"PCI\VEN_1022&DEV_15B6\B";
        const string legacy = "DEV_15B6";
        string[] connected = [first, second];

        Assert.Equal(first, DeviceTweakerImodRuntime.SelectUsbProfileMatch(first, connected, [legacy, first, second], out bool ambiguous));
        Assert.False(ambiguous);
        Assert.Equal(second, DeviceTweakerImodRuntime.SelectUsbProfileMatch(second, connected, [first, second], out ambiguous));
        Assert.False(ambiguous);
        Assert.Null(DeviceTweakerImodRuntime.SelectUsbProfileMatch(second, connected, [legacy, first], out ambiguous));
        Assert.True(ambiguous);
        Assert.Equal(legacy, DeviceTweakerImodRuntime.SelectUsbProfileMatch(second, [second], [legacy], out ambiguous));
        Assert.False(ambiguous);
        Assert.Null(DeviceTweakerImodRuntime.SelectUsbProfileMatch(first + "0", [first + "0"], [first], out ambiguous));
        Assert.False(ambiguous);
    }

    [Fact]
    public void ProfileRoundTripKeepsEveryValue()
    {
        string json = MainForm.BuildStartupProfile(Sample(), DriverPath).ToJson();
        MainForm.ImodConfig restored = MainForm.ConvertStartupProfile(StartupProfile.FromJson(json));

        Assert.Equal(json, MainForm.BuildStartupProfile(restored, DriverPath).ToJson());
        Assert.Equal(0x80000000UL, restored.NicItrEntries.Single().OrBits);
        Assert.Equal(0xFA0u, restored.Overrides.Single(e => e.Hwid == "VEN_8086&DEV_7AE0").RoleIntervals!["Audio"]);
    }

    [Fact]
    public void ProfileIsReadableJsonWithHexRegisters()
    {
        string json = MainForm.BuildStartupProfile(Sample(), DriverPath).ToJson();
        using JsonDocument document = JsonDocument.Parse(json);
        JsonElement root = document.RootElement;

        Assert.Equal(1, root.GetProperty("schemaVersion").GetInt32());
        Assert.Equal(DriverPath, root.GetProperty("driver").GetProperty("path").GetString());
        Assert.Equal("DeviceTweakerImodRuntime.cs", root.GetProperty("runtime").GetProperty("file").GetString());
        Assert.Matches("^[0-9A-F]{64}$", root.GetProperty("runtime").GetProperty("sha256").GetString());
        Assert.Equal(@"Loader\kdu.exe", root.GetProperty("loader").GetProperty("kdu").GetProperty("file").GetString());
        Assert.True(root.GetProperty("usb").GetProperty("apply").GetBoolean());
        Assert.Equal("0xC8", root.GetProperty("usb").GetProperty("interval").GetString());
        Assert.Equal("0x80000000", root.GetProperty("nicItr")[0].GetProperty("orBits").GetString());
        Assert.Equal("Mouse=0x0, Audio=0xFA0",
            root.GetProperty("usb").GetProperty("controllers").EnumerateArray()
                .Single(c => c.GetProperty("match").GetString() == "VEN_8086&DEV_7AE0").GetProperty("roleIntervals").GetString());
    }

    [Fact]
    public void ProfileWithUnknownSchemaIsRejected()
    {
        string json = MainForm.BuildStartupProfile(Sample(), DriverPath).ToJson().Replace("\"schemaVersion\": 1", "\"schemaVersion\": 2");

        Assert.Throws<InvalidDataException>(() => StartupProfile.FromJson(json));
    }

    [Fact]
    public void LegacyScriptSettingsBecomeTheSameProfile()
    {
        const string legacy = """
            $imodScriptVersion = 30
            $imodSettingsBegin = $true
            $ImodDriverPath = 'C:\Windows\DTIMOD.sys'
            $globalInterval = 0xC8
            $globalHCSPARAMSOffset = 0x4
            $globalRTSOFF = 0x18
            $applyUsbImod = $true
            $userDefinedData = @{
                "VEN_1022&DEV_15B6" = @{
                    "INTERVALS" = @(0x0, 0x64, 0xC8)
                    "HCSPARAMS_OFFSET" = 0x4
                    "RTSOFF" = 0x18
                }
                "VEN_8086&DEV_7AE0" = @{
                    "ENABLED" = $true
                    "INTERVAL" = 0x0
                    "ADAPTIVE_ROLE_BINDING" = $true
                    "ROLE_INTERVALS" = 'Mouse=0x0, Audio=0xFA0'
                }
            }
            $nicItrData = @(
                @{
                    "HWID" = 'VEN_8086&DEV_125C'
                    "FAMILY" = 'Intel I225/I226 (EITR)'
                    "BASE_OFFSET" = 0x1680
                    "STRIDE" = 0x4
                    "QUEUES" = 5
                    "WIDTH" = 32
                    "MASK" = 0x7FFC
                    "OR_BITS" = 0x80000000
                    "VALUES" = @(0x0, 0x28, 0x28, 0x28, 0x28)
                }
            )
            $imodSettingsEnd = $true
            OLD ENGINE BODY
            """;

        MainForm.ImodConfig parsed = MainForm.ParseImodScriptText(legacy);

        Assert.Equal(
            MainForm.BuildStartupProfile(Sample(), DriverPath).ToJson(),
            MainForm.BuildStartupProfile(parsed, DriverPath).ToJson());
    }

    // Backups of real machines (Invoke-Qa passes their folder) hold whole scripts.
    [Theory]
    [MemberData(nameof(BackupRestoreTests.ExistingBackupFiles), MemberType = typeof(BackupRestoreTests))]
    public void ExistingBackupScriptsConvertWithoutLosingValues(string path)
    {
        if (path == "synthetic")
        {
            return;
        }

        DeviceTweakerBackup backup = JsonSerializer.Deserialize<DeviceTweakerBackup>(File.ReadAllText(path))!;
        if (backup.ImodScript is not { Exists: true, Text: { } text })
        {
            return;
        }

        MainForm.ImodConfig parsed = MainForm.ParseImodScriptText(text);
        string json = MainForm.BuildStartupProfile(parsed, DriverPath).ToJson();
        MainForm.ImodConfig converted = MainForm.ConvertStartupProfile(StartupProfile.FromJson(json));

        Assert.Equal(json, MainForm.BuildStartupProfile(converted, DriverPath).ToJson());
        Assert.Equal(parsed.Overrides.Count(e => !string.IsNullOrWhiteSpace(e.Hwid)), converted.Overrides.Count);
        Assert.Equal(parsed.NicItrEntries.Count(e => e.Values.Count > 0), converted.NicItrEntries.Count);
        Assert.Equal(parsed.GlobalInterval, converted.GlobalInterval);
    }
}

public sealed class StartupTaskTests
{
    private static readonly XNamespace Ns = "http://schemas.microsoft.com/windows/2004/02/mit/task";

    [Fact]
    public void TaskRunsScriptAsSystemAtLogonResumeAndDeviceStart()
    {
        const string script = @"C:\ProgramData\DEVICE TWEAKER\Startup\ApplyIMOD.ps1";
        XDocument task = XDocument.Parse(StartupTask.BuildXml(script));

        Assert.Equal("S-1-5-18", task.Descendants(Ns + "UserId").Single().Value);
        Assert.Equal(@"\DEVICE TWEAKER\IMOD", task.Descendants(Ns + "URI").Single().Value);
        Assert.Single(task.Descendants(Ns + "LogonTrigger"));
        List<string> subscriptions = task.Descendants(Ns + "Subscription").Select(e => e.Value).ToList();
        Assert.Contains(subscriptions, s => s.Contains("Microsoft-Windows-Power-Troubleshooter", StringComparison.Ordinal) && s.Contains("EventID=1]", StringComparison.Ordinal));
        Assert.Contains(subscriptions, s => s.Contains("Microsoft-Windows-Kernel-PnP/Configuration", StringComparison.Ordinal) && s.Contains("EventID=410", StringComparison.Ordinal));
        Assert.Equal("Event/EventData/Data[@Name='DeviceInstanceId']", task.Descendants(Ns + "Value").Single().Value);
        Assert.Equal("IgnoreNew", task.Descendants(Ns + "MultipleInstancesPolicy").Single().Value);
        Assert.Equal("false", task.Descendants(Ns + "Hidden").Single().Value);
        string arguments = task.Descendants(Ns + "Arguments").Single().Value;
        Assert.Contains($"-File \"{script}\"", arguments);
        Assert.Contains("-TriggerDevice \"$(Device)\"", arguments);
        Assert.EndsWith(@"WindowsPowerShell\v1.0\powershell.exe", task.Descendants(Ns + "Command").Single().Value);
    }

    [Fact]
    public void ShippedStartupFilesSatisfyThePersistenceContract()
    {
        Assert.True(MainForm.ValidateImodStartupPersistenceContract(out string error), error);
    }

    [Fact]
    public void ScriptPathWithXmlCharactersStaysIntact()
    {
        const string script = @"C:\Data & Co\O'Brien <test>\ApplyIMOD.ps1";
        XDocument task = XDocument.Parse(StartupTask.BuildXml(script));

        Assert.Contains($"-File \"{script}\"", task.Descendants(Ns + "Arguments").Single().Value);
    }

    private const string Script = @"C:\ProgramData\DEVICE TWEAKER\Startup\ApplyIMOD.ps1";

    private static string Registered(Action<XElement> change)
    {
        XDocument task = XDocument.Parse(StartupTask.BuildXml(Script));
        change(task.Descendants(Ns + "Actions").Single());
        return task.ToString();
    }

    [Fact]
    public void RegisteredTaskMatchingTheBuiltXmlIsAccepted()
    {
        Assert.Null(StartupTask.VerifyRegisteredXml(StartupTask.BuildXml(Script), Script));
    }

    [Fact]
    public void RegisteredTaskRequiresTheIntendedAccountAndDeviceTrigger()
    {
        XDocument task = XDocument.Parse(StartupTask.BuildXml(Script));
        XElement account = task.Descendants(Ns + "UserId").Single();
        account.Value = "S-1-5-19";
        Assert.Contains("does not run as SYSTEM", StartupTask.VerifyRegisteredXml(task.ToString(), Script));

        account.Value = "S-1-5-18";
        task.Descendants(Ns + "ValueQueries").Single().Remove();
        Assert.Contains("device query is missing", StartupTask.VerifyRegisteredXml(task.ToString(), Script));
    }

    [Fact]
    public void NicProfileBindsToAnExactInstanceOrAUniqueLegacyMatch()
    {
        const string key = "VEN_8086&DEV_15F3";
        const string first = @"PCI\VEN_8086&DEV_15F3\1";
        const string second = @"PCI\VEN_8086&DEV_15F3\2";
        Assert.Equal(second, DeviceTweakerImodRuntime.SelectNicProfileTarget(second.ToLowerInvariant(), [first, second], out bool ambiguous));
        Assert.False(ambiguous);
        Assert.Null(DeviceTweakerImodRuntime.SelectNicProfileTarget(first, [second], out ambiguous));
        Assert.False(ambiguous);
        Assert.Equal(second, DeviceTweakerImodRuntime.SelectNicProfileTarget(key, [second], out ambiguous));
        Assert.False(ambiguous);
        Assert.Null(DeviceTweakerImodRuntime.SelectNicProfileTarget(key, [first, second], out ambiguous));
        Assert.True(ambiguous);
    }

    [Fact]
    public void RegisteredTaskWithAnAddedActionIsRejected()
    {
        string xml = Registered(actions => actions.Add(new XElement(Ns + "Exec",
            new XElement(Ns + "Command", @"C:\Windows\System32\cmd.exe"),
            new XElement(Ns + "Arguments", "/c " + Script))));

        Assert.Contains("2 actions", StartupTask.VerifyRegisteredXml(xml, Script));
    }

    [Fact]
    public void RegisteredTaskWithANonExecActionIsRejected()
    {
        string xml = Registered(actions =>
        {
            actions.RemoveNodes();
            actions.Add(new XElement(Ns + "ComHandler", new XElement(Ns + "ClassId", "{00000000-0000-0000-0000-000000000000}")));
        });

        Assert.Contains("expected one Exec", StartupTask.VerifyRegisteredXml(xml, Script));
    }

    [Fact]
    public void RegisteredTaskWithChangedCommandOrArgumentsIsRejected()
    {
        string command = Registered(actions => actions.Descendants(Ns + "Command").Single().Value = @"C:\Temp\powershell.exe");
        string arguments = Registered(actions => actions.Descendants(Ns + "Arguments").Single().Value
            = StartupTask.BuildArguments(Script) + " -Command calc");

        Assert.Contains("unexpected command", StartupTask.VerifyRegisteredXml(command, Script));
        Assert.Contains("unexpected arguments", StartupTask.VerifyRegisteredXml(arguments, Script));
    }

    [Fact]
    public void RegisteredTaskForAnotherScriptIsRejected()
    {
        string other = StartupTask.BuildXml(@"C:\Users\Public\ApplyIMOD.ps1");

        Assert.Contains("unexpected arguments", StartupTask.VerifyRegisteredXml(other, Script));
    }

    [Fact]
    public void MalformedRegisteredTaskXmlIsRejected()
    {
        Assert.Contains("could not be parsed", StartupTask.VerifyRegisteredXml("<Task>", Script));
        Assert.Contains("0 action lists", StartupTask.VerifyRegisteredXml($"<Task xmlns=\"{Ns}\"/>", Script));
    }
}
