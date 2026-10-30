using Microsoft.Win32;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;

namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private const uint ImodDefaultInterval = 0xC8;
    private const uint ImodDefaultHcsparamsOffset = 0x4;
    private const uint ImodDefaultRtsoff = 0x18;
    private const string ImodScriptFileName = "ApplyIMOD.ps1";
    private const string ImodStartupRunKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
    private const string ImodStartupRunValueName = "DEVICE TWEAKER IMOD";
    private const string ImodDriverName = "DTIMOD.sys";
    private const bool ImodStartupScriptLoggingEnabled = true;
    private const bool ImodStartupScriptVerboseLoggingEnabled = false;

    private ImodConfig? _imodConfigCache;
    private bool _imodConfigLoaded;

    internal sealed class ImodConfigEntry
    {
        public required string Hwid { get; set; }
        public uint? Interval { get; set; }
        public List<uint>? Intervals { get; set; }
        public bool? AdaptiveRoleBinding { get; set; }
        public Dictionary<string, uint>? RoleIntervals { get; set; }
        public uint? HcsparamsOffset { get; set; }
        public uint? Rtsoff { get; set; }
        public bool? Enabled { get; set; }
    }

    internal sealed class NicItrConfigEntry
    {
        public required string Hwid { get; set; }
        public string FamilyName { get; set; } = string.Empty;
        public uint BaseOffset { get; set; }
        public uint Stride { get; set; }
        public int Queues { get; set; }
        public int Width { get; set; }
        public ulong Mask { get; set; }
        public ulong OrBits { get; set; }
        public List<ulong> Values { get; set; } = [];
    }

    internal sealed class ImodConfig
    {
        public uint GlobalInterval { get; set; } = ImodDefaultInterval;
        public uint GlobalHcsparamsOffset { get; set; } = ImodDefaultHcsparamsOffset;
        public uint GlobalRtsoff { get; set; } = ImodDefaultRtsoff;
        public List<ImodConfigEntry> Overrides { get; } = [];
        public List<NicItrConfigEntry> NicItrEntries { get; } = [];
        public bool HasScript { get; set; }
    }

    private enum ImodApplyOutcome
    {
        Applied,
        AppliedWithWarnings,
        SkippedNoUsb,
        SkippedNoController,
        SkippedNoConfig,
        Failed,
    }

    private sealed record ImodInput(uint? Interval, List<uint>? Intervals, Dictionary<string, uint>? RoleIntervals, bool IsDefault);

    private void InvalidateImodCache()
    {
        _imodConfigCache = null;
        _imodConfigLoaded = false;
    }

    private void EnsureImodConfigLoaded()
    {
        if (_imodConfigLoaded)
        {
            return;
        }

        _imodConfigLoaded = true;
        _imodConfigCache = LoadStartupConfig();
    }

    // Backups may hold the current profile or a script saved by an older layout.
    private static bool IsManagedImodScriptPath(string path)
    {
        try
        {
            string fullPath = Path.GetFullPath(path);
            return new[] { GetStartupProfilePath(), GetLegacyAppDataImodScriptPath(), GetLegacyImodStartupPath() }
                .Any(candidate => string.Equals(fullPath, Path.GetFullPath(candidate), StringComparison.OrdinalIgnoreCase));
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
        {
            return false;
        }
    }

    internal static bool ValidateImodStartupPersistenceContract(out string error)
    {
        string root = GetStartupRoot();
        string programData = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);
        if (!root.StartsWith(programData + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
        {
            error = $"startup folder is outside %ProgramData%: {root}";
            return false;
        }

        string task = StartupTask.BuildXml(GetStartupScriptPath());
        string[] taskRequirements = ["S-1-5-18", "<LogonTrigger>", "Power-Troubleshooter", "EventID=410", "-TriggerDevice", GetStartupScriptPath()];
        string? missingTask = taskRequirements.FirstOrDefault(token => !task.Contains(token, StringComparison.Ordinal));
        if (missingTask is not null)
        {
            error = $"startup task is missing: {missingTask}";
            return false;
        }

        string script = ReadStartupResource(StartupScriptResource);
        string runtime = ReadStartupResource(StartupRuntimeResource);
        string[] scriptRequirements = ["Import-ImodProfile", "Test-ImodFileHash -name 'DeviceTweakerImodRuntime.cs'", "-role 'Audio'"];
        string? missingScript = scriptRequirements.FirstOrDefault(token => !script.Contains(token, StringComparison.Ordinal));
        if (missingScript is not null)
        {
            error = $"startup script is missing: {missingScript}";
            return false;
        }

        if (!runtime.Contains("registers.TryRead(address, 4, out readBack", StringComparison.Ordinal)
            || !runtime.Contains("IMOD readback mismatch", StringComparison.Ordinal))
        {
            error = "startup runtime does not read IMOD back after writing";
            return false;
        }

        error = string.Empty;
        return true;
    }

    private static string GetImodDriverSystemPath()
    {
        string windows = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
        if (string.IsNullOrWhiteSpace(windows))
        {
            windows = Path.GetPathRoot(Environment.SystemDirectory) ?? string.Empty;
            if (string.IsNullOrWhiteSpace(windows))
            {
                windows = AppContext.BaseDirectory;
                if (string.IsNullOrWhiteSpace(windows))
                {
                    windows = Environment.CurrentDirectory;
                }
            }
        }

        return Path.Combine(windows, ImodDriverName);
    }

    private static bool EnsureImodStartupKduPayload(out string kduPath, out string dbPath, out string? error)
    {
        string directory = GetImodStartupKduDirectory();
        kduPath = Path.Combine(directory, ImodKduFileName);
        dbPath = Path.Combine(directory, ImodKduDatabaseFileName);

        try
        {
            SecureStartupFolder(directory);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            error = ex.Message;
            return false;
        }

        if (!TryWriteEmbeddedImodResource("DeviceTweakerCS.IMOD.Loader.kdu.exe", ".kdu.exe", kduPath, out error))
        {
            return false;
        }

        if (!TryWriteEmbeddedImodResource("DeviceTweakerCS.IMOD.Loader.drv64.dll", ".drv64.dll", dbPath, out error))
        {
            return false;
        }
        return TryVerifyEmbeddedImodResourceFile("DeviceTweakerCS.IMOD.Loader.kdu.exe", ".kdu.exe", kduPath, out error)
            && TryVerifyEmbeddedImodResourceFile("DeviceTweakerCS.IMOD.Loader.drv64.dll", ".drv64.dll", dbPath, out error);
    }

    private static bool IsImodDriverSystemPath(string? path)
    {
        string systemPath = GetImodDriverSystemPath();
        if (string.IsNullOrWhiteSpace(path) || string.IsNullOrWhiteSpace(systemPath))
        {
            return false;
        }

        try
        {
            return string.Equals(Path.GetFullPath(path), Path.GetFullPath(systemPath), StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return string.Equals(path, systemPath, StringComparison.OrdinalIgnoreCase);
        }
    }

    private void RemoveImodPersistenceFiles(OperationReport? report = null)
    {
        DemandRealWrite("RemoveImodPersistenceFiles");
        RestoreNicDriverModerationForSavedProfiles(report);
        RemoveStartupBundle(report);
        DeleteFileIfExists(Path.Combine(GetScriptRoot(), "dtimod.sys"), "IMOD.DRIVER.LEGACY", report);
        DeleteFileIfExists(Path.Combine(GetScriptRoot(), ImodDriverName), "IMOD.DRIVER.LEGACY", report);
        WriteLog($"IMOD.DRIVER: keep staged system driver {GetImodDriverSystemPath()}");
    }

    // HKCU\...\Run started the script before the scheduled task replaced it.
    private void DeleteImodStartupRegistration(OperationReport? report = null)
    {
        DemandRealWrite("DeleteImodStartupRegistration");
        try
        {
            using RegistryKey? key = Registry.CurrentUser.OpenSubKey(ImodStartupRunKeyPath, writable: true);
            if (key?.GetValue(ImodStartupRunValueName, null, RegistryValueOptions.DoNotExpandEnvironmentNames) is null)
            {
                return;
            }

            key.DeleteValue(ImodStartupRunValueName, throwOnMissingValue: false);
            WriteLog($"IMOD.STARTUP: deleted HKCU\\{ImodStartupRunKeyPath} value=\"{ImodStartupRunValueName}\"");
        }
        catch (Exception ex)
        {
            WriteLog($"IMOD.STARTUP: failed to delete registration: {ex.Message}");
            report?.AddError("IMOD.STARTUP", $"failed to delete startup registration: {ex.Message}");
        }
    }

    private void DeleteFileIfExists(string? path, string label, OperationReport? report = null)
    {
        DemandRealWrite("DeleteFileIfExists");
        if (string.IsNullOrWhiteSpace(path))
        {
            return;
        }

        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
                WriteLog($"{label}: deleted {path}");
            }
        }
        catch (Exception ex)
        {
            WriteLog($"{label}: failed to delete {path}: {ex.Message}");
            report?.AddError(label, $"failed to delete {path}: {ex.Message}");
        }
    }

    private void DeleteDirectoryIfExists(string? path, string label)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return;
        }

        try
        {
            if (Directory.Exists(path))
            {
                Directory.Delete(path, recursive: true);
                WriteLog($"{label}: deleted {path}");
            }
        }
        catch (Exception ex)
        {
            WriteLog($"{label}: failed to delete {path}: {ex.Message}");
        }
    }

    // Reads the settings block of an ApplyIMOD.ps1 saved by an older layout.
    private static ImodConfig ParseImodScriptFile(string path)
        => ParseImodScriptText(File.ReadAllText(path, Encoding.UTF8));

    internal static ImodConfig ParseImodScriptText(string text)
    {
        ImodConfig config = new();
        ImodConfigEntry? currentDevice = null;
        NicItrConfigEntry? currentNicItr = null;
        bool inOverrides = false;
        bool inNicItr = false;

        string[] lines = text.Replace("\r\n", "\n", StringComparison.Ordinal).Split('\n');
        foreach (string raw in lines)
        {
            string line = StripInlineComment(raw).Trim();
            if (line.Length > 0 && line[0] == '\uFEFF')
            {
                line = line.TrimStart('\uFEFF');
            }
            if (line.Length == 0)
            {
                continue;
            }

            if (inNicItr)
            {
                if (currentNicItr is null)
                {
                    if (line.StartsWith(")", StringComparison.Ordinal))
                    {
                        inNicItr = false;
                        continue;
                    }

                    if (line.StartsWith("@{", StringComparison.Ordinal))
                    {
                        currentNicItr = new NicItrConfigEntry { Hwid = string.Empty };
                    }

                    continue;
                }

                if (line.StartsWith("}", StringComparison.Ordinal))
                {
                    if (!string.IsNullOrWhiteSpace(currentNicItr.Hwid)
                        && currentNicItr.Queues > 0
                        && (currentNicItr.Width == 16 || currentNicItr.Width == 32)
                        && currentNicItr.Values.Count > 0)
                    {
                        config.NicItrEntries.Add(currentNicItr);
                    }

                    currentNicItr = null;
                    continue;
                }

                if (!TryParseQuotedAssignment(line, out string nicKeyName, out string nicValueText))
                {
                    continue;
                }

                string nicKey = nicKeyName.Trim().ToUpperInvariant();
                if (nicKey == "HWID")
                {
                    currentNicItr.Hwid = UnquotePowerShellString(nicValueText);
                }
                else if (nicKey == "FAMILY")
                {
                    currentNicItr.FamilyName = UnquotePowerShellString(nicValueText);
                }
                else if (nicKey == "BASE_OFFSET" && TryParseUInt32Flexible(nicValueText, out uint baseOffset))
                {
                    currentNicItr.BaseOffset = baseOffset;
                }
                else if (nicKey == "STRIDE" && TryParseUInt32Flexible(nicValueText, out uint stride))
                {
                    currentNicItr.Stride = stride;
                }
                else if (nicKey == "QUEUES" && TryParseUInt32Flexible(nicValueText, out uint queues))
                {
                    currentNicItr.Queues = (int)Math.Min(queues, 1024);
                }
                else if (nicKey == "WIDTH" && TryParseUInt32Flexible(nicValueText, out uint width))
                {
                    currentNicItr.Width = (int)width;
                }
                else if (nicKey == "MASK" && TryParseUInt64Flexible(nicValueText, out ulong mask))
                {
                    currentNicItr.Mask = mask;
                }
                else if ((nicKey == "OR_BITS" || nicKey == "ORBITS") && TryParseUInt64Flexible(nicValueText, out ulong orBits))
                {
                    currentNicItr.OrBits = orBits;
                }
                else if (nicKey == "VALUES" && TryParseUInt64List(nicValueText, out List<ulong> nicValues))
                {
                    currentNicItr.Values = nicValues;
                }

                continue;
            }

            if (TryParseAssignment(line, "$globalInterval", out string valueText)
                && TryParseUInt32Flexible(valueText, out uint parsedGlobal))
            {
                config.GlobalInterval = parsedGlobal;
                continue;
            }

            if (TryParseAssignment(line, "$globalHCSPARAMSOffset", out valueText)
                && TryParseUInt32Flexible(valueText, out uint parsedHcsparams))
            {
                config.GlobalHcsparamsOffset = parsedHcsparams;
                continue;
            }

            if (TryParseAssignment(line, "$globalRTSOFF", out valueText)
                && TryParseUInt32Flexible(valueText, out uint parsedRtsoff))
            {
                config.GlobalRtsoff = parsedRtsoff;
                continue;
            }

            if (line.StartsWith("$nicItrData", StringComparison.OrdinalIgnoreCase))
            {
                inNicItr = true;
                currentNicItr = null;
                string compact = new(line.Where(ch => !char.IsWhiteSpace(ch)).ToArray());
                if (compact.Contains("@()"))
                {
                    inNicItr = false;
                }
                continue;
            }

            if (line.StartsWith("$userDefinedData", StringComparison.OrdinalIgnoreCase))
            {
                inOverrides = true;
                currentDevice = null;
                string compact = new(line.Where(ch => !char.IsWhiteSpace(ch)).ToArray());
                if (compact.Contains("@{}"))
                {
                    inOverrides = false;
                }
                continue;
            }

            if (!inOverrides)
            {
                continue;
            }

            if (currentDevice is null)
            {
                if (line.StartsWith("}", StringComparison.Ordinal))
                {
                    inOverrides = false;
                    continue;
                }

                if (TryParseQuotedKey(line, out string hwidKey))
                {
                    currentDevice = new ImodConfigEntry { Hwid = hwidKey };
                    config.Overrides.Add(currentDevice);
                }

                continue;
            }

            if (line.StartsWith("}", StringComparison.Ordinal))
            {
                currentDevice = null;
                continue;
            }

            if (!TryParseQuotedAssignment(line, out string keyName, out valueText))
            {
                continue;
            }

            string key = keyName.Trim().ToUpperInvariant();
            if (key == "INTERVALS")
            {
                if (TryParseImodIntervalList(valueText, out List<uint> parsedValues) && parsedValues.Count > 0)
                {
                    currentDevice.Intervals = parsedValues;
                }

                continue;
            }

            if (key == "ROLE_INTERVALS")
            {
                if (TryParseImodRoleIntervals(valueText, out Dictionary<string, uint> roleValues) && roleValues.Count > 0)
                {
                    currentDevice.RoleIntervals = roleValues;
                    currentDevice.AdaptiveRoleBinding = true;
                }

                continue;
            }

            if (key == "ADAPTIVE_ROLE_BINDING")
            {
                if (TryParseBoolFlexible(valueText, out bool adaptiveValue))
                {
                    currentDevice.AdaptiveRoleBinding = adaptiveValue;
                }

                continue;
            }

            if (!TryParseUInt32Flexible(valueText, out uint parsedValue))
            {
                if (key == "ENABLED" && TryParseBoolFlexible(valueText, out bool enabledValue))
                {
                    currentDevice.Enabled = enabledValue;
                }
                continue;
            }

            if (key == "INTERVAL")
            {
                currentDevice.Interval = parsedValue;
            }
            else if (key == "ENABLED")
            {
                currentDevice.Enabled = parsedValue != 0;
            }
            else if (key == "HCSPARAMS_OFFSET" || key == "HCSPARAPS_OFFSET")
            {
                currentDevice.HcsparamsOffset = parsedValue;
            }
            else if (key == "RTSOFF")
            {
                currentDevice.Rtsoff = parsedValue;
            }
        }

        if (config.GlobalInterval == 0)
        {
            config.GlobalInterval = ImodDefaultInterval;
        }

        return config;
    }

    private static bool TryParseAssignment(string line, string key, out string valueText)
    {
        valueText = string.Empty;
        if (!line.StartsWith(key, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        int eqPos = line.IndexOf('=');
        if (eqPos < 0)
        {
            return false;
        }

        valueText = line[(eqPos + 1)..].Trim();
        return valueText.Length > 0;
    }

    private static bool TryParseQuotedKey(string line, out string key)
    {
        key = string.Empty;
        int firstQuote = line.IndexOf('"');
        if (firstQuote < 0)
        {
            return false;
        }

        int secondQuote = line.IndexOf('"', firstQuote + 1);
        if (secondQuote <= firstQuote)
        {
            return false;
        }

        key = line.Substring(firstQuote + 1, secondQuote - firstQuote - 1).Trim();
        return !string.IsNullOrWhiteSpace(key);
    }

    private static bool TryParseQuotedAssignment(string line, out string key, out string valueText)
    {
        key = string.Empty;
        valueText = string.Empty;
        int firstQuote = line.IndexOf('"');
        if (firstQuote < 0)
        {
            return false;
        }

        int secondQuote = line.IndexOf('"', firstQuote + 1);
        if (secondQuote <= firstQuote)
        {
            return false;
        }

        key = line.Substring(firstQuote + 1, secondQuote - firstQuote - 1).Trim();
        int eqPos = line.IndexOf('=', secondQuote + 1);
        if (eqPos < 0)
        {
            return false;
        }

        valueText = line[(eqPos + 1)..].Trim();
        return !string.IsNullOrWhiteSpace(key) && valueText.Length > 0;
    }

    private static string StripInlineComment(string value)
    {
        int hashPos = value.IndexOf('#');
        int semiPos = value.IndexOf(';');
        int cut = -1;
        if (hashPos >= 0)
        {
            cut = hashPos;
        }
        if (semiPos >= 0)
        {
            cut = cut < 0 ? semiPos : Math.Min(cut, semiPos);
        }
        return cut >= 0 ? value[..cut] : value;
    }

    private static bool TryParseUInt32Flexible(string text, out uint value)
    {
        value = 0;
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        string trimmed = text.Trim();
        if (trimmed.StartsWith("0x", StringComparison.OrdinalIgnoreCase))
        {
            return uint.TryParse(trimmed[2..], NumberStyles.HexNumber, CultureInfo.InvariantCulture, out value);
        }

        return uint.TryParse(trimmed, NumberStyles.Integer, CultureInfo.InvariantCulture, out value);
    }

    private static bool TryParseImodInterval(string text, uint fallback, out uint value)
    {
        value = fallback;
        if (string.IsNullOrWhiteSpace(text))
        {
            return true;
        }

        string trimmed = text.Trim();
        if (trimmed.Equals("default", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return TryParseUInt32Flexible(trimmed, out value);
    }

    private static bool TryParseImodIntervalList(string text, out List<uint> values)
    {
        values = [];
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        string cleaned = text.Trim();
        if (cleaned.StartsWith("@(", StringComparison.Ordinal) && cleaned.EndsWith(")", StringComparison.Ordinal))
        {
            cleaned = cleaned[2..^1];
        }
        else if (cleaned.StartsWith("[", StringComparison.Ordinal) && cleaned.EndsWith("]", StringComparison.Ordinal))
        {
            cleaned = cleaned[1..^1];
        }

        string[] parts = cleaned.Split([',', ';', ' ', '\t'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length == 0)
        {
            return false;
        }

        foreach (string part in parts)
        {
            if (!TryParseUInt32Flexible(part, out uint parsed))
            {
                values.Clear();
                return false;
            }

            values.Add(parsed);
        }

        return values.Count > 0;
    }

    private static bool TryParseUInt64List(string text, out List<ulong> values)
    {
        values = [];
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        string cleaned = text.Trim();
        if (cleaned.StartsWith("@(", StringComparison.Ordinal) && cleaned.EndsWith(")", StringComparison.Ordinal))
        {
            cleaned = cleaned[2..^1];
        }
        else if (cleaned.StartsWith("[", StringComparison.Ordinal) && cleaned.EndsWith("]", StringComparison.Ordinal))
        {
            cleaned = cleaned[1..^1];
        }

        string[] parts = cleaned.Split([',', ';', ' ', '\t'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length == 0)
        {
            return false;
        }

        foreach (string part in parts)
        {
            if (!TryParseUInt64Flexible(part, out ulong parsed))
            {
                values.Clear();
                return false;
            }

            values.Add(parsed);
        }

        return values.Count > 0;
    }

    private static bool TryParseImodInput(string text, uint fallback, out ImodInput input)
    {
        input = new ImodInput(fallback, null, null, true);
        if (string.IsNullOrWhiteSpace(text))
        {
            return true;
        }

        string trimmed = text.Trim();
        if (trimmed.Equals("default", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        if (TryParseImodRoleIntervals(trimmed, out Dictionary<string, uint> roleValues) && roleValues.Count > 0)
        {
            input = new ImodInput(null, null, roleValues, false);
            return true;
        }

        if (TryParseImodIntervalList(trimmed, out List<uint> vector))
        {
            if (vector.Count > 1)
            {
                input = new ImodInput(null, vector, null, false);
            }
            else
            {
                uint single = vector[0];
                input = new ImodInput(single, null, null, single == fallback);
            }

            return true;
        }

        if (TryParseUInt32Flexible(trimmed, out uint interval))
        {
            input = new ImodInput(interval, null, null, interval == fallback);
            return true;
        }

        return false;
    }

    private static bool TryParseImodRoleIntervals(string text, out Dictionary<string, uint> roleValues)
    {
        roleValues = new Dictionary<string, uint>(StringComparer.OrdinalIgnoreCase);
        string trimmed = UnquotePowerShellString(text.Trim());
        if (trimmed.Length == 0 || (!trimmed.Contains('=') && !trimmed.Contains(':')))
        {
            return false;
        }

        string[] parts = trimmed.Split([',', ';'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        foreach (string part in parts)
        {
            int eq = part.IndexOf('=');
            int colon = part.IndexOf(':');
            int split = eq >= 0 && colon >= 0 ? Math.Min(eq, colon) : Math.Max(eq, colon);
            if (split <= 0 || split >= part.Length - 1)
            {
                roleValues.Clear();
                return false;
            }

            string rawRole = part[..split].Trim();
            string rawValue = part[(split + 1)..].Trim();
            if (!TryNormalizeImodRoleName(rawRole, out string role)
                || !TryParseImodInterval(rawValue, ImodDefaultInterval, out uint value))
            {
                roleValues.Clear();
                return false;
            }

            roleValues[role] = value & 0xFFFF;
        }

        return roleValues.Count > 0;
    }

    private static bool TryNormalizeImodRoleName(string text, out string role)
    {
        role = string.Empty;
        string compact = new string(text.Where(char.IsLetterOrDigit).ToArray()).ToUpperInvariant();
        role = compact switch
        {
            "MOUSE" => "Mouse",
            "KEYBOARD" => "Keyboard",
            "AUDIO" or "SPEAKER" or "SPEAKERS" or "MIC" or "MICROPHONE" => "Audio",
            "GAMEPAD" or "PAD" or "JOYSTICK" => "Gamepad",
            "WEBCAM" or "CAMERA" => "Webcam",
            _ => string.Empty,
        };

        return role.Length > 0;
    }

    private static string UnquotePowerShellString(string text)
    {
        string value = text.Trim();
        if (value.Length >= 2
            && ((value[0] == '"' && value[^1] == '"') || (value[0] == '\'' && value[^1] == '\'')))
        {
            value = value[1..^1];
        }

        return value.Replace("`\"", "\"", StringComparison.Ordinal)
            .Replace("`'", "'", StringComparison.Ordinal)
            .Trim();
    }

    private static string FormatImodRoleIntervals(IReadOnlyDictionary<string, uint> roleIntervals)
    {
        string[] preferredOrder = ["Mouse", "Keyboard", "Audio", "Gamepad", "Webcam"];
        List<string> parts = [];
        HashSet<string> emitted = new(StringComparer.OrdinalIgnoreCase);

        foreach (string role in preferredOrder)
        {
            if (roleIntervals.TryGetValue(role, out uint value))
            {
                parts.Add($"{role}={FormatImodValue(value)}");
                emitted.Add(role);
            }
        }

        foreach (KeyValuePair<string, uint> pair in roleIntervals.OrderBy(kvp => kvp.Key, StringComparer.OrdinalIgnoreCase))
        {
            if (emitted.Contains(pair.Key))
            {
                continue;
            }

            parts.Add($"{pair.Key}={FormatImodValue(pair.Value)}");
        }

        return string.Join(", ", parts);
    }

    private static bool TryParseBoolFlexible(string text, out bool value)
    {
        value = false;
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        string trimmed = text.Trim();
        if (trimmed.StartsWith("$", StringComparison.Ordinal))
        {
            trimmed = trimmed[1..].Trim();
        }

        if (trimmed.Equals("true", StringComparison.OrdinalIgnoreCase))
        {
            value = true;
            return true;
        }

        if (trimmed.Equals("false", StringComparison.OrdinalIgnoreCase))
        {
            value = false;
            return true;
        }

        if (TryParseUInt32Flexible(trimmed, out uint parsed))
        {
            value = parsed != 0;
            return true;
        }

        return false;
    }

    private static string FormatImodValue(uint value)
    {
        return $"0x{value:X}";
    }

    private static string FormatImodVector(IReadOnlyList<uint> values)
    {
        return string.Join(", ", values.Select(FormatImodValue));
    }

    private static string GetImodOverrideKey(string instanceId)
    {
        if (string.IsNullOrWhiteSpace(instanceId))
        {
            return string.Empty;
        }

        int index = instanceId.IndexOf("DEV_", StringComparison.OrdinalIgnoreCase);
        if (index >= 0 && index + 8 <= instanceId.Length)
        {
            string candidate = instanceId.Substring(index, 8).ToUpperInvariant();
            if (candidate.Length == 8 && candidate[0..4] == "DEV_")
            {
                bool isHex = true;
                foreach (char ch in candidate.AsSpan(4, 4))
                {
                    if (!IsHexDigit(ch))
                    {
                        isHex = false;
                        break;
                    }
                }

                if (isHex)
                {
                    return candidate;
                }
            }
        }

        return instanceId;
    }

    private static bool IsHexDigit(char ch)
    {
        return (ch >= '0' && ch <= '9') || (ch >= 'A' && ch <= 'F') || (ch >= 'a' && ch <= 'f');
    }

    private uint GetEffectiveImodInterval(string instanceId, ImodConfig config)
    {
        uint interval = config.GlobalInterval;
        foreach (ImodConfigEntry entry in config.Overrides)
        {
            if (!string.IsNullOrWhiteSpace(entry.Hwid)
                && instanceId.Contains(entry.Hwid, StringComparison.OrdinalIgnoreCase)
                && (!entry.Enabled.HasValue || entry.Enabled.Value)
                && entry.Interval.HasValue)
            {
                interval = entry.Interval.Value;
            }
        }

        return interval;
    }

    private static ImodConfigEntry? FindImodOverride(string instanceId, ImodConfig config)
    {
        ImodConfigEntry? match = null;
        foreach (ImodConfigEntry entry in config.Overrides)
        {
            if (!string.IsNullOrWhiteSpace(entry.Hwid)
                && instanceId.Contains(entry.Hwid, StringComparison.OrdinalIgnoreCase))
            {
                match = entry;
            }
        }

        return match;
    }

    private static string ComputeEmbeddedImodResourceHash(string exactName, string suffix)
    {
        using Stream? stream = OpenManifestResourceStreamExactOrSuffix(exactName, suffix);
        return stream is null
            ? string.Empty
            : Convert.ToHexString(SHA256.HashData(stream));
    }

    private static bool HasActiveImod(ImodConfig config)
    {
        if (config.HasScript)
        {
            return true;
        }

        if (config.NicItrEntries.Count > 0)
        {
            return true;
        }

        if (config.GlobalInterval != ImodDefaultInterval
            || config.GlobalHcsparamsOffset != ImodDefaultHcsparamsOffset
            || config.GlobalRtsoff != ImodDefaultRtsoff)
        {
            return true;
        }

        foreach (ImodConfigEntry entry in config.Overrides)
        {
            if (string.IsNullOrWhiteSpace(entry.Hwid))
            {
                continue;
            }

            if (entry.Enabled.HasValue && !entry.Enabled.Value)
            {
                continue;
            }

            if (entry.Interval.HasValue
                || entry.Intervals is { Count: > 0 }
                || entry.RoleIntervals is { Count: > 0 }
                || entry.HcsparamsOffset.HasValue
                || entry.Rtsoff.HasValue)
            {
                return true;
            }

            if (entry.Enabled.HasValue && entry.Enabled.Value)
            {
                return true;
            }
        }

        return false;
    }

    private static bool HasCustomImod(ImodConfig config)
    {
        return HasCustomUsbImod(config) || config.NicItrEntries.Count > 0;
    }

    private static bool HasCustomUsbImod(ImodConfig config)
    {
        if (config.GlobalInterval != ImodDefaultInterval
            || config.GlobalHcsparamsOffset != ImodDefaultHcsparamsOffset
            || config.GlobalRtsoff != ImodDefaultRtsoff)
        {
            return true;
        }

        foreach (ImodConfigEntry entry in config.Overrides)
        {
            if (string.IsNullOrWhiteSpace(entry.Hwid))
            {
                continue;
            }

            if (entry.Enabled.HasValue && !entry.Enabled.Value)
            {
                return true;
            }

            if (entry.Interval.HasValue && entry.Interval.Value != ImodDefaultInterval)
            {
                return true;
            }

            if (entry.Intervals is { Count: > 0 } intervals
                && intervals.Any(value => value != ImodDefaultInterval))
            {
                return true;
            }

            if (entry.RoleIntervals is { Count: > 0 } roleIntervals
                && roleIntervals.Values.Any(value => value != ImodDefaultInterval))
            {
                return true;
            }

            if (entry.HcsparamsOffset.HasValue && entry.HcsparamsOffset.Value != ImodDefaultHcsparamsOffset)
            {
                return true;
            }

            if (entry.Rtsoff.HasValue && entry.Rtsoff.Value != ImodDefaultRtsoff)
            {
                return true;
            }
        }

        return false;
    }

    private ImodApplyOutcome ApplyImodSettings(out string? note, out string? technicalDetails)
    {
        note = null;
        technicalDetails = null;
        List<DeviceBlock> xhciBlocks = _blocks
            .Where(b => b.Kind == DeviceKind.USB && b.Device.UsbIsXhci && b.Device.UsbHasDevices && !b.Device.IsTestDevice)
            .ToList();
        if (xhciBlocks.Count == 0)
        {
            return ImodApplyOutcome.SkippedNoUsb;
        }

        List<DeviceBlock> targetBlocks = xhciBlocks.Where(b => IsUsbImodTarget(b.Device)).ToList();

        ImodConfig config = ReadSavedStartupConfig(out _);
        List<string> invalidInputs = [];

        foreach (DeviceBlock block in targetBlocks)
        {
            string instanceId = block.Device.InstanceId;
            config.Overrides.RemoveAll(e =>
                string.Equals(e.Hwid, instanceId, StringComparison.OrdinalIgnoreCase)
                && e.Enabled == false
                && !e.Interval.HasValue
                && e.Intervals is not { Count: > 0 }
                && e.RoleIntervals is not { Count: > 0 }
                && !e.AdaptiveRoleBinding.HasValue
                && !e.HcsparamsOffset.HasValue
                && !e.Rtsoff.HasValue);

            string text = block.ImodBox.Text ?? string.Empty;
            text = text.Trim();
            string hwid = GetImodOverrideKey(block.Device.InstanceId);
            ImodConfigEntry? existing = null;
            foreach (ImodConfigEntry entry in config.Overrides)
            {
                if (string.Equals(entry.Hwid, hwid, StringComparison.OrdinalIgnoreCase))
                {
                    existing = entry;
                }
            }

            if (text.Length == 0)
            {
                config.Overrides.RemoveAll(e =>
                    string.Equals(e.Hwid, hwid, StringComparison.OrdinalIgnoreCase));
                block.ImodBox.Text = FormatImodValue(config.GlobalInterval);
                continue;
            }

            if (!TryParseImodInput(text, config.GlobalInterval, out ImodInput parsedInput))
            {
                string shortPnp = GetShortPnpId(block.Device.InstanceId);
                string label = string.IsNullOrWhiteSpace(shortPnp) ? block.Device.Name : $"{block.Device.Name} ({shortPnp})";
                invalidInputs.Add(label);
                parsedInput = new ImodInput(config.GlobalInterval, null, null, true);
            }

            if (parsedInput.IsDefault)
            {
                config.Overrides.RemoveAll(e =>
                    string.Equals(e.Hwid, hwid, StringComparison.OrdinalIgnoreCase));
                block.ImodBox.Text = FormatImodValue(config.GlobalInterval);
                continue;
            }

            block.ImodBox.Text = parsedInput.RoleIntervals is { Count: > 0 } roleIntervals
                ? FormatImodRoleIntervals(roleIntervals)
                : parsedInput.Intervals is { Count: > 0 } intervals
                    ? FormatImodVector(intervals)
                    : FormatImodValue(parsedInput.Interval ?? config.GlobalInterval);

            if (existing is not null)
            {
                config.Overrides.RemoveAll(e =>
                    !ReferenceEquals(e, existing)
                    && string.Equals(e.Hwid, hwid, StringComparison.OrdinalIgnoreCase));

                existing.Enabled = true;
                existing.Interval = parsedInput.Interval;
                existing.Intervals = parsedInput.Intervals;
                existing.RoleIntervals = parsedInput.RoleIntervals;
                existing.AdaptiveRoleBinding = parsedInput.RoleIntervals is { Count: > 0 };
            }
            else
            {
                config.Overrides.Add(new ImodConfigEntry
                {
                    Hwid = hwid,
                    Enabled = true,
                    Interval = parsedInput.Interval,
                    Intervals = parsedInput.Intervals,
                    RoleIntervals = parsedInput.RoleIntervals,
                    AdaptiveRoleBinding = parsedInput.RoleIntervals is { Count: > 0 },
                });
            }
        }

        foreach (DeviceBlock block in xhciBlocks.Where(b => !IsUsbImodTarget(b.Device)))
        {
            string hwid = block.Device.InstanceId;
            ImodConfigEntry? existing = null;
            foreach (ImodConfigEntry entry in config.Overrides)
            {
                if (string.Equals(entry.Hwid, hwid, StringComparison.OrdinalIgnoreCase))
                {
                    existing = entry;
                }
            }

            if (existing is null)
            {
                existing = new ImodConfigEntry { Hwid = hwid };
                config.Overrides.Add(existing);
            }
            else
            {
                config.Overrides.RemoveAll(e =>
                    !ReferenceEquals(e, existing)
                    && string.Equals(e.Hwid, hwid, StringComparison.OrdinalIgnoreCase));
            }

            existing.Enabled = false;
            existing.Interval = null;
            existing.Intervals = null;
            existing.AdaptiveRoleBinding = null;
            existing.RoleIntervals = null;
            existing.HcsparamsOffset = null;
            existing.Rtsoff = null;
            WriteLog($"IMOD.CONFIG: skip non-hid {block.Device.InstanceId} roles=\"{block.Device.UsbRoles}\"");
        }

        if (invalidInputs.Count > 0)
        {
            string shown = string.Join(", ", invalidInputs.Take(3));
            string suffix = invalidInputs.Count > 3 ? " ..." : string.Empty;
            note = $"{invalidInputs.Count} invalid value(s) were reset to {FormatImodValue(config.GlobalInterval)}.";
            technicalDetails = $"Invalid IMOD values: {shown}{suffix}";
        }

        bool hasCustomUsb = HasCustomUsbImod(config);
        bool hasCustom = hasCustomUsb || config.NicItrEntries.Count > 0;
        bool shouldApplyUsbLive = hasCustomUsb || !hasCustom;
        string startupPath = GetStartupProfilePath();
        ImodApplyStats stats = new();
        if (shouldApplyUsbLive && !TryApplyImod(config, hasCustom, out stats, out string? applyError))
        {
            note = FormatImodUnavailableUserMessage(applyError, includeNotChanged: true);
            technicalDetails = applyError ?? "Unknown IMOD driver error.";
            WriteLog($"IMOD: failed: {technicalDetails}");
            return ImodApplyOutcome.Failed;
        }
        if (!shouldApplyUsbLive)
        {
            WriteLog("IMOD: USB live apply skipped (no custom USB IMOD config; NIC ITR startup config preserved)");
        }

        if (hasCustom)
        {
            try
            {
                SaveStartupProfile(config);
                WriteLog($"IMOD.CONFIG: startup profile saved {startupPath}");
            }
            catch (Exception ex)
            {
                note = "USB IMOD is active for this session, but startup persistence was not saved.";
                technicalDetails = ex.ToString();
                WriteLog($"IMOD.CONFIG: write failed {startupPath}: {FlattenLogText(ex.ToString())}");
                return ImodApplyOutcome.AppliedWithWarnings;
            }
        }
        else
        {
            RemoveImodPersistenceFiles();
            config.HasScript = false;
        }

        _imodConfigCache = config;
        _imodConfigLoaded = true;

        if (!shouldApplyUsbLive)
        {
            note = "IMOD startup script saved. USB live apply skipped (no custom USB IMOD config).";
            WriteLog($"IMOD: {note}");
            return ImodApplyOutcome.Applied;
        }

        if (stats.ControllersFound == 0)
        {
            note = "IMOD skipped (no XHCI controllers found).";
            WriteLog($"IMOD: {note}");
            return ImodApplyOutcome.SkippedNoController;
        }

        if (stats.ControllersApplied == 0)
        {
            if (stats.ControllersPartiallyApplied > 0)
            {
                note = $"USB IMOD was only partially written ({stats.WritesSucceeded}/{stats.WritesAttempted} writes).";
                technicalDetails = $"controllersPartial={stats.ControllersPartiallyApplied}; controllersFailed={stats.ControllersFailed}; readFailures={stats.ReadFailures}; writeFailures={stats.WriteFailures}";
                WriteLog($"IMOD: {note} {technicalDetails}");
                return ImodApplyOutcome.AppliedWithWarnings;
            }

            if (stats.WriteFailures > 0 || stats.ControllersFailed > 0 || stats.ReadFailures > 0 || stats.WritesAttempted > 0)
            {
                note = "USB IMOD was not applied.";
                technicalDetails = $"controllersFailed={stats.ControllersFailed}; readFailures={stats.ReadFailures}; writeFailures={stats.WriteFailures}; writesAttempted={stats.WritesAttempted}";
                WriteLog($"IMOD: {note} {technicalDetails}");
                return ImodApplyOutcome.Failed;
            }

            note = "IMOD skipped (no eligible USB controllers).";
            WriteLog($"IMOD: {note}");
            return ImodApplyOutcome.SkippedNoController;
        }

        if (stats.ReadFailures > 0 || stats.WriteFailures > 0 || stats.ControllersPartiallyApplied > 0 || stats.ControllersFailed > 0)
        {
            note = $"USB IMOD was applied with warnings ({stats.WritesSucceeded}/{stats.WritesAttempted} writes).";
            technicalDetails = $"controllersApplied={stats.ControllersApplied}; controllersPartial={stats.ControllersPartiallyApplied}; controllersFailed={stats.ControllersFailed}; readFailures={stats.ReadFailures}; writeFailures={stats.WriteFailures}";
            return ImodApplyOutcome.AppliedWithWarnings;
        }
        else
        {
            note = $"IMOD applied to {stats.ControllersApplied} USB controller(s).";
        }

        return ImodApplyOutcome.Applied;
    }

    private static void AddImodResultToReport(
        OperationReport report,
        ImodApplyOutcome outcome,
        string? userMessage,
        string? technicalDetails,
        string component = "USB IMOD")
    {
        switch (outcome)
        {
            case ImodApplyOutcome.Failed:
                report.AddError(component, userMessage ?? "USB IMOD was not applied.", technicalDetails);
                break;
            case ImodApplyOutcome.AppliedWithWarnings:
                report.AddWarning(component, userMessage ?? "USB IMOD was applied with warnings.", technicalDetails);
                break;
            case ImodApplyOutcome.SkippedNoUsb:
            case ImodApplyOutcome.SkippedNoController:
            case ImodApplyOutcome.SkippedNoConfig:
                report.AddSkipped(component, userMessage ?? "No eligible target was found.", technicalDetails);
                break;
            default:
                report.AddSuccess(component, userMessage ?? "Applied");
                break;
        }
    }
}

