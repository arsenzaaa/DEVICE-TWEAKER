using System.Text.RegularExpressions;

namespace DeviceTweakerCS;

internal static class ImodStartupUpgrade
{
    internal static bool UpdateManaged(string original, string upgraded, Action preparePayload, Action saveOriginal, Action<string> write)
    {
        // A current script may still have missing/stale payload files: prepare them even when text matches.
        preparePayload();
        if (string.Equals(original, upgraded, StringComparison.Ordinal)) return false;
        saveOriginal();
        write(upgraded);
        return true;
    }

    internal static string Upgrade(string text, string currentTemplate, string driverHash, string? kduHash = null, string? databaseHash = null,
        string? driverPath = null, string? kduPath = null, string? databasePath = null)
    {
        bool legacy = text.Contains("private const string ImodDriverServiceName = \"DeviceTweakerImod2\"", StringComparison.Ordinal);
        Match version = Regex.Match(text, @"(?m)^\s*\$imodScriptVersion\s*=\s*(\d+)\s*$");
        if (!legacy && !version.Success) return text;
        if (version.Success && int.Parse(version.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture) > 32)
            throw new InvalidDataException("The saved IMOD script is newer than this application; it was retained.");
        const string begin = "$imodSettingsBegin = $true", end = "$imodSettingsEnd = $true";
        int start = text.IndexOf(begin, StringComparison.Ordinal), finish = text.IndexOf(end, StringComparison.Ordinal);
        if (start < 0 || finish < start || text.LastIndexOf(begin, StringComparison.Ordinal) != start
            || text.LastIndexOf(end, StringComparison.Ordinal) != finish)
            throw new InvalidDataException("Could not locate the saved IMOD / NIC ITR configuration for migration.");
        if (!Regex.IsMatch(driverHash, "\\A[0-9A-Fa-f]{64}\\z")) throw new InvalidDataException("Missing current DTIMOD hash.");
        const string token = "{{IMOD_CONFIG_BLOCK}}";
        if (currentTemplate.IndexOf(token, StringComparison.Ordinal) < 0
            || currentTemplate.IndexOf(token, StringComparison.Ordinal) != currentTemplate.LastIndexOf(token, StringComparison.Ordinal))
            throw new InvalidDataException("Invalid current startup template.");
        string config = text[start..(finish + end.Length)];
        Regex hashLine = new(@"(?m)^[ \t]*\$ImodDriverSha256[ \t]*=[^\r\n]*");
        if (hashLine.Matches(config).Count != 1) throw new InvalidDataException("Missing or duplicated saved DTIMOD hash.");
        config = hashLine.Replace(config, _ => "$ImodDriverSha256 = '" + driverHash + "'");
        foreach ((string name, string? hash) in new[] { ("ImodKduSha256", kduHash), ("ImodKduDbSha256", databaseHash) })
        {
            if (hash is null) continue;
            if (!Regex.IsMatch(hash, "\\A[0-9A-Fa-f]{64}\\z")) throw new InvalidDataException("Invalid loader payload hash.");
            Regex line = new(@"(?m)^[ \t]*\$" + name + @"[ \t]*=[^\r\n]*");
            if (line.Matches(config).Count != 1) throw new InvalidDataException("Missing or duplicated saved " + name + ".");
            config = line.Replace(config, _ => "$" + name + " = '" + hash + "'");
        }
        foreach ((string name, string? path) in new[] { ("ImodDriverPath", driverPath), ("ImodKduPath", kduPath), ("ImodKduDbPath", databasePath) })
        {
            if (path is null) continue;
            Regex line = new(@"(?m)^[ \t]*\$" + name + @"[ \t]*=[^\r\n]*");
            if (line.Matches(config).Count > 1) throw new InvalidDataException("Duplicated saved " + name + ".");
            string replacement = "$" + name + " = '" + path.Replace("'", "''", StringComparison.Ordinal) + "'";
            config = line.IsMatch(config) ? line.Replace(config, _ => replacement) : config.Replace(end, replacement + "\n" + end, StringComparison.Ordinal);
        }
        return currentTemplate.Replace("{{IMOD_CONFIG_BLOCK}}", config, StringComparison.Ordinal);
    }
}
