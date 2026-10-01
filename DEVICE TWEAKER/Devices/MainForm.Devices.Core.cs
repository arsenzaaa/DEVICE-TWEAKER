using System.Text.RegularExpressions;

namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private string NormalizeInstanceId(string? id)
    {
        if (string.IsNullOrWhiteSpace(id))
        {
            return string.Empty;
        }

        return id.Replace("\\\\", "\\").Trim().ToUpperInvariant();
    }

    private string? GetParentId(string id)
    {
        return NativeCfgMgr32.TryGetParentInstanceId(id, out string? parent) ? parent : null;
    }

    private static bool IsWiFiDevice(string pnpid, string name, string service)
    {
        if (Regex.IsMatch(name, "(?i)Wi-?Fi|Wireless|802\\.11|WLAN"))
        {
            return true;
        }

        if (Regex.IsMatch(pnpid, "(?i)\\\\VEN_14C3\\\\"))
        {
            return true;
        }

        if (Regex.IsMatch(service, "(?i)netwtw|rtwl|rtwlane|rtwlanu|athw|athw10|ath|bcmwl|qcwlan|qwlan|iwl|iwlwifi|mtk"))
        {
            return true;
        }

        return false;
    }

    private static List<string> GetUsbControllerKeys(string id, Func<string, string> normalize)
    {
        List<string> keys = [];
        string norm = normalize(id);
        if (string.IsNullOrWhiteSpace(norm))
        {
            return keys;
        }

        keys.Add(norm);
        string[] parts = norm.Split('\\');
        if (parts.Length >= 2)
        {
            string root = $"{parts[0]}\\{parts[1]}";
            if (!string.IsNullOrWhiteSpace(root) && !keys.Contains(root, StringComparer.OrdinalIgnoreCase))
            {
                keys.Add(root);
            }
        }

        return keys;
    }

    private static bool IsUsbInfrastructureDevice(string instanceId)
    {
        if (string.IsNullOrWhiteSpace(instanceId))
        {
            return false;
        }

        return instanceId.Contains("ROOT_HUB", StringComparison.OrdinalIgnoreCase);
    }

    private static string GetShortPnpId(string id)
    {
        if (string.IsNullOrWhiteSpace(id))
        {
            return string.Empty;
        }

        Match m = Regex.Match(id, "^(?<bus>PCI)\\\\VEN_(?<ven>[0-9A-Fa-f]{4})&DEV_(?<dev>[0-9A-Fa-f]{4})", RegexOptions.CultureInvariant);
        if (m.Success)
        {
            return $"{m.Groups["bus"].Value.ToUpperInvariant()}_VEN_{m.Groups["ven"].Value.ToUpperInvariant()}_DEV_{m.Groups["dev"].Value.ToUpperInvariant()}";
        }

        string[] parts = id.Split('\\');
        return parts.Length > 0 ? parts[0] : id;
    }

    private static string GetDisplayRegPath(string instanceId)
    {
        const string root = @"HKLM\SYSTEM\CurrentControlSet\Enum";
        if (string.IsNullOrWhiteSpace(instanceId))
        {
            return root;
        }

        Match m = Regex.Match(instanceId, "^(?<bus>[^\\\\]+)\\\\VEN_(?<ven>[0-9A-Fa-f]{4})&DEV_(?<dev>[0-9A-Fa-f]{4})", RegexOptions.CultureInvariant);
        if (m.Success)
        {
            return $"{root}\\{m.Groups["bus"].Value.ToUpperInvariant()}\\VEN_{m.Groups["ven"].Value.ToUpperInvariant()}&DEV_{m.Groups["dev"].Value.ToUpperInvariant()}";
        }

        string[] parts = instanceId.Split('\\');
        if (parts.Length >= 2)
        {
            return $"{root}\\{parts[0]}\\{parts[1]}";
        }

        return $"{root}\\{instanceId}";
    }

    private static string GetFullRegPath(string regBase)
    {
        if (string.IsNullOrWhiteSpace(regBase))
        {
            return string.Empty;
        }

        string p = regBase;
        p = Regex.Replace(p, "^Microsoft\\.PowerShell\\.Core\\\\Registry::", string.Empty, RegexOptions.CultureInvariant);
        p = Regex.Replace(p, "^HKLM:", "HKLM", RegexOptions.CultureInvariant);
        return p;
    }

    private static bool IsDisplayHdmiaudio(string pnpid, string desc)
    {
        if (Regex.IsMatch(pnpid, "(?i)\\\\VEN_10DE&(DEV_1A[0-9A-F]{2}|DEV_22[0-9A-F]{2}|DEV_26[0-9A-F]{2}|DEV_28[0-9A-F]{2})"))
        {
            return true;
        }

        if (Regex.IsMatch(pnpid, "(?i)\\\\VEN_1002&DEV_AA[0-9A-F]{2}"))
        {
            return true;
        }

        if (Regex.IsMatch(pnpid, "(?i)\\\\VEN_8086&DEV_28[0-9A-F]{2}"))
        {
            return true;
        }

        if (Regex.IsMatch(desc, "(?i)Display Audio|HDMI Audio|NVIDIA High Definition Audio|AMD High Definition Audio"))
        {
            return true;
        }

        return false;
    }

    private static bool IsDisplayAudioEndpointsText(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        string t = text.Trim();
        if (t == "[UNKNOWN]")
        {
            return false;
        }

        string lower = t.ToLowerInvariant();
        if (Regex.IsMatch(lower, "(?i)hdmi|display audio|monitor|displayport|digital audio|dp\\b"))
        {
            return true;
        }

        if (Regex.IsMatch(t, "^[A-Z0-9]{3,}$", RegexOptions.CultureInvariant) && Regex.IsMatch(t, "\\d", RegexOptions.CultureInvariant))
        {
            return true;
        }

        return false;
    }

    private static bool IsDisplayAudioDevice(string pnpid, string name, string audioText)
    {
        _ = audioText;
        return IsDisplayHdmiaudio(pnpid, name);
    }

    private static string CleanDeviceDisplayName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            return string.Empty;
        }

        string n = name;
        const string vendor = "(?:microsoft|майкрософт)";
        n = Regex.Replace(n, $"\\s*\\(\\s*{vendor}\\s*\\)", string.Empty, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        n = Regex.Replace(n, $"\\s*\\-\\s*{vendor}\\b", string.Empty, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        n = Regex.Replace(n, $"\\b{vendor}\\b", string.Empty, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        n = Regex.Replace(n, "\\s{2,}", " ", RegexOptions.CultureInvariant);
        return n.Trim();
    }
}
