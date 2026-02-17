using Microsoft.Win32;
using System.Text.RegularExpressions;

namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private HashSet<string> GetDisabledAudioEndpointNames()
    {
        HashSet<string> disabled = new(StringComparer.OrdinalIgnoreCase);

        string[] roots =
        [
            @"SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render",
            @"SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Capture",
        ];

        foreach (string root in roots)
        {
            try
            {
                using RegistryKey? rk = Registry.LocalMachine.OpenSubKey(root);
                if (rk is null)
                {
                    continue;
                }

                foreach (string sub in rk.GetSubKeyNames())
                {
                    try
                    {
                        using RegistryKey? devKey = rk.OpenSubKey(sub);
                        if (devKey is null)
                        {
                            continue;
                        }

                        object? stateObj = devKey.GetValue("DeviceState");
                        if (stateObj is not int state)
                        {
                            continue;
                        }

                        bool isDisabled = (state & 0x2) != 0 || (state & 0x4) != 0;
                        if (!isDisabled)
                        {
                            continue;
                        }

                        using RegistryKey? propsKey = devKey.OpenSubKey("Properties");
                        if (propsKey is null)
                        {
                            continue;
                        }

                        string fnKey = "{a45c254e-df1c-4efd-8020-67d146a850e0},2";
                        string? friendly = propsKey.GetValue(fnKey) as string;
                        if (!string.IsNullOrWhiteSpace(friendly))
                        {
                            disabled.Add(friendly.ToLowerInvariant());
                        }
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
            }
        }

        return disabled;
    }

    private string? FindAudioControllerFor(string childId, Dictionary<string, WmiPnPDevice> deviceLookup)
    {
        string curr = childId;
        for (int i = 0; i < 25; i++)
        {
            string? parent = GetParentId(curr);
            if (string.IsNullOrWhiteSpace(parent))
            {
                break;
            }

            string parentKey = parent.ToUpperInvariant();
            WmiPnPDevice? dev = deviceLookup.TryGetValue(parentKey, out WmiPnPDevice? cached) ? cached : WmiInterop.TryGetPnPDeviceById(parentKey);
            if (dev is not null)
            {
                string n = dev.Name ?? parentKey;
                bool isAudio =
                    Regex.IsMatch(n, "(?i)audio|sound")
                    || string.Equals(dev.Class, "MEDIA", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(dev.Class, "Audio", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(dev.Class, "AudioEndpoint", StringComparison.OrdinalIgnoreCase);

                if (parentKey.StartsWith("PCI\\", StringComparison.OrdinalIgnoreCase) && isAudio)
                {
                    return parentKey;
                }
            }

            curr = parentKey;
        }

        return null;
    }

    private Dictionary<string, List<string>> AudioControllerEndpoints(List<WmiPnPDevice> deviceCache, Dictionary<string, WmiPnPDevice> deviceLookup)
    {
        Dictionary<string, List<string>> map = new(StringComparer.OrdinalIgnoreCase);
        HashSet<string> disabledNames = GetDisabledAudioEndpointNames();

        IEnumerable<WmiPnPDevice> endpoints = deviceCache.Where(d => string.Equals(d.Class, "AudioEndpoint", StringComparison.OrdinalIgnoreCase));
        foreach (WmiPnPDevice d in endpoints)
        {
            bool present = string.Equals(d.Status, "OK", StringComparison.OrdinalIgnoreCase) && d.ConfigManagerErrorCode != 22;
            if (!present)
            {
                continue;
            }

            if (string.IsNullOrWhiteSpace(d.InstanceId))
            {
                continue;
            }

            string name = !string.IsNullOrWhiteSpace(d.Name) ? d.Name : d.InstanceId;
            string? ctrl = FindAudioControllerFor(d.InstanceId, deviceLookup);
            if (string.IsNullOrWhiteSpace(ctrl))
            {
                continue;
            }

            if (disabledNames.Contains(name.ToLowerInvariant()))
            {
                continue;
            }

            if (!map.TryGetValue(ctrl, out List<string>? list))
            {
                list = [];
                map[ctrl] = list;
            }

            if (!list.Contains(name, StringComparer.OrdinalIgnoreCase))
            {
                list.Add(name);
            }
        }

        foreach (string k in map.Keys)
        {
            WriteLog($"AUDIO.MAP: {k} -> [{string.Join("; ", map[k])}]");
        }

        return map;
    }

    private static string? GetAudioEndpointRole(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            return null;
        }

        string n = name.ToLowerInvariant();
        if (Regex.IsMatch(n, "(?i)hdmi|display\\s*audio|monitor|монитор|displayport|\\bdp\\b"))
        {
            return "HDMI Audio";
        }

        if (Regex.IsMatch(n, "(?i)speakers?|speaker|динамик|динамики|колонки|колонка"))
        {
            return "Speakers";
        }

        if (Regex.IsMatch(n, "(?i)headphones?|headset|наушники|наушник|гарнитура"))
        {
            return "Headphones";
        }

        if (Regex.IsMatch(n, "(?i)microphone| mic\\b|микрофон|микроф\\b"))
        {
            return "Microphone";
        }

        if (Regex.IsMatch(n, "(?i)line in|line\\-in"))
        {
            return "Line-in";
        }

        if (Regex.IsMatch(n, "(?i)line out|line\\-out"))
        {
            return "Line-out";
        }

        return null;
    }

    private static string CleanAudioEndpointName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            return string.Empty;
        }

        string n = name;
        n = Regex.Replace(n, "\\s*\\(\\s*(?:\\d+\\s*[-:]\\s*)?(high definition audio device|audio device)\\s*\\)$", string.Empty, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        n = Regex.Replace(n, "\\s*[-:]\\s*(?:\\d+\\s*[-:]\\s*)?(high definition audio device|audio device)$", string.Empty, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        return n.Trim();
    }

    private string FormatAudioEndpointsSummary(IEnumerable<string> names)
    {
        List<string> clean = names
            .Where(n => !string.IsNullOrWhiteSpace(n))
            .Select(CleanAudioEndpointName)
            .Where(n => !string.IsNullOrWhiteSpace(n))
            .ToList();

        if (clean.Count == 0)
        {
            return string.Empty;
        }

        List<string> roles = [];
        foreach (string n in clean)
        {
            string? role = GetAudioEndpointRole(n);
            if (!string.IsNullOrWhiteSpace(role))
            {
                roles.Add(role);
            }
        }

        if (roles.Count == 0)
        {
            string primary = clean[0];
            if (Regex.IsMatch(primary, "(?i)^high definition audio device$"))
            {
                return "HDMI AUDIO";
            }

            if (Regex.IsMatch(primary, "(?i)^nvidia high definition audio$"))
            {
                return "HDMI AUDIO";
            }

            if (Regex.IsMatch(primary, "(?i)display\\s*audio|hdmi|displayport|\\bdp\\b|monitor|монитор"))
            {
                return "HDMI AUDIO";
            }

            if (clean.Count == 1 && Regex.IsMatch(primary, "^[A-Za-z]{2,}\\d{2,}"))
            {
                return $"HDMI AUDIO {primary}";
            }

            return primary;
        }

        if (roles.Contains("Microphone", StringComparer.OrdinalIgnoreCase))
        {
            return "Microphone";
        }

        int hdmiCount = roles.Count(r => r == "HDMI Audio");
        List<string> result = [];
        int hdmiIndex = 0;
        foreach (string role in roles)
        {
            if (role == "HDMI Audio" && hdmiCount > 1)
            {
                hdmiIndex++;
                string label = $"{role} #{hdmiIndex}";
                if (!result.Contains(label, StringComparer.OrdinalIgnoreCase))
                {
                    result.Add(label);
                }
            }
            else
            {
                if (!result.Contains(role, StringComparer.OrdinalIgnoreCase))
                {
                    result.Add(role);
                }
            }
        }

        return string.Join(", ", result);
    }

    private static string DetectDisplayTransport(IEnumerable<string> names, string deviceName)
    {
        List<string> candidates = [];
        candidates.AddRange(names.Where(n => !string.IsNullOrWhiteSpace(n)));
        if (!string.IsNullOrWhiteSpace(deviceName))
        {
            candidates.Add(deviceName);
        }

        foreach (string n in candidates)
        {
            string ln = n.ToLowerInvariant();
            if (Regex.IsMatch(ln, "(?i)displayport|display port|\\bdp\\b|dp audio"))
            {
                return "DP";
            }

            if (Regex.IsMatch(ln, "(?i)hdmi"))
            {
                return "HDMI";
            }
        }

        return "HDMI/DP";
    }
}
