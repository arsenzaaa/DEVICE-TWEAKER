using System.Text.RegularExpressions;

using System.Globalization;

namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private Dictionary<string, List<string>> GetHidControllerRoleMap(Dictionary<string, WmiPnPDevice> deviceLookup, List<(string ControllerId, string DependentId)> usbPairs)
    {
        Dictionary<string, List<string>> map = new(StringComparer.OrdinalIgnoreCase);
        WriteLog("USBROLE.HID: start HID scan");

        foreach (string devicePath in HidInterop.EnumerateHidDevicePaths())
        {
            _ = HidInterop.TryReadProductAndUsage(devicePath, out string product, out int? usagePage, out int? usageId);
            string? productRole = GetHidProductType(product);

            string? inst = HidInterop.TryParseInstanceIdFromDevicePath(devicePath);
            inst = NormalizeInstanceId(inst);

            WmiPnPDevice? pnpDev = null;
            if (!string.IsNullOrWhiteSpace(inst))
            {
                deviceLookup.TryGetValue(inst, out pnpDev);
            }

            string? usageRole = GetHidUsageRole(usagePage, usageId);
            string? pnpRole = GetPnpHidRole(pnpDev);
            string? role = ResolveHidRole(product, inst, productRole, usageRole, pnpRole);

            if (string.IsNullOrWhiteSpace(inst) || string.IsNullOrWhiteSpace(role))
            {
                WriteLog($"USBROLE.HID: skip no-inst-or-role path={devicePath} prod=\"{product}\" role={role}");
                continue;
            }

            if (deviceLookup.Count > 0 && !deviceLookup.ContainsKey(inst))
            {
                WriteLog($"USBROLE.HID: skip absent in DeviceLookup inst={inst} prod=\"{product}\" role={role}");
            }

            foreach ((string controllerId, string dependentId) in usbPairs)
            {
                if (!dependentId.Contains(inst, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                foreach (string key in GetUsbControllerKeys(controllerId, NormalizeInstanceId))
                {
                    if (!map.TryGetValue(key, out List<string>? roles))
                    {
                        roles = [];
                        map[key] = roles;
                    }

                    if (!roles.Contains(role, StringComparer.OrdinalIgnoreCase))
                    {
                        roles.Add(role);
                        WriteLog($"USBROLE.HID: mapped controller {key} -> {role} (inst={inst} usagePage={usagePage} usage={usageId} prod=\"{product}\")");
                    }
                }
            }
        }

        WriteLog($"USBROLE.HID: end HID scan, controllers mapped={map.Keys.Count}");
        return map;
    }

    private List<UsbControllerInfo> FindUsbControllersByVidPid(string id, List<(string ControllerId, string DependentId)> usbPairs)
    {
        List<UsbControllerInfo> list = [];
        if (string.IsNullOrWhiteSpace(id))
        {
            return list;
        }

        Match m = Regex.Match(id, "VID_[0-9A-Fa-f]{4}&PID_[0-9A-Fa-f]{4}", RegexOptions.CultureInvariant);
        if (!m.Success)
        {
            return list;
        }

        string needle = m.Value.ToUpperInvariant();
        HashSet<string> seen = new(StringComparer.OrdinalIgnoreCase);
        foreach ((string controllerId, string dependentId) in usbPairs)
        {
            if (!dependentId.Contains(needle, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            string ctrlId = NormalizeInstanceId(controllerId);
            if (string.IsNullOrWhiteSpace(ctrlId))
            {
                continue;
            }

            if (seen.Add(ctrlId))
            {
                list.Add(new UsbControllerInfo(ctrlId, ctrlId));
            }
        }

        return list;
    }

    private Dictionary<string, UsbPollingRateInfo> BuildUsbPollingRateLookup()
    {
        Dictionary<string, UsbPollingRateInfo> lookup = new(StringComparer.OrdinalIgnoreCase);
        List<UsbEndpointInfo> endpoints;
        try
        {
            endpoints = UsbTopologyInterop.EnumerateEndpoints();
        }
        catch (Exception ex)
        {
            WriteLog($"USBPOLL: endpoint enumeration failed: {ex.Message}");
            return lookup;
        }

        WriteLog($"USBPOLL: endpoints={endpoints.Count}");

        foreach (UsbEndpointInfo endpoint in endpoints)
        {
            if (!IsPollingEndpointCandidate(endpoint)
                || !TryCalculatePollingRate(endpoint.Speed, endpoint.BInterval, out double hertz))
            {
                continue;
            }

            string vidPid = $"{endpoint.VendorId}:{endpoint.ProductId}".ToUpperInvariant();
            UsbPollingRateInfo info = new(
                vidPid,
                hertz,
                endpoint.Speed,
                endpoint.BInterval,
                FormatPollingRateTag(hertz));

            SetBestPollingRate(lookup, vidPid, info);

            if (endpoint.InterfaceNumber >= 0)
            {
                SetBestPollingRate(lookup, $"{vidPid}:MI_{endpoint.InterfaceNumber:X2}", info);
            }

            string? endpointRole = GetUsbEndpointRole(endpoint);
            if (!string.IsNullOrWhiteSpace(endpointRole))
            {
                SetBestPollingRate(lookup, $"{vidPid}:{endpointRole}", info);
            }

            WriteLog(
                $"USBPOLL.ENDPOINT: {vidPid} iface={endpoint.InterfaceNumber} class={endpoint.InterfaceClass} protocol={endpoint.InterfaceProtocol} dir={endpoint.Direction} type={endpoint.TransferType} speed={endpoint.Speed} bInterval={endpoint.BInterval} endpointMax={info.Tag}");
        }

        foreach (KeyValuePair<string, UsbPollingRateInfo> kvp in lookup
            .Where(k => !k.Key.EndsWith(":Keyboard", StringComparison.OrdinalIgnoreCase)
                && !k.Key.EndsWith(":Mouse", StringComparison.OrdinalIgnoreCase)
                && !Regex.IsMatch(k.Key, ":MI_[0-9A-F]{2}$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
            .OrderBy(k => k.Key, StringComparer.OrdinalIgnoreCase))
        {
            UsbPollingRateInfo info = kvp.Value;
            WriteLog($"USBPOLL: {kvp.Key} endpointMax={info.Tag} speed={info.Speed} bInterval={info.BInterval} hz={info.Hertz.ToString("0.###", CultureInfo.InvariantCulture)}");
        }

        return lookup;
    }

    private static bool IsPollingEndpointCandidate(UsbEndpointInfo endpoint)
    {
        return !string.IsNullOrWhiteSpace(endpoint.VendorId)
            && !string.IsNullOrWhiteSpace(endpoint.ProductId)
            && string.Equals(endpoint.TransferType, "Interrupt", StringComparison.OrdinalIgnoreCase)
            && string.Equals(endpoint.Direction, "IN", StringComparison.OrdinalIgnoreCase)
            && endpoint.AlternateSetting == 0
            && endpoint.BInterval > 0;
    }

    private static void SetBestPollingRate(
        Dictionary<string, UsbPollingRateInfo> lookup,
        string key,
        UsbPollingRateInfo candidate)
    {
        if (!lookup.TryGetValue(key, out UsbPollingRateInfo? existing)
            || candidate.Hertz > existing.Hertz)
        {
            lookup[key] = candidate;
        }
    }

    private static string? GetUsbEndpointRole(UsbEndpointInfo endpoint)
    {
        if (!string.Equals(endpoint.InterfaceClass, "0x03", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        return endpoint.InterfaceProtocol.ToUpperInvariant() switch
        {
            "0X01" => "Keyboard",
            "0X02" => "Mouse",
            _ => null,
        };
    }

    private static bool TryCalculatePollingRate(string speed, int bInterval, out double hertz)
    {
        hertz = 0;
        if (bInterval <= 0)
        {
            return false;
        }

        if (speed.StartsWith("High", StringComparison.OrdinalIgnoreCase)
            || speed.StartsWith("Super", StringComparison.OrdinalIgnoreCase))
        {
            if (bInterval > 16)
            {
                return false;
            }

            double microframes = Math.Pow(2, bInterval - 1);
            hertz = 8000d / microframes;
            return hertz > 0;
        }

        hertz = 1000d / bInterval;
        return hertz > 0;
    }

    private static string FormatPollingRateTag(double hertz)
    {
        if (hertz >= 1000d)
        {
            double khz = hertz / 1000d;
            double rounded = Math.Round(khz);
            return Math.Abs(khz - rounded) < 0.05d
                ? $"{rounded.ToString("0", CultureInfo.InvariantCulture)}K"
                : $"{khz.ToString("0.#", CultureInfo.InvariantCulture)}K";
        }

        return $"{hertz.ToString(hertz >= 100d ? "0" : "0.#", CultureInfo.InvariantCulture)}Hz";
    }

    private static bool TryGetVidPidKey(string? instanceId, out string vidPid)
    {
        vidPid = string.Empty;
        if (string.IsNullOrWhiteSpace(instanceId))
        {
            return false;
        }

        Match match = Regex.Match(
            instanceId,
            "VID_([0-9A-Fa-f]{4})&PID_([0-9A-Fa-f]{4})",
            RegexOptions.CultureInvariant);

        if (!match.Success)
        {
            return false;
        }

        vidPid = $"{match.Groups[1].Value}:{match.Groups[2].Value}".ToUpperInvariant();
        return true;
    }

    private string AddPollingRateTag(
        string role,
        string? instanceId,
        IReadOnlyDictionary<string, UsbPollingRateInfo> pollingRateLookup)
    {
        bool isMouse = IsUsbRoleText(role, "Mouse");
        bool isKeyboard = IsUsbRoleText(role, "Keyboard");

        if ((!isMouse && !isKeyboard)
            || !TryGetVidPidKey(instanceId, out string vidPid))
        {
            return role;
        }

        if (isMouse)
        {
            return $"{role} scanning";
        }

        if (pollingRateLookup.Count == 0)
        {
            return role;
        }

        UsbPollingRateInfo? info = null;
        if (TryGetMiKey(instanceId, out string miKey))
        {
            _ = pollingRateLookup.TryGetValue($"{vidPid}:{miKey}", out info);
        }

        if (info is null
            && !pollingRateLookup.TryGetValue($"{vidPid}:{role}", out info)
            && !pollingRateLookup.TryGetValue(vidPid, out info))
        {
            return role;
        }

        string tag = FormatStaticPollingTag(role, info.Tag);
        return string.IsNullOrWhiteSpace(tag) ? role : $"{role} {tag}";
    }

    private static bool TryGetMiKey(string? instanceId, out string miKey)
    {
        miKey = string.Empty;
        if (string.IsNullOrWhiteSpace(instanceId))
        {
            return false;
        }

        Match match = Regex.Match(instanceId, "&MI_([0-9A-Fa-f]{2})", RegexOptions.CultureInvariant);
        if (!match.Success)
        {
            return false;
        }

        miKey = "MI_" + match.Groups[1].Value.ToUpperInvariant();
        return true;
    }

    private static string FormatStaticPollingTag(string role, string tag)
    {
        return IsUsbRoleText(role, "Mouse") ? "scanning" : tag;
    }

    private static string FormatUsbPollingRoleSummary(IEnumerable<string> roles)
    {
        List<string> items = [];
        foreach (string role in roles)
        {
            Match match = Regex.Match(
                role.Trim(),
                "^(Mouse|Keyboard)\\s+(.+)$",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

            if (match.Success)
            {
                items.Add($"{match.Groups[1].Value}: {match.Groups[2].Value}");
            }
        }

        return string.Join(", ", items.Distinct(StringComparer.OrdinalIgnoreCase));
    }

    private static bool IsUsbRoleText(string text, string role)
    {
        string trimmed = text.Trim();
        return string.Equals(trimmed, role, StringComparison.OrdinalIgnoreCase)
            || trimmed.StartsWith(role + " ", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsKnownInternalVendorHid(string productName)
    {
        if (string.IsNullOrWhiteSpace(productName))
        {
            return false;
        }

        return Regex.IsMatch(
            productName,
            "(?i)^ite\\s+device\\(\\d+\\)$",
            RegexOptions.CultureInvariant);
    }

    private static bool IsCollectionLevelHid(string? instanceId)
    {
        if (string.IsNullOrWhiteSpace(instanceId))
        {
            return false;
        }

        return Regex.IsMatch(instanceId, "(?i)&COL\\d+", RegexOptions.CultureInvariant);
    }

    private string? GetHidProductType(string productName)
    {
        if (string.IsNullOrWhiteSpace(productName))
        {
            return null;
        }

        if (IsKnownInternalVendorHid(productName))
        {
            return null;
        }

        string name = productName.ToLowerInvariant();
        if (Regex.IsMatch(name, "(?i)headset|headphone|earbud|earphone|microphone|\\bmic\\b|usb audio|audio|sound card|soundcard|speaker|speakers|dac|surround\\s*sound|virtual\\s*surround|blackshark|kraken|barracuda|nari|seiren|hammerhead|tiamat|thresher|manowar|opus", RegexOptions.CultureInvariant))
        {
            return null;
        }
        if (name.Contains("samson", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        if (name == "usb receiver")
        {
            return "Mouse";
        }

        if (name == "usb device" || name == "<none>")
        {
            return "Keyboard";
        }

        if (name == "wireless-receiver")
        {
            return "Mouse";
        }

        if (Regex.IsMatch(name, "(?i)usb gaming keyboard|\\bctl\\b"))
        {
            return null;
        }

        string[] keyboardPatterns =
        [
            "keyboard", "kbd", "kb", "he", "68", "75", "80", "63", "irok", "87", "96", "104", "820", "none",
            "60%", "65%", "tkl", "varmilo", "blackwidow", "keypad", "mechanical", "comard", "ak820",
            "cherry mx", "gateron", "keychron", "ducky", "leopold", "filco", "akko", "85",
            "gmmk", "iqunix", "nuphy", "apex pro", "k70", "k95", "optical switch", "rs",
            "75%", "fullsize", "tenkeyless", "macro pad", "keymap", "keycap", "switch",
        ];

        string[] mousePatterns =
        [
            "mouse", "ms", "8k", "2.4g", "4k", "pulsefire", "haste", "deathadder", "helios",
            "viper", "ajazz", "model o", "model d", "g pro", "g502", "g703", "g903", "mad",
            "pulsar", "glorious", "zowie", "trackball", "sensor", "dpi", "gaming mouse",
            "g-wolves", "xm1", "skoll", "hsk", "viper mini", "orca", "superlight", "mchose",
            "scroll wheel", "side button", "ergonomic", "ambidextrous", "fingertip", "major",
            "palm grip", "claw grip", "lod", "ips", "polling rate", "wlmouse", "xd",
        ];

        string[] webcamPatterns =
        [
            "webcam", "web cam", "web camera", "camera", "streamcam", "stream cam", "usb video",
            "brio", "c920", "c922", "c930", "c925", "c270", "c310", "c615", "c525",
        ];

        string[] gamepadPatterns =
        [
            "gamepad", "game pad", "game controller", "wireless controller", "xbox", "xinput",
            "dualshock", "dualsense", "playstation", "ps4", "ps5", "ps3",
            "8bitdo", "joycon", "joy-con", "switch pro", "pro controller", "steam controller", "nintendo",
        ];

        Dictionary<string, string> brandMapping = new(StringComparer.OrdinalIgnoreCase)
        {
            ["varmilo"] = "Keyboard",
            ["ajazz"] = "Mouse",
            ["lamzu"] = "Mouse",
            ["razer"] = "Mouse",
            ["logitech"] = "Mouse",
            ["steelseries"] = "Mouse",
            ["endgame"] = "Mouse",
            ["finalmouse"] = "Mouse",
            ["keychron"] = "Keyboard",
            ["hexgears"] = "Keyboard",
            ["ducky"] = "Keyboard",
            ["leopold"] = "Keyboard",
            ["filco"] = "Keyboard",
            ["akko"] = "Keyboard",
            ["iqunix"] = "Keyboard",
            ["nuphy"] = "Keyboard",
            ["corsair"] = "Keyboard",
            ["hyperx"] = "Keyboard",
            ["asus"] = "Keyboard",
            ["msi"] = "Keyboard",
            ["bloody"] = "Mouse",
            ["roccat"] = "Mouse",
            ["coolermaster"] = "Keyboard",
        };

        Dictionary<string, int> patternMatches = new(StringComparer.OrdinalIgnoreCase);
        void AddMatches(IEnumerable<string> patterns)
        {
            foreach (string p in patterns)
            {
                try
                {
                    int count = Regex.Matches(name, Regex.Escape(p), RegexOptions.IgnoreCase | RegexOptions.CultureInvariant).Count;
                    if (count > 0)
                    {
                        patternMatches[p] = patternMatches.TryGetValue(p, out int old) ? old + count : count;
                    }
                }
                catch
                {
                }
            }
        }

        AddMatches(keyboardPatterns);
        AddMatches(mousePatterns);
        AddMatches(webcamPatterns);
        AddMatches(gamepadPatterns);
        AddMatches(brandMapping.Keys);

        int countKeyboard = 0;
        int countMouse = 0;
        int countWebcam = 0;
        int countGamepad = 0;
        foreach ((string key, int count) in patternMatches)
        {
            if (keyboardPatterns.Contains(key, StringComparer.OrdinalIgnoreCase))
            {
                countKeyboard += count;
            }
            else if (mousePatterns.Contains(key, StringComparer.OrdinalIgnoreCase))
            {
                countMouse += count;
            }
            else if (webcamPatterns.Contains(key, StringComparer.OrdinalIgnoreCase))
            {
                countWebcam += count;
            }
            else if (gamepadPatterns.Contains(key, StringComparer.OrdinalIgnoreCase))
            {
                countGamepad += count;
            }
            else if (brandMapping.TryGetValue(key, out string? mapped))
            {
                if (string.Equals(mapped, "Keyboard", StringComparison.OrdinalIgnoreCase))
                {
                    countKeyboard += count;
                }
                else
                {
                    countMouse += count;
                }
            }
        }

        if (countWebcam > countMouse && countWebcam > countKeyboard && countWebcam > countGamepad)
        {
            return "Webcam";
        }

        if (countGamepad > countMouse && countGamepad > countKeyboard && countGamepad > countWebcam)
        {
            return "Gamepad";
        }

        if (countMouse > countKeyboard && countMouse > countGamepad && countMouse > countWebcam)
        {
            return "Mouse";
        }

        if (countKeyboard > countMouse && countKeyboard > countGamepad && countKeyboard > countWebcam)
        {
            return "Keyboard";
        }

        string[] keyboardEvidence = keyboardPatterns.Where(p => patternMatches.ContainsKey(p)).ToArray();
        string[] mouseEvidence = mousePatterns.Where(p => patternMatches.ContainsKey(p)).ToArray();
        string[] webcamEvidence = webcamPatterns.Where(p => patternMatches.ContainsKey(p)).ToArray();
        string[] gamepadEvidence = gamepadPatterns.Where(p => patternMatches.ContainsKey(p)).ToArray();
        string[] brandEvidence = brandMapping.Keys.Where(p => patternMatches.ContainsKey(p)).ToArray();

        if (keyboardEvidence.Length > 0 && mouseEvidence.Length == 0 && gamepadEvidence.Length == 0 && webcamEvidence.Length == 0)
        {
            return "Keyboard";
        }

        if (mouseEvidence.Length > 0 && keyboardEvidence.Length == 0 && gamepadEvidence.Length == 0 && webcamEvidence.Length == 0)
        {
            return "Mouse";
        }

        if (webcamEvidence.Length > 0 && keyboardEvidence.Length == 0 && mouseEvidence.Length == 0 && gamepadEvidence.Length == 0)
        {
            return "Webcam";
        }

        if (gamepadEvidence.Length > 0 && keyboardEvidence.Length == 0 && mouseEvidence.Length == 0 && webcamEvidence.Length == 0)
        {
            return "Gamepad";
        }

        foreach (string brand in brandEvidence)
        {
            if (brandMapping.TryGetValue(brand, out string? mapped))
            {
                if (mapped == "Keyboard")
                {
                    return "Keyboard";
                }

                if (mapped == "Mouse")
                {
                    return "Mouse";
                }
            }
        }

        foreach (string p in keyboardPatterns)
        {
            if (name.Contains(p, StringComparison.OrdinalIgnoreCase))
            {
                return "Keyboard";
            }
        }

        foreach (string p in mousePatterns)
        {
            if (name.Contains(p, StringComparison.OrdinalIgnoreCase))
            {
                return "Mouse";
            }
        }

        foreach (string p in webcamPatterns)
        {
            if (name.Contains(p, StringComparison.OrdinalIgnoreCase))
            {
                return "Webcam";
            }
        }

        foreach (string p in gamepadPatterns)
        {
            if (name.Contains(p, StringComparison.OrdinalIgnoreCase))
            {
                return "Gamepad";
            }
        }

        return null;
    }

    private static string? GetHidUsageRole(int? usagePage, int? usageId)
    {
        if (usagePage != 1 || usageId is null)
        {
            return null;
        }

        return usageId.Value switch
        {
            6 => "Keyboard",
            2 => "Mouse",
            4 or 5 or 8 => "Gamepad",
            _ => null,
        };
    }

    private static bool IsKeyboardMouseComboProduct(string productName)
    {
        if (string.IsNullOrWhiteSpace(productName))
        {
            return false;
        }

        return Regex.IsMatch(
            productName,
            "(?i)touchpad|trackpad|touch pad|track pad|touch keyboard|keyboard.*touch|combo|all-in-one|keyboard\\s*mouse|keyboard\\s*\\+\\s*mouse|kb\\s*\\+\\s*mouse",
            RegexOptions.CultureInvariant);
    }

    private static string? ResolveHidRole(string productName, string? instanceId, string? productRole, string? usageRole, string? pnpRole)
    {
        if (IsCollectionLevelHid(instanceId)
            && string.IsNullOrWhiteSpace(usageRole)
            && string.IsNullOrWhiteSpace(pnpRole))
        {
            return null;
        }

        if (IsKnownInternalVendorHid(productName)
            && string.IsNullOrWhiteSpace(usageRole)
            && string.IsNullOrWhiteSpace(pnpRole))
        {
            return null;
        }

        if (!string.IsNullOrWhiteSpace(productRole))
        {
            if (string.IsNullOrWhiteSpace(usageRole)
                || string.Equals(productRole, usageRole, StringComparison.OrdinalIgnoreCase))
            {
                return productRole;
            }

            if (string.Equals(productRole, "Keyboard", StringComparison.OrdinalIgnoreCase)
                && string.Equals(usageRole, "Mouse", StringComparison.OrdinalIgnoreCase)
                && IsKeyboardMouseComboProduct(productName))
            {
                return usageRole;
            }

            if (string.Equals(productRole, "Mouse", StringComparison.OrdinalIgnoreCase)
                && string.Equals(usageRole, "Keyboard", StringComparison.OrdinalIgnoreCase)
                && IsKeyboardMouseComboProduct(productName))
            {
                return usageRole;
            }

            if ((string.Equals(productRole, "Mouse", StringComparison.OrdinalIgnoreCase)
                    && string.Equals(usageRole, "Keyboard", StringComparison.OrdinalIgnoreCase))
                || (string.Equals(productRole, "Keyboard", StringComparison.OrdinalIgnoreCase)
                    && string.Equals(usageRole, "Mouse", StringComparison.OrdinalIgnoreCase)))
            {
                return null;
            }

            return productRole;
        }

        if (!string.IsNullOrWhiteSpace(pnpRole)
            && (string.Equals(pnpRole, "Webcam", StringComparison.OrdinalIgnoreCase)
                || string.Equals(pnpRole, "Gamepad", StringComparison.OrdinalIgnoreCase)))
        {
            return pnpRole;
        }

        return !string.IsNullOrWhiteSpace(usageRole) ? usageRole : pnpRole;
    }

    private static string? GetPnpHidRole(WmiPnPDevice? pnpDev)
    {
        if (pnpDev is null)
        {
            return null;
        }

        string name = pnpDev.Name ?? string.Empty;
        if (string.Equals(pnpDev.Class, "Camera", StringComparison.OrdinalIgnoreCase)
            || string.Equals(pnpDev.Class, "Image", StringComparison.OrdinalIgnoreCase)
            || Regex.IsMatch(name, "(?i)web\\s*cam|webcam|camera|streamcam|usb video", RegexOptions.CultureInvariant))
        {
            return "Webcam";
        }

        if (string.Equals(pnpDev.Class, "Keyboard", StringComparison.OrdinalIgnoreCase))
        {
            return "Keyboard";
        }

        if (string.Equals(pnpDev.Class, "Mouse", StringComparison.OrdinalIgnoreCase))
        {
            return "Mouse";
        }

        if (!string.IsNullOrWhiteSpace(name)
            && Regex.IsMatch(name, "(?i)game\\s*controller|gamepad", RegexOptions.CultureInvariant))
        {
            return "Gamepad";
        }

        return null;
    }

    private List<HidDeviceInfo> GetHidDevicesWithUsbControllers(Dictionary<string, WmiPnPDevice> deviceLookup, List<(string ControllerId, string DependentId)> usbPairs)
    {
        List<HidDeviceInfo> results = [];

        foreach (string devicePath in HidInterop.EnumerateHidDevicePaths())
        {
            _ = HidInterop.TryReadProductAndUsage(devicePath, out string product, out int? usagePage, out int? usageId);
            string? productRole = GetHidProductType(product);

            string? inst = HidInterop.TryParseInstanceIdFromDevicePath(devicePath);
            inst = NormalizeInstanceId(inst);

            WmiPnPDevice? pnpDev = null;
            if (!string.IsNullOrWhiteSpace(inst))
            {
                deviceLookup.TryGetValue(inst, out pnpDev);
                pnpDev ??= WmiInterop.TryGetPnPDeviceById(inst);
            }

            string? usageRole = GetHidUsageRole(usagePage, usageId);
            string? pnpRole = GetPnpHidRole(pnpDev);
            string? role = ResolveHidRole(product, inst, productRole, usageRole, pnpRole);

            string? usbAncestor = null;
            string? currId = inst;
            for (int i = 0; i < 10 && !string.IsNullOrWhiteSpace(currId); i++)
            {
                string? parent = GetParentId(currId);
                if (string.IsNullOrWhiteSpace(parent))
                {
                    break;
                }

                if (parent.StartsWith("USB\\", StringComparison.OrdinalIgnoreCase))
                {
                    usbAncestor = parent.ToUpperInvariant();
                    break;
                }

                currId = parent.ToUpperInvariant();
            }

            List<UsbControllerInfo> controllers = [];
            if (!string.IsNullOrWhiteSpace(inst))
            {
                string needle = usbAncestor ?? inst;

                HashSet<string> seen = new(StringComparer.OrdinalIgnoreCase);
                foreach ((string controllerId, string dependentId) in usbPairs)
                {
                    if (!dependentId.Contains(needle, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    string cid = NormalizeInstanceId(controllerId);
                    if (seen.Add(cid))
                    {
                        controllers.Add(new UsbControllerInfo(cid, cid));
                    }
                }

                if (controllers.Count == 0)
                {
                    string? ctrl = FindUsbControllerFor(inst, deviceLookup);
                    if (!string.IsNullOrWhiteSpace(ctrl))
                    {
                        ctrl = NormalizeInstanceId(ctrl);
                        controllers.Add(new UsbControllerInfo(ctrl, ctrl));
                    }
                }

                if (controllers.Count == 0)
                {
                    controllers.AddRange(FindUsbControllersByVidPid(inst, usbPairs));
                }
            }

            results.Add(new HidDeviceInfo
            {
                ProductString = product,
                DeviceType = role,
                DevicePath = devicePath,
                DeviceInstanceId = inst,
                UsbControllers = controllers.DistinctBy(c => c.ControllerPNPID, StringComparer.OrdinalIgnoreCase).ToList(),
                UsagePage = usagePage,
                UsageId = usageId,
            });
        }

        return results;
    }

    private string? FindUsbControllerFor(string childId, Dictionary<string, WmiPnPDevice> deviceLookup)
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
                if (parentKey.StartsWith("PCI\\", StringComparison.OrdinalIgnoreCase)
                    && (string.Equals(dev.Class, "USB", StringComparison.OrdinalIgnoreCase)
                        || Regex.IsMatch(n, "(?i)xHCI|Host Controller")))
                {
                    return parentKey;
                }
            }

            curr = parentKey;
        }

        return null;
    }

    private Dictionary<string, List<string>> UsbControllerRoles(
        List<WmiPnPDevice> deviceCache,
        Dictionary<string, WmiPnPDevice> deviceLookup,
        List<(string ControllerId, string DependentId)> usbPairs,
        IReadOnlyDictionary<string, UsbPollingRateInfo> pollingRateLookup)
    {
        Dictionary<string, List<string>> map = new(StringComparer.OrdinalIgnoreCase);

        List<HidDeviceInfo> hidDevices = GetHidDevicesWithUsbControllers(deviceLookup, usbPairs);
        if (hidDevices.Count > 0)
        {
            foreach (HidDeviceInfo dev in hidDevices)
            {
                string? role = dev.DeviceType;
                if (role is not ("Keyboard" or "Mouse" or "Gamepad" or "Webcam"))
                {
                    continue;
                }

                string displayRole = AddPollingRateTag(role, dev.DeviceInstanceId, pollingRateLookup);
                List<UsbControllerInfo> controllers = dev.UsbControllers;
                if (controllers.Count == 0)
                {
                    string? ctrlFallback = FindUsbControllerFor(dev.DeviceInstanceId ?? string.Empty, deviceLookup);
                    if (!string.IsNullOrWhiteSpace(ctrlFallback))
                    {
                        ctrlFallback = NormalizeInstanceId(ctrlFallback);
                        controllers =
                        [
                            new UsbControllerInfo(ctrlFallback, ctrlFallback),
                        ];
                    }
                    else
                    {
                        WriteLog($"USBROLE: HID device has no controllers prod=\"{dev.ProductString}\" inst={dev.DeviceInstanceId} role={dev.DeviceType}");
                        continue;
                    }
                }

                foreach (UsbControllerInfo ctrl in controllers)
                {
                    string primaryKey = NormalizeInstanceId(ctrl.ControllerPNPID);
                    foreach (string ctrlKey in GetUsbControllerKeys(ctrl.ControllerPNPID, NormalizeInstanceId))
                    {
                        if (!map.TryGetValue(ctrlKey, out List<string>? roles))
                        {
                            roles = [];
                            map[ctrlKey] = roles;
                        }

                        if (!roles.Contains(displayRole, StringComparer.OrdinalIgnoreCase))
                        {
                            roles.Add(displayRole);
                            if (string.Equals(ctrlKey, primaryKey, StringComparison.OrdinalIgnoreCase))
                            {
                                WriteLog($"USBROLE: HID map {displayRole} -> {ctrlKey} (product=\"{dev.ProductString}\" usagePage={dev.UsagePage} usage={dev.UsageId} inst={dev.DeviceInstanceId})");
                            }
                        }
                    }
                }
            }

            foreach (HidDeviceInfo dev in hidDevices)
            {
                string? role = dev.DeviceType;
                if (role is not ("Keyboard" or "Mouse" or "Gamepad" or "Webcam"))
                {
                    continue;
                }

                bool alreadyMapped = map.Values.Any(v => v.Any(entry => IsUsbRoleText(entry, role)));
                if (alreadyMapped)
                {
                    continue;
                }

                string? ctrlParent = FindUsbControllerFor(dev.DeviceInstanceId ?? string.Empty, deviceLookup);
                if (string.IsNullOrWhiteSpace(ctrlParent))
                {
                    continue;
                }

                string displayRole = AddPollingRateTag(role, dev.DeviceInstanceId, pollingRateLookup);
                string primaryKey = NormalizeInstanceId(ctrlParent);
                foreach (string ctrlKey in GetUsbControllerKeys(ctrlParent, NormalizeInstanceId))
                {
                    if (!map.TryGetValue(ctrlKey, out List<string>? roles))
                    {
                        roles = [];
                        map[ctrlKey] = roles;
                    }

                    if (!roles.Contains(displayRole, StringComparer.OrdinalIgnoreCase))
                    {
                        roles.Add(displayRole);
                        if (string.Equals(ctrlKey, primaryKey, StringComparison.OrdinalIgnoreCase))
                        {
                            WriteLog($"USBROLE: HID parent-walk {displayRole} -> {ctrlKey} (product=\"{dev.ProductString}\" inst={dev.DeviceInstanceId})");
                        }
                    }
                }
            }
        }

        foreach (WmiPnPDevice d in deviceCache)
        {
            if (string.IsNullOrWhiteSpace(d.InstanceId))
            {
                continue;
            }

            bool present = string.Equals(d.Status, "OK", StringComparison.OrdinalIgnoreCase) && d.ConfigManagerErrorCode != 22;
            if (!present)
            {
                continue;
            }

            string role = string.Empty;
            string name = d.Name ?? d.InstanceId;

            if (string.Equals(d.Class, "AudioEndpoint", StringComparison.OrdinalIgnoreCase))
            {
                role = "Audio";
                string? roleHint = GetAudioEndpointRole(name);
                if (roleHint == "MIC")
                {
                    role = "Microphone";
                }
            }
            else if (Regex.IsMatch(name, "(?i)microphone|микрофон"))
            {
                role = "Microphone";
            }
            else if (string.Equals(d.Class, "Camera", StringComparison.OrdinalIgnoreCase)
                || string.Equals(d.Class, "Image", StringComparison.OrdinalIgnoreCase)
                || Regex.IsMatch(name, "(?i)web\\s*cam|webcam|camera|streamcam|usb video", RegexOptions.CultureInvariant))
            {
                role = "Webcam";
            }

            if (string.IsNullOrWhiteSpace(role))
            {
                continue;
            }

            string? ctrl = FindUsbControllerFor(d.InstanceId, deviceLookup);
            if (string.IsNullOrWhiteSpace(ctrl))
            {
                List<UsbControllerInfo> extraCtrls = FindUsbControllersByVidPid(d.InstanceId, usbPairs);
                foreach (UsbControllerInfo c in extraCtrls)
                {
                    string ctrlKey = c.ControllerPNPID;
                    if (!map.TryGetValue(ctrlKey, out List<string>? roles))
                    {
                        roles = [];
                        map[ctrlKey] = roles;
                    }

                    if (!roles.Contains(role, StringComparer.OrdinalIgnoreCase))
                    {
                        roles.Add(role);
                        WriteLog($"USBROLE: PNP fallback {role} -> {ctrlKey} (inst={d.InstanceId})");
                    }
                }

                continue;
            }

            string primaryKey = NormalizeInstanceId(ctrl);
            foreach (string ctrlKey in GetUsbControllerKeys(ctrl, NormalizeInstanceId))
            {
                if (!map.TryGetValue(ctrlKey, out List<string>? roles))
                {
                    roles = [];
                    map[ctrlKey] = roles;
                }

                if (!roles.Contains(role, StringComparer.OrdinalIgnoreCase))
                {
                    roles.Add(role);
                    if (string.Equals(ctrlKey, primaryKey, StringComparison.OrdinalIgnoreCase))
                    {
                        WriteLog($"USBROLE: PNP fallback {role} -> {ctrlKey} (inst={d.InstanceId})");
                    }
                }
            }
        }

        return map;
    }
}
