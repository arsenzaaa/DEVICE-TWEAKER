using Microsoft.Win32;
using System.Management;
using System.Text.RegularExpressions;

namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private string? GetClassKeyForDevice(string instanceId)
    {
        string enumPath = $@"SYSTEM\CurrentControlSet\Enum\{instanceId}";
        try
        {
            using RegistryKey? enumKey = Registry.LocalMachine.OpenSubKey(enumPath);
            if (enumKey is null)
            {
                return null;
            }

            string? driver = enumKey.GetValue("Driver") as string;
            if (string.IsNullOrWhiteSpace(driver))
            {
                return null;
            }

            string classKeyPath = $@"SYSTEM\CurrentControlSet\Control\Class\{driver}";
            using RegistryKey? ck = Registry.LocalMachine.OpenSubKey(classKeyPath);
            return ck is not null ? classKeyPath : null;
        }
        catch
        {
            return null;
        }
    }

    private bool TestNdisRssBasePresent(string instanceId)
    {
        string enumPath = $@"SYSTEM\CurrentControlSet\Enum\{instanceId}";
        try
        {
            string? ckPath = GetClassKeyForDevice(instanceId);
            if (!string.IsNullOrWhiteSpace(ckPath))
            {
                try
                {
                    using RegistryKey? ck = Registry.LocalMachine.OpenSubKey(ckPath);
                    if (ck is not null)
                    {
                        if (ck.GetValueNames().Any(n => n == "*RssBaseProcNumber"))
                        {
                            return true;
                        }
                    }
                }
                catch
                {
                }

                try
                {
                    using RegistryKey? ndiParams = Registry.LocalMachine.OpenSubKey(ckPath + "\\Ndi\\Params");
                    if (ndiParams is not null)
                    {
                        foreach (string sub in ndiParams.GetSubKeyNames())
                        {
                            if (Regex.IsMatch(sub, "^\\*RssBaseProcNumber$", RegexOptions.CultureInvariant))
                            {
                                return true;
                            }
                        }
                    }
                }
                catch
                {
                }
            }

            using RegistryKey? ek = Registry.LocalMachine.OpenSubKey(enumPath);
            if (ek is not null)
            {
                if (ek.GetValueNames().Any(n => n == "*RssBaseProcNumber"))
                {
                    return true;
                }
            }
        }
        catch
        {
            return false;
        }

        return false;
    }

    private int? GetNdisBaseCore(string instanceId)
    {
        string enumPath = $@"SYSTEM\CurrentControlSet\Enum\{instanceId}";
        string? ckPath = GetClassKeyForDevice(instanceId);
        if (!string.IsNullOrWhiteSpace(ckPath))
        {
            try
            {
                using RegistryKey? ck = Registry.LocalMachine.OpenSubKey(ckPath);
                object? val = ck?.GetValue("*RssBaseProcNumber");
                if (val is int i)
                {
                    return i;
                }
            }
            catch
            {
            }
        }

        try
        {
            using RegistryKey? ek = Registry.LocalMachine.OpenSubKey(enumPath);
            object? val = ek?.GetValue("*RssBaseProcNumber");
            if (val is int i)
            {
                return i;
            }
        }
        catch
        {
        }

        return null;
    }

    private int? GetNdisRssQueues(string instanceId)
    {
        string enumPath = $@"SYSTEM\CurrentControlSet\Enum\{instanceId}";
        string? ckPath = GetClassKeyForDevice(instanceId);
        if (!string.IsNullOrWhiteSpace(ckPath))
        {
            try
            {
                using RegistryKey? ck = Registry.LocalMachine.OpenSubKey(ckPath);
                object? val = ck?.GetValue("*NumRssQueues");
                if (TryParseRegistryInt(val, out int result))
                {
                    return result;
                }
            }
            catch
            {
            }
        }

        try
        {
            using RegistryKey? ek = Registry.LocalMachine.OpenSubKey(enumPath);
            object? val = ek?.GetValue("*NumRssQueues");
            if (TryParseRegistryInt(val, out int result))
            {
                return result;
            }
        }
        catch
        {
        }

        return null;
    }

    private void SetNdisBaseCore(string instanceId, int baseCore)
    {
        if (baseCore < 0)
        {
            baseCore = 0;
        }

        string? ckPath = GetClassKeyForDevice(instanceId);
        if (!string.IsNullOrWhiteSpace(ckPath))
        {
            try
            {
                using RegistryKey? ck = Registry.LocalMachine.CreateSubKey(ckPath);
                ck?.SetValue("*RssBaseProcNumber", baseCore, RegistryValueKind.DWord);
                WriteLog($"RSS.SET: {instanceId} -> *RssBaseProcNumber={baseCore} (class key)");
            }
            catch
            {
            }
        }
        else
        {
            WriteLog($"RSS.SET.SKIP: {instanceId} class key not found; *RssBaseProcNumber not written");
        }
    }

    private void SetNdisRssQueues(string instanceId, int queues)
    {
        if (queues < 1)
        {
            queues = 1;
        }

        string? ckPath = GetClassKeyForDevice(instanceId);
        if (!string.IsNullOrWhiteSpace(ckPath))
        {
            try
            {
                using RegistryKey? ck = Registry.LocalMachine.CreateSubKey(ckPath);
                ck?.SetValue("*NumRssQueues", queues, RegistryValueKind.DWord);
                WriteLog($"RSS.SET: {instanceId} -> *NumRssQueues={queues} (class key)");
            }
            catch
            {
            }
        }
        else
        {
            WriteLog($"RSS.SET.SKIP: {instanceId} class key not found; *NumRssQueues not written");
        }
    }

    private static bool TryParseRegistryInt(object? value, out int result)
    {
        result = 0;
        if (value is null)
        {
            return false;
        }

        switch (value)
        {
            case int i:
                result = i;
                return true;
            case uint ui:
                result = unchecked((int)ui);
                return true;
            case long l:
                result = unchecked((int)l);
                return true;
            case ulong ul:
                result = unchecked((int)ul);
                return true;
            case short s:
                result = s;
                return true;
            case ushort us:
                result = us;
                return true;
            case byte b:
                result = b;
                return true;
            case sbyte sb:
                result = sb;
                return true;
            case string str when int.TryParse(str, out int parsed):
                result = parsed;
                return true;
            default:
                return false;
        }
    }

    private static bool HasNonZeroAssignmentOverride(object? value)
    {
        if (value is null)
        {
            return false;
        }

        switch (value)
        {
            case byte[] bytes:
                return bytes.Any(b => b != 0);
            case int i:
                return i != 0;
            case uint ui:
                return ui != 0;
            case long l:
                return l != 0;
            case ulong ul:
                return ul != 0;
            case short s:
                return s != 0;
            case ushort us:
                return us != 0;
            case byte b:
                return b != 0;
            case sbyte sb:
                return sb != 0;
            case string str when ulong.TryParse(str, out ulong parsed):
                return parsed != 0;
            default:
                return false;
        }
    }

    private bool HasNonDefaultInterruptAffinity(string regBase)
    {
        string affPath = regBase + @"\Device Parameters\Interrupt Management\Affinity Policy";
        try
        {
            using RegistryKey? affKey = Registry.LocalMachine.OpenSubKey(affPath);
            if (affKey is null)
            {
                return false;
            }

            if (TryParseRegistryInt(affKey.GetValue("DevicePolicy"), out int policy) && policy != 0)
            {
                return true;
            }

            if (TryParseRegistryInt(affKey.GetValue("DevicePriority"), out int priority) && priority != 2)
            {
                return true;
            }

            return HasNonZeroAssignmentOverride(affKey.GetValue("AssignmentSetOverride"));
        }
        catch
        {
            return false;
        }
    }

    private void ClearNdisBaseCore(string instanceId)
    {
        string? ckPath = GetClassKeyForDevice(instanceId);
        if (string.IsNullOrWhiteSpace(ckPath))
        {
            return;
        }

        try
        {
            using RegistryKey? ck = Registry.LocalMachine.OpenSubKey(ckPath, writable: true);
            ck?.DeleteValue("*RssBaseProcNumber", throwOnMissingValue: false);
        }
        catch
        {
        }
    }

    private void ClearNdisRssQueues(string instanceId)
    {
        string? ckPath = GetClassKeyForDevice(instanceId);
        if (string.IsNullOrWhiteSpace(ckPath))
        {
            return;
        }

        try
        {
            using RegistryKey? ck = Registry.LocalMachine.OpenSubKey(ckPath, writable: true);
            ck?.DeleteValue("*NumRssQueues", throwOnMissingValue: false);
        }
        catch
        {
        }
    }

    private DeviceKind GetNetDeviceKind(WmiPnPDevice device)
    {
        bool hasCx = false;
        bool filterHit = false;
        bool svcHit = false;

        string? svcName = device.Service;
        if (string.IsNullOrWhiteSpace(svcName))
        {
            try
            {
                using RegistryKey? enumKey = Registry.LocalMachine.OpenSubKey($@"SYSTEM\CurrentControlSet\Enum\{device.InstanceId}");
                svcName = enumKey?.GetValue("Service") as string;
            }
            catch
            {
            }
        }

        if (!string.IsNullOrWhiteSpace(svcName))
        {
            try
            {
                using RegistryKey? svcKey = Registry.LocalMachine.OpenSubKey($@"SYSTEM\CurrentControlSet\Services\{svcName}");
                if (svcKey is not null)
                {
                    string? displayName = svcKey.GetValue("DisplayName") as string;
                    string? imagePath = svcKey.GetValue("ImagePath") as string;
                    string[] depends = svcKey.GetValue("DependOnService") as string[] ?? [];

                    bool netadapter =
                        Regex.IsMatch(displayName ?? string.Empty, "(?i)netadapter")
                        || Regex.IsMatch(imagePath ?? string.Empty, "(?i)netadaptercx|netadapter\\.sys")
                        || depends.Any(d => Regex.IsMatch(d, "(?i)netadapter"));

                    if (netadapter)
                    {
                        hasCx = true;
                        svcHit = true;
                    }
                }
            }
            catch
            {
            }
        }

        if (!hasCx)
        {
            try
            {
                using RegistryKey? enumKey = Registry.LocalMachine.OpenSubKey($@"SYSTEM\CurrentControlSet\Enum\{device.InstanceId}");
                if (enumKey is not null)
                {
                    foreach (string name in new[] { "UpperFilters", "LowerFilters" })
                    {
                        string[] values = enumKey.GetValue(name) as string[] ?? [];
                        if (values.Any(v => Regex.IsMatch(v, "(?i)netadapter")))
                        {
                            hasCx = true;
                            filterHit = true;
                            break;
                        }
                    }
                }
            }
            catch
            {
            }
        }

        if (hasCx)
        {
            WriteLog($"NET.KIND: {device.InstanceId} -> NET_CX (svc={svcName} filterHit={filterHit} svcHit={svcHit})");
            return DeviceKind.NET_CX;
        }

        bool rssPresent = false;
        try
        {
            rssPresent = TestNdisRssBasePresent(device.InstanceId);
        }
        catch
        {
            rssPresent = false;
        }

        if (rssPresent)
        {
            WriteLog($"NET.KIND: {device.InstanceId} -> NET_NDIS (RSS base present)");
            return DeviceKind.NET_NDIS;
        }

        WriteLog($"NET.KIND: {device.InstanceId} -> NET_NDIS (default, svc={svcName})");
        return DeviceKind.NET_NDIS;
    }

    private List<DeviceInfo> GetDeviceList()
    {
        WriteLog("SCAN: Get-DeviceList start");

        List<DeviceInfo> SortDevices(List<DeviceInfo> list)
        {
            return list
                .OrderBy(d => d.Kind switch
                {
                    DeviceKind.USB => 1,
                    DeviceKind.GPU => 2,
                    DeviceKind.AUDIO => 3,
                    DeviceKind.NET_NDIS => 4,
                    DeviceKind.NET_CX => 4,
                    DeviceKind.STOR => 5,
                    _ => 6,
                })
                .ThenBy(d => d.Name, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        if (_testDevicesEnabled && _testDevicesOnly && _testDevices.Count > 0)
        {
            WriteLog($"SCAN.TEST: using test devices only count={_testDevices.Count}");
            List<DeviceInfo> testOnly = _testDevices.ToList();
            testOnly = SortDevices(testOnly);
            WriteLog($"SCAN: Get-DeviceList done, count={testOnly.Count}");
            return testOnly;
        }

        List<DeviceInfo> devices = [];
        List<WmiPnPDevice> raw = WmiInterop.GetPnPDevices();
        Dictionary<string, WmiPnPDevice> deviceLookup = new(StringComparer.OrdinalIgnoreCase);
        foreach (WmiPnPDevice r in raw)
        {
            string key = NormalizeInstanceId(r.InstanceId);
            if (!string.IsNullOrWhiteSpace(key) && !deviceLookup.ContainsKey(key))
            {
                deviceLookup[key] = r;
            }
        }

        List<(string ControllerId, string DependentId)> usbPairs = WmiInterop.GetUsbControllerDevicePairs(NormalizeInstanceId);
        HashSet<string> usbControllersWithDevice = new(StringComparer.OrdinalIgnoreCase);
        foreach ((string controllerId, string dependentId) in usbPairs)
        {
            if (string.IsNullOrWhiteSpace(controllerId) || string.IsNullOrWhiteSpace(dependentId))
            {
                continue;
            }

            if (IsUsbInfrastructureDevice(dependentId))
            {
                continue;
            }

            usbControllersWithDevice.Add(controllerId);
        }

        Dictionary<string, List<string>> usbRoles = UsbControllerRoles(raw, deviceLookup, usbPairs);
        Dictionary<string, List<string>> audioEndpoints = AudioControllerEndpoints(raw, deviceLookup);
        List<WmiPhysicalDisk> physicalDisks = WmiInterop.GetPhysicalDisks();

        string[] skipPatterns =
        [
            "^ACPI\\\\AMDI00(10|30)\\\\",
            "^ACPI\\\\PNP0103",
            "^ACPI\\\\PNP0501",
            "^PCI\\\\VEN_1022&DEV_14DB",
            "^PCI\\\\VEN_1022&DEV_14DD",
            "^ACPI\\\\INTC1055\\\\",
            "^PCI\\\\VEN_8086&DEV_51E8",
            "^PCI\\\\VEN_8086&DEV_A73E",
        ];

        foreach (WmiPnPDevice d in raw)
        {
            if (string.IsNullOrWhiteSpace(d.InstanceId))
            {
                continue;
            }

            bool present = string.Equals(d.Status, "OK", StringComparison.OrdinalIgnoreCase) && d.ConfigManagerErrorCode != 22;
            if (!present)
            {
                WriteLog($"SCAN: skipped non-present/disabled device {d.InstanceId} class={d.Class} status={d.Status} cmErr={d.ConfigManagerErrorCode}");
                continue;
            }

            if (d.InstanceId.StartsWith(@"ACPI\PNP0100", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (skipPatterns.Any(p => Regex.IsMatch(d.InstanceId, p, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)))
            {
                WriteLog($"SCAN: skipped filtered device {d.InstanceId}");
                continue;
            }

            string name = !string.IsNullOrWhiteSpace(d.Name) ? d.Name : d.InstanceId;
            string service = d.Service ?? string.Empty;
            if (string.IsNullOrWhiteSpace(service))
            {
                try
                {
                    using RegistryKey? enumKey = Registry.LocalMachine.OpenSubKey($@"SYSTEM\CurrentControlSet\Enum\{d.InstanceId}");
                    service = enumKey?.GetValue("Service") as string ?? string.Empty;
                }
                catch
                {
                    service = string.Empty;
                }
            }

            if (string.Equals(d.Class, "System", StringComparison.OrdinalIgnoreCase) && Regex.IsMatch(name, "(?i)ACPI"))
            {
                continue;
            }

            if (string.Equals(d.Class, "Ports", StringComparison.OrdinalIgnoreCase) || Regex.IsMatch(name, "(?i)\\b(COM\\d+|Serial Port|Communications Port)"))
            {
                WriteLog($"SCAN: skipped noisy port device {d.InstanceId} class={d.Class} name=\"{name}\"");
                continue;
            }

            if (string.Equals(d.Class, "Bluetooth", StringComparison.OrdinalIgnoreCase) || Regex.IsMatch(name, "(?i)Bluetooth"))
            {
                WriteLog($"SCAN: skipped Bluetooth device {d.InstanceId} class={d.Class} name=\"{name}\"");
                continue;
            }

            if (string.Equals(d.Class, "Net", StringComparison.OrdinalIgnoreCase) && Regex.IsMatch(name, "(?i)(virtual|loopback|hyper-v|vmware|wan miniport|tap|tun|isatap|teredo|npcap)"))
            {
                WriteLog($"SCAN: skipped virtual/net-miniport {d.InstanceId} class={d.Class} name=\"{name}\"");
                continue;
            }

            if (string.Equals(d.Class, "SoftwareDevice", StringComparison.OrdinalIgnoreCase) || Regex.IsMatch(name, "(?i)(vb-?audio|voicemeeter|virtual cable|virtual audio|broadcast|software device)"))
            {
                WriteLog($"SCAN: skipped software/virtual audio device {d.InstanceId} class={d.Class} name=\"{name}\"");
                continue;
            }

            if (string.Equals(d.Class, "System", StringComparison.OrdinalIgnoreCase)
                && Regex.IsMatch(name, "(?i)(PCI Express Root Port|PCI-to-PCI Bridge|PCI standard .*bridge|SMBus|LPC Controller|ISA Bridge|I2C Controller|SPI Controller|GPIO Controller|PS/2 Controller)"))
            {
                WriteLog($"SCAN: skipped chipset/bridge device {d.InstanceId} class={d.Class} name=\"{name}\"");
                continue;
            }

            if (Regex.IsMatch(name, "(?i)Intel\\s*(?:\\(R\\))?\\s*RST\\s*VMD\\s*Controller\\b"))
            {
                WriteLog($"SCAN: skipped filtered device (Intel RST VMD controller) {d.InstanceId} class={d.Class} name=\"{name}\"");
                continue;
            }

            if (Regex.IsMatch(name, "(?i)Intel\\s*(?:\\(R\\))?\\s*UHD\\s*Graphics\\b"))
            {
                WriteLog($"SCAN: skipped filtered device (Intel UHD Graphics) {d.InstanceId} class={d.Class} name=\"{name}\"");
                continue;
            }

            if (string.Equals(d.Class, "Display", StringComparison.OrdinalIgnoreCase)
                && (Regex.IsMatch(name, "(?i)VirtualBox") || Regex.IsMatch(d.InstanceId, "(?i)VEN_80EE")))
            {
                WriteLog($"SCAN: skipped virtual GPU device {d.InstanceId} class={d.Class} name=\"{name}\"");
                continue;
            }

            DeviceKind kind = DeviceKind.OTHER;
            if (string.Equals(d.Class, "USB", StringComparison.OrdinalIgnoreCase) || Regex.IsMatch(name, "(?i)xHCI|Host Controller"))
            {
                kind = DeviceKind.USB;
            }
            else if (string.Equals(d.Class, "Display", StringComparison.OrdinalIgnoreCase))
            {
                kind = DeviceKind.GPU;
            }
            else if (string.Equals(d.Class, "Net", StringComparison.OrdinalIgnoreCase))
            {
                kind = GetNetDeviceKind(d);
            }
            else if (new[] { "SCSIAdapter", "HDC", "IDE", "Storage" }.Any(c => string.Equals(d.Class, c, StringComparison.OrdinalIgnoreCase)))
            {
                kind = DeviceKind.STOR;
            }
            else if (string.Equals(d.Class, "MEDIA", StringComparison.OrdinalIgnoreCase)
                || string.Equals(d.Class, "AudioEndpoint", StringComparison.OrdinalIgnoreCase)
                || Regex.IsMatch(name, "(?i)High Definition Audio|HD Audio|Audio Controller"))
            {
                kind = DeviceKind.AUDIO;
            }

            bool isWifi = IsWiFiDevice(d.InstanceId, name, service);
            bool usbIsXhci = false;
            bool usbHasDevices = false;
            string idKey = NormalizeInstanceId(d.InstanceId);

            if (kind == DeviceKind.USB)
            {
                usbIsXhci = Regex.IsMatch(name, "(?i)xHCI|Host Controller")
                    || string.Equals(service, "USBXHCI", StringComparison.OrdinalIgnoreCase);

                if (!string.IsNullOrWhiteSpace(idKey))
                {
                    foreach (string key in GetUsbControllerKeys(idKey, NormalizeInstanceId))
                    {
                        if (usbControllersWithDevice.Contains(key))
                        {
                            usbHasDevices = true;
                            break;
                        }
                    }
                }
            }

            if (kind == DeviceKind.OTHER)
            {
                WriteLog($"SCAN: skipped device (kind OTHER) {d.InstanceId} class={d.Class} name=\"{name}\"");
                continue;
            }

            if (kind == DeviceKind.USB && Regex.IsMatch(d.InstanceId, "(?i)\\\\VEN_10DE\\\\"))
            {
                WriteLog($"SCAN: skipped NVIDIA USB controller {d.InstanceId} name=\"{name}\"");
                continue;
            }

            string regBase = $@"SYSTEM\CurrentControlSet\Enum\{d.InstanceId}";
            string intBase = regBase + @"\Device Parameters\Interrupt Management";
            try
            {
                using RegistryKey? intKey = Registry.LocalMachine.OpenSubKey(intBase);
                if (intKey is null)
                {
                    continue;
                }
            }
            catch
            {
                continue;
            }

            string usbText = string.Empty;
            if (kind == DeviceKind.USB && !string.IsNullOrWhiteSpace(idKey))
            {
                foreach (string k in GetUsbControllerKeys(idKey, NormalizeInstanceId))
                {
                    if (usbRoles.TryGetValue(k, out List<string>? roles))
                    {
                        usbText = string.Join(", ", roles.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(r => r));
                        break;
                    }
                }

                if (string.IsNullOrWhiteSpace(usbText))
                {
                    if (!HasNonDefaultInterruptAffinity(regBase))
                    {
                        WriteLog($"SCAN: skipped USB controller (no attached roles) {d.InstanceId} name=\"{name}\"");
                        continue;
                    }

                    WriteLog($"SCAN: keeping USB controller (no attached roles, custom affinity present) {d.InstanceId} name=\"{name}\"");
                }
            }

            string audioText = string.Empty;
            List<string>? rawList = null;
            bool isSpdifOnlyAudio = false;
            if (kind == DeviceKind.AUDIO)
            {
                if (audioEndpoints.TryGetValue(d.InstanceId, out List<string>? names))
                {
                    rawList = names.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(x => x).ToList();
                }

                if (rawList is not null && rawList.Count > 0)
                {
                    bool hasSpdifEndpoint = rawList.Any(IsSpdifAudioEndpointsText);
                    bool hasNonSpdifEndpoint = rawList.Any(endpoint => !IsSpdifAudioEndpointsText(endpoint));
                    isSpdifOnlyAudio = hasSpdifEndpoint && !hasNonSpdifEndpoint;

                    audioText = FormatAudioEndpointsSummary(rawList);
                    bool isDisplayAudio = !isSpdifOnlyAudio && (IsDisplayHdmiaudio(d.InstanceId, name) || IsDisplayAudioEndpointsText(audioText));
                    if (isDisplayAudio && !string.IsNullOrWhiteSpace(audioText))
                    {
                        audioText = Regex.Replace(audioText, "^(?i)HDMI AUDIO(?:\\s*#\\d+)?\\s*-?\\s*", string.Empty, RegexOptions.CultureInvariant);
                        string transport = DetectDisplayTransport(rawList, name);
                        string transportLabel = transport is not null && transport != "HDMI/DP" ? $"Monitor {transport}" : "Monitor";
                        audioText = string.IsNullOrWhiteSpace(audioText) ? transportLabel : $"{transportLabel} - {audioText}";
                    }
                }
            }

            if (kind == DeviceKind.AUDIO && isSpdifOnlyAudio)
            {
                string endpointsLog = rawList is not null && rawList.Count > 0
                    ? string.Join("; ", rawList)
                    : audioText;
                WriteLog($"SCAN: skipped AUDIO device (filtered S/PDIF endpoint) {d.InstanceId} class={d.Class} name=\"{name}\" endpoints=\"{endpointsLog}\"");
                continue;
            }

            if (kind == DeviceKind.AUDIO && (rawList is null || rawList.Count == 0))
            {
                WriteLog($"SCAN: skipped AUDIO device (no endpoints) {d.InstanceId} class={d.Class} name=\"{name}\"");
                continue;
            }

            string displayName = CleanDeviceDisplayName(name);
            if (string.IsNullOrWhiteSpace(displayName))
            {
                displayName = name;
            }

            string storageTag = string.Empty;
            if (kind == DeviceKind.STOR)
            {
                storageTag = GetStorageTagForDevice(displayName, physicalDisks);
            }

            DeviceInfo devInfo = new()
            {
                Name = displayName,
                InstanceId = d.InstanceId,
                Class = d.Class ?? string.Empty,
                RegBase = regBase,
                Kind = kind,
                UsbRoles = usbText,
                AudioEndpoints = audioText,
                StorageTag = storageTag,
                Wifi = isWifi,
                UsbIsXhci = usbIsXhci,
                UsbHasDevices = usbHasDevices,
            };

            devices.Add(devInfo);
            WriteLog($"SCAN: device {d.InstanceId} kind={kind} class={d.Class} name=\"{displayName}\" reg=HKLM\\{regBase} usbRoles=\"{usbText}\" audio=\"{audioText}\"");
        }

        devices = devices
            .ToList();

        if (_testDevicesEnabled && _testDevices.Count > 0)
        {
            devices.AddRange(_testDevices);
            WriteLog($"SCAN.TEST: appended test devices count={_testDevices.Count}");
        }

        devices = SortDevices(devices);

        WriteLog($"SCAN: Get-DeviceList done, count={devices.Count}");
        return devices;
    }

    private static string GetStorageTagForDevice(string deviceName, IReadOnlyList<WmiPhysicalDisk> physicalDisks)
    {
        if (string.IsNullOrWhiteSpace(deviceName))
        {
            return string.Empty;
        }

        bool isNvme = Regex.IsMatch(deviceName, "(?i)NVM\\s*Express|NVMe");
        bool isSata = Regex.IsMatch(deviceName, "(?i)\\bSATA\\b|\\bAHCI\\b");

        if (physicalDisks.Count == 0)
        {
            return isNvme ? "SSD" : string.Empty;
        }

        List<ushort> busTypes = [];
        if (isNvme)
        {
            busTypes.Add(17);
        }
        else if (isSata)
        {
            busTypes.Add(11);
            busTypes.Add(3);
        }

        if (busTypes.Count == 0)
        {
            return isNvme ? "SSD" : string.Empty;
        }

        bool anySsd = false;
        bool anyHdd = false;
        foreach (WmiPhysicalDisk d in physicalDisks)
        {
            if (!busTypes.Contains(d.BusType))
            {
                continue;
            }

            anyHdd |= d.MediaType == 3;
            anySsd |= d.MediaType == 4 || d.MediaType == 5;
        }

        if (anySsd && !anyHdd)
        {
            return "SSD";
        }

        if (anyHdd && !anySsd)
        {
            return "HDD";
        }

        if (anySsd && anyHdd)
        {
            return "SSD+HDD";
        }

        return isNvme ? "SSD" : string.Empty;
    }

    private Dictionary<string, int> GetDeviceIrqCounts()
    {
        Dictionary<string, int> irqCounts = new(StringComparer.OrdinalIgnoreCase);
        try
        {
            using ManagementObjectSearcher searcher = new(
                "root\\CIMV2",
                "SELECT Antecedent, Dependent FROM Win32_PnPAllocatedResource");

            foreach (ManagementObject mo in searcher.Get())
            {
                try
                {
                    string? dependent = mo["Dependent"] as string;
                    string? antecedent = mo["Antecedent"] as string;
                    if (string.IsNullOrWhiteSpace(dependent) || string.IsNullOrWhiteSpace(antecedent))
                    {
                        continue;
                    }

                    if (!antecedent.Contains("Win32_IRQResource", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    Match match = Regex.Match(dependent, "DeviceID=\"([^\"]+)\"", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                    if (!match.Success)
                    {
                        continue;
                    }

                    string deviceId = match.Groups[1].Value.Replace("\\\\", "\\");
                    if (deviceId.Contains("ACPI", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    string formattedId = GetShortPnpId(deviceId);
                    if (string.IsNullOrWhiteSpace(formattedId))
                    {
                        continue;
                    }

                    irqCounts.TryGetValue(formattedId, out int count);
                    irqCounts[formattedId] = count + 1;
                }
                catch
                {
                }
            }
        }
        catch
        {
        }

        foreach (string k in irqCounts.Keys)
        {
            WriteLog($"IRQ.COUNT: {k} -> {irqCounts[k]}");
        }

        return irqCounts;
    }
}
