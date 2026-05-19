using System.Management;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32;
using System.Diagnostics;
using System.Globalization;

namespace DeviceTweakerCS;

public sealed partial class MainForm : Form
{
    private readonly List<DeviceBlock> _blocks = [];
    private readonly Dictionary<string, NdisRssRuntimeState> _ndisRssRuntimeCache = new(StringComparer.OrdinalIgnoreCase);

    private Panel _devicesHost = null!;
    private Panel _devicesPanel = null!;
    private ThemedScrollBar _devicesScroll = null!;
    private Panel? _reservedCpuPanel;

    private Button _btnLog = null!;
    private int _suppressReservedCpuEvents;
    private bool _testCpuActive;
    private readonly List<DeviceInfo> _testDevices = [];
    private readonly HashSet<string> _testHiddenDeviceIds = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, string> _testHiddenDeviceLabels = new(StringComparer.OrdinalIgnoreCase);
    private bool _testDevicesEnabled;
    private bool _testDevicesOnly;
    private bool _testAutoDryRun;
    private int _testDeviceSequence;
    private string _testCpuName = string.Empty;
    private Label? _cpuHeaderLabel;
    private Label? _htPrefixLabel;
    private Label? _htStatusLabel;
    private Label? _hybridCpuPrefixLabel;
    private Label? _hybridCpuStatusLabel;
    private Label? _cppcPrefixLabel;
    private Label? _cppcStatusLabel;
    private Label? _dualCcdPrefixLabel;
    private Label? _dualCcdStatusLabel;
    private FlowLayoutPanel? _cpuFlagsPanel;

    private bool _detailedLogEnabled;
    private string? _detailedLogPath;
    private bool _syncingScroll;
    private bool? _lastGpuDriverDetected;
    private bool _pendingGpuDriverWarning;
    private string? _imodKernelCiBlockStatus;
    private string? _imodKernelCiBlockDetail;
    private DateTime _imodKernelCiBlockStatusUtc;
    private int _imodReadbackGeneration;
    private System.Windows.Forms.Timer? _layoutRefreshTimer;
    private int _lastLayoutViewportWidth;
    private int _lastLayoutDpi;
    private bool _expandingMainWindowForLayout;
    private bool _initialDeviceViewportHeightAdjusted;
    private int _irqRefreshGeneration;

    public MainForm()
    {
        UpdateUiScale();
        AutoScaleMode = AutoScaleMode.None;

        EnableDetailedLog();
        InitializeCpu();
        InitializeGui();
        ApplyAppIcon();
    }

    protected override void OnShown(EventArgs e)
    {
        base.OnShown(e);
        ApplyTitleBarTheme();
        InitializeRawPolling();
        BeginInvoke(new Action(() => RefreshBlocks()));
        if (_pendingGpuDriverWarning)
        {
            _pendingGpuDriverWarning = false;
            ShowMissingGpuDriverWarning();
        }
    }

    private const int DwmwaCaptionColor = 35;
    private const int DwmwaTextColor = 36;

    [DllImport("dwmapi.dll")]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attribute, ref int attributeValue, int attributeSize);

    private void ApplyTitleBarTheme()
    {
        try
        {
            if (!OperatingSystem.IsWindowsVersionAtLeast(10, 0, 17763))
            {
                return;
            }

            int bg = ColorToColorRef(_bgForm);
            _ = DwmSetWindowAttribute(Handle, DwmwaCaptionColor, ref bg, sizeof(int));

            int fg = ColorToColorRef(_fgMain);
            _ = DwmSetWindowAttribute(Handle, DwmwaTextColor, ref fg, sizeof(int));
        }
        catch
        {
        }
    }

    private void ShowMissingGpuDriverWarning()
    {
        const string message = "NVIDIA/AMD video driver not detected.\nInstall the GPU driver and press REFRESH.";
        WriteLog("WARN: NVIDIA/AMD GPU driver not detected");
        ShowThemedInfo(message);
    }

    private void ApplyAppIcon()
    {
        if (_appIcon is not null)
        {
            Icon = _appIcon;
            return;
        }

        _appIcon = LoadEmbeddedAppIcon() ?? TryExtractExeIcon();
        if (_appIcon is not null)
        {
            Icon = _appIcon;
        }
    }

    private static Icon? LoadEmbeddedAppIcon()
    {
        try
        {
            Assembly asm = typeof(MainForm).Assembly;
            using Stream? stream = asm.GetManifestResourceStream("DeviceTweakerCS.AppIcon");
            if (stream is null)
            {
                return null;
            }

            using Icon icon = new(stream);
            return (Icon)icon.Clone();
        }
        catch
        {
            return null;
        }
    }

    private static Icon? TryExtractExeIcon()
    {
        try
        {
            return Icon.ExtractAssociatedIcon(Application.ExecutablePath);
        }
        catch
        {
            return null;
        }
    }

    private static int ColorToColorRef(Color color)
    {
        return color.R | (color.G << 8) | (color.B << 16);
    }

    private string GetScriptRoot()
    {
        try
        {
            return AppContext.BaseDirectory.TrimEnd(Path.DirectorySeparatorChar);
        }
        catch
        {
            return Environment.CurrentDirectory;
        }
    }

    private void InitializeDetailedLogFile()
    {
        if (!string.IsNullOrWhiteSpace(_detailedLogPath))
        {
            return;
        }

        string root = GetScriptRoot();
        string stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        string path = Path.Combine(root, $"DeviceTweaker_{stamp}.log");

        try
        {
            File.WriteAllText(path, string.Empty);
        }
        catch
        {
        }

        _detailedLogPath = path;
    }

    private void WriteLog(string message)
    {
        if (!_detailedLogEnabled)
        {
            return;
        }

        if (string.IsNullOrWhiteSpace(message))
        {
            return;
        }

        if (string.IsNullOrWhiteSpace(_detailedLogPath))
        {
            InitializeDetailedLogFile();
        }

        if (string.IsNullOrWhiteSpace(_detailedLogPath))
        {
            return;
        }

        string line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}";
        try
        {
            File.AppendAllText(_detailedLogPath, line + Environment.NewLine, Encoding.UTF8);
        }
        catch
        {
        }
    }

    private void EnableDetailedLog()
    {
        if (_detailedLogEnabled)
        {
            return;
        }

        _detailedLogEnabled = true;
        InitializeDetailedLogFile();

        WriteLog("LOG: detailed logging ENABLED");
        WriteLog($"LOG.VERSION: {GetAppVersion()}");

        try
        {
            WriteLog(
                $"BOOT: DotNet={Environment.Version} PID={Environment.ProcessId} Arch={RuntimeInformation.ProcessArchitecture} Dir={GetScriptRoot()}");
        }
        catch
        {
        }

        try
        {
            using ManagementObjectSearcher osSearcher = new(
                "root\\CIMV2",
                "SELECT Caption, Version, LastBootUpTime FROM Win32_OperatingSystem");

            foreach (ManagementObject mo in osSearcher.Get())
            {
                string caption = mo["Caption"] as string ?? "Windows";
                string version = mo["Version"] as string ?? string.Empty;
                string? lastBoot = mo["LastBootUpTime"] as string;

                if (!string.IsNullOrWhiteSpace(lastBoot))
                {
                    DateTime bootTime = ManagementDateTimeConverter.ToDateTime(lastBoot);
                    TimeSpan uptime = DateTime.Now - bootTime;
                    WriteLog($"OS: {caption} {version} Uptime={uptime.TotalDays:N1}d");
                }
                else
                {
                    WriteLog($"OS: {caption} {version}");
                }

                break;
            }
        }
        catch
        {
        }

        if (_cpuInfo is not null)
        {
            WriteLog($"CPU.SUMMARY: logical={_cpuInfo.Topology.Logical} maxVisual={_maxLogical} groups={_cpuGroupCount} ccd={_cpuInfo.CcdMap.Values.Distinct().Count()} ccx={_cpuInfo.CcxMap.Values.Distinct().Count()}");
            WriteLog("CPU.TOPO: snapshot from CpuInfo");

            foreach (CpuLpInfo e in _cpuInfo.Topology.LPs.OrderBy(x => x.LP))
            {
                bool isSecondaryThread = false;
                int coreKey = CpuTopology.MakeCoreKey(e.Group, e.Core);
                if (_cpuInfo.Topology.ByCore.TryGetValue(coreKey, out List<CpuLpInfo>? coreGroup) && coreGroup.Count > 1)
                {
                    if (coreGroup[0].LP != e.LP)
                    {
                        isSecondaryThread = true;
                    }
                }

                int ccdId = _cpuInfo.CcdMap.TryGetValue(e.LP, out int cid) ? cid : 0;
                int ccxId = _cpuInfo.CcxMap.TryGetValue(e.LP, out int xid) ? xid : 0;
                string localText = e.LocalIndex >= 0 ? $" Local={e.LocalIndex}" : string.Empty;
                string idText = e.CpuSetId >= 0 ? $" Id={e.CpuSetId}" : string.Empty;
                string typeText = IsEfficiencyCore(e) ? "E" : isSecondaryThread ? "P/HT" : "P";
                string cppcText = _cppcRanks.TryGetValue(e.LP, out int rank)
                    ? _cppcRatings.TryGetValue(e.LP, out int rating)
                        ? $" CPPC=R{rating}/#{rank} Preferred={(rank == 1 ? 1 : 0)}"
                        : $" CPPC=#{rank} Preferred={(rank == 1 ? 1 : 0)}"
                    : " CPPC=unavailable Preferred=0";
                WriteLog($"CPU.ENTRY: G{e.Group} L{e.LP}{localText}{idText} Core={e.Core} Type={typeText} CCD={ccdId} CCX={ccxId} SMT={isSecondaryThread} EffClass={e.EffClass}{cppcText}");
            }

            Dictionary<int, List<int>> ccdGroups = new();
            foreach (KeyValuePair<int, int> kvp in _cpuInfo.CcdMap)
            {
                if (!ccdGroups.TryGetValue(kvp.Value, out List<int>? list))
                {
                    list = [];
                    ccdGroups[kvp.Value] = list;
                }

                list.Add(kvp.Key);
            }

            foreach (int ccdId in ccdGroups.Keys.OrderBy(x => x))
            {
                List<int> lps = ccdGroups[ccdId].OrderBy(x => x).ToList();
                WriteLog($"CCD.MAP: CCD{ccdId} -> LPs=[{string.Join(',', lps)}]");
            }

            Dictionary<int, List<int>> ccxGroups = new();
            foreach (KeyValuePair<int, int> kvp in _cpuInfo.CcxMap)
            {
                if (!ccxGroups.TryGetValue(kvp.Value, out List<int>? list))
                {
                    list = [];
                    ccxGroups[kvp.Value] = list;
                }

                list.Add(kvp.Key);
            }

            foreach (int ccxId in ccxGroups.Keys.OrderBy(x => x))
            {
                List<int> lps = ccxGroups[ccxId].OrderBy(x => x).ToList();
                WriteLog($"CCX.MAP: CCX{ccxId} -> LPs=[{string.Join(',', lps)}]");
            }
        }

        if (_blocks.Count > 0)
        {
            foreach (DeviceBlock b in _blocks)
            {
                string msi = b.MsiCombo.SelectedItem?.ToString() ?? "(none)";
                string prio = b.PrioCombo.SelectedItem?.ToString() ?? "(none)";
                string policy = b.PolicyCombo.SelectedItem?.ToString() ?? "(none)";
                string name = !string.IsNullOrWhiteSpace(b.Device.Name) ? b.Device.Name : b.Device.InstanceId;
                string cls = b.Device.Class ?? string.Empty;
                string usb = b.Device.UsbRoles ?? string.Empty;
                string usbPolling = b.Device.UsbPollingRates ?? string.Empty;
                string audio = b.Device.AudioEndpoints ?? string.Empty;

                WriteLog(
                    $"BOOT.DEV: {b.Device.InstanceId} Kind={b.Kind} Class={cls} Name=\"{name}\" MSI={msi} Prio={prio} Policy={policy} Mask=0x{b.AffinityMask:X} UsbRoles=\"{usb}\" UsbPolling=\"{usbPolling}\" Audio=\"{audio}\"");
            }
        }
    }

    private static string FlattenLogText(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        return value
            .Replace("\r\n", " | ")
            .Replace("\n", " | ")
            .Replace("\r", " | ")
            .Trim();
    }

    private static string GetAppVersion()
    {
        try
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            AssemblyInformationalVersionAttribute? info =
                asm.GetCustomAttribute<AssemblyInformationalVersionAttribute>();
            if (!string.IsNullOrWhiteSpace(info?.InformationalVersion))
            {
                return info.InformationalVersion;
            }

            Version? ver = asm.GetName().Version;
            return ver?.ToString() ?? "unknown";
        }
        catch
        {
            return "unknown";
        }
    }

    private static string FormatIndexList(List<int> values)
    {
        return values.Count == 0 ? "none" : string.Join(',', values);
    }

    private static string SanitizeLogValue(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        return FlattenLogText(value).Replace("\"", "'");
    }

    private static int? TryParseLeadingInt(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        string text = value.Trim();
        int i = 0;
        while (i < text.Length && char.IsDigit(text[i]))
        {
            i++;
        }

        if (i == 0)
        {
            return null;
        }

        return int.TryParse(text[..i], out int result) ? result : null;
    }

    private static int? ParseLimitText(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return 0;
        }

        string text = value.Trim();
        if (string.Equals(text, "0", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }

        if (string.Equals(text, "unlimited", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }

        if (int.TryParse(text, out int parsed))
        {
            return parsed < 0 ? 0 : parsed;
        }

        return null;
    }

    private static ulong BuildUiMask(DeviceBlock block)
    {
        ulong mask = 0;
        for (int i = 0; i < block.CpuBoxes.Count; i++)
        {
            if (block.CpuBoxes[i].Checked)
            {
                mask |= 1UL << i;
            }
        }

        return mask;
    }

    private static int? MapPrioText(string? text)
    {
        return text switch
        {
            "Low" => 1,
            "Normal" => 2,
            "High" => 3,
            "Undefined" => null,
            _ => null,
        };
    }

    private static int? MapPolicyText(string? text)
    {
        return text switch
        {
            "MachineDefault" => 0,
            "All" => 1,
            "Single" => 2,
            "AllClose" => 3,
            "SpecCPU" => 4,
            "SpreadMessages" => 5,
            _ => null,
        };
    }

    private static string FormatRegistryValue(object? value)
    {
        if (value is null)
        {
            return string.Empty;
        }

        return value switch
        {
            string s => s,
            string[] arr => string.Join(";", arr),
            byte[] bytes => string.Join(" ", bytes.Select(b => b.ToString("X2"))),
            int i => i.ToString(),
            uint ui => ui.ToString(),
            long l => l.ToString(),
            ulong ul => ul.ToString(),
            short s16 => s16.ToString(),
            ushort u16 => u16.ToString(),
            byte b => b.ToString(),
            sbyte sb => sb.ToString(),
            _ => value.ToString() ?? string.Empty,
        };
    }

    private static string FormatNumericRegistryValue(object? value)
    {
        if (value is null)
        {
            return string.Empty;
        }

        return value switch
        {
            int i => $"{i} (0x{i:X})",
            uint ui => $"{ui} (0x{ui:X})",
            long l => $"{l} (0x{l:X})",
            ulong ul => $"{ul} (0x{ul:X})",
            short s16 => $"{s16} (0x{s16:X})",
            ushort u16 => $"{u16} (0x{u16:X})",
            byte b => $"{b} (0x{b:X2})",
            sbyte sb => $"{sb} (0x{sb:X2})",
            _ => FormatRegistryValue(value),
        };
    }

    private static string ReadRegValue(RegistryKey? key, string name)
    {
        if (key is null)
        {
            return string.Empty;
        }

        try
        {
            return FormatRegistryValue(key.GetValue(name));
        }
        catch
        {
            return string.Empty;
        }
    }

    private static string ReadRegNumeric(RegistryKey? key, string name)
    {
        if (key is null)
        {
            return string.Empty;
        }

        try
        {
            return FormatNumericRegistryValue(key.GetValue(name));
        }
        catch
        {
            return string.Empty;
        }
    }

    private static int? TryGetRegInt(RegistryKey? key, string name)
    {
        if (key is null)
        {
            return null;
        }

        try
        {
            object? value = key.GetValue(name);
            return value switch
            {
                int i => i,
                uint ui => unchecked((int)ui),
                long l => unchecked((int)l),
                ulong ul => unchecked((int)ul),
                short s16 => s16,
                ushort u16 => u16,
                byte b => b,
                sbyte sb => sb,
                string s when int.TryParse(s, out int parsed) => parsed,
                _ => null,
            };
        }
        catch
        {
            return null;
        }
    }

    private static ulong? TryGetAssignmentMask(RegistryKey? key)
    {
        if (key is null)
        {
            return null;
        }

        try
        {
            object? raw = key.GetValue("AssignmentSetOverride");
            if (raw is byte[] bytes)
            {
                if (bytes.Length >= 8)
                {
                    return BitConverter.ToUInt64(bytes, 0);
                }

                if (bytes.Length >= 4)
                {
                    return BitConverter.ToUInt32(bytes, 0);
                }

                return null;
            }

            return raw switch
            {
                int i => (uint)i,
                uint ui => ui,
                long l => (ulong)l,
                ulong ul => ul,
                _ => null,
            };
        }
        catch
        {
            return null;
        }
    }

    private static string FormatAssignmentOverride(object? value)
    {
        if (value is byte[] bytes)
        {
            string hex = string.Join(" ", bytes.Select(b => b.ToString("X2")));
            if (bytes.Length >= 8)
            {
                ulong mask = BitConverter.ToUInt64(bytes, 0);
                return $"{mask} (0x{mask:X}) bytes=[{hex}]";
            }

            if (bytes.Length >= 4)
            {
                uint mask = BitConverter.ToUInt32(bytes, 0);
                return $"{mask} (0x{mask:X}) bytes=[{hex}]";
            }

            return $"bytes=[{hex}]";
        }

        return FormatNumericRegistryValue(value);
    }

    private static string ExtractImagePath(string rawPath)
    {
        if (string.IsNullOrWhiteSpace(rawPath))
        {
            return string.Empty;
        }

        string text = rawPath.Trim();
        if (text.StartsWith("\"", StringComparison.Ordinal))
        {
            int end = text.IndexOf('"', 1);
            if (end > 1)
            {
                return text.Substring(1, end - 1);
            }
        }

        int space = text.IndexOf(' ');
        if (space > 0)
        {
            text = text[..space];
        }

        return text;
    }

    private static string ExpandSystemRootAlias(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return string.Empty;
        }

        string text = path.Trim();
        if (text.StartsWith(@"\SystemRoot\", StringComparison.OrdinalIgnoreCase))
        {
            string winDir = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
            string tail = text.Substring(@"\SystemRoot\".Length);
            return Path.Combine(winDir, tail);
        }

        return text;
    }

    private static string StripDevicePathPrefix(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return string.Empty;
        }

        string text = path.Trim();
        if (text.StartsWith(@"\??\", StringComparison.OrdinalIgnoreCase))
        {
            text = text.Substring(@"\??\".Length);
        }

        if (text.StartsWith(@"\\?\"))
        {
            text = text.Substring(@"\\?\".Length);
        }

        return text;
    }

    private static string EnsureAbsoluteWindowsPath(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return string.Empty;
        }

        string text = path.Trim();
        if (Path.IsPathRooted(text))
        {
            return text;
        }

        string winDir = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
        string systemDir = Environment.SystemDirectory;
        string trimmed = text.TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

        if (trimmed.StartsWith("System32\\", StringComparison.OrdinalIgnoreCase) ||
            trimmed.StartsWith("SysWOW64\\", StringComparison.OrdinalIgnoreCase))
        {
            return Path.Combine(winDir, trimmed);
        }

        if (trimmed.StartsWith("Drivers\\", StringComparison.OrdinalIgnoreCase) ||
            trimmed.StartsWith("DriverStore\\", StringComparison.OrdinalIgnoreCase))
        {
            return Path.Combine(systemDir, trimmed);
        }

        return Path.Combine(winDir, trimmed);
    }

    private static string ResolveImagePath(string? rawPath)
    {
        if (string.IsNullOrWhiteSpace(rawPath))
        {
            return string.Empty;
        }

        string expanded = Environment.ExpandEnvironmentVariables(rawPath.Trim());
        string extracted = ExtractImagePath(expanded);
        string stripped = StripDevicePathPrefix(extracted);
        string expandedRoot = ExpandSystemRootAlias(stripped);
        return EnsureAbsoluteWindowsPath(expandedRoot);
    }

    private static FileVersionInfo? TryGetFileVersionInfo(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return null;
        }

        try
        {
            if (File.Exists(path))
            {
                return FileVersionInfo.GetVersionInfo(path);
            }
        }
        catch
        {
        }

        return null;
    }

    private static string FormatWmiDate(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        try
        {
            DateTime dt = ManagementDateTimeConverter.ToDateTime(value);
            return dt.ToString("yyyy-MM-dd");
        }
        catch
        {
            return value;
        }
    }

    private static string? GetWmiString(ManagementBaseObject mo, string name)
    {
        try
        {
            return mo[name]?.ToString();
        }
        catch
        {
            return null;
        }
    }

    private static string GetWmiStringArray(ManagementBaseObject mo, string name)
    {
        try
        {
            if (mo[name] is string[] arr)
            {
                return string.Join(";", arr);
            }
        }
        catch
        {
        }

        try
        {
            return mo[name]?.ToString() ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private sealed class SignedDriverInfo
    {
        public string? DeviceId { get; init; }
        public string? DeviceName { get; init; }
        public string? DeviceClass { get; init; }
        public string? ClassGuid { get; init; }
        public string? Manufacturer { get; init; }
        public string? DriverVersion { get; init; }
        public string? DriverDate { get; init; }
        public string? DriverProviderName { get; init; }
        public string? DriverName { get; init; }
        public string? InfName { get; init; }
        public string? FriendlyName { get; init; }
        public string? Description { get; init; }
        public string? Location { get; init; }
        public string? IsSigned { get; init; }
        public string? Signer { get; init; }
        public string? HardwareIds { get; init; }
        public string? CompatibleIds { get; init; }
    }

    private Dictionary<string, SignedDriverInfo> BuildSignedDriverInfoMap(out string? error)
    {
        error = null;
        Dictionary<string, SignedDriverInfo> map = new(StringComparer.OrdinalIgnoreCase);
        try
        {
            using ManagementObjectSearcher searcher = new(
                "root\\CIMV2",
                "SELECT DeviceID, DeviceName, DeviceClass, ClassGuid, Manufacturer, DriverVersion, DriverDate, DriverProviderName, DriverName, InfName, FriendlyName, Description, Location, IsSigned, Signer, HardWareID, CompatID FROM Win32_PnPSignedDriver");

            foreach (ManagementObject mo in searcher.Get())
            {
                string? deviceId = GetWmiString(mo, "DeviceID");
                if (string.IsNullOrWhiteSpace(deviceId))
                {
                    continue;
                }

                string key = NormalizeInstanceId(deviceId);
                SignedDriverInfo info = new()
                {
                    DeviceId = deviceId,
                    DeviceName = GetWmiString(mo, "DeviceName"),
                    DeviceClass = GetWmiString(mo, "DeviceClass"),
                    ClassGuid = GetWmiString(mo, "ClassGuid"),
                    Manufacturer = GetWmiString(mo, "Manufacturer"),
                    DriverVersion = GetWmiString(mo, "DriverVersion"),
                    DriverDate = FormatWmiDate(GetWmiString(mo, "DriverDate")),
                    DriverProviderName = GetWmiString(mo, "DriverProviderName"),
                    DriverName = GetWmiString(mo, "DriverName"),
                    InfName = GetWmiString(mo, "InfName"),
                    FriendlyName = GetWmiString(mo, "FriendlyName"),
                    Description = GetWmiString(mo, "Description"),
                    Location = GetWmiString(mo, "Location"),
                    IsSigned = GetWmiString(mo, "IsSigned"),
                    Signer = GetWmiString(mo, "Signer"),
                    HardwareIds = GetWmiStringArray(mo, "HardWareID"),
                    CompatibleIds = GetWmiStringArray(mo, "CompatID"),
                };

                map[key] = info;
            }
        }
        catch (Exception ex)
        {
            error = ex.Message;
        }

        return map;
    }

    private void LogGuiSnapshot(string reason)
    {
        if (!_detailedLogEnabled)
        {
            return;
        }

        int issueCount = 0;
        Dictionary<string, SignedDriverInfo> signedDriverMap = BuildSignedDriverInfoMap(out string? wmiError);
        bool wmiMapEmpty = signedDriverMap.Count == 0;
        if (!string.IsNullOrWhiteSpace(wmiError))
        {
            WriteLog($"GUI.ISSUE: reason=wmiQueryFailed error=\"{SanitizeLogValue(wmiError)}\"");
            issueCount++;
        }
        else if (wmiMapEmpty)
        {
            WriteLog("GUI.ISSUE: reason=wmiQueryEmpty");
            issueCount++;
        }

        WriteLog($"GUI.WMI: signedDrivers={signedDriverMap.Count}");

        string safeReason = string.IsNullOrWhiteSpace(reason) ? "unknown" : reason.Trim();
        WriteLog($"GUI.SNAPSHOT: start reason={safeReason}");

        string cpuHeader = _cpuHeaderLabel?.Text ?? _cpuHeaderText;
        string htPrefix = _htPrefixLabel?.Text ?? string.Empty;
        string htStatus = _htStatusLabel?.Text ?? string.Empty;
        string hybridCpuStatus = _hybridCpuStatusLabel?.Text ?? string.Empty;
        string cppcStatus = _cppcStatusLabel?.Text ?? string.Empty;
        string dualCcdStatus = _dualCcdStatusLabel?.Text ?? string.Empty;

        WriteLog(
            $"GUI.HEADER: cpuHeader=\"{FlattenLogText(cpuHeader)}\" smt=\"{FlattenLogText(htPrefix)} {FlattenLogText(htStatus)}\" hybridCpu=\"{FlattenLogText(hybridCpuStatus)}\" cppc=\"{FlattenLogText(cppcStatus)}\" dualCcd=\"{FlattenLogText(dualCcdStatus)}\" smtText=\"{FlattenLogText(_smtText)}\"");
        WriteLog(
            $"GUI.STATE: blocks={_blocks.Count} maxLogical={_maxLogical} groupCount={_cpuGroupCount} testCpu={_testCpuActive} testDevicesEnabled={_testDevicesEnabled} testDevicesOnly={_testDevicesOnly} hiddenRealDevices={_testHiddenDeviceIds.Count} autoDryRun={_testAutoDryRun}");
        WriteLog(
            $"GUI.LAYOUT.VIEWPORT: form={FormatGuiBounds(this)} host={FormatGuiBounds(_devicesHost)} panel={FormatGuiBounds(_devicesPanel)} scrollVisible={_devicesScroll.Visible} scrollWidth={_devicesScroll.Width} scrollValue={_devicesScroll.Value} dpi={GetCurrentWindowDpi()}");

        for (int i = 0; i < _blocks.Count; i++)
        {
            DeviceBlock b = _blocks[i];
            string title = BuildDeviceBlockTitle(b.Device);
            string logTitle = b.Kind == DeviceKind.STOR ? $"{title} {StorageAffinityNoteText}" : title;
            WriteLog(
                $"GUI.BLOCK: idx={i} title=\"{FlattenLogText(logTitle)}\" kind={b.Kind} id={b.Device.InstanceId} name=\"{FlattenLogText(b.Device.Name)}\" class=\"{FlattenLogText(b.Device.Class)}\" test={b.Device.IsTestDevice} wifi={b.Device.Wifi} roles=\"{FlattenLogText(b.Device.UsbRoles)}\" usbPolling=\"{FlattenLogText(b.Device.UsbPollingRates)}\" audio=\"{FlattenLogText(b.Device.AudioEndpoints)}\" storage=\"{FlattenLogText(b.Device.StorageTag)}\"");

            List<int> selected = [];
            for (int cpu = 0; cpu < b.CpuBoxes.Count; cpu++)
            {
                if (b.CpuBoxes[cpu].Checked)
                {
                    selected.Add(cpu);
                }
            }

            string msi = b.MsiCombo.SelectedItem?.ToString() ?? "(none)";
            string prio = b.PrioCombo.SelectedItem?.ToString() ?? "(none)";
            string policy = b.PolicyCombo.SelectedItem?.ToString() ?? "(none)";
            string limit = b.LimitBox.Text?.Trim() ?? string.Empty;
            string affinityText = FlattenLogText(b.AffinityLabel.Text);
            string irqText = FlattenLogText(b.IrqLabel.Text);
            string policyLabel = FlattenLogText(b.PolicyLabel.Text);

            WriteLog(
                $"GUI.BLOCK.STATE: idx={i} msi={msi} limit={limit} prio={prio} policy={policy} policyLabel=\"{policyLabel}\" policyEnabled={b.PolicyCombo.Enabled} mask=0x{b.AffinityMask:X} affinityText=\"{affinityText}\" irqText=\"{irqText}\" cpuChecked=[{FormatIndexList(selected)}]");

            string imodValue = b.ImodBox.Text?.Trim() ?? string.Empty;
            string imodDefault = FlattenLogText(b.ImodDefaultLabel.Text);
            string imodCurrent = FlattenLogText(b.ImodCurrentLabel.Text);
            string imodMap = FlattenLogText(b.ImodMapLabel?.Text);
            string imodMapDetail = FlattenLogText(b.ImodMapLabel?.Tag as string);
            WriteLog(
                $"GUI.BLOCK.IMOD: idx={i} visible={b.ImodAutoCheck.Visible} checked={b.ImodAutoCheck.Checked} value=\"{imodValue}\" default=\"{imodDefault}\" current=\"{imodCurrent}\" map=\"{imodMap}\" detail=\"{imodMapDetail}\"");

            if (b.NdisModeCombo is not null)
            {
                string ndisMode = b.NdisModeCombo.SelectedItem?.ToString() ?? string.Empty;
                string rssQueues = b.RssQueueBox is not null ? ((int)b.RssQueueBox.Value).ToString(CultureInfo.InvariantCulture) : string.Empty;
                string activeRss = b.NdisRssRuntime is null ? "Unknown" : FormatNdisRuntimeBool(b.NdisRssRuntime.Enabled);
                WriteLog($"GUI.BLOCK.NDIS: idx={i} mode={ndisMode} rssBase={(b.RssBaseCore?.ToString(CultureInfo.InvariantCulture) ?? string.Empty)} rssQueues={rssQueues} activeRss={activeRss}");
            }

            if (b.NicItrBox is not null)
            {
                string nicItrValue = b.NicItrBox.Text?.Trim() ?? string.Empty;
                string nicItrStatus = FlattenLogText(b.NicItrStatusLabel?.Text);
                string nicItrTime = FlattenLogText(b.NicItrTimeLabel?.Text);
                WriteLog($"GUI.BLOCK.NICITR: idx={i} value=\"{nicItrValue}\" status=\"{nicItrStatus}\" time=\"{nicItrTime}\"");
            }

            string infoText = FlattenLogText(b.InfoLabel.Text);
            string infoReg = FlattenLogText(b.InfoLabel.Tag as string);
            WriteLog($"GUI.BLOCK.INFO: idx={i} text=\"{infoText}\" reg=\"{infoReg}\"");

            issueCount += LogGuiBlockLayout(b, i);
            if (i == 0 && _devicesHost is not null)
            {
                int bottomSlack = _devicesHost.ClientSize.Height - b.Group.Bottom;
                WriteLog($"GUI.LAYOUT.FIRSTFIT: idx=0 blockBottom={b.Group.Bottom} viewportHeight={_devicesHost.ClientSize.Height} bottomSlack={bottomSlack}");
                if (bottomSlack < UiScale(1))
                {
                    WriteLog($"GUI.LAYOUT.ISSUE: idx=0 reason=firstBlockViewport bottomSlack={bottomSlack}");
                    issueCount++;
                }
            }

            if (i > 0)
            {
                DeviceBlock prev = _blocks[i - 1];
                int gap = b.Group.Top - prev.Group.Bottom;
                WriteLog($"GUI.LAYOUT.GAP: prev={i - 1} idx={i} gap={gap}");
                if (gap < UiScale(8))
                {
                    WriteLog($"GUI.LAYOUT.ISSUE: idx={i} reason=blockGap prev={i - 1} gap={gap}");
                    issueCount++;
                }
            }

            issueCount += LogGuiBlockDetails(b, i, signedDriverMap, wmiMapEmpty);
        }

        if (_reservedCpuPanel?.Tag is ReservedCpuPanelTag tag)
        {
            List<int> reserved = tag.Meta
                .Where(m => m.Control.Checked)
                .Select(m => m.Index)
                .OrderBy(x => x)
                .ToList();
            string valueText = tag.ValueLabel.Text;
            byte[] bytes = BuildReservedCpuSetBytes(reserved);
            string bytesText = bytes.Length == 0 ? "none" : string.Join(" ", bytes.Select(b => b.ToString("X2")));
            WriteLog($"GUI.RESERVED: count={tag.Meta.Count} set=[{FormatIndexList(reserved)}] bytes=[{bytesText}] value=\"{SanitizeLogValue(valueText)}\"");
        }
        else
        {
            WriteLog("GUI.RESERVED: none");
        }

        WriteLog($"GUI.ISSUE.SUMMARY: count={issueCount}");
        WriteLog($"GUI.SNAPSHOT: end reason={safeReason}");
    }

    private int LogGuiBlockLayout(DeviceBlock block, int index)
    {
        int issues = 0;
        Control group = block.Group;
        Control? header = block.HeaderPanel;
        Control? divider = block.Divider;
        Control? cpuTitle = block.CpuTitleLabel;
        Control? cpuPanel = block.CpuPanel;
        Control? settingsPanel = block.SettingsPanel;
        Control info = block.InfoLabel;

        int headerAvailable = header?.ClientSize.Width ?? 0;
        int headerPreferred = GetFlowPreferredWidth(header);
        bool headerClip = headerAvailable > 0 && headerPreferred > headerAvailable + UiScale(2);

        (int cpuRight, int cpuBottom) = GetVisibleContentBounds(cpuPanel);
        bool cpuClip = cpuPanel is not null
            && (cpuRight > cpuPanel.ClientSize.Width + UiScale(2)
                || cpuBottom > cpuPanel.ClientSize.Height + UiScale(2));
        int cpuHorizontalSlack = cpuPanel is null ? 0 : cpuPanel.ClientSize.Width - cpuRight;
        bool cpuTight = cpuPanel is not null && cpuHorizontalSlack >= 0 && cpuHorizontalSlack < UiScale(18);
        (int minCpuCellWidth, int maxCpuCellWidth) = GetCpuCellWidthRange(block);
        bool cpuCellMismatch = maxCpuCellWidth > 0 && maxCpuCellWidth - minCpuCellWidth > UiScale(2);

        (int settingsRight, int settingsBottom) = GetVisibleContentBounds(settingsPanel);
        bool settingsClip = settingsPanel is not null
            && (settingsRight > settingsPanel.ClientSize.Width + UiScale(2)
                || settingsBottom > settingsPanel.ClientSize.Height + UiScale(2));
        int settingsHorizontalSlack = settingsPanel is null ? 0 : settingsPanel.ClientSize.Width - settingsRight;
        bool settingsTight = settingsPanel is not null && settingsHorizontalSlack >= 0 && settingsHorizontalSlack < UiScale(12);

        (int groupRight, int groupBottom) = GetVisibleContentBounds(group);
        bool groupClip = groupRight > group.ClientSize.Width + UiScale(2)
            || groupBottom > group.ClientSize.Height + UiScale(2);

        int infoPreferredHeight = GetPreferredTextHeight(info, Math.Max(1, info.ClientSize.Width));
        bool infoClip = infoPreferredHeight > info.ClientSize.Height + UiScale(2);

        string cpuScroll = cpuPanel is ScrollableControl scrollable
            ? $" autoScroll={scrollable.AutoScroll} scrollMin={FormatGuiSize(scrollable.AutoScrollMinSize)}"
            : string.Empty;

        WriteLog(
            $"GUI.LAYOUT: idx={index} group={FormatGuiBounds(group)} header={FormatGuiBounds(header)} headerPref={headerPreferred} headerAvail={headerAvailable} headerClip={headerClip} divider={FormatGuiBounds(divider)} cpuTitle={FormatGuiBounds(cpuTitle)} cpuPanel={FormatGuiBounds(cpuPanel)} cpuContent={cpuRight}x{cpuBottom} cpuSlack={cpuHorizontalSlack} cpuCellWidth={minCpuCellWidth}-{maxCpuCellWidth} cpuCellMismatch={cpuCellMismatch} cpuClip={cpuClip} cpuTight={cpuTight}{cpuScroll} settings={FormatGuiBounds(settingsPanel)} settingsContent={settingsRight}x{settingsBottom} settingsSlack={settingsHorizontalSlack} settingsClip={settingsClip} settingsTight={settingsTight} info={FormatGuiBounds(info)} infoPrefH={infoPreferredHeight} infoClip={infoClip} groupContent={groupRight}x{groupBottom} groupClip={groupClip}");

        List<string> reasons = [];
        if (headerClip)
        {
            reasons.Add($"header:{headerPreferred}>{headerAvailable}");
        }

        if (cpuClip && cpuPanel is not null)
        {
            reasons.Add($"cpu:{cpuRight}x{cpuBottom}>{cpuPanel.ClientSize.Width}x{cpuPanel.ClientSize.Height}");
        }
        else if (cpuTight)
        {
            reasons.Add($"cpuTight:slack={cpuHorizontalSlack}");
        }

        if (cpuCellMismatch)
        {
            reasons.Add($"cpuCellWidth:{minCpuCellWidth}-{maxCpuCellWidth}");
        }

        if (settingsClip && settingsPanel is not null)
        {
            reasons.Add($"settings:{settingsRight}x{settingsBottom}>{settingsPanel.ClientSize.Width}x{settingsPanel.ClientSize.Height}");
        }
        else if (settingsTight)
        {
            reasons.Add($"settingsTight:slack={settingsHorizontalSlack}");
        }

        if (infoClip)
        {
            reasons.Add($"infoH:{infoPreferredHeight}>{info.ClientSize.Height}");
        }

        if (groupClip)
        {
            reasons.Add($"group:{groupRight}x{groupBottom}>{group.ClientSize.Width}x{group.ClientSize.Height}");
        }

        if (reasons.Count > 0)
        {
            WriteLog($"GUI.LAYOUT.ISSUE: idx={index} reason={string.Join(';', reasons)} title=\"{SanitizeLogValue(BuildDeviceBlockTitle(block.Device))}\"");
            issues++;
        }

        return issues;
    }

    private static string FormatGuiBounds(Control? control)
    {
        if (control is null)
        {
            return "null";
        }

        return $"{control.Left},{control.Top},{control.Width}x{control.Height}";
    }

    private static string FormatGuiSize(Size size)
    {
        return $"{size.Width}x{size.Height}";
    }

    private static int GetFlowPreferredWidth(Control? control)
    {
        if (control is null)
        {
            return 0;
        }

        int total = 0;
        foreach (Control child in control.Controls)
        {
            if (!child.Visible)
            {
                continue;
            }

            total += child.PreferredSize.Width + child.Margin.Horizontal;
        }

        return total;
    }

    private static (int Right, int Bottom) GetVisibleContentBounds(Control? control)
    {
        if (control is null)
        {
            return (0, 0);
        }

        int right = 0;
        int bottom = 0;
        foreach (Control child in control.Controls)
        {
            if (!child.Visible)
            {
                continue;
            }

            right = Math.Max(right, child.Right + child.Margin.Right);
            bottom = Math.Max(bottom, child.Bottom + child.Margin.Bottom);
        }

        return (right, bottom);
    }

    private static (int Min, int Max) GetCpuCellWidthRange(DeviceBlock block)
    {
        int min = int.MaxValue;
        int max = 0;
        foreach (CheckBox box in block.CpuBoxes)
        {
            if (!box.Visible)
            {
                continue;
            }

            min = Math.Min(min, box.Width);
            max = Math.Max(max, box.Width);
        }

        return max == 0 ? (0, 0) : (min, max);
    }

    private static int GetPreferredTextHeight(Control control, int width)
    {
        string text = control.Text ?? string.Empty;
        if (string.IsNullOrWhiteSpace(text))
        {
            return control.PreferredSize.Height;
        }

        Size proposed = new(Math.Max(1, width), int.MaxValue);
        return TextRenderer.MeasureText(text, control.Font, proposed, TextFormatFlags.WordBreak).Height;
    }

    private int LogGuiBlockDetails(DeviceBlock block, int index, Dictionary<string, SignedDriverInfo> signedDriverMap, bool wmiMapEmpty)
    {
        int issues = 0;
        try
        {
            string instanceId = block.Device.InstanceId;
            string normalizedId = NormalizeInstanceId(instanceId);
            string shortId = GetShortPnpId(instanceId);
            string parentId = GetParentId(instanceId) ?? string.Empty;

            string enumPath = $@"SYSTEM\CurrentControlSet\Enum\{instanceId}";
            using RegistryKey? enumKey = Registry.LocalMachine.OpenSubKey(enumPath);

            string enumServiceRaw = ReadRegValue(enumKey, "Service");
            string enumDriverRaw = ReadRegValue(enumKey, "Driver");
            string enumClassRaw = ReadRegValue(enumKey, "Class");
            string enumClassGuidRaw = ReadRegValue(enumKey, "ClassGUID");
            string enumMfgRaw = ReadRegValue(enumKey, "Mfg");
            string enumFriendlyRaw = ReadRegValue(enumKey, "FriendlyName");
            string enumDescRaw = ReadRegValue(enumKey, "DeviceDesc");
            string enumLocationRaw = ReadRegValue(enumKey, "LocationInformation");
            string enumLocationPathsRaw = ReadRegValue(enumKey, "LocationPaths");
            string enumHardwareRaw = ReadRegValue(enumKey, "HardwareID");
            string enumCompatibleRaw = ReadRegValue(enumKey, "CompatibleIDs");
            string enumUpperFiltersRaw = ReadRegValue(enumKey, "UpperFilters");
            string enumLowerFiltersRaw = ReadRegValue(enumKey, "LowerFilters");
            string enumContainerRaw = ReadRegValue(enumKey, "ContainerID");
            string enumParentPrefixRaw = ReadRegValue(enumKey, "ParentIdPrefix");
            string enumCapabilitiesRaw = ReadRegNumeric(enumKey, "Capabilities");
            string enumConfigFlagsRaw = ReadRegNumeric(enumKey, "ConfigFlags");
            string enumProblemRaw = ReadRegNumeric(enumKey, "Problem");
            string enumUINumberRaw = ReadRegNumeric(enumKey, "UINumber");

            string classKeyPathRaw = GetClassKeyForDevice(instanceId) ?? string.Empty;
            using RegistryKey? classKey = string.IsNullOrWhiteSpace(classKeyPathRaw)
                ? null
                : Registry.LocalMachine.OpenSubKey(classKeyPathRaw);

            string classDriverDescRaw = ReadRegValue(classKey, "DriverDesc");
            string classProviderRaw = ReadRegValue(classKey, "ProviderName");
            string classDriverVersionRaw = ReadRegValue(classKey, "DriverVersion");
            string classDriverDateRaw = ReadRegValue(classKey, "DriverDate");
            string classInfPathRaw = ReadRegValue(classKey, "InfPath");
            string classInfSectionRaw = ReadRegValue(classKey, "InfSection");
            string classInfSectionExtRaw = ReadRegValue(classKey, "InfSectionExt");
            string classCatalogRaw = ReadRegValue(classKey, "CatalogFile");
            string classMatchIdRaw = ReadRegValue(classKey, "MatchingDeviceId");
            string classClassRaw = ReadRegValue(classKey, "Class");
            string classClassGuidRaw = ReadRegValue(classKey, "ClassGUID");
            string classNetCfgRaw = ReadRegValue(classKey, "NetCfgInstanceId");
            string classComponentIdRaw = ReadRegValue(classKey, "ComponentId");

            if (string.IsNullOrWhiteSpace(enumServiceRaw))
            {
                enumServiceRaw = ReadRegValue(classKey, "Service");
            }

            string serviceKeyPathRaw = string.IsNullOrWhiteSpace(enumServiceRaw)
                ? string.Empty
                : $@"SYSTEM\CurrentControlSet\Services\{enumServiceRaw}";
            using RegistryKey? svcKey = string.IsNullOrWhiteSpace(serviceKeyPathRaw)
                ? null
                : Registry.LocalMachine.OpenSubKey(serviceKeyPathRaw);

            string svcDisplayRaw = ReadRegValue(svcKey, "DisplayName");
            string svcImageRaw = ReadRegValue(svcKey, "ImagePath");
            string svcGroupRaw = ReadRegValue(svcKey, "Group");
            string svcStartRaw = ReadRegNumeric(svcKey, "Start");
            string svcTypeRaw = ReadRegNumeric(svcKey, "Type");
            string svcErrorRaw = ReadRegNumeric(svcKey, "ErrorControl");
            string svcDescRaw = ReadRegValue(svcKey, "Description");
            string svcImageResolved = ResolveImagePath(svcImageRaw);

            FileVersionInfo? drvInfo = TryGetFileVersionInfo(svcImageResolved);

            string intBase = block.Device.RegBase + @"\Device Parameters\Interrupt Management";
            string msiPath = intBase + @"\MessageSignaledInterruptProperties";
            string affPath = intBase + @"\Affinity Policy";

            using RegistryKey? msiKey = Registry.LocalMachine.OpenSubKey(msiPath);
            using RegistryKey? affKey = Registry.LocalMachine.OpenSubKey(affPath);

            string regMsiSupportedRaw = ReadRegNumeric(msiKey, "MSISupported");
            string regMessageLimitRaw = ReadRegNumeric(msiKey, "MessageNumberLimit");
            string regDevicePriorityRaw = ReadRegNumeric(affKey, "DevicePriority");
            string regDevicePolicyRaw = ReadRegNumeric(affKey, "DevicePolicy");
            string regAssignmentRaw = FormatAssignmentOverride(affKey?.GetValue("AssignmentSetOverride"));

            WriteLog($"GUI.BLOCK.REG.PATHS: idx={index} shortId=\"{SanitizeLogValue(shortId)}\" normId=\"{SanitizeLogValue(normalizedId)}\" parentId=\"{SanitizeLogValue(parentId)}\" enumPath=\"HKLM\\{SanitizeLogValue(enumPath)}\" classKey=\"{SanitizeLogValue(classKeyPathRaw)}\" serviceKey=\"{SanitizeLogValue(serviceKeyPathRaw)}\"");
            WriteLog($"GUI.BLOCK.REG.META: idx={index} service=\"{SanitizeLogValue(enumServiceRaw)}\" driverKey=\"{SanitizeLogValue(enumDriverRaw)}\" class=\"{SanitizeLogValue(enumClassRaw)}\" classGuid=\"{SanitizeLogValue(enumClassGuidRaw)}\" mfg=\"{SanitizeLogValue(enumMfgRaw)}\" friendly=\"{SanitizeLogValue(enumFriendlyRaw)}\" desc=\"{SanitizeLogValue(enumDescRaw)}\" location=\"{SanitizeLogValue(enumLocationRaw)}\" parentPrefix=\"{SanitizeLogValue(enumParentPrefixRaw)}\" containerId=\"{SanitizeLogValue(enumContainerRaw)}\"");
            WriteLog($"GUI.BLOCK.REG.HW: idx={index} hardwareIds=[{SanitizeLogValue(enumHardwareRaw)}] compatibleIds=[{SanitizeLogValue(enumCompatibleRaw)}] upperFilters=[{SanitizeLogValue(enumUpperFiltersRaw)}] lowerFilters=[{SanitizeLogValue(enumLowerFiltersRaw)}] locationPaths=[{SanitizeLogValue(enumLocationPathsRaw)}]");
            WriteLog($"GUI.BLOCK.REG.FLAGS: idx={index} capabilities={SanitizeLogValue(enumCapabilitiesRaw)} configFlags={SanitizeLogValue(enumConfigFlagsRaw)} problem={SanitizeLogValue(enumProblemRaw)} uiNumber={SanitizeLogValue(enumUINumberRaw)}");
            WriteLog($"GUI.BLOCK.REG.CLASS: idx={index} driverDesc=\"{SanitizeLogValue(classDriverDescRaw)}\" provider=\"{SanitizeLogValue(classProviderRaw)}\" version=\"{SanitizeLogValue(classDriverVersionRaw)}\" date=\"{SanitizeLogValue(classDriverDateRaw)}\" infPath=\"{SanitizeLogValue(classInfPathRaw)}\" infSection=\"{SanitizeLogValue(classInfSectionRaw)}\" infSectionExt=\"{SanitizeLogValue(classInfSectionExtRaw)}\" catalog=\"{SanitizeLogValue(classCatalogRaw)}\" matchId=\"{SanitizeLogValue(classMatchIdRaw)}\" netCfg=\"{SanitizeLogValue(classNetCfgRaw)}\" componentId=\"{SanitizeLogValue(classComponentIdRaw)}\" class=\"{SanitizeLogValue(classClassRaw)}\" classGuid=\"{SanitizeLogValue(classClassGuidRaw)}\"");
            WriteLog($"GUI.BLOCK.REG.SVC: idx={index} name=\"{SanitizeLogValue(enumServiceRaw)}\" display=\"{SanitizeLogValue(svcDisplayRaw)}\" imagePath=\"{SanitizeLogValue(svcImageRaw)}\" imagePathResolved=\"{SanitizeLogValue(svcImageResolved)}\" group=\"{SanitizeLogValue(svcGroupRaw)}\" start={SanitizeLogValue(svcStartRaw)} type={SanitizeLogValue(svcTypeRaw)} errorControl={SanitizeLogValue(svcErrorRaw)} description=\"{SanitizeLogValue(svcDescRaw)}\"");

            if (!string.IsNullOrWhiteSpace(svcImageResolved))
            {
                if (drvInfo is not null)
                {
                    WriteLog($"GUI.BLOCK.DRVFILE: idx={index} path=\"{SanitizeLogValue(svcImageResolved)}\" fileVersion=\"{SanitizeLogValue(drvInfo.FileVersion)}\" productVersion=\"{SanitizeLogValue(drvInfo.ProductVersion)}\" description=\"{SanitizeLogValue(drvInfo.FileDescription)}\" company=\"{SanitizeLogValue(drvInfo.CompanyName)}\" originalName=\"{SanitizeLogValue(drvInfo.OriginalFilename)}\"");
                }
                else
                {
                    WriteLog($"GUI.BLOCK.DRVFILE: idx={index} path=\"{SanitizeLogValue(svcImageResolved)}\" missing=1");
                }
            }

            WriteLog($"GUI.BLOCK.REG.IM: idx={index} msiSupported={SanitizeLogValue(regMsiSupportedRaw)} messageLimit={SanitizeLogValue(regMessageLimitRaw)} devicePriority={SanitizeLogValue(regDevicePriorityRaw)} devicePolicy={SanitizeLogValue(regDevicePolicyRaw)} assignmentOverride=\"{SanitizeLogValue(regAssignmentRaw)}\"");

            if (block.Kind == DeviceKind.NET_NDIS)
            {
                int? rssBase = GetNdisBaseCore(instanceId);
                string rssText = rssBase.HasValue ? rssBase.Value.ToString() : string.Empty;
                int? rssQueues = GetNdisRssQueues(instanceId);
                string rssQueueText = rssQueues.HasValue ? rssQueues.Value.ToString() : string.Empty;
                string rssBaseGroupRaw = ReadRegNumeric(classKey, "*RssBaseProcGroup");
                string rssMaxProcessorsRaw = ReadRegNumeric(classKey, "*MaxRssProcessors");
                string rssMaxGroupRaw = ReadRegNumeric(classKey, "*RSSMaxProcGroup");
                string rssMaxProcRaw = ReadRegNumeric(classKey, "*RssMaxProcNumber");
                string rssNumaRaw = ReadRegNumeric(classKey, "*NumaNodeId");
                WriteLog($"GUI.BLOCK.RSS: idx={index} baseCore={rssText} queues={rssQueueText} baseGroup={SanitizeLogValue(rssBaseGroupRaw)} maxProcessors={SanitizeLogValue(rssMaxProcessorsRaw)} maxGroup={SanitizeLogValue(rssMaxGroupRaw)} maxProc={SanitizeLogValue(rssMaxProcRaw)} numa={SanitizeLogValue(rssNumaRaw)}");
                NdisRssRuntimeState runtime = GetNdisRssRuntimeState(instanceId);
                block.NdisRssRuntime = runtime;
                WriteLog($"GUI.BLOCK.RSS.ACTIVE: idx={index} {FormatNdisRssRuntimeState(runtime)}");
                LogNdisRssComparison("GUI.BLOCK", instanceId, runtime, rssBase, rssQueues, TryParseLeadingInt(rssMaxProcessorsRaw));
                issues += LogGuiBlockIssues(block, index, normalizedId, signedDriverMap, wmiMapEmpty, enumProblemRaw, svcImageRaw, svcImageResolved, drvInfo, msiKey, affKey, rssBase);
            }
            else
            {
                issues += LogGuiBlockIssues(block, index, normalizedId, signedDriverMap, wmiMapEmpty, enumProblemRaw, svcImageRaw, svcImageResolved, drvInfo, msiKey, affKey, null);
            }

            WriteLog($"GUI.BLOCK.EXTRA: idx={index} usbIsXhci={block.Device.UsbIsXhci} usbHasDevices={block.Device.UsbHasDevices}");

            if (signedDriverMap.TryGetValue(normalizedId, out SignedDriverInfo? signed))
            {
                WriteLog($"GUI.BLOCK.WMI: idx={index} deviceName=\"{SanitizeLogValue(signed.DeviceName)}\" class=\"{SanitizeLogValue(signed.DeviceClass)}\" classGuid=\"{SanitizeLogValue(signed.ClassGuid)}\" manufacturer=\"{SanitizeLogValue(signed.Manufacturer)}\" driverVersion=\"{SanitizeLogValue(signed.DriverVersion)}\" driverDate=\"{SanitizeLogValue(signed.DriverDate)}\" provider=\"{SanitizeLogValue(signed.DriverProviderName)}\" driverName=\"{SanitizeLogValue(signed.DriverName)}\" infName=\"{SanitizeLogValue(signed.InfName)}\" friendly=\"{SanitizeLogValue(signed.FriendlyName)}\" description=\"{SanitizeLogValue(signed.Description)}\" location=\"{SanitizeLogValue(signed.Location)}\" isSigned={SanitizeLogValue(signed.IsSigned)} signer=\"{SanitizeLogValue(signed.Signer)}\" hardwareIds=[{SanitizeLogValue(signed.HardwareIds)}] compatibleIds=[{SanitizeLogValue(signed.CompatibleIds)}]");
            }
            else
            {
                WriteLog($"GUI.BLOCK.WMI: idx={index} missing=1");
            }
        }
        catch (Exception ex)
        {
            WriteLog($"GUI.BLOCK.DETAIL.ERROR: idx={index} error={ex.Message}");
            issues++;
        }

        return issues;
    }

    private int LogGuiBlockIssues(
        DeviceBlock block,
        int index,
        string normalizedId,
        Dictionary<string, SignedDriverInfo> signedDriverMap,
        bool wmiMapEmpty,
        string enumProblemRaw,
        string svcImageRaw,
        string svcImageResolved,
        FileVersionInfo? drvInfo,
        RegistryKey? msiKey,
        RegistryKey? affKey,
        int? rssBase)
    {
        int issues = 0;

        void LogIssue(string reason, string? detail = null)
        {
            if (string.IsNullOrWhiteSpace(detail))
            {
                WriteLog($"GUI.ISSUE: idx={index} id={block.Device.InstanceId} reason={reason}");
            }
            else
            {
                WriteLog($"GUI.ISSUE: idx={index} id={block.Device.InstanceId} reason={reason} {detail}");
            }

            issues++;
        }

        if (block.Device.IsTestDevice)
        {
            WriteLog($"GUI.TEST.BLOCK: idx={index} id={block.Device.InstanceId} registryCheck=skipped");
            return issues;
        }

        int? problemCode = TryParseLeadingInt(enumProblemRaw);
        if (problemCode.HasValue && problemCode.Value != 0)
        {
            LogIssue("cmProblem", $"code={SanitizeLogValue(enumProblemRaw)}");
        }

        if (!string.IsNullOrWhiteSpace(svcImageRaw) && string.IsNullOrWhiteSpace(svcImageResolved))
        {
            LogIssue("svcImageUnresolved", $"raw=\"{SanitizeLogValue(svcImageRaw)}\"");
        }
        else if (!string.IsNullOrWhiteSpace(svcImageResolved) && drvInfo is null)
        {
            LogIssue("driverFileMissing", $"path=\"{SanitizeLogValue(svcImageResolved)}\"");
        }

        if (!wmiMapEmpty && !signedDriverMap.ContainsKey(normalizedId))
        {
            LogIssue("wmiSignedDriverMissing");
        }

        if (block.Kind == DeviceKind.NET_NDIS)
        {
            NdisAffinityMode ndisMode = GetSelectedNdisAffinityMode(block);
            bool expectsRss = ndisMode is NdisAffinityMode.Rss or NdisAffinityMode.Both;
            bool expectsIrq = ndisMode is NdisAffinityMode.IrqPolicy or NdisAffinityMode.Both;

            if (expectsRss && !rssBase.HasValue)
            {
                LogIssue("rssBaseMissing");
            }
            else if (!expectsRss && rssBase.HasValue)
            {
                LogIssue("rssBaseUnexpected", $"mode={FormatNdisAffinityMode(ndisMode)} base={rssBase.Value}");
            }

            int rssQueues = GetNdisRssQueues(block.Device.InstanceId) ?? 1;
            if (rssQueues < 1)
            {
                rssQueues = 1;
            }

            if (rssQueues > _maxLogical)
            {
                rssQueues = _maxLogical;
            }

            int selectedCount = 0;
            for (int i = 0; i < block.CpuBoxes.Count; i++)
            {
                if (!block.CpuBoxes[i].Checked)
                {
                    continue;
                }

                selectedCount++;
            }

            if (expectsRss && selectedCount != rssQueues)
            {
                LogIssue("rssBaseSelection", $"selected={selectedCount} expected={rssQueues}");
            }
            else if (expectsRss && rssBase.HasValue && rssBase.Value >= 0 && rssBase.Value < block.CpuBoxes.Count && !block.CpuBoxes[rssBase.Value].Checked)
            {
                LogIssue("rssBaseMismatch", $"uiMissingBase={rssBase.Value}");
            }

            if (expectsIrq && !block.Device.IsTestDevice)
            {
                int? irqPolicy = TryGetRegInt(affKey, "DevicePolicy");
                ulong irqMask = TryGetAssignmentMask(affKey) ?? 0;
                if (irqPolicy != 4 || irqMask == 0)
                {
                    LogIssue("rssIrqPolicyMissing", $"mode={FormatNdisAffinityMode(ndisMode)} policy={(irqPolicy?.ToString(CultureInfo.InvariantCulture) ?? "missing")} mask=0x{irqMask:X}");
                }
            }

            if (block.PolicyCombo.Enabled)
            {
                LogIssue("rssPolicyEnabled");
            }
        }
        else if (block.Kind == DeviceKind.STOR)
        {
            if (block.PolicyCombo.Enabled)
            {
                LogIssue("storPolicyEnabled");
            }
        }

        string uiMsiText = block.MsiCombo.SelectedItem?.ToString() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(uiMsiText))
        {
            LogIssue("msiUiMissing");
        }

        string uiPrioText = block.PrioCombo.SelectedItem?.ToString() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(uiPrioText))
        {
            LogIssue("prioUiMissing");
        }

        if (block.Kind != DeviceKind.NET_NDIS)
        {
            string uiPolicyText = block.PolicyCombo.SelectedItem?.ToString() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(uiPolicyText))
            {
                LogIssue("policyUiMissing");
            }
        }

        if (!block.Device.IsTestDevice)
        {
            int uiMsi = uiMsiText == "Enabled" ? 1 : 0;
            int? regMsi = TryGetRegInt(msiKey, "MSISupported");
            if (regMsi.HasValue && regMsi.Value != uiMsi)
            {
                LogIssue("msiMismatch", $"ui={SanitizeLogValue(uiMsiText)} reg={regMsi.Value}");
            }
            else if (!regMsi.HasValue && uiMsi != 0)
            {
                LogIssue("msiMismatch", $"ui={SanitizeLogValue(uiMsiText)} reg=missing");
            }

            int? uiLimit = ParseLimitText(block.LimitBox.Text);
            int? regLimitRaw = TryGetRegInt(msiKey, "MessageNumberLimit");
            int regLimit = regLimitRaw ?? 0;
            if (!uiLimit.HasValue)
            {
                LogIssue("msiLimitInvalid", $"ui=\"{SanitizeLogValue(block.LimitBox.Text)}\"");
            }
            else if (uiLimit.Value != regLimit)
            {
                string suffix = regLimitRaw.HasValue ? string.Empty : " regMissing=1";
                LogIssue("msiLimitMismatch", $"ui={uiLimit.Value} reg={regLimit}{suffix}");
            }

            int? uiPrio = MapPrioText(uiPrioText);
            int? regPrioRaw = TryGetRegInt(affKey, "DevicePriority");
            if (uiPrio.HasValue)
            {
                if (!regPrioRaw.HasValue || regPrioRaw.Value != uiPrio.Value)
                {
                    string regText = regPrioRaw.HasValue ? regPrioRaw.Value.ToString(CultureInfo.InvariantCulture) : "missing";
                    LogIssue("prioMismatch", $"ui={SanitizeLogValue(uiPrioText)} reg={regText}");
                }
            }
            else if (string.Equals(uiPrioText, "Undefined", StringComparison.OrdinalIgnoreCase))
            {
                if (regPrioRaw.HasValue)
                {
                    LogIssue("prioMismatch", $"ui=Undefined reg={regPrioRaw.Value}");
                }
            }
            else
            {
                LogIssue("prioUiInvalid", $"ui={SanitizeLogValue(uiPrioText)}");
            }

            if (block.Kind != DeviceKind.NET_NDIS)
            {
                string uiPolicyText = block.PolicyCombo.SelectedItem?.ToString() ?? string.Empty;
                int? uiPolicy = MapPolicyText(uiPolicyText);
                if (uiPolicy is null)
                {
                    LogIssue("policyUiInvalid", $"ui={SanitizeLogValue(uiPolicyText)}");
                }
                else
                {
                    int? regPolicyRaw = TryGetRegInt(affKey, "DevicePolicy");
                    int regPolicy = regPolicyRaw ?? 0;
                    if (regPolicy != uiPolicy.Value)
                    {
                        string suffix = regPolicyRaw.HasValue ? string.Empty : " regMissing=1";
                        LogIssue("policyMismatch", $"ui={SanitizeLogValue(uiPolicyText)} reg={regPolicy}{suffix}");
                    }
                }

                ulong uiMask = BuildUiMask(block);
                if (uiMask != block.AffinityMask)
                {
                    LogIssue("uiMaskOutOfSync", $"ui=0x{uiMask:X} state=0x{block.AffinityMask:X}");
                }

                ulong? regMaskRaw = TryGetAssignmentMask(affKey);
                ulong regMask = regMaskRaw ?? 0;
                if (uiMask != regMask)
                {
                    string suffix = regMaskRaw.HasValue ? string.Empty : " regMissing=1";
                    LogIssue("maskMismatch", $"ui=0x{uiMask:X} reg=0x{regMask:X}{suffix}");
                }
            }
        }

        return issues;
    }

    private void DisableDetailedLog()
    {
        if (!_detailedLogEnabled)
        {
            return;
        }

        WriteLog("LOG: detailed logging DISABLED");
        _detailedLogEnabled = false;
    }
}
