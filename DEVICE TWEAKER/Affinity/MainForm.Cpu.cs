using System.Management;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Runtime.Intrinsics.X86;
using System.Text.RegularExpressions;

namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private CpuInfo? _cpuInfo;
    private readonly Dictionary<int, CpuLpInfo> _cpuLpByIndex = new();
    private readonly Dictionary<int, int> _cpuSetIdByIndex = new();
    private readonly Dictionary<int, int> _cpuIndexByCpuSetId = new();
    private readonly HashSet<int> _effClassP = new();
    private readonly HashSet<int> _effClassE = new();
    private readonly Dictionary<int, int> _cppcRatings = new();
    private readonly Dictionary<int, int> _cppcRanks = new();
    private bool _cppcEnabled;
    private int _maxLogical;
    private int _grpHeight;
    private int _cpuGroupCount = 1;
    private string _smtText = string.Empty;
    private string _cpuHeaderText = "CPU: Unknown";
    private static int MaxAffinityBits => IntPtr.Size * 8;

    private void InitializeCpu()
    {
        CpuTopology? cpuRaw = QueryCpuCpuSet();
        if (cpuRaw is null)
        {
            cpuRaw = QueryCpuGlpi();
        }

        CpuVendorInfo cpuVendor = DetectCpuVendor();
        bool htEnabled = cpuRaw.ByCore.Values.Any(g => g.Count > 1);

        _smtText = string.Empty;
        if (cpuVendor.Vendor.Contains("Intel", StringComparison.OrdinalIgnoreCase))
        {
            _smtText = htEnabled ? "Hyper-Threading: Enabled" : "Hyper-Threading: Disabled";
        }
        else if (cpuVendor.Vendor.Contains("AMD", StringComparison.OrdinalIgnoreCase))
        {
            _smtText = htEnabled ? "SMT: Enabled" : "SMT: Disabled";
        }

        _cpuHeaderText = $"CPU: {cpuVendor.Name}";

        Dictionary<int, int> ccdMap = BuildCcdMap(cpuRaw, cpuVendor);
        Dictionary<int, int> ccxMap = BuildCcxMap(cpuRaw);
        _cpuInfo = new CpuInfo
        {
            Topology = cpuRaw,
            CcdMap = ccdMap,
            CcxMap = ccxMap,
        };
        UpdateEfficiencyClassMap(cpuRaw);
        LoadCppcRatings(cpuRaw.Logical);

        _cpuGroupCount = Math.Max(1, cpuRaw.LPs.Select(lp => lp.Group).Distinct().Count());
        _cpuLpByIndex.Clear();
        _cpuSetIdByIndex.Clear();
        _cpuIndexByCpuSetId.Clear();
        foreach (CpuLpInfo lp in cpuRaw.LPs)
        {
            _cpuLpByIndex[lp.LP] = lp;
            int cpuSetId = lp.CpuSetId >= 0 ? lp.CpuSetId : lp.LP;
            _cpuSetIdByIndex[lp.LP] = cpuSetId;
            _cpuIndexByCpuSetId.TryAdd(cpuSetId, lp.LP);
        }

        int group0Count = cpuRaw.LPs.Count(lp => lp.Group == 0);
        if (group0Count <= 0)
        {
            group0Count = cpuRaw.Logical;
        }

        _maxLogical = Math.Min(group0Count, MaxAffinityBits);
        _grpHeight = UiScale(120) + (_maxLogical * UiScale(24)) + UiScale(160);

        int ccdCount = ccdMap.Values.Distinct().Count();
        int ccxCount = ccxMap.Values.Distinct().Count();
        WriteLog($"CPU.SUMMARY: logical={cpuRaw.Logical} physical={cpuRaw.PhysicalCores} groups={_cpuGroupCount} ccd={ccdCount} ccx={ccxCount} group0={group0Count} maxAffinity={_maxLogical}");
        if (_cpuGroupCount > 1)
        {
            WriteLog($"CPU.GROUPS: using group0 for affinity UI (KAFFINITY max {MaxAffinityBits})");
        }
        WriteLog($"CPU.IDENT: {cpuVendor.Name} | Vendor={cpuVendor.Vendor} | SMT/HT={_smtText}");
    }

    private void LoadCppcRatings(int logicalCount)
    {
        _cppcRatings.Clear();
        _cppcRanks.Clear();
        _cppcEnabled = false;

        try
        {
            string xmlText = QueryKernelProcessorPowerEvents(Math.Max(logicalCount * 4, 16));
            if (string.IsNullOrWhiteSpace(xmlText))
            {
                WriteLog("CPU.CPPC: no Event ID 55 data");
                return;
            }

            Dictionary<int, int> collected = [];
            foreach (Match eventMatch in Regex.Matches(xmlText, "<Event\\b.*?</Event>", RegexOptions.Singleline | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
            {
                string eventXml = eventMatch.Value;
                if (!TryReadEventDataInt(eventXml, "Number", out int processor)
                    || !TryReadEventDataInt(eventXml, "MaximumPerformancePercent", out int performance))
                {
                    continue;
                }

                if (processor >= 0 && processor < logicalCount)
                {
                    collected.TryAdd(processor, performance);
                }

                if (collected.Count >= logicalCount)
                {
                    break;
                }
            }

            if (collected.Count == 0)
            {
                WriteLog("CPU.CPPC: Event ID 55 present but ratings were not parsed");
                return;
            }

            List<int> uniqueRatings = collected.Values.Distinct().OrderByDescending(v => v).ToList();
            if (uniqueRatings.Count <= 1)
            {
                WriteLog($"CPU.CPPC: disabled, all parsed cores share rating={uniqueRatings.FirstOrDefault()} count={collected.Count}");
                return;
            }

            int rank = 1;
            foreach (int rating in uniqueRatings)
            {
                foreach (KeyValuePair<int, int> item in collected.Where(kvp => kvp.Value == rating))
                {
                    _cppcRatings[item.Key] = item.Value;
                    _cppcRanks[item.Key] = rank;
                }

                rank++;
            }

            _cppcEnabled = _cppcRanks.Count > 0;
            string ratingsText = string.Join(
                " ",
                _cppcRatings
                    .OrderBy(kvp => kvp.Key)
                    .Select(kvp => $"CPU{kvp.Key}=R{kvp.Value}/#{_cppcRanks[kvp.Key]}"));
            WriteLog($"CPU.CPPC: enabled count={_cppcRanks.Count} {ratingsText}");
        }
        catch (Exception ex)
        {
            WriteLog($"CPU.CPPC: unavailable: {ex.Message}");
            _cppcEnabled = false;
        }
    }

    private bool HasHybridCpu()
    {
        if (_cpuInfo?.Topology is null)
        {
            return false;
        }

        if (_effClassE.Count > 0)
        {
            return true;
        }

        return _cpuInfo.Topology.LPs
            .Select(lp => lp.EffClass)
            .Where(eff => eff >= 0)
            .Distinct()
            .Skip(1)
            .Any();
    }

    private bool HasDualCcdCpu()
    {
        return _cpuInfo?.CcdMap.Values
            .Distinct()
            .Skip(1)
            .Any() == true;
    }

    private bool HasVisibleCcxSplit()
    {
        if (_cpuInfo is null || _cpuInfo.CcxMap.Count == 0)
        {
            return false;
        }

        return _cpuInfo.CcdMap
            .GroupBy(kvp => kvp.Value)
            .Any(group => group
                .Select(kvp => _cpuInfo.CcxMap.TryGetValue(kvp.Key, out int ccx) ? ccx : 0)
                .Distinct()
                .Skip(1)
                .Any());
    }

    private static string QueryKernelProcessorPowerEvents(int maxEvents)
    {
        using Process process = new();
        process.StartInfo = new ProcessStartInfo
        {
            FileName = "wevtutil.exe",
            Arguments = $"qe System /q:\"*[System[Provider[@Name='Microsoft-Windows-Kernel-Processor-Power'] and EventID=55]]\" /c:{maxEvents.ToString(CultureInfo.InvariantCulture)} /rd:true /f:xml",
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
        };

        process.Start();
        string output = process.StandardOutput.ReadToEnd();
        _ = process.StandardError.ReadToEnd();
        if (!process.WaitForExit(2500))
        {
            try
            {
                process.Kill(entireProcessTree: true);
            }
            catch
            {
            }

            return string.Empty;
        }

        return process.ExitCode == 0 ? output : string.Empty;
    }

    private static bool TryReadEventDataInt(string eventXml, string name, out int value)
    {
        value = 0;
        Match match = Regex.Match(
            eventXml,
            $"<Data\\s+Name=['\"]{Regex.Escape(name)}['\"]>(?<value>[^<]+)</Data>",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        return match.Success
            && int.TryParse(match.Groups["value"].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out value);
    }

    private void UpdateEfficiencyClassMap(CpuTopology topo)
    {
        _effClassP.Clear();
        _effClassE.Clear();

        List<(int EffClass, bool HasSmt)> cores = topo.ByCore.Values
            .Where(g => g.Count > 0)
            .Select(g => (g[0].EffClass, g.Count > 1))
            .ToList();
        if (cores.Count == 0)
        {
            return;
        }

        List<int> classes = cores.Select(x => x.EffClass).Distinct().OrderBy(x => x).ToList();
        if (classes.Count == 0)
        {
            return;
        }

        List<int> smtClasses = cores.Where(x => x.HasSmt).Select(x => x.EffClass).Distinct().OrderBy(x => x).ToList();
        List<int> nonSmtClasses = cores.Where(x => !x.HasSmt).Select(x => x.EffClass).Distinct().OrderBy(x => x).ToList();

        if (smtClasses.Count == 1 && nonSmtClasses.Count == 1 && smtClasses[0] != nonSmtClasses[0])
        {
            _effClassP.Add(smtClasses[0]);
            _effClassE.Add(nonSmtClasses[0]);
            WriteLog($"CPU.EFFCLASS: SMT class={smtClasses[0]} NonSMT class={nonSmtClasses[0]} -> P={smtClasses[0]} E={nonSmtClasses[0]}");
            return;
        }

        int perfClass = classes.Contains(0) ? 0 : classes[0];
        _effClassP.Add(perfClass);
        foreach (int cls in classes)
        {
            if (cls != perfClass)
            {
                _effClassE.Add(cls);
            }
        }

        WriteLog($"CPU.EFFCLASS: classes=[{string.Join(',', classes)}] perfClass={perfClass} eClasses=[{string.Join(',', _effClassE)}]");
    }

    private bool IsEfficiencyClass(int effClass)
    {
        if (_effClassP.Count > 0 || _effClassE.Count > 0)
        {
            if (_effClassE.Contains(effClass))
            {
                return true;
            }

            if (_effClassP.Contains(effClass))
            {
                return false;
            }
        }

        return effClass > 0;
    }

    private bool IsEfficiencyCore(CpuLpInfo lpInfo)
    {
        return IsEfficiencyClass(lpInfo.EffClass);
    }

    private CpuTopology? QueryCpuCpuSet()
    {
        try
        {
            _ = NativeCpuSet.GetSystemCpuSetInformation(IntPtr.Zero, 0, out int len, IntPtr.Zero, 0);
            if (len <= 0)
            {
                return null;
            }

            IntPtr buf = Marshal.AllocHGlobal(len);
            try
            {
                bool ok = NativeCpuSet.GetSystemCpuSetInformation(buf, len, out len, IntPtr.Zero, 0);
                if (!ok)
                {
                    return null;
                }

                int offset = 0;
                List<(int Group, int LocalIndex, int Core, int LLC, int NUMA, int EffClass, int CpuSetId)> raw = [];
                while (offset < len)
                {
                    NativeCpuSet.SystemCpuSetInformation item = Marshal.PtrToStructure<NativeCpuSet.SystemCpuSetInformation>(buf + offset);
                    if (item.Size < 1)
                    {
                        break;
                    }

                    raw.Add((
                        Group: item.Group,
                        LocalIndex: item.LogicalProcessorIndex,
                        Core: item.CoreIndex,
                        LLC: item.LastLevelCacheIndex,
                        NUMA: item.NumaNodeIndex,
                        EffClass: item.EfficiencyClass,
                        CpuSetId: item.Id));

                    offset += item.Size;
                }

                List<CpuLpInfo> entries = [];
                int globalIndex = 0;
                foreach (IGrouping<int, (int Group, int LocalIndex, int Core, int LLC, int NUMA, int EffClass, int CpuSetId)> group
                    in raw.GroupBy(x => x.Group).OrderBy(x => x.Key))
                {
                    foreach (var item in group.OrderBy(x => x.LocalIndex).ThenBy(x => x.CpuSetId))
                    {
                        entries.Add(new CpuLpInfo(
                            Group: item.Group,
                            LP: globalIndex,
                            Core: item.Core,
                            LLC: item.LLC,
                            NUMA: item.NUMA,
                            EffClass: item.EffClass,
                            LocalIndex: item.LocalIndex,
                            CpuSetId: item.CpuSetId));
                        globalIndex++;
                    }
                }

                CpuTopology topo = new(entries.OrderBy(x => x.LP).ToList());

                WriteLog("CPU.TOPO: source=CpuSet");
                foreach (CpuLpInfo e in topo.LPs.OrderBy(x => x.LP))
                {
                    int coreKey = CpuTopology.MakeCoreKey(e.Group, e.Core);
                    bool smt = topo.ByCore.TryGetValue(coreKey, out List<CpuLpInfo>? coreGroup) && coreGroup.Count > 1;
                    string localText = e.LocalIndex >= 0 ? $" Local={e.LocalIndex}" : string.Empty;
                    string idText = e.CpuSetId >= 0 ? $" Id={e.CpuSetId}" : string.Empty;
                    WriteLog($"CPU.ENTRY: G{e.Group} L{e.LP}{localText}{idText} Core={e.Core} SMT={smt} NUMA={e.NUMA} LLC={e.LLC} EffClass={e.EffClass}");
                }

                return topo;
            }
            finally
            {
                Marshal.FreeHGlobal(buf);
            }
        }
        catch
        {
            return null;
        }
    }

    private CpuTopology QueryCpuGlpi()
    {
        int envLP = Environment.ProcessorCount;
        List<CpuLpInfo> list = [];
        for (int i = 0; i < envLP; i++)
        {
            list.Add(new CpuLpInfo(
                Group: 0,
                LP: i,
                Core: i,
                LLC: 0,
                NUMA: 0,
                EffClass: 0,
                LocalIndex: i,
                CpuSetId: i));
        }

        WriteLog("CPU.TOPO: source=GLPI (fallback)");
        foreach (CpuLpInfo e in list)
        {
            WriteLog($"CPU.ENTRY: G0 L{e.LP} Local={e.LocalIndex} Id={e.CpuSetId} Core={e.LP} SMT=0 NUMA=0 LLC=0 EffClass=0");
        }

        return new CpuTopology(list);
    }

    private CpuVendorInfo DetectCpuVendor()
    {
        try
        {
            using ManagementObjectSearcher searcher = new(
                "root\\CIMV2",
                "SELECT Name, Caption, Manufacturer FROM Win32_Processor");

            foreach (ManagementObject mo in searcher.Get())
            {
                string name = (mo["Name"] as string) ?? (mo["Caption"] as string) ?? "Unknown";
                string vendor = mo["Manufacturer"] as string ?? "Unknown";
                name = name.Trim();
                vendor = vendor.Trim();
                return new CpuVendorInfo(name, vendor);
            }
        }
        catch
        {
        }

        return new CpuVendorInfo("Unknown", "Unknown");
    }

    private Dictionary<int, int> BuildCcdMap(CpuTopology cpu, CpuVendorInfo? vendorOverride = null)
    {
        Dictionary<int, int> map = new();

        List<KeyValuePair<int, List<CpuLpInfo>>> llcGroups = cpu.ByLLC
            .Where(g => g.Key >= 0)
            .OrderBy(g => g.Key)
            .ToList();

        bool perLpLlc = llcGroups.Count == cpu.Logical && llcGroups.All(g => g.Value.Count == 1);
        if (llcGroups.Count == 0 || perLpLlc)
        {
            foreach (CpuLpInfo lp in cpu.LPs.OrderBy(x => x.LP))
            {
                map.TryAdd(lp.LP, 0);
            }

            if (perLpLlc)
            {
                WriteLog("CPU.CCD: LLC map is per-LP; using single CCD group");
            }

            return map;
        }

        int ccdIndex = 0;
        bool pairCcx = ShouldPairAmdCcxGroups(llcGroups, vendorOverride);
        if (pairCcx)
        {
            WriteLog("CPU.CCD: AMD family 17h detected; pairing LLC groups into CCDs");
            foreach (IGrouping<int, KeyValuePair<int, List<CpuLpInfo>>> group in llcGroups
                .GroupBy(g => ExtractCpuGroupFromLlcKey(g.Key))
                .OrderBy(g => g.Key))
            {
                int pairIndex = 0;
                foreach (KeyValuePair<int, List<CpuLpInfo>> g in group.OrderBy(x => x.Key))
                {
                    int targetCcd = ccdIndex + (pairIndex / 2);
                    foreach (CpuLpInfo lp in g.Value)
                    {
                        map.TryAdd(lp.LP, targetCcd);
                    }

                    pairIndex++;
                }

                ccdIndex += pairIndex / 2;
            }

            return map;
        }

        ccdIndex = 0;
        foreach (KeyValuePair<int, List<CpuLpInfo>> g in llcGroups)
        {
            foreach (CpuLpInfo lp in g.Value)
            {
                map.TryAdd(lp.LP, ccdIndex);
            }

            ccdIndex++;
        }

        return map;
    }

    private Dictionary<int, int> BuildCcxMap(CpuTopology cpu)
    {
        Dictionary<int, int> map = new();

        List<KeyValuePair<int, List<CpuLpInfo>>> llcGroups = cpu.ByLLC
            .Where(g => g.Key >= 0)
            .OrderBy(g => g.Key)
            .ToList();

        bool perLpLlc = llcGroups.Count == cpu.Logical && llcGroups.All(g => g.Value.Count == 1);
        if (llcGroups.Count == 0 || perLpLlc)
        {
            foreach (CpuLpInfo lp in cpu.LPs.OrderBy(x => x.LP))
            {
                map.TryAdd(lp.LP, 0);
            }

            return map;
        }

        int ccxIndex = 0;
        foreach (IGrouping<int, KeyValuePair<int, List<CpuLpInfo>>> group in llcGroups
            .GroupBy(g => ExtractCpuGroupFromLlcKey(g.Key))
            .OrderBy(g => g.Key))
        {
            foreach (KeyValuePair<int, List<CpuLpInfo>> llcGroup in group.OrderBy(x => x.Key))
            {
                foreach (CpuLpInfo lp in llcGroup.Value)
                {
                    map.TryAdd(lp.LP, ccxIndex);
                }

                ccxIndex++;
            }
        }

        foreach (CpuLpInfo lp in cpu.LPs.OrderBy(x => x.LP))
        {
            map.TryAdd(lp.LP, 0);
        }

        return map;
    }

    private static int ExtractCpuGroupFromLlcKey(int llcKey)
    {
        return (llcKey >> 16) & 0xFFFF;
    }

    private bool ShouldPairAmdCcxGroups(
        List<KeyValuePair<int, List<CpuLpInfo>>> llcGroups,
        CpuVendorInfo? vendorOverride)
    {
        CpuVendorInfo vendor = vendorOverride ?? DetectCpuVendor();
        if (!vendor.Vendor.Contains("AMD", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (!TryGetCpuFamilyModel(out int family, out _))
        {
            return false;
        }

        if (family != 0x17)
        {
            return false;
        }

        List<IGrouping<int, KeyValuePair<int, List<CpuLpInfo>>>> groups = llcGroups
            .GroupBy(g => ExtractCpuGroupFromLlcKey(g.Key))
            .ToList();

        if (groups.Count == 0)
        {
            return false;
        }

        return groups.All(g => g.Count() % 2 == 0);
    }

    private static bool TryGetCpuFamilyModel(out int family, out int model)
    {
        family = -1;
        model = -1;

        try
        {
            if (!X86Base.IsSupported)
            {
                return false;
            }

            var regs = X86Base.CpuId(1, 0);
            int eax = regs.Eax;

            int baseFamily = (eax >> 8) & 0xF;
            int baseModel = (eax >> 4) & 0xF;
            int extFamily = (eax >> 20) & 0xFF;
            int extModel = (eax >> 16) & 0xF;

            family = baseFamily == 0xF ? baseFamily + extFamily : baseFamily;
            model = baseModel | (extModel << 4);

            return true;
        }
        catch
        {
            return false;
        }
    }

    private void StyleCpuCheckbox(CheckBox cb, int lpIndex)
    {
        if (_cpuInfo is null)
        {
            return;
        }

        if (!_cpuLpByIndex.TryGetValue(lpIndex, out CpuLpInfo? lpInfo))
        {
            return;
        }

        bool isHyper = false;
        int coreKey = CpuTopology.MakeCoreKey(lpInfo.Group, lpInfo.Core);
        if (_cpuInfo.Topology.ByCore.TryGetValue(coreKey, out List<CpuLpInfo>? coreGroup) && coreGroup.Count > 1)
        {
            if (coreGroup[0].LP != lpInfo.LP)
            {
                isHyper = true;
            }
        }

        int ccdId = _cpuInfo.CcdMap.TryGetValue(lpIndex, out int cid) ? cid : 0;
        int ccxId = _cpuInfo.CcxMap.TryGetValue(lpIndex, out int xid) ? xid : 0;

        string coreType = "P-Core";
        Color textColor = _cpuTextP;
        if (IsEfficiencyCore(lpInfo))
        {
            coreType = "E-Core";
            textColor = _cpuTextE;
        }
        else if (isHyper)
        {
            coreType = "P-Core/HT";
            textColor = _cpuTextSmt;
        }

        bool showCcd = HasDualCcdCpu();
        bool showCcx = HasVisibleCcxSplit();
        List<string> cpuLabelParts = [coreType];
        if (showCcd)
        {
            cpuLabelParts.Add($"CCD{ccdId}");
        }

        if (showCcx)
        {
            cpuLabelParts.Add($"CCX{ccxId}");
        }

        if (_cpuGroupCount > 1)
        {
            string localText = lpInfo.LocalIndex >= 0 ? $"/L{lpInfo.LocalIndex}" : string.Empty;
            cpuLabelParts.Add($"G{lpInfo.Group}{localText}");
        }

        string cppcTooltip = "CPPC: unavailable";
        if (_cppcEnabled && _cppcRanks.TryGetValue(lpIndex, out int rank))
        {
            if (_cppcRatings.TryGetValue(lpIndex, out int rating))
            {
                cpuLabelParts.Add(rank == 1 ? $"R{rating}, Pref" : $"R{rating}, #{rank}");
                cppcTooltip = rank == 1
                    ? $"CPPC: rating {rating}, preferred rank #1"
                    : $"CPPC: rating {rating}, rank #{rank}";
            }
            else
            {
                cpuLabelParts.Add(rank == 1 ? "Pref" : $"#{rank}");
                cppcTooltip = rank == 1 ? "CPPC: preferred rank #1" : $"CPPC: rank #{rank}";
            }
        }

        cb.Text = $"CPU {lpIndex} ({string.Join(", ", cpuLabelParts)})";
        cb.AutoSize = true;
        cb.FlatStyle = FlatStyle.Standard;
        cb.UseVisualStyleBackColor = false;
        cb.BackColor = showCcx
            ? _cpuCcxBackColors[Math.Abs(ccxId) % _cpuCcxBackColors.Length]
            : showCcd && ccdId == 1 ? Color.FromArgb(70, 30, 30) : _bgGroup;
        cb.ForeColor = textColor;
        cb.Padding = new Padding(UiScale(2), 0, 0, 0);
        cb.Margin = Padding.Empty;
        string ccdTooltip = showCcd ? $", CCD {ccdId}" : string.Empty;
        string ccxTooltip = showCcx ? $", CCX {ccxId}" : string.Empty;
        _copyToolTip?.SetToolTip(
            cb,
            $"CPU {lpIndex}: {(IsEfficiencyCore(lpInfo) ? "E-core" : isHyper ? "P-core SMT sibling" : "P-core")}{ccdTooltip}{ccxTooltip}, Group {lpInfo.Group}, Core {lpInfo.Core}, Local {lpInfo.LocalIndex}. {cppcTooltip}");
    }
}
