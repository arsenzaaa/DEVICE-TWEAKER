using System.Management;
using System.Runtime.InteropServices;
using System.Runtime.Intrinsics.X86;

namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private CpuInfo? _cpuInfo;
    private readonly Dictionary<int, CpuLpInfo> _cpuLpByIndex = new();
    private readonly Dictionary<int, int> _cpuSetIdByIndex = new();
    private readonly Dictionary<int, int> _cpuIndexByCpuSetId = new();
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
        _cpuInfo = new CpuInfo
        {
            Topology = cpuRaw,
            CcdMap = ccdMap,
        };

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
        _grpHeight = 120 + (_maxLogical * 24) + 160;

        WriteLog($"CPU.SUMMARY: logical={cpuRaw.Logical} physical={cpuRaw.PhysicalCores} groups={_cpuGroupCount} group0={group0Count} maxAffinity={_maxLogical}");
        if (_cpuGroupCount > 1)
        {
            WriteLog($"CPU.GROUPS: using group0 for affinity UI (KAFFINITY max {MaxAffinityBits})");
        }
        WriteLog($"CPU.IDENT: {cpuVendor.Name} | Vendor={cpuVendor.Vendor} | SMT/HT={_smtText}");
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

        string suffix = "P";
        Color textColor = _cpuTextP;
        if (lpInfo.EffClass > 0)
        {
            suffix = "E";
            textColor = _cpuTextE;
        }
        else if (isHyper)
        {
            suffix = "T";
            textColor = _cpuTextSmt;
        }

        string groupText = string.Empty;
        if (_cpuGroupCount > 1)
        {
            string localText = lpInfo.LocalIndex >= 0 ? $"/L{lpInfo.LocalIndex}" : string.Empty;
            groupText = $", G{lpInfo.Group}{localText}";
        }

        cb.Text = $"CPU {lpIndex} ({suffix}, CCD {ccdId}{groupText})";
        cb.AutoSize = true;
        cb.FlatStyle = FlatStyle.Standard;
        cb.UseVisualStyleBackColor = false;
        cb.BackColor = ccdId == 1 ? Color.FromArgb(70, 30, 30) : _bgGroup;
        cb.ForeColor = textColor;
        cb.Padding = new Padding(2, 0, 0, 0);
        cb.Margin = Padding.Empty;
    }
}
