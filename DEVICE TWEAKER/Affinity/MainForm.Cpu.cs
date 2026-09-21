using System.Management;
using System.Diagnostics;
using System.Diagnostics.Eventing.Reader;
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
    private bool _cpuTopologyReliable;
    private bool _cpuSetIdsReliable;
    private int _maxLogical;
    private int _grpHeight;
    private int _cpuGroupCount = 1;
    private string _smtText = string.Empty;
    private string _cpuHeaderText = "CPU: Unknown";
    private static int MaxAffinityBits => IntPtr.Size * 8;

    private void InitializeCpu()
    {
        _cpuTopologyReliable = false;
        _cpuSetIdsReliable = false;
        CpuTopology? cpuRaw = QueryCpuCpuSet();
        if (cpuRaw is null)
        {
            cpuRaw = QueryCpuGlpi();
        }
        if (cpuRaw is null)
        {
            cpuRaw = QueryCpuMinimal();
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
        LoadCppcRatings(cpuRaw);

        _cpuGroupCount = Math.Max(1, cpuRaw.LPs.Select(lp => lp.Group).Distinct().Count());
        _cpuLpByIndex.Clear();
        _cpuSetIdByIndex.Clear();
        _cpuIndexByCpuSetId.Clear();
        foreach (CpuLpInfo lp in cpuRaw.LPs)
        {
            _cpuLpByIndex[lp.LP] = lp;
            if (lp.CpuSetId >= 0)
            {
                _cpuSetIdByIndex[lp.LP] = lp.CpuSetId;
                _cpuIndexByCpuSetId.TryAdd(lp.CpuSetId, lp.LP);
            }
        }

        int group0Count = cpuRaw.LPs.Count(lp => lp.Group == 0);
        if (group0Count <= 0)
        {
            group0Count = cpuRaw.Logical;
        }

        _maxLogical = Math.Min(group0Count, MaxAffinityBits);
        _grpHeight = UiScale(280);

        int ccdCount = ccdMap.Values.Distinct().Count();
        int ccxCount = ccxMap.Values.Distinct().Count();
        WriteLog($"CPU.SUMMARY: logical={cpuRaw.Logical} physical={cpuRaw.PhysicalCores} groups={_cpuGroupCount} ccd={ccdCount} ccx={ccxCount} group0={group0Count} maxAffinity={_maxLogical}");
        if (_cpuGroupCount > 1)
        {
            WriteLog($"CPU.GROUPS: using group0 for affinity UI (KAFFINITY max {MaxAffinityBits})");
        }
        WriteLog($"CPU.IDENT: {cpuVendor.Name} | Vendor={cpuVendor.Vendor} | SMT/HT={_smtText}");
    }

    private void LoadCppcRatings(CpuTopology topology)
    {
        _cppcRatings.Clear();
        _cppcRanks.Clear();
        _cppcEnabled = false;

        try
        {
            List<(int Group, int Number, int Performance)> events = QueryKernelProcessorPowerEvents(Math.Max(topology.Logical * 4, 16));
            if (events.Count == 0)
            {
                WriteLog("CPU.CPPC: no Event ID 55 data");
                return;
            }

            Dictionary<(int Group, int Number), int> lpByGroupAndNumber = topology.LPs
                .Where(lp => lp.Group >= 0 && lp.LocalIndex >= 0)
                .GroupBy(lp => (lp.Group, lp.LocalIndex))
                .ToDictionary(group => group.Key, group => group.First().LP);
            Dictionary<int, int> collected = [];
            foreach (var (group, processor, performance) in events)
            {
                if (lpByGroupAndNumber.TryGetValue((group, processor), out int globalLp))
                {
                    collected.TryAdd(globalLp, performance);
                }

                if (collected.Count >= topology.Logical)
                {
                    break;
                }
            }

            if (collected.Count == 0)
            {
                WriteLog("CPU.CPPC: Event ID 55 present but ratings were not parsed");
                return;
            }

            int[] requiredLps = topology.LPs
                .Where(lp => lp.Group == 0)
                .Select(lp => lp.LP)
                .Distinct()
                .OrderBy(lp => lp)
                .ToArray();
            int[] missingLps = requiredLps.Where(lp => !collected.ContainsKey(lp)).ToArray();
            if (missingLps.Length > 0)
            {
                WriteLog($"CPU.CPPC: disabled, incomplete group0 data parsed={collected.Count} required={requiredLps.Length} missing=[{string.Join(',', missingLps)}]");
                return;
            }

            collected = collected
                .Where(item => requiredLps.Contains(item.Key))
                .ToDictionary(item => item.Key, item => item.Value);

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

    private static List<(int Group, int Number, int Performance)> QueryKernelProcessorPowerEvents(int maxEvents)
    {
        List<(int Group, int Number, int Performance)> events = [];
        try
        {
            string query = "*[System[Provider[@Name='Microsoft-Windows-Kernel-Processor-Power'] and EventID=55]]";
            EventLogQuery logQuery = new("System", PathType.LogName, query)
            {
                ReverseDirection = true,
            };

            using EventLogReader reader = new(logQuery);
            for (int i = 0; i < maxEvents; i++)
            {
                using EventRecord? record = reader.ReadEvent();
                if (record is null)
                {
                    break;
                }

                string xml = record.ToXml();
                if (TryReadEventDataInt(xml, "Number", out int processor)
                    && TryReadEventDataInt(xml, "MaximumPerformancePercent", out int performance))
                {
                    int group = TryReadEventDataInt(xml, "Group", out int parsedGroup) ? parsedGroup : 0;
                    events.Add((group, processor, performance));
                }
            }

            if (events.Count > 0)
            {
                return events;
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"CPU.CPPC: native EventLogReader unavailable, falling back to wevtutil: {ex.Message}");
        }

        try
        {
            string xmlText = QueryKernelProcessorPowerEventsViaWevtutil(maxEvents);
            if (!string.IsNullOrWhiteSpace(xmlText))
            {
                foreach (Match eventMatch in Regex.Matches(xmlText, "<Event\\b.*?</Event>", RegexOptions.Singleline | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                {
                    string eventXml = eventMatch.Value;
                    if (TryReadEventDataInt(eventXml, "Number", out int processor)
                        && TryReadEventDataInt(eventXml, "MaximumPerformancePercent", out int performance))
                    {
                        int group = TryReadEventDataInt(eventXml, "Group", out int parsedGroup) ? parsedGroup : 0;
                        events.Add((group, processor, performance));
                    }
                }
            }
        }
        catch
        {
        }

        return events;
    }

    private static string QueryKernelProcessorPowerEventsViaWevtutil(int maxEvents)
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

        if (!process.Start())
        {
            return string.Empty;
        }

        Task<string> stdout = process.StandardOutput.ReadToEndAsync();
        Task<string> stderr = process.StandardError.ReadToEndAsync();
        if (!process.WaitForExit(2500))
        {
            try
            {
                process.Kill(entireProcessTree: true);
            }
            catch
            {
            }

            try
            {
                _ = process.WaitForExit(1000);
            }
            catch
            {
            }

            try
            {
                _ = Task.WaitAll([stdout, stderr], 1000);
            }
            catch
            {
            }

            return string.Empty;
        }

        try
        {
            _ = Task.WaitAll([stdout, stderr], 1000);
        }
        catch
        {
        }

        return process.ExitCode == 0 && stdout.IsCompletedSuccessfully
            ? stdout.Result
            : string.Empty;
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

        // Windows defines larger EfficiencyClass values as faster and less
        // power-efficient. SMT remains the strongest hybrid hint above, while
        // this branch also works on hybrid CPUs with HT disabled.
        int perfClass = classes.Max();
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

        return false;
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
                int recordSize = Marshal.SizeOf<NativeCpuSet.SystemCpuSetInformation>();
                if (recordSize != 32)
                {
                    WriteLog($"CPU.TOPO: invalid managed CpuSet ABI size={recordSize}, expected=32");
                    return null;
                }

                List<(int Group, int LocalIndex, int Core, int LLC, int NUMA, int EffClass, int CpuSetId)> raw = [];
                while (offset <= len - (sizeof(int) * 2))
                {
                    IntPtr record = IntPtr.Add(buf, offset);
                    int size = Marshal.ReadInt32(record);
                    int type = Marshal.ReadInt32(record, sizeof(int));
                    if (size < sizeof(int) * 2 || size > len - offset)
                    {
                        WriteLog($"CPU.TOPO: invalid CpuSet record offset={offset} size={size} remaining={len - offset}");
                        break;
                    }

                    if (type == 0 && size >= recordSize)
                    {
                        NativeCpuSet.SystemCpuSetInformation item = Marshal.PtrToStructure<NativeCpuSet.SystemCpuSetInformation>(record);
                        raw.Add((
                            Group: item.Group,
                            LocalIndex: item.LogicalProcessorIndex,
                            Core: item.CoreIndex,
                            LLC: item.LastLevelCacheIndex,
                            NUMA: item.NumaNodeIndex,
                            EffClass: item.EfficiencyClass,
                            CpuSetId: checked((int)item.Id)));
                    }

                    offset += size;
                }

                if (raw.Count == 0)
                {
                    WriteLog("CPU.TOPO: CpuSet returned no processor records, falling back to GLPI");
                    return null;
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

                _cpuTopologyReliable = true;
                _cpuSetIdsReliable = true;

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

    private CpuTopology? QueryCpuGlpi()
    {
        int length = 0;
        _ = NativeLogicalProcessor.GetLogicalProcessorInformationEx(
            NativeLogicalProcessor.RelationAll,
            IntPtr.Zero,
            ref length);
        if (length <= 0)
        {
            WriteLog($"CPU.TOPO: GLPIEx size query failed error={Marshal.GetLastWin32Error()}");
            return null;
        }

        IntPtr buffer = Marshal.AllocHGlobal(length);
        try
        {
            if (!NativeLogicalProcessor.GetLogicalProcessorInformationEx(
                    NativeLogicalProcessor.RelationAll,
                    buffer,
                    ref length))
            {
                WriteLog($"CPU.TOPO: GLPIEx query failed error={Marshal.GetLastWin32Error()}");
                return null;
            }

            List<(IntPtr Address, int Relationship, int Size)> records = [];
            int offset = 0;
            while (offset <= length - NativeLogicalProcessor.RecordHeaderSize)
            {
                IntPtr record = IntPtr.Add(buffer, offset);
                int relationship = Marshal.ReadInt32(record);
                int size = Marshal.ReadInt32(record, sizeof(int));
                if (size < NativeLogicalProcessor.RecordHeaderSize || size > length - offset)
                {
                    WriteLog($"CPU.TOPO: invalid GLPIEx record offset={offset} size={size} remaining={length - offset}");
                    return null;
                }

                records.Add((record, relationship, size));
                offset += size;
            }

            if (offset != length)
            {
                WriteLog($"CPU.TOPO: GLPIEx record length mismatch parsed={offset} returned={length}");
                return null;
            }

            Dictionary<(int Group, int Local), GlpiLpBuilder> builders = [];
            Dictionary<int, int> nextCoreByGroup = [];
            foreach (var record in records.Where(r => r.Relationship == NativeLogicalProcessor.RelationProcessorCore))
            {
                byte efficiencyClass = Marshal.ReadByte(record.Address, NativeLogicalProcessor.RecordHeaderSize + 1);
                int groupCount = Marshal.ReadInt16(record.Address, NativeLogicalProcessor.ProcessorGroupCountOffset) & 0xFFFF;
                if (groupCount <= 0)
                {
                    continue;
                }

                if (NativeLogicalProcessor.ProcessorGroupMasksOffset
                    + (groupCount * NativeLogicalProcessor.GroupAffinitySize) > record.Size)
                {
                    WriteLog($"CPU.TOPO: invalid processor-core record size={record.Size} groups={groupCount}");
                    return null;
                }

                for (int maskIndex = 0; maskIndex < groupCount; maskIndex++)
                {
                    IntPtr groupMask = IntPtr.Add(
                        record.Address,
                        NativeLogicalProcessor.ProcessorGroupMasksOffset + (maskIndex * NativeLogicalProcessor.GroupAffinitySize));
                    ulong mask = NativeLogicalProcessor.ReadAffinityMask(groupMask);
                    int group = Marshal.ReadInt16(groupMask, IntPtr.Size) & 0xFFFF;
                    int core = nextCoreByGroup.TryGetValue(group, out int nextCore) ? nextCore : 0;
                    nextCoreByGroup[group] = core + 1;
                    foreach (int local in EnumerateAffinityBits(mask))
                    {
                        builders[(group, local)] = new GlpiLpBuilder(group, local, core, efficiencyClass);
                    }
                }
            }

            if (builders.Count == 0)
            {
                WriteLog("CPU.TOPO: GLPIEx returned no processor-core relationships");
                return null;
            }

            foreach (var record in records.Where(r => r.Relationship is NativeLogicalProcessor.RelationNumaNode or NativeLogicalProcessor.RelationNumaNodeEx))
            {
                int node = Marshal.ReadInt32(record.Address, NativeLogicalProcessor.NumaNodeNumberOffset);
                int groupCount = record.Relationship == NativeLogicalProcessor.RelationNumaNode
                    ? 1
                    : Math.Max(1, Marshal.ReadInt16(record.Address, NativeLogicalProcessor.NumaGroupCountOffset) & 0xFFFF);
                ApplyGlpiMasks(record, NativeLogicalProcessor.NumaGroupMasksOffset, groupCount, builders,
                    builder => builder.Numa = node);
            }

            Dictionary<int, int> nextCacheByGroup = [];
            Dictionary<(int Group, int Local), int> cacheLevelByLp = [];
            foreach (var record in records.Where(r => r.Relationship == NativeLogicalProcessor.RelationCache))
            {
                int level = Marshal.ReadByte(record.Address, NativeLogicalProcessor.CacheLevelOffset);
                int type = Marshal.ReadInt32(record.Address, NativeLogicalProcessor.CacheTypeOffset);
                if (level <= 0 || type is not (0 or 1))
                {
                    continue;
                }

                int groupCount = Math.Max(1, Marshal.ReadInt16(record.Address, NativeLogicalProcessor.CacheGroupCountOffset) & 0xFFFF);
                if (NativeLogicalProcessor.CacheGroupMasksOffset
                    + (groupCount * NativeLogicalProcessor.GroupAffinitySize) > record.Size)
                {
                    WriteLog($"CPU.TOPO: invalid cache record size={record.Size} groups={groupCount}");
                    return null;
                }

                for (int maskIndex = 0; maskIndex < groupCount; maskIndex++)
                {
                    IntPtr groupMask = IntPtr.Add(
                        record.Address,
                        NativeLogicalProcessor.CacheGroupMasksOffset + (maskIndex * NativeLogicalProcessor.GroupAffinitySize));
                    ulong mask = NativeLogicalProcessor.ReadAffinityMask(groupMask);
                    int group = Marshal.ReadInt16(groupMask, IntPtr.Size) & 0xFFFF;
                    int cache = nextCacheByGroup.TryGetValue(group, out int nextCache) ? nextCache : 0;
                    nextCacheByGroup[group] = cache + 1;
                    foreach (int local in EnumerateAffinityBits(mask))
                    {
                        var key = (group, local);
                        if (builders.TryGetValue(key, out GlpiLpBuilder? builder)
                            && (!cacheLevelByLp.TryGetValue(key, out int oldLevel) || level >= oldLevel))
                        {
                            builder.Llc = cache;
                            cacheLevelByLp[key] = level;
                        }
                    }
                }
            }

            List<CpuLpInfo> entries = [];
            int globalIndex = 0;
            foreach (GlpiLpBuilder builder in builders.Values.OrderBy(b => b.Group).ThenBy(b => b.Local))
            {
                entries.Add(new CpuLpInfo(
                    builder.Group,
                    globalIndex++,
                    builder.Core,
                    builder.Llc,
                    builder.Numa,
                    builder.EfficiencyClass,
                    builder.Local,
                    CpuSetId: -1));
            }

            CpuTopology topology = new(entries);
            _cpuTopologyReliable = true;
            _cpuSetIdsReliable = false;
            WriteLog("CPU.TOPO: source=GLPIEx fallback cpuSetIds=unavailable");
            foreach (CpuLpInfo entry in topology.LPs)
            {
                int coreKey = CpuTopology.MakeCoreKey(entry.Group, entry.Core);
                bool smt = topology.ByCore.TryGetValue(coreKey, out List<CpuLpInfo>? siblings) && siblings.Count > 1;
                WriteLog($"CPU.ENTRY: G{entry.Group} L{entry.LP} Local={entry.LocalIndex} Id=n/a Core={entry.Core} SMT={(smt ? 1 : 0)} NUMA={entry.NUMA} LLC={entry.LLC} EffClass={entry.EffClass}");
            }

            return topology;
        }
        catch (Exception ex)
        {
            WriteLog($"CPU.TOPO: GLPIEx fallback failed: {ex.Message}");
            return null;
        }
        finally
        {
            Marshal.FreeHGlobal(buffer);
        }
    }

    private static IEnumerable<int> EnumerateAffinityBits(ulong mask)
    {
        for (int bit = 0; bit < 64; bit++)
        {
            if ((mask & (1UL << bit)) != 0)
            {
                yield return bit;
            }
        }
    }

    private static void ApplyGlpiMasks(
        (IntPtr Address, int Relationship, int Size) record,
        int masksOffset,
        int groupCount,
        IReadOnlyDictionary<(int Group, int Local), GlpiLpBuilder> builders,
        Action<GlpiLpBuilder> apply)
    {
        for (int maskIndex = 0; maskIndex < groupCount; maskIndex++)
        {
            int offset = masksOffset + (maskIndex * NativeLogicalProcessor.GroupAffinitySize);
            if (offset + NativeLogicalProcessor.GroupAffinitySize > record.Size)
            {
                break;
            }

            IntPtr groupMask = IntPtr.Add(record.Address, offset);
            ulong mask = NativeLogicalProcessor.ReadAffinityMask(groupMask);
            int group = Marshal.ReadInt16(groupMask, IntPtr.Size) & 0xFFFF;
            foreach (int local in EnumerateAffinityBits(mask))
            {
                if (builders.TryGetValue((group, local), out GlpiLpBuilder? builder))
                {
                    apply(builder);
                }
            }
        }
    }

    private CpuTopology QueryCpuMinimal()
    {
        int logicalCount = Math.Max(1, Environment.ProcessorCount);
        List<CpuLpInfo> entries = Enumerable.Range(0, logicalCount)
            .Select(index => new CpuLpInfo(0, index, index, 0, 0, 0, index, -1))
            .ToList();
        _cpuTopologyReliable = false;
        _cpuSetIdsReliable = false;
        WriteLog($"CPU.TOPO.ERROR: all topology APIs failed; using display-only minimal map logical={logicalCount}; AUTO affinity disabled");
        return new CpuTopology(entries);
    }

    private sealed class GlpiLpBuilder(int group, int local, int core, int efficiencyClass)
    {
        public int Group { get; } = group;
        public int Local { get; } = local;
        public int Core { get; } = core;
        public int EfficiencyClass { get; } = efficiencyClass;
        public int Llc { get; set; }
        public int Numa { get; set; }
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
            int primaryLp = coreGroup.Min(x => x.LP);
            if (lpInfo.LP != primaryLp)
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
        cb.Font = _blockFont;
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
