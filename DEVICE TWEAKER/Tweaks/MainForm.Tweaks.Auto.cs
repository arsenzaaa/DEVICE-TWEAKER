using System.Text.RegularExpressions;

namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private enum AutoAffinityRole
    {
        Gpu,
        InputMouse,
        InputController,
        Keyboard,
        Nic,
        Audio,
        OtherUsb,
        Other,
    }

    private sealed class AutoAffinityPlanSlot
    {
        public required DeviceBlock Block { get; init; }
        public required AutoAffinityRole Role { get; init; }
        public required int Need { get; init; }
        public required int OriginalIndex { get; init; }
        public List<int> Lps { get; } = [];
        public string Reason { get; set; } = string.Empty;
    }

    private sealed class AutoCpuUnit
    {
        public required int Id { get; init; }
        public required int PrimaryLp { get; init; }
        public required List<int> Lps { get; init; }
        public required bool IsECore { get; init; }
        public required int Ccd { get; init; }
        public required int Ccx { get; init; }
        public required int Rating { get; init; }
        public required int Rank { get; init; }

        public bool HasCore0 => Lps.Contains(0);
        public bool HasSmt => Lps.Count > 1;
    }

    private (List<int> P, List<int> E)? GetAutoCpuSets()
    {
        if (_cpuInfo is null)
        {
            return null;
        }

        List<int> primaryP = [];
        List<int> primaryE = [];

        foreach (int core in _cpuInfo.Topology.ByCore.Keys.OrderBy(x => x))
        {
            List<CpuLpInfo> group = _cpuInfo.Topology.ByCore[core];
            CpuLpInfo? primary = group.OrderBy(x => x.LP).FirstOrDefault();
            if (primary is null)
            {
                continue;
            }

            if (IsEfficiencyCore(primary))
            {
                primaryE.Add(primary.LP);
            }
            else
            {
                primaryP.Add(primary.LP);
            }
        }

        return (SortAutoCpuCandidates(primaryP, preferStrongest: true), SortAutoCpuCandidates(primaryE, preferStrongest: false));
    }

    private List<AutoCpuUnit> BuildAutoCpuUnits(IReadOnlyCollection<int>? allowedLps)
    {
        if (_cpuInfo is null)
        {
            return [];
        }

        HashSet<int>? allowed = allowedLps is { Count: > 0 }
            ? new HashSet<int>(allowedLps)
            : null;

        List<AutoCpuUnit> units = [];
        foreach (KeyValuePair<int, List<CpuLpInfo>> pair in _cpuInfo.Topology.ByCore.OrderBy(kvp => kvp.Value.Min(lp => lp.LP)))
        {
            List<CpuLpInfo> members = pair.Value
                .Where(lp => lp.LP >= 0 && lp.LP < _maxLogical)
                .Where(lp => allowed is null || allowed.Contains(lp.LP))
                .OrderBy(lp => lp.LP)
                .ToList();

            if (members.Count == 0)
            {
                continue;
            }

            CpuLpInfo primary = members[0];
            int rating = members
                .Select(lp => _cppcRatings.TryGetValue(lp.LP, out int value) ? value : 0)
                .DefaultIfEmpty(0)
                .Max();
            int rank = members
                .Select(lp => _cppcRanks.TryGetValue(lp.LP, out int value) ? value : int.MaxValue)
                .DefaultIfEmpty(int.MaxValue)
                .Min();
            int ccd = _cpuInfo.CcdMap.TryGetValue(primary.LP, out int ccdValue) ? ccdValue : 0;
            int ccx = _cpuInfo.CcxMap.TryGetValue(primary.LP, out int ccxValue) ? ccxValue : 0;

            units.Add(new AutoCpuUnit
            {
                Id = pair.Key,
                PrimaryLp = primary.LP,
                Lps = members.Select(lp => lp.LP).ToList(),
                IsECore = IsEfficiencyCore(primary),
                Ccd = ccd,
                Ccx = ccx,
                Rating = rating,
                Rank = rank,
            });
        }

        return units;
    }

    private List<AutoCpuUnit> SortAutoCpuUnits(IEnumerable<AutoCpuUnit> units, bool preferStrongest)
    {
        List<AutoCpuUnit> list = units
            .Where(unit => unit.PrimaryLp >= 0 && unit.PrimaryLp < _maxLogical)
            .GroupBy(unit => unit.Id)
            .Select(group => group.First())
            .ToList();

        if (!_cppcEnabled)
        {
            return list.OrderBy(unit => unit.PrimaryLp).ToList();
        }

        List<AutoCpuUnit> sorted = list
            .OrderBy(unit => unit.Rank)
            .ThenByDescending(unit => unit.Rating)
            .ThenBy(unit => unit.PrimaryLp)
            .ToList();

        if (!preferStrongest)
        {
            sorted.Reverse();
        }

        return sorted;
    }

    private string FormatAutoCpuUnits(IEnumerable<AutoCpuUnit> units)
    {
        return string.Join(
            ',',
            units.Select(unit =>
            {
                string smt = unit.HasSmt ? $"/smt=[{string.Join('+', unit.Lps)}]" : string.Empty;
                string type = unit.IsECore ? "E" : "P";
                string ccx = HasVisibleCcxSplit() ? $"/ccx={unit.Ccx}" : string.Empty;
                if (_cppcEnabled && unit.Rank != int.MaxValue)
                {
                    return $"{unit.PrimaryLp}:{type}:R{unit.Rating}/#{unit.Rank}/ccd={unit.Ccd}{ccx}{smt}";
                }

                return $"{unit.PrimaryLp}:{type}/ccd={unit.Ccd}{ccx}{smt}";
            }));
    }

    private List<int> SortAutoCpuCandidates(IEnumerable<int> logicalProcessors, bool preferStrongest)
    {
        List<int> lps = logicalProcessors
            .Where(lp => lp >= 0 && lp < _maxLogical)
            .Distinct()
            .ToList();

        if (!_cppcEnabled)
        {
            return lps.OrderBy(lp => lp).ToList();
        }

        IOrderedEnumerable<int> ordered = lps
            .OrderBy(lp => _cppcRanks.TryGetValue(lp, out int rank) ? rank : int.MaxValue)
            .ThenByDescending(lp => _cppcRatings.TryGetValue(lp, out int rating) ? rating : int.MinValue)
            .ThenBy(lp => lp);

        List<int> sorted = ordered.ToList();
        if (!preferStrongest)
        {
            sorted.Reverse();
        }

        return sorted;
    }

    private string FormatAutoCpuCandidates(IEnumerable<int> logicalProcessors)
    {
        return string.Join(
            ',',
            logicalProcessors.Select(lp =>
            {
                if (_cppcEnabled && _cppcRanks.TryGetValue(lp, out int rank))
                {
                    int rating = _cppcRatings.TryGetValue(lp, out int value) ? value : 0;
                    return $"{lp}:R{rating}/#{rank}";
                }

            return lp.ToString();
        }));
    }

    private int SelectAutoTargetCcd(IReadOnlyList<int> ccdIds, IEnumerable<int> primaryP, IEnumerable<int> primaryE, out string reason)
    {
        reason = "default";
        if (_cpuInfo is null || ccdIds.Count == 0)
        {
            return 0;
        }

        List<int> candidateLps = primaryP
            .Concat(primaryE)
            .Where(lp => lp >= 0 && lp < _maxLogical)
            .Distinct()
            .ToList();

        if (candidateLps.Count == 0)
        {
            candidateLps = _cpuInfo.CcdMap.Keys
                .Where(lp => lp >= 0 && lp < _maxLogical)
                .Distinct()
                .ToList();
        }

        if (_cppcEnabled)
        {
            var scored = ccdIds
                .Select(ccd =>
                {
                    List<int> lps = candidateLps
                        .Where(lp => _cpuInfo.CcdMap.TryGetValue(lp, out int value) && value == ccd)
                        .ToList();

                    int bestRank = lps
                        .Select(lp => _cppcRanks.TryGetValue(lp, out int rank) ? rank : int.MaxValue)
                        .DefaultIfEmpty(int.MaxValue)
                        .Min();
                    int bestRating = lps
                        .Select(lp => _cppcRatings.TryGetValue(lp, out int rating) ? rating : int.MinValue)
                        .DefaultIfEmpty(int.MinValue)
                        .Max();
                    double averageTopRating = lps
                        .Select(lp => _cppcRatings.TryGetValue(lp, out int rating) ? rating : int.MinValue)
                        .Where(rating => rating != int.MinValue)
                        .OrderByDescending(rating => rating)
                        .Take(4)
                        .DefaultIfEmpty(int.MinValue)
                        .Average();

                    return new
                    {
                        Ccd = ccd,
                        BestRank = bestRank,
                        BestRating = bestRating,
                        AverageTopRating = averageTopRating,
                    };
                })
                .OrderBy(item => item.BestRank)
                .ThenByDescending(item => item.BestRating)
                .ThenByDescending(item => item.AverageTopRating)
                .ThenBy(item => item.Ccd)
                .FirstOrDefault();

            if (scored is not null && (scored.BestRank != int.MaxValue || scored.BestRating != int.MinValue))
            {
                reason = $"cppc rank={scored.BestRank} rating={scored.BestRating}";
                return scored.Ccd;
            }
        }

        int fallback = ccdIds.OrderBy(x => x).First();
        reason = "default-lowest-ccd";
        return fallback;
    }

    private static int GetAutoAssignmentPriority(DeviceKind kind)
    {
        return kind switch
        {
            DeviceKind.USB => 1,
            DeviceKind.GPU => 2,
            DeviceKind.NET_NDIS or DeviceKind.NET_CX => 3,
            DeviceKind.AUDIO => 4,
            _ => 5,
        };
    }

    private static AutoAffinityRole GetAutoAffinityRole(DeviceBlock block)
    {
        return block.Kind switch
        {
            DeviceKind.GPU => AutoAffinityRole.Gpu,
            DeviceKind.NET_NDIS or DeviceKind.NET_CX => AutoAffinityRole.Nic,
            DeviceKind.AUDIO => AutoAffinityRole.Audio,
            DeviceKind.USB => GetAutoUsbAffinityRole(block.Device.UsbRoles ?? string.Empty),
            _ => AutoAffinityRole.Other,
        };
    }

    private static AutoAffinityRole GetAutoUsbAffinityRole(string roles)
    {
        if (HasRoleText(roles, "Mouse"))
        {
            return AutoAffinityRole.InputMouse;
        }

        if (HasRoleText(roles, "Gamepad") || HasRoleText(roles, "Controller"))
        {
            return AutoAffinityRole.InputController;
        }

        if (HasRoleText(roles, "Keyboard"))
        {
            return AutoAffinityRole.Keyboard;
        }

        if (HasRoleText(roles, "Audio")
            || HasRoleText(roles, "Microphone")
            || HasRoleText(roles, "Speaker"))
        {
            return AutoAffinityRole.Audio;
        }

        return AutoAffinityRole.OtherUsb;
    }

    private static int GetAutoAffinityRolePriority(AutoAffinityRole role)
    {
        return role switch
        {
            AutoAffinityRole.InputMouse => 0,
            AutoAffinityRole.InputController => 1,
            AutoAffinityRole.Gpu => 2,
            AutoAffinityRole.Nic => 3,
            AutoAffinityRole.Keyboard => 4,
            AutoAffinityRole.Audio => 5,
            AutoAffinityRole.OtherUsb => 6,
            _ => 7,
        };
    }

    private static bool IsInputAffinityRole(AutoAffinityRole role)
    {
        return role is AutoAffinityRole.InputMouse or AutoAffinityRole.InputController or AutoAffinityRole.Keyboard;
    }

    private static string BuildAutoUsbImodRoleProfile(string roles)
    {
        Dictionary<string, uint> roleValues = new(StringComparer.OrdinalIgnoreCase);
        if (HasRoleText(roles, "Mouse"))
        {
            roleValues["Mouse"] = 0x0;
        }

        if (HasRoleText(roles, "Keyboard"))
        {
            roleValues["Keyboard"] = ImodDefaultInterval;
        }

        if (HasRoleText(roles, "Audio") || HasRoleText(roles, "Microphone") || HasRoleText(roles, "Speaker"))
        {
            roleValues["Audio"] = 0xFA0;
        }

        if (HasRoleText(roles, "Gamepad"))
        {
            roleValues["Gamepad"] = ImodDefaultInterval;
        }

        if (HasRoleText(roles, "Webcam"))
        {
            roleValues["Webcam"] = ImodDefaultInterval;
        }

        return roleValues.Count > 0
            ? FormatImodRoleIntervals(roleValues)
            : GetDefaultRoleImodText();
    }

    private static bool HasRoleText(string roles, string role)
    {
        return !string.IsNullOrWhiteSpace(roles)
            && Regex.IsMatch(roles, $@"(?i)\b{Regex.Escape(role)}\b");
    }

    private static string FormatAutoResultRole(AutoAffinityRole role)
    {
        return role switch
        {
            AutoAffinityRole.InputMouse => "Input mouse",
            AutoAffinityRole.InputController => "Input controller",
            AutoAffinityRole.Keyboard => "Keyboard",
            AutoAffinityRole.Gpu => "GPU",
            AutoAffinityRole.Nic => "Network",
            AutoAffinityRole.Audio => "Audio",
            AutoAffinityRole.OtherUsb => "Other USB",
            _ => "Other",
        };
    }

    private static string FormatAutoResultKind(DeviceKind kind)
    {
        return kind switch
        {
            DeviceKind.NET_NDIS => "NET_NDIS",
            DeviceKind.NET_CX => "NET_CX",
            _ => kind.ToString(),
        };
    }

    private static string FormatAutoResultPicker(ThemedDropDownPicker? picker)
    {
        return picker?.SelectedItem?.ToString() ?? "n/a";
    }

    private static string FormatAutoResultDeviceName(DeviceBlock block)
    {
        string name = SanitizeLogValue(block.Device.Name);
        if (string.IsNullOrWhiteSpace(name))
        {
            name = block.Device.InstanceId;
        }

        string roles = SanitizeLogValue(block.Device.UsbRoles);
        if (block.Kind == DeviceKind.USB && !string.IsNullOrWhiteSpace(roles))
        {
            return $"{name} roles=\"{roles}\"";
        }

        return name;
    }

    private void WriteAutoOptimizationResultSummary(
        bool optimizeUsbImod,
        bool usingP,
        IReadOnlyList<int> primaryP,
        IReadOnlyList<int> primaryE,
        IReadOnlyList<int> targetCcdLps,
        IReadOnlyList<DeviceBlock> usbBlocks,
        IReadOnlyList<DeviceBlock> gpuBlocks,
        IReadOnlyList<DeviceBlock> integratedGpuBlocks,
        int netCount,
        IReadOnlyList<DeviceBlock> audioBlocks,
        int storCount,
        IReadOnlyCollection<string> wifiIds,
        IReadOnlyCollection<string> msiOnlyGpuIds,
        IReadOnlyDictionary<string, string> skipReasons,
        IReadOnlyList<AutoAffinityPlanSlot> planSlots,
        IReadOnlyCollection<int> consumedCores,
        IReadOnlyCollection<int> inputShareCores)
    {
        List<int> assignedLps = planSlots
            .SelectMany(slot => slot.Lps)
            .Distinct()
            .OrderBy(lp => lp)
            .ToList();

        List<int> reservedLps = consumedCores
            .Where(lp => !assignedLps.Contains(lp))
            .Distinct()
            .OrderBy(lp => lp)
            .ToList();

        int assigned = planSlots.Count(slot => slot.Lps.Count > 0);
        int skipped = planSlots.Count(slot => slot.Lps.Count == 0)
            + wifiIds.Count
            + msiOnlyGpuIds.Count
            + skipReasons.Count
            + storCount;

        string imodState = optimizeUsbImod
            ? "enabled for eligible XHCI"
            : "skipped by user/no eligible target";

        WriteLog($"AUTO.RESULT.MODE: {(_testAutoDryRun ? "dry-run preview, nothing saved" : "apply mode")} | CPU={(usingP ? "P-cores" : "E-cores")} | primaryP=[{string.Join(',', primaryP)}] primaryE=[{string.Join(',', primaryE)}] targetCCD=[{string.Join(',', targetCcdLps)}]");
        WriteLog($"AUTO.RESULT.DETECTED: USB={usbBlocks.Count} GPU={gpuBlocks.Count} iGPU={integratedGpuBlocks.Count} NET={netCount} AUDIO={audioBlocks.Count} STOR={storCount} WiFi={wifiIds.Count} | USB IMOD={imodState}");

        foreach (AutoAffinityPlanSlot slot in planSlots.Where(slot => slot.Lps.Count > 0))
        {
            DeviceBlock block = slot.Block;
            string lps = string.Join(',', slot.Lps);
            string name = FormatAutoResultDeviceName(block);
            string role = FormatAutoResultRole(slot.Role);
            string kind = FormatAutoResultKind(block.Kind);
            string policy = FormatAutoResultPicker(block.PolicyCombo);
            string msi = FormatAutoResultPicker(block.MsiCombo);
            string prio = FormatAutoResultPicker(block.PrioCombo);
            string extra = string.Empty;

            if (block.Kind == DeviceKind.NET_NDIS)
            {
                string mode = FormatAutoResultPicker(block.NdisModeCombo);
                string queues = block.RssQueueBox?.Value.ToString() ?? "n/a";
                extra = $" | NDIS={mode} queues={queues}";
            }
            else if (block.Kind == DeviceKind.USB && IsUsbImodTarget(block.Device))
            {
                string imod = SanitizeLogValue(block.ImodBox.Text);
                extra = $" | IMOD={(string.IsNullOrWhiteSpace(imod) ? "n/a" : imod)}";
            }

            WriteLog($"AUTO.RESULT.APPLIED: {role} | {kind} | \"{name}\" -> CPU=[{lps}] mask=0x{block.AffinityMask:X} MSI={msi} prio={prio} policy={policy}{extra} | reason={slot.Reason}");
        }

        foreach (AutoAffinityPlanSlot slot in planSlots.Where(slot => slot.Lps.Count == 0))
        {
            DeviceBlock block = slot.Block;
            WriteLog($"AUTO.RESULT.SKIPPED: {FormatAutoResultRole(slot.Role)} | {FormatAutoResultKind(block.Kind)} | \"{FormatAutoResultDeviceName(block)}\" -> no safe CPU core");
        }

        foreach (DeviceBlock block in _blocks.Where(block => wifiIds.Contains(block.Device.InstanceId)))
        {
            string wifiAction = block.Device.IsTestDevice ? "test affinity cleared, MSI/prio only" : "affinity preserved, MSI/prio only";
            WriteLog($"AUTO.RESULT.SKIPPED: WiFi | {FormatAutoResultKind(block.Kind)} | \"{FormatAutoResultDeviceName(block)}\" -> {wifiAction}");
        }

        foreach (DeviceBlock block in _blocks.Where(block => msiOnlyGpuIds.Contains(block.Device.InstanceId)))
        {
            WriteLog($"AUTO.RESULT.SKIPPED: Integrated GPU | {FormatAutoResultKind(block.Kind)} | \"{FormatAutoResultDeviceName(block)}\" -> affinity preserved, MSI only");
        }

        foreach (DeviceBlock block in _blocks.Where(block => skipReasons.ContainsKey(block.Device.InstanceId)))
        {
            WriteLog($"AUTO.RESULT.SKIPPED: {FormatAutoResultKind(block.Kind)} | \"{FormatAutoResultDeviceName(block)}\" -> {skipReasons[block.Device.InstanceId]}");
        }

        if (storCount > 0)
        {
            WriteLog($"AUTO.RESULT.SKIPPED: Storage devices={storCount} -> AUTO does not touch storage affinity");
        }

        WriteLog($"AUTO.RESULT.FINAL: assigned={assigned} skippedOrPreserved={skipped} usedLPs=[{string.Join(',', assignedLps)}] reservedSpacingLPs=[{string.Join(',', reservedLps)}] inputShareLPs=[{string.Join(',', inputShareCores.Distinct().OrderBy(lp => lp))}]");
    }

    private void InvokeAutoOptimization(bool optimizeUsbImod)
    {
        if (_blocks.Count == 0)
        {
            return;
        }

        WriteLog("AUTO: Invoke-AutoOptimization start");
        if (_testAutoDryRun)
        {
            WriteLog("AUTO.THROTTLE: dry-run -> skipped");
        }
        else
        {
            ApplyAutoRawMouseThrottle();
        }

        if (_testAutoDryRun)
        {
            ResetReservedCpuSetsPreview();
        }
        else
        {
            ResetReservedCpuSets();
        }

        (List<int> P, List<int> E)? cpuSets = GetAutoCpuSets();
        if (cpuSets is null)
        {
            return;
        }

        List<int> primaryP = cpuSets.Value.P.Where(lp => lp >= 0 && lp < _maxLogical).ToList();
        List<int> primaryE = cpuSets.Value.E.Where(lp => lp >= 0 && lp < _maxLogical).ToList();
        List<int> performanceP = primaryP.ToList();
        List<int> performanceE = primaryE.ToList();
        List<int> targetCcdLps = [];
        List<int> ccdIdsUnique = [];
        if (_cpuInfo is not null && _cpuInfo.CcdMap.Count > 0)
        {
            ccdIdsUnique = _cpuInfo.CcdMap.Values.Distinct().OrderBy(x => x).ToList();
            if (ccdIdsUnique.Count >= 2)
            {
                int targetCcd = SelectAutoTargetCcd(ccdIdsUnique, primaryP, primaryE, out string targetCcdReason);
                targetCcdLps = _cpuInfo.CcdMap.Where(kvp => kvp.Value == targetCcd).Select(kvp => kvp.Key).OrderBy(x => x).ToList();
                performanceP = primaryP.Where(lp => targetCcdLps.Contains(lp)).ToList();
                performanceE = primaryE.Where(lp => targetCcdLps.Contains(lp)).ToList();

                WriteLog(
                    $"AUTO: CCD map ids=[{string.Join(',', ccdIdsUnique)}] targetCCD={targetCcd} reason=\"{targetCcdReason}\" targetLPs=[{string.Join(',', targetCcdLps)}] performanceP=[{string.Join(',', performanceP)}] performanceE=[{string.Join(',', performanceE)}]");

                if (performanceP.Count + performanceE.Count == 0 && targetCcdLps.Count > 0)
                {
                    performanceP = targetCcdLps.OrderBy(x => x).ToList();
                    WriteLog($"AUTO: CCD fallback -> using all target CCD LPs as performance P: [{string.Join(',', performanceP)}]");
                }
            }
        }

        List<int> allP = [];
        List<int> allE = [];
        if (_cpuInfo is not null)
        {
            allP = _cpuInfo.Topology.LPs
                .Where(lp => !IsEfficiencyCore(lp))
                .Select(lp => lp.LP)
                .OrderBy(x => x)
                .ToList();

            allE = _cpuInfo.Topology.LPs
                .Where(lp => IsEfficiencyCore(lp))
                .Select(lp => lp.LP)
                .OrderBy(x => x)
                .ToList();
        }

        if (primaryP.Count == 0 && allP.Count > 0)
        {
            primaryP = SortAutoCpuCandidates(allP, preferStrongest: true);
            WriteLog($"AUTO: P fallback -> using all P cores from map: [{string.Join(',', primaryP)}]");
        }

        if (primaryE.Count == 0 && allE.Count > 0)
        {
            primaryE = SortAutoCpuCandidates(allE, preferStrongest: false);
        }

        primaryP = SortAutoCpuCandidates(primaryP, preferStrongest: true);
        primaryE = SortAutoCpuCandidates(primaryE, preferStrongest: false);
        if (performanceP.Count == 0 && targetCcdLps.Count == 0)
        {
            performanceP = primaryP.ToList();
        }

        if (performanceE.Count == 0 && targetCcdLps.Count == 0)
        {
            performanceE = primaryE.ToList();
        }

        performanceP = SortAutoCpuCandidates(performanceP, preferStrongest: true);
        performanceE = SortAutoCpuCandidates(performanceE, preferStrongest: false);
        allP = SortAutoCpuCandidates(allP, preferStrongest: true);
        allE = SortAutoCpuCandidates(allE, preferStrongest: false);

        int pCount = performanceP.Count;
        if (pCount <= 0)
        {
            return;
        }

        WriteLog($"AUTO: CPU primary P=[{string.Join(',', primaryP)}] E=[{string.Join(',', primaryE)}] performanceP=[{string.Join(',', performanceP)}] performanceE=[{string.Join(',', performanceE)}]");
        if (_cppcEnabled)
        {
            WriteLog($"AUTO.CPPC: primaryP=[{FormatAutoCpuCandidates(primaryP)}] primaryE=[{FormatAutoCpuCandidates(primaryE)}] performanceP=[{FormatAutoCpuCandidates(performanceP)}] performanceE=[{FormatAutoCpuCandidates(performanceE)}] allP=[{FormatAutoCpuCandidates(allP)}] allE=[{FormatAutoCpuCandidates(allE)}]");
        }
        if (HasVisibleCcxSplit() && _cpuInfo is not null)
        {
            string ccxText = string.Join(
                " | ",
                _cpuInfo.CcxMap
                    .GroupBy(kvp => kvp.Value)
                    .OrderBy(g => g.Key)
                    .Select(g => $"CCX{g.Key}=[{string.Join(',', g.Select(kvp => kvp.Key).OrderBy(x => x))}]"));
            WriteLog($"AUTO.CCX: enabled {ccxText}");
        }

        HashSet<int> primaryPSet = performanceP.ToHashSet();
        HashSet<int> primaryESet = performanceE.ToHashSet();
        List<AutoCpuUnit> allCpuUnits = BuildAutoCpuUnits(null);
        List<AutoCpuUnit> primaryPUnits = SortAutoCpuUnits(
            allCpuUnits.Where(unit => !unit.IsECore && primaryPSet.Contains(unit.PrimaryLp)),
            preferStrongest: true);
        List<AutoCpuUnit> primaryEUnits = SortAutoCpuUnits(
            allCpuUnits.Where(unit => unit.IsECore && primaryESet.Contains(unit.PrimaryLp)),
            preferStrongest: false);

        if (primaryPUnits.Count == 0)
        {
            primaryPUnits = SortAutoCpuUnits(allCpuUnits.Where(unit => !unit.IsECore), preferStrongest: true);
        }

        if (primaryEUnits.Count == 0)
        {
            primaryEUnits = SortAutoCpuUnits(allCpuUnits.Where(unit => unit.IsECore), preferStrongest: false);
        }

        bool usingP = primaryPUnits.Count > 0;
        List<AutoCpuUnit> preferredUnits = usingP ? primaryPUnits : primaryEUnits;
        HashSet<int> preferredUnitIds = preferredUnits.Select(unit => unit.Id).ToHashSet();
        List<AutoCpuUnit> nonTargetBackgroundUnits = targetCcdLps.Count > 0
            ? SortAutoCpuUnits(
                allCpuUnits.Where(unit => !targetCcdLps.Contains(unit.PrimaryLp) && !unit.HasCore0),
                preferStrongest: false)
            : [];
        List<AutoCpuUnit> eBackgroundUnits = SortAutoCpuUnits(
            allCpuUnits.Where(unit => unit.IsECore && !unit.HasCore0),
            preferStrongest: false);
        List<AutoCpuUnit> spareBackgroundUnits = SortAutoCpuUnits(
            allCpuUnits.Where(unit => !preferredUnitIds.Contains(unit.Id) && !unit.HasCore0),
            preferStrongest: false);
        List<AutoCpuUnit> fallbackBackgroundUnits = SortAutoCpuUnits(
            allCpuUnits.Where(unit => !unit.HasCore0),
            preferStrongest: false);
        List<AutoCpuUnit> backgroundUnits = nonTargetBackgroundUnits
            .Concat(eBackgroundUnits)
            .Concat(spareBackgroundUnits)
            .Concat(fallbackBackgroundUnits)
            .DistinctBy(unit => unit.Id)
            .ToList();
        Dictionary<int, AutoCpuUnit> unitById = allCpuUnits
            .GroupBy(unit => unit.Id)
            .ToDictionary(group => group.Key, group => group.First());
        Dictionary<int, AutoCpuUnit> unitByLp = allCpuUnits
            .SelectMany(unit => unit.Lps.Select(lp => new { lp, unit }))
            .GroupBy(item => item.lp)
            .ToDictionary(group => group.Key, group => group.First().unit);
        HashSet<int> consumedCores = new();
        List<int> inputShareCores = [];
        HashSet<int> reservedUnitIds = new();
        Dictionary<int, List<AutoAffinityPlanSlot>> assignedSlotsByUnit = new();
        List<int> inputShareUnitIds = [];

        foreach (IGrouping<DeviceKind, DeviceBlock> g in _blocks.GroupBy(b => b.Kind))
        {
            WriteLog($"AUTO: blocks kind={g.Key} count={g.Count()}");
        }

        HashSet<string> wifiIds = new(StringComparer.OrdinalIgnoreCase);
        foreach (DeviceBlock nb in _blocks.Where(b => b.Kind is DeviceKind.NET_NDIS or DeviceKind.NET_CX))
        {
            if (nb.Device.Wifi)
            {
                wifiIds.Add(nb.Device.InstanceId);
                WriteLog($"AUTO.WIFI.DETECT: {nb.Device.InstanceId} name=\"{nb.Device.Name}\"");
            }
        }

        List<DeviceBlock> usbBlocks = _blocks.Where(b => b.Kind == DeviceKind.USB).ToList();
        List<DeviceBlock> netNdisBlocks = _blocks.Where(b => b.Kind == DeviceKind.NET_NDIS && !wifiIds.Contains(b.Device.InstanceId)).ToList();
        List<DeviceBlock> netCxBlocks = _blocks.Where(b => b.Kind == DeviceKind.NET_CX && !wifiIds.Contains(b.Device.InstanceId)).ToList();
        List<DeviceBlock> gpuBlocks = _blocks.Where(b => b.Kind == DeviceKind.GPU).ToList();
        List<DeviceBlock> integratedGpuBlocks = gpuBlocks.Where(b => b.Device.IsIntegratedGpu).ToList();
        List<DeviceBlock> audioBlocks = _blocks.Where(b => b.Kind == DeviceKind.AUDIO).ToList();
        int storCount = _blocks.Count(b => b.Kind == DeviceKind.STOR);
        HashSet<string> skipAutoIds = new(StringComparer.OrdinalIgnoreCase);
        Dictionary<string, string> skipReasons = new(StringComparer.OrdinalIgnoreCase);
        HashSet<string> msiOnlyGpuIds = integratedGpuBlocks
            .Select(b => b.Device.InstanceId)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        foreach (DeviceBlock usbBlock in usbBlocks)
        {
            if (string.IsNullOrWhiteSpace(usbBlock.Device.UsbRoles))
            {
                skipAutoIds.Add(usbBlock.Device.InstanceId);
                skipReasons[usbBlock.Device.InstanceId] = "USB controller has no detected HID roles, left for manual/reset only";
                WriteLog($"AUTO.SKIP.USB: {usbBlock.Device.InstanceId} no HID roles (manual/reset only)");
            }
        }

        int netCount = netNdisBlocks.Count + netCxBlocks.Count;
        bool hasWiFiOnly = wifiIds.Count > 0 && netCount == 0;
        WriteLog(
            $"AUTO.SUMMARY: GPU={gpuBlocks.Count} GPUi={integratedGpuBlocks.Count} NET={netCount} USB={usbBlocks.Count} AUDIO={audioBlocks.Count} STOR={storCount} WIFI={wifiIds.Count} WiFiOnly={hasWiFiOnly} targetCCD=[{string.Join(',', targetCcdLps)}] primaryP=[{string.Join(',', primaryP)}] primaryE=[{string.Join(',', primaryE)}] performanceP=[{string.Join(',', performanceP)}] performanceE=[{string.Join(',', performanceE)}]");
        if (hasWiFiOnly)
        {
            WriteLog("AUTO.WIFI-ONLY: no wired NET adapters found; skipping NET affinity.");
        }

        foreach (DeviceBlock audioBlock in audioBlocks)
        {
            string pnpId = audioBlock.Device.InstanceId;
            string desc = audioBlock.Device.Name;
            string audioText = audioBlock.Device.AudioEndpoints;

            bool isSpdif = IsSpdifAudioEndpointsText(audioText);
            bool isDisplay = IsDisplayHdmiaudio(pnpId, desc) || IsDisplayAudioEndpointsText(audioText);
            if (!isDisplay && !isSpdif)
            {
                continue;
            }

            skipAutoIds.Add(pnpId);
            string reason = isSpdif ? "digital S/PDIF audio" : "display/HDMI audio";
            skipReasons[pnpId] = $"{reason}, left untouched by AUTO";
            WriteLog($"AUTO.SKIP.AUDIO: {pnpId} classified as {reason} (name=\"{desc}\" endpoints=\"{audioText}\")");
        }

        foreach (DeviceBlock block in _blocks)
        {
            bool isSkipAuto = skipAutoIds.Contains(block.Device.InstanceId);
            bool isMsiOnlyGpu = msiOnlyGpuIds.Contains(block.Device.InstanceId);
            bool isWifi = wifiIds.Contains(block.Device.InstanceId);
            ulong beforeMask = block.AffinityMask;
            string beforePolicy = block.PolicyCombo.SelectedItem?.ToString() ?? "(none)";
            WriteLog($"AUTO.RESET: {block.Device.InstanceId} Kind={block.Kind} maskBefore=0x{beforeMask:X} policyBefore={beforePolicy}");

            if (isMsiOnlyGpu)
            {
                string msiBefore = block.MsiCombo.SelectedItem?.ToString() ?? "(none)";
                block.MsiCombo.SelectedItem = "Enabled";
                string msiAfter = block.MsiCombo.SelectedItem?.ToString() ?? "(none)";
                WriteLog($"AUTO.SKIP.GPU: {block.Device.InstanceId} integrated=1 msiBefore={msiBefore} msiAfter={msiAfter} reason=integratedGpuAutoSkip");
                continue;
            }

            if (isWifi)
            {
                if (block.Device.IsTestDevice)
                {
                    block.SuppressCpuEvents++;
                    try
                    {
                        foreach (CheckBox cb in block.CpuBoxes)
                        {
                            cb.Checked = false;
                        }
                    }
                    finally
                    {
                        block.SuppressCpuEvents--;
                    }

                    block.AffinityMask = 0;
                    block.RssBaseCore = null;
                }

                block.MsiCombo.SelectedItem = "Enabled";
                block.PrioCombo.SelectedItem = "High";
                string affinityState = block.Device.IsTestDevice ? "test affinity cleared" : "affinity/limit preserved";
                WriteLog($"AUTO.WIFI.SKIP: {block.Device.InstanceId} -> MSI=Enabled Prio=High ({affinityState})");
                continue;
            }

            block.SuppressCpuEvents++;
            try
            {
                foreach (CheckBox cb in block.CpuBoxes)
                {
                    cb.Checked = false;
                }
            }
            finally
            {
                block.SuppressCpuEvents--;
            }

            block.AffinityMask = 0;

            if (isSkipAuto)
            {
                block.MsiCombo.SelectedItem = "Enabled";
                block.LimitBox.Text = "0";
                block.PrioCombo.SelectedItem = "Undefined";
                if (block.Kind != DeviceKind.NET_NDIS && block.PolicyCombo.Enabled)
                {
                    block.PolicyCombo.SelectedItem = "MachineDefault";
                }
            }
            else
            {
                block.MsiCombo.SelectedItem = "Enabled";
                block.LimitBox.Text = "0";
                block.PrioCombo.SelectedItem = "High";
                if (block.Kind != DeviceKind.NET_NDIS && block.PolicyCombo.Enabled)
                {
                    block.PolicyCombo.SelectedItem = "MachineDefault";
                }
            }

            RecalcAffinityMask(block);

            ulong afterMask = block.AffinityMask;
            string afterPolicy = block.PolicyCombo.SelectedItem?.ToString() ?? "(none)";
            WriteLog($"AUTO.RESET: {block.Device.InstanceId} Kind={block.Kind} maskAfter=0x{afterMask:X} policyAfter={afterPolicy} skipAuto={isSkipAuto}");
            if (isSkipAuto)
            {
                string reason = IsSpdifAudioEndpointsText(block.Device.AudioEndpoints) ? "digital S/PDIF audio" : "display/HDMI audio";
                WriteLog($"AUTO.RESET.SKIP: {block.Device.InstanceId} Kind={block.Kind} reason={reason}");
            }
        }

        List<DeviceBlock> imodTargets = usbBlocks.Where(b => IsUsbImodTarget(b.Device)).ToList();
        if (!optimizeUsbImod)
        {
            WriteLog("AUTO.IMOD: skipped by user choice");
        }
        else if (imodTargets.Count == 0)
        {
            WriteLog("AUTO.IMOD: no eligible XHCI controllers -> skipping");
        }
        else
        {
            foreach (DeviceBlock block in imodTargets)
            {
                string roles = block.Device.UsbRoles ?? string.Empty;
                bool hasKeyboard = Regex.IsMatch(roles, "(?i)\\bKeyboard\\b");
                bool hasMouse = Regex.IsMatch(roles, "(?i)\\bMouse\\b");

                string before = block.ImodBox.Text?.Trim() ?? string.Empty;
                string roleProfile = BuildAutoUsbImodRoleProfile(roles);
                block.ImodBox.Text = roleProfile;
                block.ImodAutoCheck.Checked = true;
                UpdateImodSelectorsFromText(block);
                WriteLog(
                    $"AUTO.IMOD: {block.Device.InstanceId} -> role-profile {roleProfile} (roles=\"{roles}\", hasMouse={hasMouse}, hasKeyboard={hasKeyboard}, prev={before})");
            }
        }

        List<AutoAffinityPlanSlot> planSlots = _blocks
            .Where(b => b.Kind != DeviceKind.STOR)
            .Where(b => !wifiIds.Contains(b.Device.InstanceId))
            .Where(b => !skipAutoIds.Contains(b.Device.InstanceId))
            .Where(b => !msiOnlyGpuIds.Contains(b.Device.InstanceId))
            .Select((block, index) => new { block, index })
            .Select(x => new AutoAffinityPlanSlot
            {
                Block = x.block,
                Role = GetAutoAffinityRole(x.block),
                Need = x.block.Kind == DeviceKind.GPU ? 2 : 1,
                OriginalIndex = x.index,
            })
            .OrderBy(x => GetAutoAffinityRolePriority(x.Role))
            .ThenBy(x => x.OriginalIndex)
            .ToList();

        WriteLog($"AUTO.PLAN.CPU: using={(usingP ? "P" : "E")} preferredUnits=[{FormatAutoCpuUnits(preferredUnits)}] backgroundUnits=[{FormatAutoCpuUnits(backgroundUnits)}] nonTargetBackground=[{FormatAutoCpuUnits(nonTargetBackgroundUnits)}] eBackground=[{FormatAutoCpuUnits(eBackgroundUnits)}]");
        WriteLog($"AUTO.PLAN.SLOTS: [{string.Join(" | ", planSlots.Select(s => $"{s.Role}:{s.Block.Kind}:{s.Block.Device.InstanceId}"))}]");

        int inputPreferredNeed = planSlots.Count(s => s.Role is AutoAffinityRole.InputMouse or AutoAffinityRole.InputController);
        int gpuPreferredNeed = planSlots.Where(s => s.Role == AutoAffinityRole.Gpu).Sum(s => s.Need);
        int nicPreferredNeed = planSlots.Count(s => s.Role == AutoAffinityRole.Nic);
        int keyboardPreferredNeed = planSlots.Count(s => s.Role == AutoAffinityRole.Keyboard);
        int audioSlots = planSlots.Count(s => s.Role == AutoAffinityRole.Audio);
        int audioPreferredNeed = Math.Max(0, audioSlots - backgroundUnits.Count);
        int requiredPreferredUnitCount = inputPreferredNeed + gpuPreferredNeed + nicPreferredNeed + keyboardPreferredNeed + audioPreferredNeed;
        int nonCore0PreferredUnits = preferredUnits.Count(unit => !unit.HasCore0);
        bool avoidCore0WhenPossible = nonCore0PreferredUnits >= Math.Min(requiredPreferredUnitCount, preferredUnits.Count);
        int spacingEligibleUnits = avoidCore0WhenPossible ? nonCore0PreferredUnits : preferredUnits.Count;
        int spacingSkips = Math.Max(0, spacingEligibleUnits - Math.Min(spacingEligibleUnits, requiredPreferredUnitCount));
        WriteLog($"AUTO.PLAN.UNITS: required={requiredPreferredUnitCount} available={preferredUnits.Count} nonCore0={nonCore0PreferredUnits} avoidCore0={avoidCore0WhenPossible} spacingEligible={spacingEligibleUnits} spacing={spacingSkips}");

        List<AutoCpuUnit> ApplyCcxPreference(List<AutoCpuUnit> ordered, int? preferCcx, int? avoidCcx)
        {
            if (!HasVisibleCcxSplit() || ordered.Count <= 1)
            {
                return ordered;
            }

            if (preferCcx.HasValue && ordered.Any(unit => unit.Ccx == preferCcx.Value))
            {
                return ordered
                    .Where(unit => unit.Ccx == preferCcx.Value)
                    .Concat(ordered.Where(unit => unit.Ccx != preferCcx.Value))
                    .ToList();
            }

            if (avoidCcx.HasValue && ordered.Any(unit => unit.Ccx != avoidCcx.Value))
            {
                return ordered
                    .Where(unit => unit.Ccx != avoidCcx.Value)
                    .Concat(ordered.Where(unit => unit.Ccx == avoidCcx.Value))
                    .ToList();
            }

            return ordered;
        }

        int? GetInputAnchorCcx()
        {
            if (!HasVisibleCcxSplit())
            {
                return null;
            }

            foreach (int unitId in inputShareUnitIds)
            {
                if (unitById.TryGetValue(unitId, out AutoCpuUnit? unit))
                {
                    return unit.Ccx;
                }
            }

            return null;
        }

        List<AutoCpuUnit> GetCandidateUnits(
            IEnumerable<AutoCpuUnit> units,
            bool allowCore0,
            bool preferWeak = false,
            int? preferCcx = null,
            int? avoidCcx = null)
        {
            List<AutoCpuUnit> ordered = units
                .Where(unit => !reservedUnitIds.Contains(unit.Id))
                .ToList();
            if (preferWeak)
            {
                ordered.Reverse();
            }

            if (!allowCore0)
            {
                return ApplyCcxPreference(ordered.Where(unit => !unit.HasCore0).ToList(), preferCcx, avoidCcx);
            }

            if (!avoidCore0WhenPossible)
            {
                return ApplyCcxPreference(ordered, preferCcx, avoidCcx);
            }

            return ApplyCcxPreference(ordered.Where(unit => !unit.HasCore0).ToList(), preferCcx, avoidCcx)
                .Concat(ApplyCcxPreference(ordered.Where(unit => unit.HasCore0).ToList(), preferCcx, avoidCcx))
                .ToList();
        }

        void ReserveUnit(AutoCpuUnit unit, string reason)
        {
            reservedUnitIds.Add(unit.Id);
            consumedCores.Add(unit.PrimaryLp);
            WriteLog($"AUTO.PLAN.UNIT.RESERVE: LP={unit.PrimaryLp} unit={unit.Id} ccd={unit.Ccd} ccx={unit.Ccx} type={(unit.IsECore ? "E" : "P")} reason={reason}");
        }

        AutoCpuUnit? TakeDedicatedUnit(
            bool allowCore0,
            bool preferWeak = false,
            bool useBackground = false,
            int? preferCcx = null,
            int? avoidCcx = null,
            string? reserveReason = null)
        {
            List<AutoCpuUnit> source = useBackground ? backgroundUnits : preferredUnits;
            AutoCpuUnit? unit = GetCandidateUnits(source, allowCore0, preferWeak, preferCcx, avoidCcx).FirstOrDefault();
            if (unit is not null)
            {
                ReserveUnit(unit, reserveReason ?? (useBackground ? "dedicated-background" : preferWeak ? "dedicated-weak" : "dedicated"));
            }

            return unit;
        }

        List<AutoCpuUnit> TakeDedicatedUnits(int count, bool allowCore0, bool preferWeak = false)
        {
            List<AutoCpuUnit> selected = GetCandidateUnits(preferredUnits, allowCore0, preferWeak)
                .Take(count)
                .ToList();

            if (selected.Count < count)
            {
                return [];
            }

            foreach (AutoCpuUnit unit in selected)
            {
                ReserveUnit(unit, "dedicated-group");
            }

            return selected;
        }

        List<AutoCpuUnit> TakeGpuDedicatedUnits(int count, bool allowCore0)
        {
            if (count != 2)
            {
                return TakeDedicatedUnits(count, allowCore0);
            }

            List<AutoCpuUnit> candidates = GetCandidateUnits(preferredUnits, allowCore0)
                .OrderBy(unit => unit.Id)
                .ToList();
            if (candidates.Count < count)
            {
                return [];
            }

            List<(AutoCpuUnit First, AutoCpuUnit Second)> pairs = [];
            for (int i = 0; i < candidates.Count - 1; i++)
            {
                AutoCpuUnit first = candidates[i];
                AutoCpuUnit second = candidates[i + 1];
                if (Math.Abs(first.Id - second.Id) != 1)
                {
                    continue;
                }

                pairs.Add((first, second));
            }

            if (pairs.Count == 0)
            {
                return TakeDedicatedUnits(count, allowCore0);
            }

            (AutoCpuUnit First, AutoCpuUnit Second) selectedPair = _cppcEnabled
                ? pairs
                    .OrderBy(pair => pair.First.Ccd == pair.Second.Ccd ? 0 : 1)
                    .ThenBy(pair => !HasVisibleCcxSplit() || pair.First.Ccx == pair.Second.Ccx ? 0 : 1)
                    .ThenBy(pair => Math.Min(pair.First.Rank, pair.Second.Rank))
                    .ThenBy(pair => (long)pair.First.Rank + pair.Second.Rank)
                    .ThenByDescending(pair => (long)pair.First.Rating + pair.Second.Rating)
                    .ThenByDescending(pair => Math.Max(pair.First.PrimaryLp, pair.Second.PrimaryLp))
                    .First()
                : pairs
                    .OrderBy(pair => pair.First.Ccd == pair.Second.Ccd ? 0 : 1)
                    .ThenBy(pair => !HasVisibleCcxSplit() || pair.First.Ccx == pair.Second.Ccx ? 0 : 1)
                    .ThenByDescending(pair => Math.Max(pair.First.PrimaryLp, pair.Second.PrimaryLp))
                    .ThenByDescending(pair => Math.Min(pair.First.PrimaryLp, pair.Second.PrimaryLp))
                    .First();

            List<AutoCpuUnit> selected = [selectedPair.First, selectedPair.Second];
            foreach (AutoCpuUnit unit in selected)
            {
                ReserveUnit(unit, "dedicated-gpu-pair");
            }

            return selected;
        }

        AutoCpuUnit? FindInputShareUnit()
        {
            foreach (int unitId in inputShareUnitIds)
            {
                if (unitById.TryGetValue(unitId, out AutoCpuUnit? unit))
                {
                    return unit;
                }
            }

            return null;
        }

        static int SelectUnitLp(AutoCpuUnit unit, bool preferSibling)
        {
            if (preferSibling && unit.Lps.Count > 1)
            {
                int sibling = unit.Lps.FirstOrDefault(lp => lp != unit.PrimaryLp);
                if (sibling >= 0)
                {
                    return sibling;
                }
            }

            return unit.PrimaryLp;
        }

        string FormatBackgroundUnitReason(string roleName, AutoCpuUnit unit)
        {
            if (targetCcdLps.Count == 0)
            {
                return $"{roleName}-background-weak-unit";
            }

            return targetCcdLps.Contains(unit.PrimaryLp)
                ? $"{roleName}-background-weak-fallback"
                : $"{roleName}-background-non-target-unit";
        }

        void AssignSlot(AutoAffinityPlanSlot slot, IEnumerable<AutoCpuUnit> units, string reason, bool preferSiblingForShared = false)
        {
            slot.Lps.Clear();
            List<AutoCpuUnit> selectedUnits = units
                .Where(unit => unit.PrimaryLp >= 0 && unit.PrimaryLp < _maxLogical)
                .DistinctBy(unit => unit.Id)
                .ToList();

            slot.Lps.AddRange(selectedUnits
                .Select(unit => SelectUnitLp(unit, preferSiblingForShared))
                .Where(lp => lp >= 0 && lp < _maxLogical)
                .Distinct());
            slot.Reason = reason;

            foreach (AutoCpuUnit unit in selectedUnits)
            {
                if (!assignedSlotsByUnit.TryGetValue(unit.Id, out List<AutoAffinityPlanSlot>? slots))
                {
                    slots = [];
                    assignedSlotsByUnit[unit.Id] = slots;
                }

                if (!slots.Contains(slot))
                {
                    slots.Add(slot);
                }
            }

            if (slot.Lps.Count > 0 && IsInputAffinityRole(slot.Role))
            {
                foreach (int lp in slot.Lps)
                {
                    if (!inputShareCores.Contains(lp))
                    {
                        inputShareCores.Add(lp);
                    }
                }

                foreach (AutoCpuUnit unit in selectedUnits)
                {
                    if (!inputShareUnitIds.Contains(unit.Id))
                    {
                        inputShareUnitIds.Add(unit.Id);
                    }
                }
            }

            foreach (int lp in slot.Lps)
            {
                consumedCores.Add(lp);
            }
        }

        void SkipSpacingUnitIfAvailable(string afterReason)
        {
            if (spacingSkips <= 0)
            {
                return;
            }

            AutoCpuUnit? unit = GetCandidateUnits(preferredUnits, allowCore0: false).FirstOrDefault();
            if (unit is null)
            {
                return;
            }

            ReserveUnit(unit, $"spacing-after-{afterReason}");
            spacingSkips--;
            WriteLog($"AUTO.PLAN.SPACING: after={afterReason} reservedUnit={unit.Id} LP={unit.PrimaryLp} remaining={spacingSkips}");
        }

        foreach (AutoAffinityPlanSlot slot in planSlots.Where(s => s.Role is AutoAffinityRole.InputMouse or AutoAffinityRole.InputController))
        {
            int? inputCcx = GetInputAnchorCcx();
            AutoCpuUnit? unit = TakeDedicatedUnit(allowCore0: true, preferCcx: inputCcx);
            if (unit is null)
            {
                WriteLog($"AUTO.PLAN.SKIP: role={slot.Role} {slot.Block.Kind} {slot.Block.Device.InstanceId} reason=no-input-unit");
                continue;
            }

            AssignSlot(slot, [unit], inputCcx.HasValue && unit.Ccx == inputCcx.Value ? "input-dedicated-unit-ccx-near" : "input-dedicated-unit");
            if (slot.Lps.Count > 0)
            {
                SkipSpacingUnitIfAvailable("input");
            }
        }

        foreach (AutoAffinityPlanSlot slot in planSlots.Where(s => s.Role == AutoAffinityRole.Gpu))
        {
            List<AutoCpuUnit> units = TakeGpuDedicatedUnits(slot.Need, allowCore0: true);
            if (units.Count < slot.Need)
            {
                WriteLog($"AUTO.PLAN.SKIP: role={slot.Role} {slot.Block.Kind} {slot.Block.Device.InstanceId} reason=no-gpu-physical-pair");
                continue;
            }

            AssignSlot(slot, units, "gpu-dedicated-high-physical-pair");
            if (slot.Lps.Count > 0)
            {
                SkipSpacingUnitIfAvailable("gpu");
            }
        }

        foreach (AutoAffinityPlanSlot slot in planSlots.Where(s => s.Role == AutoAffinityRole.Nic))
        {
            int? inputCcx = GetInputAnchorCcx();
            AutoCpuUnit? unit = TakeDedicatedUnit(allowCore0: true, avoidCcx: inputCcx);
            if (unit is null)
            {
                WriteLog($"AUTO.PLAN.SKIP: role={slot.Role} NET {slot.Block.Device.InstanceId} reason=no-safe-nic-unit");
                continue;
            }

            AssignSlot(slot, [unit], inputCcx.HasValue && unit.Ccx != inputCcx.Value ? "nic-dedicated-unit-ccx-away-from-input" : "nic-dedicated-unit");
            if (slot.Lps.Count > 0)
            {
                SkipSpacingUnitIfAvailable("nic");
            }
        }

        foreach (AutoAffinityPlanSlot slot in planSlots.Where(s => s.Role == AutoAffinityRole.Keyboard))
        {
            int? inputCcx = GetInputAnchorCcx();
            AutoCpuUnit? unit = TakeDedicatedUnit(allowCore0: true, preferCcx: inputCcx);
            if (unit is not null)
            {
                AssignSlot(slot, [unit], inputCcx.HasValue && unit.Ccx == inputCcx.Value ? "keyboard-dedicated-unit-ccx-near-input" : "keyboard-dedicated-unit");
                if (slot.Lps.Count > 0)
                {
                    SkipSpacingUnitIfAvailable("keyboard");
                }
            }
            else
            {
                AutoCpuUnit? shareUnit = FindInputShareUnit();
                if (shareUnit is not null)
                {
                    AssignSlot(slot, [shareUnit], "keyboard-shared-with-input-unit", preferSiblingForShared: true);
                }
                else
                {
                    WriteLog($"AUTO.PLAN.SKIP: role={slot.Role} USB {slot.Block.Device.InstanceId} reason=no-keyboard-unit");
                }
            }
        }

        foreach (AutoAffinityPlanSlot slot in planSlots.Where(s => s.Role == AutoAffinityRole.Audio))
        {
            AutoCpuUnit? backgroundUnit = TakeDedicatedUnit(
                allowCore0: false,
                useBackground: true,
                avoidCcx: GetInputAnchorCcx(),
                reserveReason: "dedicated-background-audio");
            if (backgroundUnit is not null)
            {
                AssignSlot(slot, [backgroundUnit], FormatBackgroundUnitReason("audio", backgroundUnit));
                continue;
            }

            AutoCpuUnit? weakUnit = TakeDedicatedUnit(allowCore0: true, preferWeak: true, avoidCcx: GetInputAnchorCcx());
            if (weakUnit is not null)
            {
                int? inputCcx = GetInputAnchorCcx();
                AssignSlot(slot, [weakUnit], inputCcx.HasValue && weakUnit.Ccx != inputCcx.Value ? "audio-dedicated-weak-unit-ccx-away-from-input" : "audio-dedicated-weak-unit");
            }
            else
            {
                WriteLog($"AUTO.PLAN.SKIP: role={slot.Role} AUDIO {slot.Block.Device.InstanceId} reason=no-audio-unit");
            }
        }

        foreach (AutoAffinityPlanSlot slot in planSlots.Where(s => s.Role is AutoAffinityRole.OtherUsb or AutoAffinityRole.Other))
        {
            AutoCpuUnit? unit = TakeDedicatedUnit(
                allowCore0: false,
                useBackground: true,
                reserveReason: "dedicated-background-other");
            if (unit is not null)
            {
                AssignSlot(slot, [unit], FormatBackgroundUnitReason("other", unit));
            }
            else
            {
                AutoCpuUnit? weakUnit = TakeDedicatedUnit(allowCore0: false, preferWeak: true);
                if (weakUnit is not null)
                {
                    AssignSlot(slot, [weakUnit], "other-dedicated-weak-unit");
                }
                else
                {
                    WriteLog($"AUTO.PLAN.SKIP: role={slot.Role} {slot.Block.Kind} {slot.Block.Device.InstanceId} reason=no-other-unit");
                }
            }
        }

        void BlockSlot(AutoAffinityPlanSlot slot, string reason)
        {
            if (slot.Lps.Count == 0)
            {
                return;
            }

            WriteLog($"AUTO.VALIDATION.BLOCK: role={slot.Role} {slot.Block.Kind} {slot.Block.Device.InstanceId} lps=[{string.Join(',', slot.Lps)}] reason={reason}");
            slot.Lps.Clear();
            slot.Reason = $"validation-blocked: {reason}";
        }

        foreach (AutoAffinityPlanSlot slot in planSlots.Where(s => s.Lps.Count > 0))
        {
            if (slot.Lps.Any(lp => lp < 0 || lp >= _maxLogical || !unitByLp.ContainsKey(lp)))
            {
                BlockSlot(slot, "invalid logical processor");
                continue;
            }

            if (slot.Role == AutoAffinityRole.Gpu)
            {
                List<AutoCpuUnit> gpuUnits = slot.Lps
                    .Select(lp => unitByLp[lp])
                    .DistinctBy(unit => unit.Id)
                    .ToList();
                if (slot.Lps.Count < slot.Need || gpuUnits.Count < slot.Need)
                {
                    BlockSlot(slot, "GPU requires two different physical P-core units");
                    continue;
                }

                if (gpuUnits.Any(unit => unit.IsECore))
                {
                    BlockSlot(slot, "GPU assigned to E-core unit");
                }
            }
        }

        foreach (KeyValuePair<int, List<AutoAffinityPlanSlot>> pair in assignedSlotsByUnit)
        {
            List<AutoAffinityPlanSlot> activeSlots = pair.Value.Where(slot => slot.Lps.Count > 0).ToList();
            if (activeSlots.Count <= 1)
            {
                continue;
            }

            bool hasGpu = activeSlots.Any(slot => slot.Role == AutoAffinityRole.Gpu);
            bool hasInput = activeSlots.Any(slot => slot.Role is AutoAffinityRole.InputMouse or AutoAffinityRole.InputController);
            bool hasNicOrAudio = activeSlots.Any(slot => slot.Role is AutoAffinityRole.Nic or AutoAffinityRole.Audio);

            if (hasGpu)
            {
                foreach (AutoAffinityPlanSlot slot in activeSlots.Where(slot => slot.Role != AutoAffinityRole.Gpu))
                {
                    BlockSlot(slot, "sharing with GPU unit is not allowed");
                }
            }

            if (hasInput && hasNicOrAudio)
            {
                foreach (AutoAffinityPlanSlot slot in activeSlots.Where(slot => slot.Role is AutoAffinityRole.Nic or AutoAffinityRole.Audio))
                {
                    BlockSlot(slot, "NIC/audio sharing with input unit is not allowed");
                }
            }
        }

        WriteLog($"AUTO.VALIDATION: assigned={planSlots.Count(slot => slot.Lps.Count > 0)} blocked={planSlots.Count(slot => slot.Reason.StartsWith("validation-blocked", StringComparison.OrdinalIgnoreCase))}");

        foreach (AutoAffinityPlanSlot slot in planSlots)
        {
            DeviceBlock block = slot.Block;
            List<int> lps = slot.Lps;
            if (lps.Count == 0)
            {
                string label = block.Kind == DeviceKind.NET_NDIS || block.Kind == DeviceKind.NET_CX ? "NET" : block.Kind.ToString();
                WriteLog($"AUTO.PLAN.SKIP: role={slot.Role} {label} {block.Device.InstanceId} reason=no-safe-core");
                continue;
            }

            block.SuppressCpuEvents++;
            try
            {
                foreach (CheckBox cb in block.CpuBoxes)
                {
                    cb.Checked = false;
                }

                foreach (int lpVal in lps)
                {
                    if (lpVal >= 0 && lpVal < block.CpuBoxes.Count)
                    {
                        block.CpuBoxes[lpVal].Checked = true;
                    }
                }
            }
            finally
            {
                block.SuppressCpuEvents--;
            }

            int? ndisAutoQueues = null;
            if (block.Kind == DeviceKind.NET_NDIS)
            {
                block.RssBaseCore = lps[0];
                NdisRssRuntimeState runtime = GetNdisRssRuntimeState(block.Device.InstanceId);
                block.NdisRssRuntime = runtime;
                if (block.RssQueueBox is not null)
                {
                    block.SuppressCpuEvents++;
                    try
                    {
                        block.RssQueueBox.Value = 1;
                    }
                    finally
                    {
                        block.SuppressCpuEvents--;
                    }

                    ndisAutoQueues = (int)block.RssQueueBox.Value;
                }

                NdisAffinityMode smartMode = ChooseSmartNdisAffinityMode(block, ndisAutoQueues ?? 1, out string ndisReason);
                SetNdisModeCombo(block, smartMode);
                WriteLog($"AUTO.NDIS.MODE: {block.Device.InstanceId} mode={FormatNdisAffinityMode(smartMode)} reason=\"{ndisReason}\" {FormatNdisRssRuntimeState(runtime)}");
            }

            if (block.Kind != DeviceKind.NET_NDIS && block.PolicyCombo.Enabled)
            {
                block.PolicyCombo.SelectedItem = "SpecCPU";
            }

            RecalcAffinityMask(block);

            string labelText = block.Kind switch
            {
                DeviceKind.NET_NDIS or DeviceKind.NET_CX => "NET",
                _ => block.Kind.ToString(),
            };
            if (block.Kind == DeviceKind.NET_NDIS)
            {
                WriteLog($"AUTO.PLAN.ASSIGN: role={slot.Role} {labelText} {block.Device.InstanceId} -> LPs=[{string.Join(',', lps)}] policy={block.PolicyCombo.SelectedItem} queues={(ndisAutoQueues ?? 1)} reason={slot.Reason}");
            }
            else
            {
                WriteLog($"AUTO.PLAN.ASSIGN: role={slot.Role} {labelText} {block.Device.InstanceId} -> LPs=[{string.Join(',', lps)}] policy={block.PolicyCombo.SelectedItem} reason={slot.Reason}");
            }
        }

        WriteLog($"AUTO.PLAN.FINAL: consumed=[{string.Join(',', consumedCores.OrderBy(x => x))}] inputShare=[{string.Join(',', inputShareCores.Distinct().OrderBy(x => x))}] assigned={planSlots.Count(s => s.Lps.Count > 0)} skipped={planSlots.Count(s => s.Lps.Count == 0)}");
        WriteAutoOptimizationResultSummary(
            optimizeUsbImod,
            usingP,
            performanceP,
            performanceE,
            targetCcdLps,
            usbBlocks,
            gpuBlocks,
            integratedGpuBlocks,
            netCount,
            audioBlocks,
            storCount,
            wifiIds,
            msiOnlyGpuIds,
            skipReasons,
            planSlots,
            consumedCores,
            inputShareCores);
        WriteLog("AUTO: Invoke-AutoOptimization done");
    }

    private void ResetAllTweaks()
    {
        WriteLog("RESET: full reset requested");
        if (_blocks.Count == 0)
        {
            RefreshBlocks();
        }

        foreach (DeviceBlock b in _blocks)
        {
            try
            {
                ResetBlockSettings(b);
                if (!b.Device.IsTestDevice)
                {
                    LoadBlockSettings(b);
                }
            }
            catch (Exception ex)
            {
                WriteLog($"RESET.ERROR: {b.Device.InstanceId} -> {ex.Message}");
            }
        }

        ResetReservedCpuSets();
        ResetImodIntervalsToDefault("reset-all");

        CalculateIrqCounts("reset-all");
        LogGuiSnapshot("reset-all");
        ShowThemedInfo(
            "All Device Tweaker changes have been cleared.\nPlease reboot your PC to fully revert device behavior.");
    }

    private void ResetImodIntervalsToDefault(string reason = "reset-imod")
    {
        string defaultText = FormatImodValue(ImodDefaultInterval);
        bool hasUsb = false;

        foreach (DeviceBlock block in _blocks)
        {
            if (IsUsbImodTarget(block.Device))
            {
                string before = block.ImodBox.Text?.Trim() ?? string.Empty;
                block.ImodBox.Text = defaultText;
                block.ImodAutoCheck.Checked = false;
                WriteLog($"RESET.IMOD.USB: {block.Device.InstanceId} {before} -> {defaultText}");
                hasUsb = true;
            }
        }

        if (hasUsb)
        {
            _ = ApplyImodSettings(out string? note);
            if (!string.IsNullOrWhiteSpace(note))
            {
                WriteLog($"RESET.IMOD.USB: {note}");
            }
        }

        RemoveImodPersistenceFiles();
        if (hasUsb)
        {
            RefreshImodCurrentValues(reason: reason);
        }
    }
}
