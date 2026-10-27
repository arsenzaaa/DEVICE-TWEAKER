using System.Text.RegularExpressions;

namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private static bool IsAutoDisplayAudioMsiOnly(DeviceBlock block)
    {
        return block.Kind == DeviceKind.AUDIO
            && (IsDisplayHdmiaudio(block.Device.InstanceId, block.Device.Name)
                || IsDisplayAudioEndpointsText(block.Device.AudioEndpoints));
    }

    private static bool IsAutoMsiOnlyDevice(DeviceBlock block)
    {
        // Integrated GPU is handled by AUTO (affinity cleared to Windows default + MSI).
        // Only display/HDMI audio still uses the MSI-only apply short-circuit.
        return IsAutoDisplayAudioMsiOnly(block);
    }

    private static string FormatAutoImodRequestState(bool optimizeUsbImod, bool hasUsbImodTarget)
    {
        return !hasUsbImodTarget
            ? "skipped: no eligible XHCI controller"
            : optimizeUsbImod
                ? "requested for eligible XHCI controller(s)"
                : "skipped: declined by user";
    }

    private static string GetAutoImodSkipReason(bool hasUsbImodTarget)
    {
        return hasUsbImodTarget ? "user-declined" : "no-eligible-xhci-controller";
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

        if (TrySelectNonVCacheCcd(ccdIds, _cpuInfo.CcdL3CacheBytes, out int nonVCacheCcd))
        {
            reason = $"non-vcache L3={_cpuInfo.CcdL3CacheBytes[nonVCacheCcd] / (1024d * 1024d):0.#}MB";
            return nonVCacheCcd;
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

    internal static bool TrySelectNonVCacheCcd(
        IReadOnlyList<int> ccdIds,
        IReadOnlyDictionary<int, long> ccdL3CacheBytes,
        out int ccd)
    {
        ccd = 0;
        List<(int Ccd, long Bytes)> known = ccdIds
            .Distinct()
            .Where(id => ccdL3CacheBytes.TryGetValue(id, out long bytes) && bytes > 0)
            .Select(id => (Ccd: id, Bytes: ccdL3CacheBytes[id]))
            .OrderBy(item => item.Bytes)
            .ThenBy(item => item.Ccd)
            .ToList();

        if (known.Count != ccdIds.Distinct().Count() || known.Count < 2)
        {
            return false;
        }

        long min = known.First().Bytes;
        long max = known.Last().Bytes;
        const long MinimumVCacheDelta = 32L * 1024L * 1024L;
        if (max - min < MinimumVCacheDelta || max < min + (min / 2))
        {
            return false;
        }

        List<(int Ccd, long Bytes)> minimum = known.Where(item => item.Bytes == min).ToList();
        if (minimum.Count != 1)
        {
            return false;
        }

        ccd = minimum[0].Ccd;
        return true;
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

    /// <summary>
    /// Mouse / Gamepad / Keyboard on an XHCI controller → external gaming input, keep affinity.
    /// Webcam / Audio-only controllers are treated as built-in and skipped.
    /// </summary>
    private static bool UsbRolesIndicateExternalInput(string roles)
    {
        return HasRoleText(roles, "Mouse")
            || HasRoleText(roles, "Gamepad")
            || HasRoleText(roles, "Controller")
            || HasRoleText(roles, "Keyboard");
    }

    private static bool ShouldSkipBuiltInUsbAffinity(string roles)
        => string.IsNullOrWhiteSpace(roles) || !UsbRolesIndicateExternalInput(roles);

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

    private static string FormatAutoResultDeviceFields(DeviceBlock block)
    {
        string name = SanitizeLogValue(block.Device.Name);
        if (string.IsNullOrWhiteSpace(name))
        {
            name = SanitizeLogValue(block.Device.InstanceId);
        }

        string roles = SanitizeLogValue(block.Device.UsbRoles);
        if (block.Kind == DeviceKind.USB && !string.IsNullOrWhiteSpace(roles))
        {
            return $"device=\"{name}\" roles=\"{roles}\"";
        }

        return $"device=\"{name}\"";
    }

    private void WriteAutoOptimizationResultSummary(
        bool optimizeUsbImod,
        bool hasUsbImodTarget,
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
            + skipReasons.Count;

        string imodState = FormatAutoImodRequestState(optimizeUsbImod, hasUsbImodTarget);

        WriteLog($"AUTO.RESULT.MODE: mode=\"{(_testAutoDryRun ? "dry-run preview, nothing saved" : "apply mode")}\" cpuType={(usingP ? "P-cores" : "E-cores")} primaryP=[{string.Join(',', primaryP)}] primaryE=[{string.Join(',', primaryE)}] targetCCD=[{string.Join(',', targetCcdLps)}]");
        WriteLog($"AUTO.RESULT.DETECTED: usb={usbBlocks.Count} gpu={gpuBlocks.Count} integratedGpu={integratedGpuBlocks.Count} network={netCount} audio={audioBlocks.Count} storage={storCount} wifi={wifiIds.Count} usbImod=\"{SanitizeLogValue(imodState)}\"");

        foreach (AutoAffinityPlanSlot slot in planSlots.Where(slot => slot.Lps.Count > 0))
        {
            DeviceBlock block = _blocks.First(b => string.Equals(b.Device.InstanceId, slot.InstanceId, StringComparison.OrdinalIgnoreCase));
            string lps = string.Join(',', slot.Lps);
            string deviceFields = FormatAutoResultDeviceFields(block);
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
                extra = $" ndisMode=\"{SanitizeLogValue(mode)}\" rssQueues={queues}";
            }
            else if (block.Kind == DeviceKind.USB && IsUsbImodTarget(block.Device))
            {
                string imod = SanitizeLogValue(block.ImodBox.Text);
                extra = $" imod=\"{(string.IsNullOrWhiteSpace(imod) ? "n/a" : imod)}\"";
            }

            WriteLog($"AUTO.RESULT.APPLIED: role=\"{SanitizeLogValue(role)}\" kind={kind} {deviceFields} cpu=[{lps}] mask=0x{block.AffinityMask:X} msi={msi} priority={prio} policy=\"{SanitizeLogValue(policy)}\"{extra} reason=\"{SanitizeLogValue(slot.Reason)}\"");
        }

        foreach (AutoAffinityPlanSlot slot in planSlots.Where(slot => slot.Lps.Count == 0))
        {
            DeviceBlock block = _blocks.First(b => string.Equals(b.Device.InstanceId, slot.InstanceId, StringComparison.OrdinalIgnoreCase));
            WriteLog($"AUTO.RESULT.SKIPPED: role=\"{SanitizeLogValue(FormatAutoResultRole(slot.Role))}\" kind={FormatAutoResultKind(block.Kind)} {FormatAutoResultDeviceFields(block)} reason=\"no safe CPU core\"");
        }

        foreach (DeviceBlock block in _blocks.Where(block => wifiIds.Contains(block.Device.InstanceId)))
        {
            WriteLog($"AUTO.RESULT.PRESERVED: role=\"WiFi\" kind={FormatAutoResultKind(block.Kind)} {FormatAutoResultDeviceFields(block)} reason=\"no changes\"");
        }

        foreach (DeviceBlock block in _blocks.Where(block => msiOnlyGpuIds.Contains(block.Device.InstanceId)))
        {
            WriteLog($"AUTO.RESULT.SKIPPED: role=\"Integrated GPU\" kind={FormatAutoResultKind(block.Kind)} {FormatAutoResultDeviceFields(block)} reason=\"affinity cleared to Windows default, MSI only\"");
        }

        foreach (DeviceBlock block in _blocks.Where(block => skipReasons.ContainsKey(block.Device.InstanceId)))
        {
            WriteLog($"AUTO.RESULT.SKIPPED: kind={FormatAutoResultKind(block.Kind)} {FormatAutoResultDeviceFields(block)} reason=\"{SanitizeLogValue(skipReasons[block.Device.InstanceId])}\"");
        }

        if (storCount > 0)
        {
            foreach (DeviceBlock block in _blocks.Where(block => block.Kind == DeviceKind.STOR))
            {
                WriteLog(
                    $"AUTO.RESULT.CONFIGURED: role=\"Storage\" kind={FormatAutoResultKind(block.Kind)} {FormatAutoResultDeviceFields(block)} " +
                    "msi=Enabled priority=High limit=unlocked affinity=WindowsDefault " +
                    "removed=[MessageNumberLimit,DevicePolicy,AssignmentSetOverride]");
            }
        }

        WriteLog($"AUTO.RESULT.FINAL: assignedCpuAffinity={assigned} skippedOrPreserved={skipped} storageConfigured={storCount} usedLPs=[{string.Join(',', assignedLps)}] reservedSpacingLPs=[{string.Join(',', reservedLps)}] inputShareLPs=[{string.Join(',', inputShareCores.Distinct().OrderBy(lp => lp))}]");
    }

    private bool InvokeAutoOptimization(bool optimizeUsbImod, bool hasUsbImodTarget, OperationReport? report = null)
    {
        if (_blocks.Count == 0)
        {
            report?.AddError("AUTO-OPTIMIZATION", "no devices to optimize");
            return false;
        }

        if (!_cpuTopologyReliable)
        {
            report?.AddError("AUTO-OPTIMIZATION", "Windows CPU topology APIs failed; no affinity changes were applied");
            WriteLog("AUTO: aborted (CPU topology is display-only and not reliable)");
            return false;
        }

        WriteLog("AUTO: Invoke-AutoOptimization start");
        (List<int> P, List<int> E)? cpuSets = GetAutoCpuSets();
        if (cpuSets is null)
        {
            report?.AddError("AUTO-OPTIMIZATION", "CPU map unavailable; affinity plan was not built");
            WriteLog("AUTO: aborted (CPU sets unavailable)");
            return false;
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
            report?.AddError("AUTO-OPTIMIZATION", "no performance P-cores available for affinity plan");
            WriteLog("AUTO: aborted (performance P count is 0)");
            return false;
        }

        // Do not touch registry-backed global settings until the CPU plan has
        // passed its preflight checks. A planning failure must leave Windows unchanged.
        ApplyAutoRawMouseThrottle(report);
        ResetReservedCpuSetsPreview();

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
        HashSet<string> msiOnlyDisplayAudioIds = new(StringComparer.OrdinalIgnoreCase);

        foreach (DeviceBlock usbBlock in usbBlocks)
        {
            string roles = usbBlock.Device.UsbRoles ?? string.Empty;
            if (string.IsNullOrWhiteSpace(roles))
            {
                skipAutoIds.Add(usbBlock.Device.InstanceId);
                skipReasons[usbBlock.Device.InstanceId] = "USB controller has no detected HID roles, left for manual/reset only";
                WriteLog($"AUTO.SKIP.USB: {usbBlock.Device.InstanceId} no HID roles (manual/reset only)");
            }
            else if (ShouldSkipBuiltInUsbAffinity(roles))
            {
                skipAutoIds.Add(usbBlock.Device.InstanceId);
                skipReasons[usbBlock.Device.InstanceId] =
                    "built-in USB (no Mouse/Gamepad/Keyboard): affinity skipped; MSI/power still applied";
                WriteLog(
                    $"AUTO.SKIP.USB: {usbBlock.Device.InstanceId} built-in roles=\"{SanitizeLogValue(roles)}\" " +
                    "reason=no-external-input-affinity-skip");
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
                // Onboard HD Audio / Speakers / Mic: built-in, do not pin IRQ affinity.
                skipAutoIds.Add(pnpId);
                skipReasons[pnpId] = "built-in audio: affinity skipped; MSI/power still applied";
                WriteLog($"AUTO.SKIP.AUDIO: {pnpId} built-in (name=\"{SanitizeLogValue(desc)}\" endpoints=\"{SanitizeLogValue(audioText)}\")");
                continue;
            }

            skipAutoIds.Add(pnpId);
            string reason = isSpdif ? "digital S/PDIF audio" : "display/HDMI audio";
            if (isDisplay)
            {
                msiOnlyDisplayAudioIds.Add(pnpId);
                skipReasons[pnpId] = $"{reason}: MSI enabled; CPU affinity, priority and MSI limit preserved";
            }
            else
            {
                skipReasons[pnpId] = $"{reason}: excluded from CPU assignment; MSI enabled, priority reset, limit unlimited, affinity reset to Windows default";
            }
            WriteLog($"AUTO.SKIP.AUDIO: {pnpId} classified as {reason} (name=\"{desc}\" endpoints=\"{audioText}\")");
        }

        foreach (DeviceBlock block in _blocks)
        {
            bool isSkipAuto = skipAutoIds.Contains(block.Device.InstanceId);
            bool isMsiOnlyGpu = msiOnlyGpuIds.Contains(block.Device.InstanceId);
            bool isMsiOnlyDisplayAudio = msiOnlyDisplayAudioIds.Contains(block.Device.InstanceId);
            bool isWifi = wifiIds.Contains(block.Device.InstanceId);
            ulong beforeMask = block.AffinityMask;
            string beforePolicy = block.PolicyCombo.SelectedItem?.ToString() ?? "(none)";
            string beforeMsi = block.MsiCombo.SelectedItem?.ToString() ?? "(none)";
            string beforeLimit = block.LimitBox.Text?.Trim() ?? string.Empty;
            string beforePriority = block.PrioCombo.SelectedItem?.ToString() ?? "(none)";
            WriteLog(
                $"AUTO.STATE.BEFORE: {block.Device.InstanceId} Kind={block.Kind} MSI={beforeMsi} " +
                $"limit=\"{beforeLimit}\" priority={beforePriority} policy={beforePolicy} mask=0x{beforeMask:X}");

            if (isMsiOnlyGpu)
            {
                string msiBefore = block.MsiCombo.SelectedItem?.ToString() ?? "(none)";
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
                block.MsiCombo.SelectedItem = "Enabled";
                block.LimitBox.Text = "Unlimited";
                block.PrioCombo.SelectedItem = "Undefined";
                if (block.PolicyCombo.Enabled)
                {
                    block.PolicyCombo.SelectedItem = "MachineDefault";
                }

                RecalcAffinityMask(block);
                string msiAfter = block.MsiCombo.SelectedItem?.ToString() ?? "(none)";
                WriteLog(
                    $"AUTO.SKIP.GPU: {block.Device.InstanceId} integrated=1 msiBefore={msiBefore} msiAfter={msiAfter} " +
                    "affinity=WindowsDefault reason=integratedGpuAutoSkip");
                continue;
            }

            if (isMsiOnlyDisplayAudio)
            {
                string msiBefore = block.MsiCombo.SelectedItem?.ToString() ?? "(none)";
                string limitBefore = block.LimitBox.Text?.Trim() ?? string.Empty;
                string priorityBefore = block.PrioCombo.SelectedItem?.ToString() ?? "(none)";
                string policyBefore = block.PolicyCombo.SelectedItem?.ToString() ?? "(none)";
                ulong maskBefore = block.AffinityMask;
                block.MsiCombo.SelectedItem = "Enabled";
                WriteLog(
                    $"AUTO.MSI-ONLY.DISPLAY-AUDIO: {block.Device.InstanceId} " +
                    $"MSI={msiBefore}->Enabled limit=\"{limitBefore}\" priority={priorityBefore} " +
                    $"policy={policyBefore} mask=0x{maskBefore:X} preserved=limit,priority,policy,affinity");
                continue;
            }

            if (isWifi)
            {
                WriteLog(
                    $"AUTO.WIFI.PRESERVE: {block.Device.InstanceId} -> no changes " +
                    $"MSI={beforeMsi} limit=\"{beforeLimit}\" priority={beforePriority} policy={beforePolicy} mask=0x{beforeMask:X}");
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
                block.LimitBox.Text = "Unlimited";
                block.PrioCombo.SelectedItem = "Undefined";
                if (block.Kind != DeviceKind.NET_NDIS && block.PolicyCombo.Enabled)
                {
                    block.PolicyCombo.SelectedItem = "MachineDefault";
                }
            }
            else
            {
                block.MsiCombo.SelectedItem = "Enabled";
                block.LimitBox.Text = "Unlimited";
                block.PrioCombo.SelectedItem = "High";
                if (block.Kind != DeviceKind.NET_NDIS && block.PolicyCombo.Enabled)
                {
                    block.PolicyCombo.SelectedItem = "MachineDefault";
                }
            }

            if (block.PowerSavingCheck is not null && !isWifi)
            {
                block.PowerSavingCheck.Checked = false;
                if (block.Kind == DeviceKind.USB)
                {
                    block.Device.UsbSelectiveSuspend = "off";
                }
                else
                {
                    block.Device.NicPowerSaving = "off";
                }

                UpdateBlockInfoText(block);
                WriteLog($"AUTO.POWER: {block.Device.InstanceId} -> Power Saving=Disabled");
            }

            RecalcAffinityMask(block);

            ulong afterMask = block.AffinityMask;
            string afterPolicy = block.PolicyCombo.SelectedItem?.ToString() ?? "(none)";
            string afterMsi = block.MsiCombo.SelectedItem?.ToString() ?? "(none)";
            string afterLimit = block.LimitBox.Text?.Trim() ?? string.Empty;
            string afterPriority = block.PrioCombo.SelectedItem?.ToString() ?? "(none)";
            WriteLog(
                $"AUTO.STATE.AFTER: {block.Device.InstanceId} Kind={block.Kind} MSI={afterMsi} " +
                $"limit=\"{afterLimit}\" priority={afterPriority} policy={afterPolicy} mask=0x{afterMask:X} " +
                $"cpuAssignmentSkipped={isSkipAuto || block.Kind == DeviceKind.STOR}");
            if (isSkipAuto)
            {
                string reason = skipReasons.TryGetValue(block.Device.InstanceId, out string? skipReason)
                    ? skipReason
                    : "skipped by AUTO policy";
                WriteLog($"AUTO.RESET.SKIP: {block.Device.InstanceId} Kind={block.Kind} reason=\"{SanitizeLogValue(reason)}\"");
            }
        }

        List<DeviceBlock> imodTargets = usbBlocks.Where(b => IsUsbImodTarget(b.Device)).ToList();
        if (!hasUsbImodTarget)
        {
            WriteLog("AUTO.IMOD: skipped reason=no-eligible-xhci-controller");
        }
        else if (!optimizeUsbImod)
        {
            WriteLog("AUTO.IMOD: skipped reason=user-declined");
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

        List<DeviceBlock> eligibleBlocks = _blocks
            .Where(b => b.Kind != DeviceKind.STOR)
            .Where(b => !wifiIds.Contains(b.Device.InstanceId))
            .Where(b => !skipAutoIds.Contains(b.Device.InstanceId))
            .Where(b => !msiOnlyGpuIds.Contains(b.Device.InstanceId))
            .ToList();
        List<AutoAffinityPlanSlot> planSlots = eligibleBlocks
            .Select((block, index) => new { block, index })
            .Select(x => new AutoAffinityPlanSlot
            {
                Kind = x.block.Kind,
                InstanceId = x.block.Device.InstanceId,
                Role = GetAutoAffinityRole(x.block),
                Need = x.block.Kind == DeviceKind.GPU ? 2 : 1,
                OriginalIndex = x.index,
            })
            .OrderBy(x => GetAutoAffinityRolePriority(x.Role))
            .ThenBy(x => x.OriginalIndex)
            .ToList();

        WriteLog($"AUTO.PLAN.CPU: using={(usingP ? "P" : "E")} preferredUnits=[{FormatAutoCpuUnits(preferredUnits)}] backgroundUnits=[{FormatAutoCpuUnits(backgroundUnits)}] nonTargetBackground=[{FormatAutoCpuUnits(nonTargetBackgroundUnits)}] eBackground=[{FormatAutoCpuUnits(eBackgroundUnits)}]");
        WriteLog($"AUTO.PLAN.SLOTS: [{string.Join(" | ", planSlots.Select(s => $"{s.Role}:{s.Kind}:{s.InstanceId}"))}]");

        AutoAllocationResult allocation = AutoAffinityPlanner.Allocate(planSlots, preferredUnits,
            backgroundUnits, allCpuUnits, targetCcdLps, _maxLogical, HasVisibleCcxSplit(),
            FormatAutoCpuUnits, WriteLog);
        HashSet<int> consumedCores = allocation.ConsumedCores;
        List<int> inputShareCores = allocation.InputShareCores;

        foreach (AutoAffinityPlanSlot slot in planSlots)
        {
            DeviceBlock block = eligibleBlocks[slot.OriginalIndex];
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
            hasUsbImodTarget,
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
        return true;
    }

    private void SafeResetTweaks(OperationReport? report = null)
    {
        report ??= new OperationReport();
        WriteLog("SAFE_RESET: requested; MSI mode/limit and Wi-Fi are preserved");
        bool ownsBusy = _devicesBusyDepth > 0;
        int blockCount = Math.Max(1, _blocks.Count);
        if (ownsBusy)
        {
            // prepare + blocks + reserved + imod + suspend/finalize
            SetDevicesBusyWork(1 + blockCount + 3, 0);
            TickDevicesBusy("Preparing RESET WINDOWS DEFAULT...", 1);
        }
        else
        {
            SetDevicesBusyStage("Preparing RESET WINDOWS DEFAULT...");
        }

        if (_blocks.Count == 0)
        {
            RefreshBlocks();
            blockCount = Math.Max(1, _blocks.Count);
            if (ownsBusy)
            {
                SetDevicesBusyWork(1 + blockCount + 3, Math.Min(1, _devicesBusyDone));
            }
        }

        if (_testAutoDryRun)
        {
            WriteLog("SAFE_RESET: dry-run -> ReservedCpuSets/IMOD persistence/registry left untouched");
            if (ownsBusy)
            {
                SetDevicesBusyWork(1 + blockCount + 1, Math.Min(_devicesBusyDone, 1 + blockCount + 1));
            }

            int dryIndex = 0;
            foreach (DeviceBlock b in _blocks)
            {
                dryIndex++;
                if (ownsBusy)
                {
                    TickDevicesBusy($"Dry-run reset ({dryIndex}/{blockCount})...", 1);
                }
                else
                {
                    SetDevicesBusyStage($"Dry-run reset ({dryIndex}/{blockCount})...");
                }

                try
                {
                    if (b.Device.Wifi)
                    {
                        WriteLog($"SAFE_RESET.WIFI.PRESERVE: {b.Device.InstanceId} -> no changes");
                        continue;
                    }

                    if (b.Device.IsTestDevice)
                    {
                        ResetBlockSettings(b, report);
                    }
                    else
                    {
                        // UI-only clear without LoadBlockSettings (would re-read live registry into UI)
                        b.SuppressCpuEvents++;
                        try
                        {
                            foreach (CheckBox cb in b.CpuBoxes)
                            {
                                cb.Checked = false;
                            }
                        }
                        finally
                        {
                            b.SuppressCpuEvents--;
                        }

                        b.AffinityMask = 0;
                        b.AffinityLabel.Text = b.Kind == DeviceKind.STOR
                            ? "Affinity Mask: Windows Default"
                            : (b.Kind == DeviceKind.AUDIO && (IsDisplayHdmiaudio(b.Device.InstanceId, b.Device.Name) || IsDisplayAudioEndpointsText(b.Device.AudioEndpoints)))
                                ? "Affinity Mask: 0x0 (Windows Default)"
                                : "Affinity Mask: 0x0";
                        b.PrioCombo.SelectedItem = "Undefined";
                        if (b.Kind == DeviceKind.NET_NDIS)
                        {
                            b.RssBaseCore = null;
                            if (b.RssQueueBox is not null)
                            {
                                b.SuppressCpuEvents++;
                                try
                                {
                                    b.RssQueueBox.Value = 1;
                                }
                                finally
                                {
                                    b.SuppressCpuEvents--;
                                }
                            }
                        }
                        else if (b.PolicyCombo.Enabled)
                        {
                            b.PolicyCombo.SelectedItem = "MachineDefault";
                        }

                        if (b.PowerSavingCheck is not null)
                        {
                            b.PowerSavingCheck.Checked = true;
                            if (b.Kind == DeviceKind.USB)
                            {
                                b.Device.UsbSelectiveSuspend = "on";
                            }
                            else if (!b.Device.Wifi)
                            {
                                b.Device.NicPowerSaving = "on";
                            }

                            UpdateBlockInfoText(b);
                        }
                    }

                    if (IsUsbImodTarget(b.Device))
                    {
                        b.ImodBox.Text = FormatImodValue(ImodDefaultInterval);
                        b.ImodAutoCheck.Checked = false;
                        if (b.Device.IsTestDevice)
                        {
                            RefreshTestImodPreview(b, "safe-reset-dry-run");
                        }
                    }
                }
                catch (Exception ex)
                {
                    WriteLog($"SAFE_RESET.ERROR: {b.Device.InstanceId} -> {ex.Message}");
                    report.AddError($"{b.Device.Name}: reset", ex.Message);
                }
            }

            if (ownsBusy)
            {
                TickDevicesBusy("Updating IRQ counts...", 1);
            }
            else
            {
                SetDevicesBusyStage("Updating IRQ counts...");
            }

            ResetReservedCpuSetsPreview();
            CalculateIrqCounts("safe-reset");
            LogGuiSnapshot("safe-reset-dry-run");
            if (ownsBusy)
            {
                _devicesBusyDone = _devicesBusyTotal;
                UpdateDevicesBusy("Ready", 100);
            }

            WriteLog($"SAFE_RESET.DONE: mode=dry-run blocks={_blocks.Count} errors={report.Errors.Count}");
            CloseDevicesBusyOverlay();
            ShowOperationResult(
                report,
                successMessage:
                    "Dry-run reset complete (UI only).\n" +
                    "No registry / IMOD startup files were changed.\n" +
                    "This developer build cannot reset Windows settings.",
                partialMessage: "RESET WINDOWS DEFAULT preview finished with errors.",
                operationName: "RESET WINDOWS DEFAULT PREVIEW");
            return;
        }

        int resetIndex = 0;
        foreach (DeviceBlock b in _blocks)
        {
            resetIndex++;
            if (ownsBusy)
            {
                TickDevicesBusy($"Resetting device settings ({resetIndex}/{blockCount})...", 1);
            }
            else
            {
                SetDevicesBusyStage($"Resetting device settings ({resetIndex}/{blockCount})...");
            }

            try
            {
                if (b.Device.Wifi)
                {
                    WriteLog($"SAFE_RESET.WIFI.PRESERVE: {b.Device.InstanceId} -> no changes");
                    continue;
                }

                ResetBlockSettings(b, report);
                if (!b.Device.IsTestDevice)
                {
                    LoadBlockSettings(b);
                }
            }
            catch (Exception ex)
            {
                WriteLog($"SAFE_RESET.ERROR: {b.Device.InstanceId} -> {ex.Message}");
                report.AddError($"{b.Device.Name}: reset", ex.Message);
            }
        }

        try
        {
            if (ownsBusy)
            {
                TickDevicesBusy("Resetting reserved CPU sets...", 1);
            }
            else
            {
                SetDevicesBusyStage("Resetting reserved CPU sets...");
            }

            ResetReservedCpuSets(report);
        }
        catch (Exception ex)
        {
            WriteLog($"SAFE_RESET.ERROR: ReservedCpuSets -> {ex.Message}");
            report.AddError("Reserved CPU sets", ex.Message);
        }

        try
        {
            if (SupportsRawMouseThrottleOs())
            {
                SetRawMouseThrottle(enabled: false, DefaultRawMouseThrottleDuration);
                RefreshRawMouseThrottleUi();
                WriteLog("SAFE_RESET.THROTTLE: custom RawMouseThrottleDuration removed");
            }
            else
            {
                WriteLog("SAFE_RESET.THROTTLE: skipped (unsupported OS)");
            }
        }
        catch (Exception ex)
        {
            WriteLog($"SAFE_RESET.ERROR: RawMouseThrottleDuration -> {ex.Message}");
            report.AddError("Raw mouse throttle", ex.Message);
        }

        try
        {
            if (ownsBusy)
            {
                TickDevicesBusy("Removing IMOD persistence...", 1);
            }
            else
            {
                SetDevicesBusyStage("Removing IMOD persistence...");
            }

            ClearImodPersistenceForSafeReset(report);
        }
        catch (Exception ex)
        {
            WriteLog($"SAFE_RESET.ERROR: IMOD persistence -> {ex.Message}");
            report.AddError("IMOD persistence", ex.Message);
        }

        try
        {
            bool anyRealUsb = _blocks.Any(b =>
                b.Kind == DeviceKind.USB
                && !b.Device.IsTestDevice
                && b.PowerSavingCheck is not null);
            if (anyRealUsb)
            {
                UsbSelectiveSuspendPolicy.SetPowerPlanEnabled(true);
                WriteLog("SAFE_RESET.SUSPEND.PLAN: USB selective suspend power plan restored to Enabled");
            }
            else
            {
                WriteLog("SAFE_RESET.SUSPEND.PLAN: skipped (no real USB blocks)");
            }
        }
        catch (Exception ex)
        {
            WriteLog($"SAFE_RESET.ERROR: USB selective suspend power plan -> {ex.Message}");
            report.AddError("USB Selective Suspend power plan", ex.Message);
        }

        if (ownsBusy)
        {
            TickDevicesBusy("Updating IRQ counts...", 1);
        }
        else
        {
            SetDevicesBusyStage("Updating IRQ counts...");
        }

        CalculateIrqCounts("safe-reset");
        LogGuiSnapshot("safe-reset");
        if (ownsBusy)
        {
            _devicesBusyDone = _devicesBusyTotal;
            UpdateDevicesBusy("Ready", 100);
        }

        WriteLog($"SAFE_RESET.DONE: mode=apply blocks={_blocks.Count} errors={report.Errors.Count}");
        CloseDevicesBusyOverlay();
        ShowOperationResult(
            report,
            successMessage:
                "Supported tweaks were reset.\n" +
                "Please reboot your PC to restore runtime hardware defaults.",
            partialMessage: "RESET WINDOWS DEFAULT finished with errors. Some settings may still be active.",
            operationName: "RESET WINDOWS DEFAULT");
    }

    private void ClearImodPersistenceForSafeReset(OperationReport? report)
    {
        foreach (DeviceBlock block in _blocks.Where(block => IsUsbImodTarget(block.Device)))
        {
            block.ImodAutoCheck.Checked = false;
        }

        if (_testAutoDryRun)
        {
            WriteLog("SAFE_RESET.IMOD: dry-run -> persistence preserved; runtime MMIO untouched");
            return;
        }

        RemoveImodPersistenceFiles(report);
        WriteLog("SAFE_RESET.IMOD: persistence removed; runtime MMIO intentionally untouched until reboot");
        RefreshImodCurrentValues(reason: "safe-reset");
    }

    private void ResetImodIntervalsToDefault(string reason = "reset-imod", OperationReport? report = null)
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
            if (_testAutoDryRun)
            {
                WriteLog($"RESET.IMOD.USB: dry-run -> skipped ApplyImodSettings/RemoveImodPersistenceFiles reason={reason}");
            }
            else
            {
                // RESET/DELETE must not pull DTIMOD via KDU. Hardware write only if
                // the driver was already loaded explicitly (CHECK / prior IMOD apply).
                if (IsImodDriverAlreadyAvailable())
                {
                    ImodApplyOutcome outcome = ApplyImodSettings(out string? note, out string? technicalDetails);
                    if (!string.IsNullOrWhiteSpace(note))
                    {
                        WriteLog($"RESET.IMOD.USB: {note}");
                    }

                    if (report is not null)
                    {
                        AddImodResultToReport(report, outcome, note, technicalDetails, "IMOD reset");
                    }
                }
                else
                {
                    WriteLog(
                        $"RESET.IMOD.USB: skipped hardware apply reason=driver-not-loaded " +
                        $"(press CHECK first); cleared UI defaults and persistence only reason={reason}");
                }

                RemoveImodPersistenceFiles(report);
                RefreshImodCurrentValues(reason: reason);
            }
        }
        else if (!_testAutoDryRun)
        {
            RemoveImodPersistenceFiles(report);
        }
        else
        {
            WriteLog($"RESET.IMOD: dry-run -> skipped RemoveImodPersistenceFiles reason={reason}");
        }
    }
}
