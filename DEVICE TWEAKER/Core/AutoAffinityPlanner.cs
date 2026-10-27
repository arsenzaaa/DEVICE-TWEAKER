namespace DeviceTweakerCS;

internal enum AutoAffinityRole
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

internal sealed class AutoAffinityPlanSlot
{
    public required DeviceKind Kind { get; init; }
    public required string InstanceId { get; init; }
    public required AutoAffinityRole Role { get; init; }
    public required int Need { get; init; }
    public required int OriginalIndex { get; init; }
    public List<int> Lps { get; } = [];
    public string Reason { get; set; } = string.Empty;
}

internal sealed class AutoCpuUnit
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

internal sealed record AutoAllocationResult(HashSet<int> ConsumedCores, List<int> InputShareCores);

internal static class AutoAffinityPlanner
{
    internal static AutoAllocationResult Allocate(
        List<AutoAffinityPlanSlot> planSlots, List<AutoCpuUnit> preferredUnits,
        List<AutoCpuUnit> backgroundUnits, List<AutoCpuUnit> allCpuUnits,
        List<int> targetCcdLps, int maxLogical, bool visibleCcxSplit,
        Func<IEnumerable<AutoCpuUnit>, string> formatUnits, Action<string> log)
    {
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
        log($"AUTO.PLAN.UNITS: required={requiredPreferredUnitCount} available={preferredUnits.Count} nonCore0={nonCore0PreferredUnits} avoidCore0={avoidCore0WhenPossible} spacingEligible={spacingEligibleUnits} spacing={spacingSkips}");

        List<AutoCpuUnit> ApplyCcxPreference(List<AutoCpuUnit> ordered, int? preferCcx, int? avoidCcx)
        {
            if (!visibleCcxSplit || ordered.Count <= 1)
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
            if (!visibleCcxSplit)
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
            log($"AUTO.PLAN.UNIT.RESERVE: LP={unit.PrimaryLp} unit={unit.Id} ccd={unit.Ccd} ccx={unit.Ccx} type={(unit.IsECore ? "E" : "P")} reason={reason}");
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

            List<AutoCpuUnit> topologyOrder = preferredUnits
                .Where(unit => !unit.IsECore)
                .OrderBy(unit => unit.Id)
                .ThenBy(unit => unit.PrimaryLp)
                .ToList();
            Dictionary<int, int> topologyIndexByUnit = topologyOrder
                .Select((unit, index) => (unit.Id, index))
                .ToDictionary(item => item.Id, item => item.index);

            HashSet<int> candidateUnitIds = GetCandidateUnits(preferredUnits, allowCore0)
                .Where(unit => !unit.IsECore)
                .Select(unit => unit.Id)
                .ToHashSet();
            if (candidateUnitIds.Count < count)
            {
                return [];
            }

            List<(AutoCpuUnit First, AutoCpuUnit Second)> pairs = [];
            for (int i = 0; i < topologyOrder.Count - 1; i++)
            {
                AutoCpuUnit first = topologyOrder[i];
                AutoCpuUnit second = topologyOrder[i + 1];
                if (!candidateUnitIds.Contains(first.Id) || !candidateUnitIds.Contains(second.Id))
                {
                    continue;
                }

                pairs.Add((first, second));
            }

            if (pairs.Count == 0)
            {
                List<AutoCpuUnit> candidates = GetCandidateUnits(preferredUnits, allowCore0)
                    .Where(unit => !unit.IsECore)
                    .OrderBy(unit => topologyIndexByUnit.TryGetValue(unit.Id, out int index) ? index : int.MaxValue)
                    .ThenBy(unit => unit.PrimaryLp)
                    .ToList();
                log($"AUTO.PLAN.GPU.PAIRS: none-adjacent fallbackCandidates=[{formatUnits(candidates)}]");
                return [];
            }

            List<(AutoCpuUnit First, AutoCpuUnit Second)> nonCore0Pairs = pairs
                .Where(pair => !pair.First.HasCore0 && !pair.Second.HasCore0)
                .ToList();
            if (nonCore0Pairs.Count > 0)
            {
                pairs = nonCore0Pairs;
            }

            string FormatGpuPair((AutoCpuUnit First, AutoCpuUnit Second) pair)
            {
                ulong mask = 0;
                foreach (int lp in new[] { pair.First.PrimaryLp, pair.Second.PrimaryLp })
                {
                    if (lp >= 0 && lp < 64)
                    {
                        mask |= 1UL << lp;
                    }
                }

                string ccx = visibleCcxSplit ? $"/ccx={pair.First.Ccx},{pair.Second.Ccx}" : string.Empty;
                return $"[{pair.First.PrimaryLp},{pair.Second.PrimaryLp}]/mask=0x{mask:X}/ccd={pair.First.Ccd},{pair.Second.Ccd}{ccx}/rank={pair.First.Rank},{pair.Second.Rank}/rating={pair.First.Rating},{pair.Second.Rating}";
            }

            log($"AUTO.PLAN.GPU.PAIRS: candidates={pairs.Count} [{string.Join("; ", pairs.Select(FormatGpuPair))}]");

            // Always prefer the highest physical LP adjacent pair on the target CCD.
            // CPPC rank must NOT steer GPU here: mid-rank pairs like [6,8] caused field
            // FPS drops on Ryzen 8940HX + RTX 5060 Laptop; the working pair was [12,14].
            (AutoCpuUnit First, AutoCpuUnit Second) selectedPair = pairs
                .OrderBy(pair => pair.First.Ccd == pair.Second.Ccd ? 0 : 1)
                .ThenBy(pair => !visibleCcxSplit || pair.First.Ccx == pair.Second.Ccx ? 0 : 1)
                .ThenByDescending(pair => Math.Max(pair.First.PrimaryLp, pair.Second.PrimaryLp))
                .ThenByDescending(pair => Math.Min(pair.First.PrimaryLp, pair.Second.PrimaryLp))
                .First();

            log($"AUTO.PLAN.GPU.SELECT: {FormatGpuPair(selectedPair)} reason=irq-safe-tail-physical-pair");
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
                .Where(unit => unit.PrimaryLp >= 0 && unit.PrimaryLp < maxLogical)
                .DistinctBy(unit => unit.Id)
                .ToList();

            slot.Lps.AddRange(selectedUnits
                .Select(unit => SelectUnitLp(unit, preferSiblingForShared))
                .Where(lp => lp >= 0 && lp < maxLogical)
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

            if (slot.Lps.Count > 0 && (slot.Role is AutoAffinityRole.InputMouse or AutoAffinityRole.InputController))
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
            log($"AUTO.PLAN.SPACING: after={afterReason} reservedUnit={unit.Id} LP={unit.PrimaryLp} remaining={spacingSkips}");
        }

        foreach (AutoAffinityPlanSlot slot in planSlots.Where(s => s.Role is AutoAffinityRole.InputMouse or AutoAffinityRole.InputController))
        {
            int? inputCcx = GetInputAnchorCcx();
            AutoCpuUnit? unit = TakeDedicatedUnit(allowCore0: true, preferCcx: inputCcx);
            if (unit is null)
            {
                log($"AUTO.PLAN.SKIP: role={slot.Role} {slot.Kind} {slot.InstanceId} reason=no-input-unit");
                continue;
            }

            AssignSlot(slot, [unit], inputCcx.HasValue && unit.Ccx == inputCcx.Value ? "input-dedicated-unit-ccx-near" : "input-dedicated-unit");
            // Do not burn spacing before GPU: spacing-after-input was reserving the
            // high-LP tail (e.g. LP14), which blocked the irq-safe GPU pair [12,14].
            if (slot.Lps.Count > 0 && gpuPreferredNeed <= 0)
            {
                SkipSpacingUnitIfAvailable("input");
            }
            else if (slot.Lps.Count > 0 && gpuPreferredNeed > 0)
            {
                log("AUTO.PLAN.SPACING: after=input deferred reason=preserve-gpu-tail-pair");
            }
        }

        foreach (AutoAffinityPlanSlot slot in planSlots.Where(s => s.Role == AutoAffinityRole.Gpu))
        {
            List<AutoCpuUnit> units = TakeGpuDedicatedUnits(slot.Need, allowCore0: true);
            if (units.Count < slot.Need)
            {
                log($"AUTO.PLAN.SKIP: role={slot.Role} {slot.Kind} {slot.InstanceId} reason=no-gpu-physical-pair");
                continue;
            }

            AssignSlot(slot, units, "gpu-dedicated-irq-safe-physical-pair");
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
                log($"AUTO.PLAN.SKIP: role={slot.Role} NET {slot.InstanceId} reason=no-safe-nic-unit");
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
                    log($"AUTO.PLAN.SKIP: role={slot.Role} USB {slot.InstanceId} reason=no-keyboard-unit");
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
                log($"AUTO.PLAN.SKIP: role={slot.Role} AUDIO {slot.InstanceId} reason=no-audio-unit");
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
                    log($"AUTO.PLAN.SKIP: role={slot.Role} {slot.Kind} {slot.InstanceId} reason=no-other-unit");
                }
            }
        }

        void BlockSlot(AutoAffinityPlanSlot slot, string reason)
        {
            if (slot.Lps.Count == 0)
            {
                return;
            }

            log($"AUTO.VALIDATION.BLOCK: role={slot.Role} {slot.Kind} {slot.InstanceId} lps=[{string.Join(',', slot.Lps)}] reason={reason}");
            slot.Lps.Clear();
            slot.Reason = $"validation-blocked: {reason}";
        }

        foreach (AutoAffinityPlanSlot slot in planSlots.Where(s => s.Lps.Count > 0))
        {
            if (slot.Lps.Any(lp => lp < 0 || lp >= maxLogical || !unitByLp.ContainsKey(lp)))
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

        log($"AUTO.VALIDATION: assigned={planSlots.Count(slot => slot.Lps.Count > 0)} blocked={planSlots.Count(slot => slot.Reason.StartsWith("validation-blocked", StringComparison.OrdinalIgnoreCase))}");

        return new(consumedCores, inputShareCores);
    }
}
