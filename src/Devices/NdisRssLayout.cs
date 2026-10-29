namespace DeviceTweakerCS;

// NDIS RSS spreads queues over physical cores and skips SMT siblings, so the
// processor range written to the class key must span `queues` cores, not LPs.
internal static class NdisRssLayout
{
    internal sealed record Layout(int BaseProcessor, int MaxProcessor, IReadOnlyList<int> Processors);

    // First logical processor of every group-0 core below maxLogical, ascending.
    internal static List<int> GetCorePrimaryProcessors(IReadOnlyCollection<CpuLpInfo>? lps, int maxLogical)
    {
        if (lps is null || lps.Count == 0)
        {
            return Enumerable.Range(0, Math.Max(0, maxLogical)).ToList();
        }

        return lps
            .Where(lp => lp.Group == 0 && lp.LP >= 0 && lp.LP < maxLogical)
            .GroupBy(lp => lp.Core)
            .Select(core => core.Min(lp => lp.LP))
            .Distinct()
            .OrderBy(lp => lp)
            .ToList();
    }

    internal static int GetMaxQueues(IReadOnlyCollection<CpuLpInfo>? lps, int maxLogical)
        => Math.Max(1, GetCorePrimaryProcessors(lps, maxLogical).Count);

    // Queue counts the adapter can actually use: the driver's *NumRssQueues enum
    // (when present), at most one queue per physical core, and a single queue
    // without MSI-X because every queue would share one interrupt vector.
    internal static List<int> GetAllowedQueueCounts(
        IEnumerable<int>? driverValues,
        PciInterruptCapabilities? capabilities,
        int maxQueues)
    {
        maxQueues = Math.Max(1, maxQueues);
        if (capabilities is not null && (capabilities.Support & PciInterruptSupport.MsiX) == 0)
        {
            return [1];
        }

        if (capabilities?.MessageMaximum is uint messages && messages > 0)
        {
            maxQueues = Math.Min(maxQueues, (int)Math.Min(messages, int.MaxValue));
        }

        List<int> driver = driverValues?.Where(value => value > 0).Distinct().ToList() ?? [];
        IEnumerable<int> candidates = driver.Count > 0 ? driver : Enumerable.Range(1, maxQueues);
        List<int> allowed = candidates.Where(value => value <= maxQueues).Append(1).Distinct().Order().ToList();
        return allowed;
    }

    // Moves a spinner value to an allowed count in the direction the user stepped.
    internal static int Snap(IReadOnlyList<int> allowed, int previous, int requested)
    {
        if (allowed.Count == 0 || allowed.Contains(requested))
        {
            return allowed.Count == 0 ? Math.Max(1, requested) : requested;
        }

        if (requested > previous)
        {
            int up = allowed.FirstOrDefault(value => value > previous);
            return up > 0 ? up : allowed[^1];
        }

        return Floor(allowed, requested);
    }

    // Largest allowed count not above the value; used for saved or AUTO values.
    internal static int Floor(IReadOnlyList<int> allowed, int value)
    {
        if (allowed.Count == 0)
        {
            return Math.Max(1, value);
        }

        int floor = allowed.LastOrDefault(candidate => candidate <= value);
        return floor > 0 ? floor : allowed[0];
    }

    internal static Layout Select(IReadOnlyCollection<CpuLpInfo>? lps, int maxLogical, int requestedBase, int queues)
    {
        List<int> primaries = GetCorePrimaryProcessors(lps, maxLogical);
        if (primaries.Count == 0)
        {
            return new Layout(0, 0, [0]);
        }

        queues = Math.Clamp(queues, 1, primaries.Count);
        int baseCorePrimary = GetCorePrimary(lps, requestedBase) ?? requestedBase;

        // Start at the core that holds the requested base; shift left when the
        // remaining cores cannot hold every queue.
        int startIndex = primaries.FindIndex(lp => lp >= baseCorePrimary);
        if (startIndex < 0)
        {
            startIndex = primaries.Count - 1;
        }

        startIndex = Math.Min(startIndex, primaries.Count - queues);
        List<int> selected = primaries.GetRange(startIndex, queues);
        return new Layout(selected[0], selected[^1], selected);
    }

    private static int? GetCorePrimary(IReadOnlyCollection<CpuLpInfo>? lps, int lp)
    {
        CpuLpInfo? owner = lps?.FirstOrDefault(x => x.Group == 0 && x.LP == lp);
        if (owner is null)
        {
            return null;
        }

        return lps!.Where(x => x.Group == 0 && x.Core == owner.Core).Min(x => x.LP);
    }
}
