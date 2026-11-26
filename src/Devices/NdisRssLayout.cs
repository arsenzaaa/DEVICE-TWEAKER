namespace DeviceTweakerCS;

// RSS queue count and processor count are separate settings. Selection uses
// physical cores, skipping SMT siblings, within a documented device CPU limit.
internal static class NdisRssLayout
{
    internal sealed record Layout(int BaseProcessor, int MaxProcessor, IReadOnlyList<int> Processors)
    {
        internal int QueueCount { get; init; } = Processors.Count;
    }

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

    // Honor advertised queues even when they share one interrupt/CPU (Intel I219).
    // Without an advertised count, retain a conservative single-queue fallback
    // for known non-MSI-X devices.
    internal static List<int> GetAllowedQueueCounts(
        IEnumerable<int>? driverValues,
        PciInterruptCapabilities? capabilities,
        int maxQueues)
    {
        maxQueues = Math.Max(1, maxQueues);
        List<int> driver = driverValues?.Where(value => value > 0).Distinct().ToList() ?? [];
        if (driver.Count > 0) return driver.Order().ToList();
        if (capabilities is not null && (capabilities.Support & PciInterruptSupport.MsiX) == 0) return [1];

        if (capabilities?.MessageMaximum is uint messages && messages > 0)
        {
            maxQueues = Math.Min(maxQueues, (int)Math.Min(messages, int.MaxValue));
        }

        return Enumerable.Range(1, maxQueues).ToList();
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

    internal static Layout Select(IReadOnlyCollection<CpuLpInfo>? lps, int maxLogical, int requestedBase, int queues,
        int? processorLimit = null)
    {
        List<int> primaries = GetCorePrimaryProcessors(lps, maxLogical);
        if (primaries.Count == 0)
        {
            return new Layout(0, 0, [0]);
        }

        int queueCount = Math.Max(1, queues);
        int processorCount = Math.Max(1, Math.Min(primaries.Count, processorLimit ?? int.MaxValue));
        queues = Math.Clamp(queueCount, 1, processorCount);
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
        return new Layout(selected[0], selected[^1], selected)
        {
            QueueCount = queueCount,
        };
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
