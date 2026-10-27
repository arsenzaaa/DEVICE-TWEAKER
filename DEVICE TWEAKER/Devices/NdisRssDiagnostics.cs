namespace DeviceTweakerCS;

internal static class NdisRssDiagnostics
{
    internal sealed record Comparison(string Status, string Detail);

    internal static string ReadStatus(NdisRssRuntimeState? state)
    {
        if (state?.Error == "wifi-preserved") return "skipped";
        if (state?.Error == "rss-not-supported") return "unsupported";
        if (state is not null && !string.IsNullOrWhiteSpace(state.Error)
            && state.Error is not ("rss-not-found" or "adapter-not-found")) return "failed";
        if (state is null || !state.AdapterFound || !state.RssFound || !state.Enabled.HasValue) return "unavailable";
        return state.Enabled.Value ? "available" : "disabled";
    }

    internal static Comparison Compare(NdisRssRuntimeState? state, int? registryBase, int? registryQueues,
        int? registryMaxProcessors, bool wifiPreserved = false)
    {
        if (wifiPreserved) return new("skipped", "Wi-Fi settings preserved");
        string status = ReadStatus(state);
        if (status != "available") return new(status, state?.Error ?? "no readback");
        List<string> differences = [];
        bool incomplete = false;
        void Check(string name, int? requested, int? actual)
        {
            if (!requested.HasValue) return;
            if (!actual.HasValue) incomplete = true;
            else if (requested != actual) differences.Add($"{name} registry={requested} active={actual}");
        }
        Check("base", registryBase, state!.BaseProcessorNumber);
        if (registryBase.HasValue) Check("baseGroup", 0, state.BaseProcessorGroup);
        Check("queues", registryQueues, state.NumberOfReceiveQueues);
        Check("maxProcessors", registryMaxProcessors, state.MaxProcessors);
        if (differences.Count > 0) return new("different", string.Join("; ", differences));
        if (incomplete) return new("unverified", "requested fields missing from readback");
        bool configured = registryBase.HasValue || registryQueues.HasValue || registryMaxProcessors.HasValue;
        return new(configured ? "ok" : "read", configured ? "requested fields match readback" : "no explicit RSS overrides");
    }
}
