namespace DeviceTweakerCS;

// Intel and Realtek drivers keep their own interrupt moderation setting
// (*InterruptModeration, Intel ITR). In adaptive mode the driver rewrites
// EITR / IntrMit while running, so a saved NIC ITR profile switches the
// driver's moderation off and leaves the register to the startup profile.
// The class-key value is read when the adapter starts, i.e. after a restart.
internal static class NicDriverModeration
{
    internal const string OffValue = "0";
    internal static readonly string[] KeywordNames = ["*InterruptModeration", "ITR"];

    // Options are the value names under Ndi\params\<keyword>\enum, mapped to
    // the driver's display text.
    internal sealed record Keyword(
        string Name,
        IReadOnlyDictionary<string, string> Options,
        string? Default,
        string? Current);

    internal sealed record Change(string Name, string? Before, string After);

    internal static List<Change> PlanDisable(IEnumerable<Keyword> keywords)
        => keywords
            .Where(keyword => keyword.Options.ContainsKey(OffValue) && keyword.Current != OffValue)
            .Select(keyword => new Change(keyword.Name, keyword.Current, OffValue))
            .ToList();

    // Only values still set to Off are returned to the INF default, so a later
    // manual choice in Device Manager is kept.
    internal static List<Change> PlanRestoreDefault(IEnumerable<Keyword> keywords)
        => keywords
            .Where(keyword => keyword.Current == OffValue
                && !string.IsNullOrWhiteSpace(keyword.Default)
                && keyword.Default != OffValue)
            .Select(keyword => new Change(keyword.Name, keyword.Current, keyword.Default!))
            .ToList();

    internal static string? Describe(IEnumerable<Keyword> keywords)
    {
        List<string> parts = [];
        foreach (Keyword keyword in keywords)
        {
            string? value = keyword.Current ?? keyword.Default;
            if (value is null)
            {
                continue;
            }

            string text = keyword.Options.TryGetValue(value, out string? label) && !string.IsNullOrWhiteSpace(label)
                ? label
                : value;
            parts.Add(keyword.Name == "ITR" ? $"ITR {text}" : text);
        }

        return parts.Count == 0 ? null : string.Join(", ", parts);
    }
}
