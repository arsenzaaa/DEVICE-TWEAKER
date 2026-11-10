namespace DeviceTweakerQa;

internal static class ProfileSelection
{
    internal static HashSet<string> Parse(string filter, IEnumerable<string> catalog)
    {
        HashSet<string> known = new(catalog, StringComparer.Ordinal);
        if (string.IsNullOrWhiteSpace(filter)) return known;
        string[] ids = filter.Split(',').Select(id => id.Trim()).ToArray();
        HashSet<string> selected = new(ids, StringComparer.Ordinal);
        if (ids.Any(string.IsNullOrEmpty) || selected.Count != ids.Length || selected.Any(id => !known.Contains(id)))
            throw new InvalidDataException("Profile selection contains empty, duplicate or unknown IDs: " + filter);
        return selected;
    }
}
