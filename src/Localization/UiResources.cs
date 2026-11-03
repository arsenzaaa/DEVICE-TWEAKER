using System.Collections;
using System.Globalization;
using System.Resources;

namespace DeviceTweakerCS;

// Resource keys stay independent of captions. The phrase adapter supports existing
// diagnostic messages while controls are moved to explicit keys incrementally.
internal static class UiResources
{
    private static readonly ResourceManager Manager = new("DeviceTweakerCS.Localization.Strings", typeof(UiResources).Assembly);
    private static readonly CultureInfo RussianCulture = CultureInfo.GetCultureInfo("ru");

    internal static string English(string key) => Required(key, CultureInfo.InvariantCulture);
    internal static string Get(string key) => Required(key, UiLanguage.IsRussian ? RussianCulture : CultureInfo.InvariantCulture);

    private static string Required(string key, CultureInfo culture)
        => Manager.GetString(key, culture) ?? throw new MissingManifestResourceException($"Missing UI resource: {key}");

    internal static IReadOnlyDictionary<string, string> ReadCulture(CultureInfo culture)
    {
        ResourceSet set = Manager.GetResourceSet(culture, true, false)
            ?? throw new MissingManifestResourceException($"Missing UI resources for {culture.Name}");
        return set.Cast<DictionaryEntry>().ToDictionary(entry => (string)entry.Key, entry => (string)entry.Value!, StringComparer.Ordinal);
    }

    internal static IReadOnlyDictionary<string, string> LoadRussianPhrases()
        => ReadCulture(CultureInfo.InvariantCulture)
            .Where(entry => !entry.Key.StartsWith("Inline.", StringComparison.Ordinal))
            .ToDictionary(entry => entry.Value, entry => Required(entry.Key, RussianCulture), StringComparer.Ordinal);

    internal static (string English, string Russian)[] LoadInlineReplacements()
        => ReadCulture(CultureInfo.InvariantCulture)
            .Where(entry => entry.Key.StartsWith("Inline.", StringComparison.Ordinal))
            .OrderBy(entry => entry.Key, StringComparer.Ordinal)
            .Select(entry => (entry.Value, Required(entry.Key, RussianCulture)))
            .ToArray();
}
