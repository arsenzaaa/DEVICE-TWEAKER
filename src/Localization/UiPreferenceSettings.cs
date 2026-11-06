using System.Text.Json;
using System.Text.Json.Serialization;

namespace DeviceTweakerCS;

// Retain unrelated settings while updating the interface language.
internal sealed class UiPreferenceSettings
{
    public string Language { get; set; } = "en";
    [JsonExtensionData]
    public Dictionary<string, JsonElement>? Extra { get; set; }

    private static readonly JsonSerializerOptions Options = new()
    { PropertyNamingPolicy = JsonNamingPolicy.CamelCase, PropertyNameCaseInsensitive = true, WriteIndented = true };

    internal static UiPreferenceSettings Read(string json)
        => JsonSerializer.Deserialize<UiPreferenceSettings>(json, Options) ?? new();

    internal static string Update(string? json, string language)
    {
        UiPreferenceSettings settings = json is null ? new() : Read(json);
        settings.Language = language;
        if (settings.Extra is not null)
        {
            foreach (string key in settings.Extra.Keys.Where(key =>
                key.Equals("schemaVersion", StringComparison.OrdinalIgnoreCase)
                || key.Equals("animationsEnabled", StringComparison.OrdinalIgnoreCase)).ToArray())
                settings.Extra.Remove(key);
        }
        return JsonSerializer.Serialize(settings, Options);
    }
}
