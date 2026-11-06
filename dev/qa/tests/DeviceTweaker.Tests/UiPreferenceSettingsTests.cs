using System.Text.Json;
using Xunit;

namespace DeviceTweakerCS.Tests;

public sealed class UiPreferenceSettingsTests
{
    [Theory]
    [InlineData("animationsEnabled")]
    [InlineData("AnimationsEnabled")]
    public void LanguageSaveDropsObsoleteMotionFlagAndRetainsUnknownSettings(string oldKey)
    {
        string old = "{\"language\":\"ru\",\"" + oldKey + "\":false,\"schemaVersion\":1,\"futurePreference\":7}";
        using JsonDocument saved = JsonDocument.Parse(UiPreferenceSettings.Update(old, "en"));
        Assert.Equal("en", saved.RootElement.GetProperty("language").GetString());
        Assert.Equal(7, saved.RootElement.GetProperty("futurePreference").GetInt32());
        Assert.False(saved.RootElement.TryGetProperty(oldKey, out _));
        Assert.False(saved.RootElement.TryGetProperty("schemaVersion", out _));
    }

    [Fact]
    public void NewSettingsContainLanguageWithoutMotionPreference()
    {
        using JsonDocument saved = JsonDocument.Parse(UiPreferenceSettings.Update(null, "ru"));
        Assert.Single(saved.RootElement.EnumerateObject());
        Assert.Equal("ru", saved.RootElement.GetProperty("language").GetString());
    }

    [Fact]
    public void MalformedPreferencesAreRejectedInsteadOfBeingOverwritten()
        => Assert.Throws<JsonException>(() => UiPreferenceSettings.Update("{broken", "en"));
}
