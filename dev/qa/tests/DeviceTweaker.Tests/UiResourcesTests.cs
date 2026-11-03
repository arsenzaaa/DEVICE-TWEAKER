using System.Globalization;
using System.Text.RegularExpressions;
using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public class UiResourcesTests
{
    [Fact]
    public void RussianSatelliteContainsEveryNeutralKeyWithoutFallback()
    {
        var english = UiResources.ReadCulture(CultureInfo.InvariantCulture);
        var russian = UiResources.ReadCulture(CultureInfo.GetCultureInfo("ru"));
        Assert.NotEmpty(english);
        Assert.Equal(english.Keys.Order(), russian.Keys.Order());
        foreach (var entry in english)
        {
            Assert.False(string.IsNullOrWhiteSpace(russian[entry.Key]), entry.Key);
            Assert.Equal(Placeholders(entry.Value), Placeholders(russian[entry.Key]));
        }
    }

    [Theory]
    [InlineData(UiText.Apply, "APPLY", "ПРИМЕНИТЬ")]
    [InlineData(UiText.ReadIrq, "READ IRQ MODE", "СЧИТАТЬ IRQ")]
    [InlineData(UiText.Changed, "Changed", "Изменено")]
    [InlineData(UiText.ProfileUnreadable, "profile: unreadable", "профиль: ошибка чтения")]
    public void StableKeysResolveCaptionsInBothLanguages(string key, string english, string russian)
    {
        Assert.Equal(english, UiResources.English(key));
        Assert.Equal(russian, UiResources.ReadCulture(CultureInfo.GetCultureInfo("ru"))[key]);
    }

    [Fact]
    public void UnknownKeysFailInsteadOfShowingAKeyToTheUser()
        => Assert.Throws<System.Resources.MissingManifestResourceException>(() => UiResources.English("Action.Unknown"));

    private static string[] Placeholders(string text)
        => Regex.Matches(text, @"(?<!\{)\{\d+(?:[^{}]*)\}(?!\})")
            .Select(match => match.Value).Order().ToArray();
}
