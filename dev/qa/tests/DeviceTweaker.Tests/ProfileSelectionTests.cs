using DeviceTweakerQa;
using Xunit;

namespace DeviceTweaker.Tests;

public sealed class ProfileSelectionTests
{
    private static readonly string[] Catalog = ["system-00", "system-01", "cpu-00"];

    [Fact]
    public void EmptyFilterPreservesTheFullCatalog()
        => Assert.True(ProfileSelection.Parse("", Catalog).SetEquals(Catalog));

    [Fact]
    public void ExplicitSubsetDoesNotAddOtherProfiles()
        => Assert.True(ProfileSelection.Parse(" system-01, cpu-00 ", Catalog).SetEquals(["system-01", "cpu-00"]));

    [Theory]
    [InlineData("system-99")]
    [InlineData("system-00,system-00")]
    [InlineData("system-00,")]
    [InlineData("SYSTEM-00")]
    public void InvalidSelectionFailsInsteadOfSilentlySkippingCoverage(string filter)
        => Assert.Throws<InvalidDataException>(() => ProfileSelection.Parse(filter, Catalog));
}
