using Xunit;

namespace DeviceTweakerCS.Tests;

public sealed class NdisRssReaderTests
{
    [Fact]
    public void ExactGuidWinsOverAnotherAdaptersPnpId()
    {
        string guid = Guid.NewGuid().ToString();
        var expected = new NdisRssReader.Adapter("{" + guid + "}", "PCI\\ONE", "one", "first");
        var other = new NdisRssReader.Adapter(Guid.NewGuid().ToString(), "PCI\\TWO", "two", "second");
        Assert.Same(expected, NdisRssReader.MatchAdapter([other, expected], guid.ToUpperInvariant(), other.PnpId));
    }
    [Fact]
    public void MissingOrStaleGuidFallsBackToExactPnpIdentity()
    {
        var expected = new NdisRssReader.Adapter(Guid.NewGuid().ToString(), "PCI\\ONE", "one", "first");
        Assert.Same(expected, NdisRssReader.MatchAdapter([expected], "unavailable", "pci\\one"));
        Assert.Null(NdisRssReader.MatchAdapter([expected], null, "PCI\\UNRELATED"));
    }
    [Fact]
    public void MissingIdentitiesCannotSelectAnAnonymousAdapter()
        => Assert.Null(NdisRssReader.MatchAdapter([new("", "", "virtual", "anonymous")], null, ""));
    [Fact]
    public void AdapterNamesAreEscapedForReadOnlyWql()
        => Assert.Equal("a\\\\b\\'c", NdisRssReader.Escape("a\\b'c"));
}
