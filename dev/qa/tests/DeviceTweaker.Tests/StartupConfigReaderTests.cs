using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public sealed class StartupConfigReaderTests : IDisposable
{
    private readonly string _directory = Path.Combine(Path.GetTempPath(), "DeviceTweakerProfile-" + Guid.NewGuid().ToString("N"));
    public StartupConfigReaderTests() => Directory.CreateDirectory(_directory);
    public void Dispose() => Directory.Delete(_directory, recursive: true);
    private string Current => Path.Combine(_directory, "profile.json");
    private string Legacy => Path.Combine(_directory, "legacy.ps1");
    private StartupConfigReadResult<MainForm.ImodConfig> Read(Func<string, string>? readCurrent = null)
        => StartupConfigReader.Read(Current, new[] { Legacy },
            path => MainForm.ConvertStartupProfile(StartupProfile.FromJson((readCurrent ?? File.ReadAllText)(path))),
            path => MainForm.ParseImodScriptText(File.ReadAllText(path)));

    [Theory]
    [InlineData("{broken")]
    [InlineData("null")]
    [InlineData("{\"schemaVersion\":2}")]
    [InlineData("{\"usb\":null}")]
    [InlineData("{\"usb\":{\"controllers\":[null]}}")]
    [InlineData("{\"nicItr\":[{\"values\":null}]}")]
    [InlineData("{\"usb\":{\"interval\":\"invalid\"}}")]
    public void CorruptCurrentProfileCannotBecomeEmptyOrFallBackToLegacy(string json)
    {
        File.WriteAllText(Current, json);
        File.WriteAllText(Legacy, "$imodInterval = 0x0");
        var result = Read();
        Assert.Equal(StartupConfigReadState.Error, result.State);
        Assert.Equal(Current, result.Source);
        Assert.Null(result.Value);
        Assert.Throws<InvalidDataException>(() => result.RequireValueOrCreate(() => new()));
        Assert.Equal(json, File.ReadAllText(Current));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AccessAndIoErrorsAreNotMissingFiles(bool accessDenied)
    {
        File.WriteAllText(Legacy, "$imodInterval = 0x0");
        var result = Read(_ => throw (accessDenied ? new UnauthorizedAccessException("denied") : new IOException("sharing violation")));
        Assert.Equal(StartupConfigReadState.Error, result.State);
        Assert.Equal(Current, result.Source);
        Assert.Throws<InvalidDataException>(() => result.RequireValueOrCreate(() => new()));
    }

    [Fact]
    public void MissingCurrentPermitsLegacyAndMissingAllPermitsNewProfile()
    {
        var result = Read();
        Assert.Equal(StartupConfigReadState.Missing, result.State);
        Assert.False(result.RequireValueOrCreate(() => new()).HasScript);
        File.WriteAllText(Legacy, "$imodInterval = 0x0");
        result = Read();
        Assert.Equal(StartupConfigReadState.Loaded, result.State);
        Assert.Equal(Legacy, result.Source);
    }

    [Fact]
    public void ValidCurrentPreservesUsbAndNicAndNeverReadsLegacy()
    {
        File.WriteAllText(Current, MainForm.BuildStartupProfile(StartupProfileTests.Sample(), @"C:\Windows\DTIMOD.sys").ToJson());
        var result = StartupConfigReader.Read(Current, new[] { Legacy },
            path => MainForm.ConvertStartupProfile(StartupProfile.FromJson(File.ReadAllText(path))),
            _ => throw new Exception("Legacy must not be read"));
        Assert.Equal(StartupConfigReadState.Loaded, result.State);
        var config = result.RequireValueOrCreate(() => new());
        Assert.Equal(2, config.Overrides.Count);
        Assert.Single(config.NicItrEntries);
    }

    [Fact]
    public void UnreadableLegacyStopsBeforeTryingOlderLayout()
    {
        int legacyReads = 0;
        var result = StartupConfigReader.Read<MainForm.ImodConfig>(Current, new[] { Legacy, "older.ps1" },
            _ => throw new FileNotFoundException(),
            _ => { legacyReads++; throw new UnauthorizedAccessException("denied"); });
        Assert.Equal(StartupConfigReadState.Error, result.State);
        Assert.Equal(1, legacyReads);
    }
}
