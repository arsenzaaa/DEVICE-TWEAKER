using System.Reflection;
using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public sealed class DiagnosticRegressionTests
{
    [Theory]
    [InlineData("LOAD.RSS.ACTIVE: adapterFound=True rssFound=False active=Unknown error=\"rss-not-found\"", "Info")]
    [InlineData("GUI.BLOCK.RSS.ACTIVE: rssFound=False error=\"wifi-preserved\"", "Info")]
    [InlineData("LOAD.RSS.ACTIVE: rssFound=False error=\"rss-not-supported\"", "Info")]
    [InlineData("LOAD.RSS.ACTIVE: rssFound=False error=\"rss-not-found then write failed\"", "Error")]
    [InlineData("LOAD.RSS.ACTIVE: adapterFound=False rssFound=False active=Unknown error=\"powershell-timeout\"", "Error")]
    [InlineData("APPLY.REG.ERROR: operation=set-msi error=\"Access denied\"", "Error")]
    public void ExpectedMissingRssIsNotAWriteFailure(string line, string expected)
    {
        object actual = typeof(AppDiagnostics).GetMethod("InferLevel", BindingFlags.NonPublic | BindingFlags.Static)!
            .Invoke(null, [line])!;
        Assert.Equal(expected, actual.ToString());
    }

    [Theory]
    [InlineData(typeof(MainForm))]
    [InlineData(typeof(AppDiagnostics))]
    public void SessionAndCrashVersionsIdentifyTheActualBuild(Type owner)
    {
        string actual = (string)owner.GetMethod("GetAppVersion", BindingFlags.NonPublic | BindingFlags.Static)!.Invoke(null, null)!;
        Assert.Contains("product=0.0.4", actual);
        Assert.Contains($"build={typeof(MainForm).Assembly.ManifestModule.ModuleVersionId:N}", actual);
        Assert.Contains("exeSha256=", actual);
    }

    [Fact]
    public void ExecutableDigestMatchesExactFileContentsAndReportsMissingFiles()
    {
        string path = Path.Combine(QaRun.OutputDirectory, "identity-probe.bin");
        Directory.CreateDirectory(QaRun.OutputDirectory);
        File.WriteAllBytes(path, [1, 2, 3, 4]);
        string expected = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(new byte[] { 1, 2, 3, 4 }));
        Assert.Equal(expected, AppBuildIdentity.HashExecutable(path));
        Assert.StartsWith("unavailable(", AppBuildIdentity.HashExecutable(path + ".missing"));
    }
}
