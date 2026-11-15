using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public sealed class ImodDriverFilesTests
{
    [Fact]
    public void ReadOnlySessionCannotStageDriverOrLoaderFiles()
    {
        Assert.False(ImodDriverFiles.EnsureOnDisk(true, _ => { }, out string path, out string? error));
        Assert.Empty(path);
        Assert.Contains("QA read-only", error);
        Assert.Throws<QaWriteBlockedException>(() => ImodDriverFiles.TryWriteEmbeddedImodResource(
            "DeviceTweakerCS.IMOD.Loader.kdu.exe", ".kdu.exe", @"C:\never-write\kdu.exe", out _));
    }

    [Fact]
    public void PayloadVerificationRetainsTheOriginalReadOnlyGuard()
    {
        const string name = "DeviceTweakerCS.IMOD.DTIMOD.sys";
        Assert.Throws<QaWriteBlockedException>(() => ImodDriverFiles.TryVerifyEmbeddedImodResourceFile(
            name, ".DTIMOD.sys", @"C:\never-read\driver.sys", out _));
    }

    [Fact]
    public void EmbeddedHashIdentifiesTheSameBundledDriver()
    {
        using Stream resource = ImodDriverFiles.OpenImodDriverResourceStream()!;
        using MemoryStream content = new();
        resource.CopyTo(content);
        Assert.True(ImodDriverFiles.HashEquals(ImodDriverFiles.ReadEmbeddedImodDriverHash(),
            ImodDriverFiles.ComputeSha256(content.ToArray())));
    }

    [Theory]
    [InlineData(@"C:\ProgramData\DEVICE TWEAKER\IMOD\Loader\kdu.exe", @"C:\Windows\DTIMOD.sys")]
    [InlineData(@"C:\test path\Loader\kdu.exe", @"C:\test path\DTIMOD.sys")]
    public void LoaderStartInfoKeepsArgumentsIsolationAndCapturedOutput(string loader, string driver)
    {
        var info = KduDriverLoader.BuildStartInfo(loader, driver);
        Assert.Equal(loader, info.FileName);
        Assert.Equal(Path.GetDirectoryName(loader), info.WorkingDirectory);
        Assert.Equal($"-scv 3 -drvn DeviceTweakerImod2 -drvr DeviceTweakerImod2 -map \"{driver}\"", info.Arguments);
        Assert.False(info.UseShellExecute);
        Assert.True(info.CreateNoWindow);
        Assert.True(info.RedirectStandardOutput);
        Assert.True(info.RedirectStandardError);
    }
}
