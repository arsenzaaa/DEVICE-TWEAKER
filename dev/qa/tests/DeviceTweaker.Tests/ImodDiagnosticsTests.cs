using System.Reflection;
using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public sealed class ImodDiagnosticsTests
{
    [Theory]
    [InlineData(2, true)]
    [InlineData(3, true)]
    [InlineData(5, false)]
    [InlineData(32, false)]
    public void MissingEndpointPromptsCheckWithoutHidingOtherOpenErrors(int code, bool needsCheck)
    {
        string error = (string)typeof(MainForm)
            .GetMethod("FormatImodDeviceOpenError", BindingFlags.Static | BindingFlags.NonPublic)!
            .Invoke(null, [code])!;
        Assert.Contains($"Win32 {code}", error);
        string status = (string)typeof(MainForm)
            .GetMethod("FormatImodReadbackStatus", BindingFlags.Static | BindingFlags.NonPublic)!
            .Invoke(null, [error])!;
        Assert.Equal(needsCheck ? "current: press CHECK" : "current: unavailable", status);
    }

    private static string Unavailable(string error) => (string)typeof(MainForm)
        .GetMethod("FormatImodUnavailableUserMessage", BindingFlags.Static | BindingFlags.NonPublic)!
        .Invoke(null, [error, true])!;

    [Theory]
    [InlineData("Service fallback: Windows error 577 (digital signature could not be verified)")]
    [InlineData("VulnerableDriverBlocklistEnable=0(off); Kernel CI: Windows error 577")]
    public void SignatureErrorDoesNotIdentifyTheBlocklist(string error)
    {
        string result = Unavailable(error);
        Assert.DoesNotContain("via Vulnerable Driver Blocklist", result, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("USB IMOD was not changed", result, StringComparison.Ordinal);
    }

    [Fact]
    public void GenericVirusErrorDoesNotIdentifyDefender()
    {
        string result = Unavailable("Windows error 225: file contains a virus or potentially unwanted software");
        Assert.DoesNotContain("Windows Defender blocked", result, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("USB IMOD was not changed", result, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("digital signature rejected; Windows error 225: virus or potentially unwanted software")]
    [InlineData("Windows error 577; quarantine")]
    public void ReadbackDetailDoesNotAttributeCombinedErrorsToDefender(string error)
    {
        string result = (string)typeof(MainForm)
            .GetMethod("FormatImodReadbackDetail", BindingFlags.Static | BindingFlags.NonPublic)!
            .Invoke(null, [error])!;
        Assert.Equal("DTIMOD.sys rejected by Windows or antivirus", result);
    }
}
