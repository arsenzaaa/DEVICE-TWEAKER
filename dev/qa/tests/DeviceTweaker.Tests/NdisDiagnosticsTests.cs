using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public sealed class NdisDiagnosticsTests
{
    private static readonly NdisRssRuntimeState Read = new(true, true, true, 0, 2, 0, 15, 4, 4, "TEST", "TEST", "TEST", "");

    [Theory]
    [InlineData("rss-not-found", "unavailable")]
    [InlineData("rss-not-supported", "unsupported")]
    [InlineData("adapter-not-found", "unavailable")]
    [InlineData("powershell-timeout", "failed")]
    [InlineData("access denied", "failed")]
    [InlineData("wifi-preserved", "skipped")]
    public void MissingReadbackNeverBecomesOk(string reason, string expected)
    {
        var state = Read with { RssFound = false, Enabled = null, Error = reason };
        Assert.Equal(expected, NdisRssDiagnostics.Compare(state, null, null, null).Status);
        Assert.Equal(expected, NdisRssDiagnostics.Compare(state, 2, 4, 4).Status);
    }

    [Fact]
    public void WifiComparisonIsSkippedEvenWithPreexistingOverridesAndConflictingReadback()
    {
        Assert.Equal("skipped", NdisRssDiagnostics.Compare(Read, 8, 1, 1, wifiPreserved: true).Status);
        Assert.Equal("different", NdisRssDiagnostics.Compare(Read, 8, 1, 1).Status);
    }

    [Fact]
    public void OkRequiresAllRequestedFieldsAndTheProcessorGroupToBeRead()
    {
        Assert.Equal("ok", NdisRssDiagnostics.Compare(Read, 2, 4, 4).Status);
        Assert.Equal("unverified", NdisRssDiagnostics.Compare(Read with { NumberOfReceiveQueues = null }, 2, 4, 4).Status);
        Assert.Equal("unverified", NdisRssDiagnostics.Compare(Read with { BaseProcessorGroup = null }, 2, 4, 4).Status);
        Assert.Equal("different", NdisRssDiagnostics.Compare(Read with { BaseProcessorGroup = 1 }, 2, 4, 4).Status);
        Assert.Equal("read", NdisRssDiagnostics.Compare(Read, null, null, null).Status);
        Assert.Equal("disabled", NdisRssDiagnostics.Compare(Read with { Enabled = false }, 2, 4, 4).Status);
        Assert.Equal("unavailable", NdisRssDiagnostics.Compare(null, null, null, null).Status);
    }

    [Theory]
    [InlineData(4, 0UL, 0)]
    [InlineData(0, 8UL, 4)]
    [InlineData(4, 8UL, 4)]
    [InlineData(1, 8UL, 1)]
    [InlineData(5, 0UL, 0)]
    public void DiagnosticPolicyUsesTheSameNormalizationAsTheCpuSelection(int registryPolicy, ulong mask, int expected)
        => Assert.Equal(expected, MainForm.NormalizePolicyForSelection(registryPolicy, mask));

    [Theory]
    [InlineData("ok", "")]
    [InlineData("missing-adapter", "adapter-not-found")]
    [InlineData("missing-rss", "rss-not-found")]
    [InlineData("adapter-denied", "adapter-query-failed: ")]
    [InlineData("rss-denied", "rss-query-failed: ")]
    public async Task RealQueryScriptSeparatesEmptyResultsFromExceptionsWithoutCallingWindowsCmdlets(string mode, string expected)
    {
        // Shadow both cmdlets in the child process. No real adapter or settings are accessed.
        string adapterBody = mode == "adapter-denied" ? "throw [UnauthorizedAccessException]::new('probe denial')"
            : mode == "missing-adapter" ? "return" : "[pscustomobject]@{Name='TEST';PnPDeviceID='TEST';InterfaceGuid='';InterfaceDescription='TEST'}";
        string rssBody = mode == "rss-denied" ? "throw [UnauthorizedAccessException]::new('probe denial')"
            : mode == "missing-rss" ? "return" : "[pscustomobject]@{Enabled=$true;BaseProcessorGroup=0;BaseProcessorNumber=2;NumberOfReceiveQueues=4}";
        string mocks = "function Get-NetAdapter { [CmdletBinding()]param([switch]$IncludeHidden) " + adapterBody + " };"
            + "function Get-NetAdapterRss { [CmdletBinding()]param([string]$Name,[switch]$IncludeHidden) " + rssBody + " };";
        using System.Diagnostics.Process process = new();
        process.StartInfo = new("powershell.exe") { UseShellExecute = false, CreateNoWindow = true,
            RedirectStandardOutput = true, RedirectStandardError = true };
        process.StartInfo.ArgumentList.Add("-NoProfile");
        process.StartInfo.ArgumentList.Add("-Command");
        process.StartInfo.ArgumentList.Add(mocks + MainForm.BuildNdisRssQueryScript(null, "TEST"));
        Assert.True(process.Start());
        Task<string> output = process.StandardOutput.ReadToEndAsync(), error = process.StandardError.ReadToEndAsync();
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(15));
        try { await process.WaitForExitAsync(timeout.Token); }
        catch { if (!process.HasExited) process.Kill(entireProcessTree: true); throw; }
        Assert.Equal(0, process.ExitCode);
        Assert.True(string.IsNullOrWhiteSpace(await error), await error);
        using var json = System.Text.Json.JsonDocument.Parse(await output);
        string reason = json.RootElement.GetProperty("Error").GetString()!;
        Assert.StartsWith(expected, reason);
        Assert.Equal(mode == "ok", json.RootElement.GetProperty("RssFound").GetBoolean());
        if (mode == "ok") Assert.Equal(string.Empty, reason);
    }
}
