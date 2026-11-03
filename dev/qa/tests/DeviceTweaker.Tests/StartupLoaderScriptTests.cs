using System.Text;
using System.Text.Json;
using Xunit;

namespace DeviceTweaker.Tests;

public sealed partial class StartupScriptTests
{
    [Theory]
    [InlineData("timeout-running", 1, true, 0)]
    [InlineData("timeout-killed", 1, true, 0)]
    [InlineData("completed-empty", 0, false, 1)]
    [InlineData("completed-ready", 0, false, 0)]
    [InlineData("start-failed", 0, false, 1)]
    public async Task ShippedLoaderKeepsTimeoutPendingAndBlocksServiceFallback(string mode, int exitCode, bool pending, int serviceCalls)
    {
        string folder = NewFolder("startup-loader");
        WriteScript(folder);
        string probe = Path.Combine(folder, "loader-probe.ps1");
        File.WriteAllText(probe, """
            param([string]$Mode)
            $ErrorActionPreference = 'Stop'
            Add-Type -TypeDefinition @'
            using System;
            using System.IO;
            using System.Diagnostics;
            public static class ProbeRuntime {
                public static bool Available;
                public static int Services;
                public static bool IsDriverDeviceAvailable() { return Available; }
                public static void EnsureDriver(string path) { Services++; Available = true; }
            }
            public static class ProbeJournal {
                public static bool Pending;
                public static void AssertLoaderIdle() { if (Pending) throw new InvalidOperationException("pending load"); }
                public static void BeginLoader() { AssertLoaderIdle(); Pending = true; }
                public static void CompleteLoader() { Pending = false; }
            }
            public sealed class ProbeProcess : IDisposable {
                public static string Mode;
                public bool HasExited;
                public int ExitCode { get { return 0; } }
                public StringReader StandardOutput = new StringReader("fake output");
                public StringReader StandardError = new StringReader("");
                public static ProbeProcess Start(ProcessStartInfo info) {
                    if (Mode == "start-failed") throw new IOException("start failed");
                    return new ProbeProcess();
                }
                public bool WaitForExit(int timeout) {
                    if (Mode.StartsWith("timeout")) return HasExited;
                    HasExited = true;
                    ProbeRuntime.Available = Mode == "completed-ready";
                    return true;
                }
                public void Kill() { HasExited = Mode == "timeout-killed"; }
                public void Dispose() { }
            }
            public sealed class ProbeGate {
                public void ReleaseMutex() { }
                public void Dispose() { }
            }
            '@
            $tokens = $null; $errors = $null
            $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot 'ApplyIMOD.ps1'), [ref]$tokens, [ref]$errors)
            if ($errors.Count) { throw ($errors | Out-String) }
            $definition = $ast.Find({ param($a) $a -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $a.Name -eq 'Invoke-ImodKduLoader' }, $true)
            $loader = $definition.Extent.Text
            if (($loader.Split(@('[System.Diagnostics.Process]::Start'), [System.StringSplitOptions]::None).Length - 1) -ne 1) { throw 'Process adapter anchor changed' }
            # Only the native/process boundaries are replaced. Branches, waits, journal calls and return values are shipped code.
            $loader = $loader.Replace('[System.Diagnostics.Process]::Start', '[ProbeProcess]::Start').Replace('[DeviceTweakerImodRuntime]', '[ProbeRuntime]').Replace('[DeviceTweakerLoaderJournal]', '[ProbeJournal]')
            Invoke-Expression $loader
            function Write-ImodLog { param([string]$text) }
            function Test-ImodFileHash { return $true }
            $ImodKduPath = Join-Path $PSScriptRoot 'fake-loader.txt'
            $ImodKduDbPath = Join-Path $PSScriptRoot 'fake-db.txt'
            $ImodDriverPath = Join-Path $PSScriptRoot 'fake-driver.txt'
            foreach ($path in @($ImodKduPath, $ImodKduDbPath, $ImodDriverPath)) { [System.IO.File]::WriteAllText($path, 'data only; never executed') }
            $ImodDriverSha256 = ''; $ImodKduSha256 = ''; $ImodKduDbSha256 = ''
            [ProbeProcess]::Mode = $Mode
            $result = @(Invoke-ImodKduLoader)
            if ($result.Count -ne 1 -or $result[0] -isnot [bool]) { throw 'Loader leaked pipeline output instead of one Boolean' }
            [ProbeJournal]::Pending = $false; [ProbeRuntime]::Available = $false; [ProbeRuntime]::Services = 0
            $applyUsbImod = $false; $nicItrData = @(); $imodLoadGate = New-Object ProbeGate
            $coordinator = $ast.Find({ param($a) $a -is [System.Management.Automation.Language.TryStatementAst] -and $a.Extent.Text.Contains('startup apply begin; driver=') }, $true)
            if ($null -eq $coordinator) { throw 'Startup coordinator missing' }
            $body = $coordinator.Extent.Text.Replace('[DeviceTweakerImodRuntime]', '[ProbeRuntime]').Replace('[DeviceTweakerLoaderJournal]', '[ProbeJournal]')
            Invoke-Expression $body
            [pscustomobject]@{ exitCode=$imodExitCode; pending=[ProbeJournal]::Pending; services=[ProbeRuntime]::Services; loaderResult=$result[0] } | ConvertTo-Json -Compress
            """, new UTF8Encoding(true));

        (int code, string output) = await RunPowerShell(probe, "-Mode", mode);
        Assert.True(code == 0, output);
        using JsonDocument json = JsonDocument.Parse(output.Trim().Split('\n')[^1]);
        Assert.Equal(exitCode, json.RootElement.GetProperty("exitCode").GetInt32());
        Assert.Equal(pending, json.RootElement.GetProperty("pending").GetBoolean());
        Assert.Equal(serviceCalls, json.RootElement.GetProperty("services").GetInt32());
        Assert.Equal(mode == "completed-ready", json.RootElement.GetProperty("loaderResult").GetBoolean());
    }
}
