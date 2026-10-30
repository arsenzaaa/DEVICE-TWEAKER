using System.Diagnostics;
using System.Text;
using System.Text.Json;
using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

// Runs the shipped startup script and runtime in Windows PowerShell 5.1, the
// interpreter Task Scheduler uses. No driver, loader or SCM call is reached.
public sealed class StartupScriptTests
{
    internal static string ReadResource(string name)
    {
        using Stream stream = typeof(MainForm).Assembly.GetManifestResourceStream(name)!;
        using StreamReader reader = new(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
        return reader.ReadToEnd();
    }

    private static string NewFolder(string name)
    {
        string folder = Path.Combine(QaRun.OutputDirectory, name + "-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        return folder;
    }

    private static string WriteScript(string folder)
    {
        string path = Path.Combine(folder, "ApplyIMOD.ps1");
        File.WriteAllText(path, ReadResource("DeviceTweakerCS.Startup.ApplyIMOD.ps1"), new UTF8Encoding(true));
        return path;
    }

    private static async Task<(int ExitCode, string Output)> RunPowerShell(string scriptPath, params string[] arguments)
    {
        ProcessStartInfo info = new("powershell.exe")
        {
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
        };
        foreach (string argument in new[] { "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", scriptPath }.Concat(arguments))
        {
            info.ArgumentList.Add(argument);
        }

        using Process process = Process.Start(info)!;
        Task<string> stdout = process.StandardOutput.ReadToEndAsync();
        Task<string> stderr = process.StandardError.ReadToEndAsync();
        if (!process.WaitForExit(60000))
        {
            process.Kill(entireProcessTree: true);
            throw new TimeoutException("Windows PowerShell did not finish");
        }

        return (process.ExitCode, await stdout + await stderr);
    }

    private static string ReadLog(string folder)
    {
        string logs = Path.Combine(folder, "logs");
        return Directory.Exists(logs)
            ? string.Join("\n", Directory.GetFiles(logs, "ApplyIMOD_*.log").Select(File.ReadAllText))
            : string.Empty;
    }

    [Fact]
    public async Task ImportImodProfileGivesTheScriptTheSavedValues()
    {
        string folder = NewFolder("startup-profile");
        File.WriteAllText(Path.Combine(folder, "profile.json"),
            MainForm.BuildStartupProfile(StartupProfileTests.Sample(), @"C:\Windows\DTIMOD.sys").ToJson());
        WriteScript(folder);
        // Loads only the two functions from the shipped script, then prints what they set.
        string probe = Path.Combine(folder, "probe.ps1");
        File.WriteAllText(probe, """
            $ErrorActionPreference = 'Stop'
            $tokens = $null; $errors = $null
            $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot 'ApplyIMOD.ps1'), [ref]$tokens, [ref]$errors)
            if ($errors.Count) { throw ($errors | Out-String) }
            foreach ($name in @('ConvertFrom-ImodNumber', 'Import-ImodProfile')) {
                $definition = $ast.Find({ param($a) $a -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $a.Name -eq $name }, $true)
                Invoke-Expression $definition.Extent.Text
            }
            $scriptRoot = $PSScriptRoot
            Import-ImodProfile -path (Join-Path $PSScriptRoot 'profile.json')
            $intel = $userDefinedData['VEN_8086&DEV_7AE0']
            $amd = $userDefinedData['VEN_1022&DEV_15B6']
            $nic = $nicItrData[0]
            [pscustomobject]@{
                apply = $applyUsbImod
                interval = [uint64]$globalInterval
                rtsoff = [uint64]$globalRTSOFF
                kdu = $ImodKduPath
                driver = $ImodDriverPath
                runtimeHashLength = $ImodRuntimeSha256.Length
                intelEnabled = $intel['ENABLED']
                intelAdaptive = $intel['ADAPTIVE_ROLE_BINDING']
                intelRoles = $intel['ROLE_INTERVALS']
                intelHasIntervals = $intel.ContainsKey('INTERVALS')
                amdIntervals = @($amd['INTERVALS'] | ForEach-Object { [uint64]$_ })
                nicCount = @($nicItrData).Count
                nicHwid = $nic['HWID']
                nicOrBits = [uint64]$nic['OR_BITS']
                nicValues = @($nic['VALUES'] | ForEach-Object { [uint64]$_ })
            } | ConvertTo-Json -Compress
            """, new UTF8Encoding(true));

        (int code, string output) = await RunPowerShell(probe);

        Assert.True(code == 0, output);
        using JsonDocument result = JsonDocument.Parse(output.Trim().Split('\n')[^1]);
        JsonElement r = result.RootElement;
        Assert.True(r.GetProperty("apply").GetBoolean());
        Assert.Equal(0xC8UL, r.GetProperty("interval").GetUInt64());
        Assert.Equal(0x18UL, r.GetProperty("rtsoff").GetUInt64());
        Assert.Equal(Path.Combine(folder, @"Loader\kdu.exe"), r.GetProperty("kdu").GetString());
        Assert.Equal(@"C:\Windows\DTIMOD.sys", r.GetProperty("driver").GetString());
        Assert.Equal(64, r.GetProperty("runtimeHashLength").GetInt32());
        Assert.True(r.GetProperty("intelEnabled").GetBoolean());
        Assert.True(r.GetProperty("intelAdaptive").GetBoolean());
        Assert.Equal("Mouse=0x0, Audio=0xFA0", r.GetProperty("intelRoles").GetString());
        Assert.False(r.GetProperty("intelHasIntervals").GetBoolean());
        Assert.Equal([0UL, 0x64UL, 0xC8UL], r.GetProperty("amdIntervals").EnumerateArray().Select(v => v.GetUInt64()));
        Assert.Equal(1, r.GetProperty("nicCount").GetInt32());
        Assert.Equal("VEN_8086&DEV_125C", r.GetProperty("nicHwid").GetString());
        Assert.Equal(0x80000000UL, r.GetProperty("nicOrBits").GetUInt64());
        Assert.Equal([0UL, 0x28UL, 0x28UL, 0x28UL, 0x28UL], r.GetProperty("nicValues").EnumerateArray().Select(v => v.GetUInt64()));
    }

    [Fact]
    public async Task NonUsbDeviceTriggerExitsBeforeReadingTheProfile()
    {
        string folder = NewFolder("startup-trigger");
        string script = WriteScript(folder);

        (int code, string output) = await RunPowerShell(script, "-TriggerDevice", @"SWD\Wintun\{880766C8-A2B0-A5E2-A0BD-771FA9005A37}");

        Assert.True(code == 0, output);
        Assert.DoesNotContain("profile.json", ReadLog(folder));
    }

    [Fact]
    public async Task MissingProfileStopsWithALoggedReason()
    {
        string folder = NewFolder("startup-missing");
        string script = WriteScript(folder);

        (int code, string output) = await RunPowerShell(script, "-TriggerDevice", "$(Device)");

        Assert.True(code == 1, output);
        Assert.Contains("error: cannot read profile.json", ReadLog(folder));
    }

    [Fact]
    public async Task UsbTriggerStopsBeforeLoaderWhenTheDriverFileIsMissing()
    {
        string folder = NewFolder("startup-driver");
        string script = WriteScript(folder);
        string missingDriver = Path.Combine(folder, "missing", "DTIMOD.sys");
        File.WriteAllText(Path.Combine(folder, "profile.json"),
            MainForm.BuildStartupProfile(StartupProfileTests.Sample(), missingDriver).ToJson());

        (int code, string output) = await RunPowerShell(script, "-TriggerDevice", @"USB\VID_046D&PID_C547\5&1234");

        Assert.True(code == 1, output);
        string log = ReadLog(folder);
        Assert.True(log.Contains("DTIMOD.sys not found", StringComparison.Ordinal)
            || log.Contains("administrator privileges required", StringComparison.Ordinal), log);
        Assert.DoesNotContain("startup loader", log);
    }

    [Fact]
    public async Task RuntimeCompilesInWindowsPowerShell51()
    {
        string folder = NewFolder("startup-runtime");
        File.WriteAllText(Path.Combine(folder, "DeviceTweakerImodRuntime.cs"), ReadResource("DeviceTweakerCS.Startup.DeviceTweakerImodRuntime.cs"));
        string probe = Path.Combine(folder, "compile.ps1");
        File.WriteAllText(probe, """
            $ErrorActionPreference = 'Stop'
            Add-Type -Language CSharp -TypeDefinition ([System.IO.File]::ReadAllText((Join-Path $PSScriptRoot 'DeviceTweakerImodRuntime.cs')))
            if (-not ('DeviceTweakerImodRuntime' -as [type])) { throw 'runtime type missing' }
            Write-Output 'RUNTIME COMPILED'
            """, new UTF8Encoding(true));

        (int code, string output) = await RunPowerShell(probe);

        Assert.True(code == 0, output);
        Assert.Contains("RUNTIME COMPILED", output);
    }
}
