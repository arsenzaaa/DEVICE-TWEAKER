using System.Diagnostics;
using System.Reflection;
using System.Text;
using System.Text.Json;
using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public sealed class StartupMigrationCompilationTests
{
    [Theory] [InlineData(false)] [InlineData(true)]
    public void UpgradedBackupScriptIsVerifiedAndRollbackKeepsOriginalEngine(bool fail)
    {
        const string original = "private const string ImodDriverServiceName = \"DeviceTweakerImod2\";\n$imodSettingsBegin = $true\n$ImodDriverSha256 = 'old'\n$globalInterval = 0xC8\n$imodSettingsEnd = $true";
        string upgraded = ImodStartupUpgrade.Upgrade(original, Template(), new string('B', 64));
        string current = original;
        DeviceTweakerBackup expected = new() { ImodScript = new() { Path = "managed", Exists = true, Text = upgraded } };
        BackupRestoreService service = new(_ => { }, _ => throw new InvalidOperationException("No registry targets"),
            _ => throw new InvalidOperationException("No registry writes"),
            () => new FileBackup { Path = "managed", Exists = true, Text = current },
            script => { current = script.Text!; if (fail && current == upgraded) throw new IOException("Injected file write error after modification"); }, _ => { });
        if (fail) Assert.Throws<IOException>(() => service.Restore(expected));
        else service.Restore(expected);
        Assert.Equal(fail ? original : upgraded, current);
    }
    private static string Template() => (string)typeof(MainForm).GetMethod("CreateImodScriptTemplate", BindingFlags.Static | BindingFlags.NonPublic)!.Invoke(null, null)!;

    [Fact]
    public async Task GeneratedStartupParsesAndCompilesOnWindowsPowerShellWithoutExecutingIt()
    {
        string output = Path.Combine(QaRun.OutputDirectory, "startup-compile-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(output);
        string template = Template().Replace("{{IMOD_CONFIG_BLOCK}}", "$imodSettingsBegin = $true\n$imodSettingsEnd = $true");
        Assert.DoesNotContain("{{DRIVER_RENAME_SOURCE}}", template);
        File.WriteAllText(Path.Combine(output, "startup.ps1"), template, new UTF8Encoding(true));
        const string validator = """
            $ErrorActionPreference = 'Stop'
            $tokens = $null; $errors = $null
            $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot 'startup.ps1'), [ref]$tokens, [ref]$errors)
            if ($errors.Count) { throw ($errors | Out-String) }
            $source = $ast.FindAll({ param($a) $a -is [System.Management.Automation.Language.StringConstantExpressionAst] -and $a.Value.Contains('public static class DeviceTweakerImodRuntime') }, $true)
            if ($source.Count -ne 1) { throw 'Expected exactly one generated runtime source' }
            Add-Type -Language CSharp -TypeDefinition $source[0].Value
            if (-not ('DeviceTweakerImodRuntime' -as [type])) { throw 'Published startup runtime missing' }
            Write-Output 'STARTUP COMPILE PASS; published version 30 compiled; no driver, SCM or startup body executed'
            """;
        File.WriteAllText(Path.Combine(output, "validate.ps1"), validator, new UTF8Encoding(true));
        ProcessStartInfo info = new("powershell.exe") { UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true };
        foreach (string arg in new[] { "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", Path.Combine(output, "validate.ps1") }) info.ArgumentList.Add(arg);
        using Process process = Process.Start(info)!;
        Task<string> stdout = process.StandardOutput.ReadToEndAsync(), stderr = process.StandardError.ReadToEndAsync();
        if (!process.WaitForExit(45000)) { process.Kill(entireProcessTree: true); throw new TimeoutException("Startup compilation timed out"); }
        string log = await stdout + await stderr;
        File.WriteAllText(Path.Combine(output, "compile.log"), log);
        Assert.True(process.ExitCode == 0, log); Assert.Contains("STARTUP COMPILE PASS", log);
    }

    [Theory] [MemberData(nameof(BackupRestoreTests.ExistingBackupFiles), MemberType = typeof(BackupRestoreTests))]
    public void ExistingBackupMigrationRetainsConfigurationAndLeavesOriginalUntouched(string path)
    {
        if (path == "synthetic")
        {
            const string old = "private const string ImodDriverServiceName = \"DeviceTweakerImod2\";\n$imodSettingsBegin = $true\n$ImodDriverSha256 = 'old'\n$globalInterval = 0xC8\n$imodSettingsEnd = $true";
            Assert.Contains("$globalInterval = 0xC8", ImodStartupUpgrade.Upgrade(old, Template(), new string('B', 64)));
            return;
        }
        byte[] original = File.ReadAllBytes(path);
        DeviceTweakerBackup backup = JsonSerializer.Deserialize<DeviceTweakerBackup>(File.ReadAllText(path))!;
        if (backup.ImodScript is not { Exists: true })
        {
            // A baseline saved before IMOD was enabled intentionally contains no startup script.
            // The restore path preserves this absence and does not call the script upgrader.
            Assert.Null(backup.ImodScript?.Text);
            Assert.Equal(original, File.ReadAllBytes(path));
            return;
        }
        Assert.NotNull(backup.ImodScript.Text);
        string text = backup.ImodScript!.Text!;
        string migrated = ImodStartupUpgrade.Upgrade(text, Template(), new string('B', 64));
        Assert.Contains("$imodScriptVersion = 30", migrated);
        string begin = "$imodSettingsBegin = $true", end = "$imodSettingsEnd = $true";
        static string Config(string s, string begin, string end) => s[s.IndexOf(begin, StringComparison.Ordinal)..(s.IndexOf(end, StringComparison.Ordinal) + end.Length)];
        string Clean(string s) => string.Join('\n', s.Replace("\r\n", "\n").Split('\n').Where(line => !line.TrimStart().StartsWith("$ImodDriverSha256", StringComparison.Ordinal)));
        Assert.Equal(Clean(Config(text, begin, end)), Clean(Config(migrated, begin, end)));
        Assert.Equal(original, File.ReadAllBytes(path));
    }
}
