param(
    [Parameter(Mandatory=$true)][string]$ExePath,
    [Parameter(Mandatory=$true)][string]$OutputDirectory,
    [ValidateRange(10,300)][int]$Seconds = 30,
    [switch]$Worker
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Private-Desktop.ps1')
$exe = [IO.Path]::GetFullPath($ExePath)
$output = [IO.Path]::GetFullPath($OutputDirectory)
if (-not $Worker) {
    if (Test-Path -LiteralPath $output) { throw 'Preserve existing measurements; choose a new output directory.' }
    $code = Invoke-QaPrivateDesktop -ScriptPath $PSCommandPath -Parameters @{Worker=$true;ExePath=$exe;OutputDirectory=$output;Seconds=$Seconds} -ReportDirectory (Join-Path $output 'isolation') -TimeoutSeconds ($Seconds+150)
    exit $code
}
Assert-QaBoundedWorkload
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class IdleCycleCounter {
    [DllImport("kernel32.dll", SetLastError=true)]
    public static extern bool QueryProcessCycleTime(IntPtr process, out ulong cycles);
}
'@
foreach ($name in @('DEVICE_TWEAKER_SHOWCASE','DEVICE_TWEAKER_SHOWCASE_FILTER','DEVICE_TWEAKER_QA_TEST_ADMIN','DEVICE_TWEAKER_QA_SANDBOX','DEVICE_TWEAKER_DEV_SCENARIO','DEVICE_TWEAKER_QA_LAYOUT_DPI')) {
    [Environment]::SetEnvironmentVariable($name,$null)
}
$env:DEVICE_TWEAKER_QA_READ_ONLY = '1'
$env:DEVICE_TWEAKER_QA_OUTPUT = $output
$env:DEVICE_TWEAKER_QA_SESSION = [Guid]::NewGuid().ToString('N')
$env:DEVICE_TWEAKER_LANGUAGE = 'EN'
$digest = (Get-FileHash -LiteralPath $exe).Hash
$owned = $null
$samples = New-Object 'Collections.Generic.List[object]'
try {
    $owned = Start-QaOwnedProcess $exe '--qa-read-only'
    $startup = [Diagnostics.Stopwatch]::StartNew()
    $log = $null
    while ($startup.Elapsed.TotalSeconds -lt 90) {
        if ($owned.HasExited) { throw 'Application exited during startup.' }
        $log = Get-ChildItem -LiteralPath (Join-Path $output 'logs') -Filter 'DeviceTweaker_*.log' -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($log) {
            $stream = [IO.File]::Open($log.FullName,[IO.FileMode]::Open,[IO.FileAccess]::Read,([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
            $reader = [IO.StreamReader]::new($stream,[Text.Encoding]::UTF8,$true)
            try { $text = $reader.ReadToEnd() } finally { $reader.Dispose() }
            if ($text.Contains('REFRESH.DONE:')) { break }
        }
        Start-Sleep -Milliseconds 50
    }
    if (-not $log -or -not $text.Contains('REFRESH.DONE:')) { throw 'Initial refresh did not complete.' }
    Start-Sleep -Seconds 3
    $clock = [Diagnostics.Stopwatch]::StartNew()
    for ($i=0; $i -le $Seconds; $i++) {
        $owned.Refresh()
        if ($owned.HasExited) { throw 'Application exited during observation.' }
        [UInt64]$cycles = 0
        if (-not [IdleCycleCounter]::QueryProcessCycleTime($owned.Handle,[ref]$cycles)) { throw 'QueryProcessCycleTime failed.' }
        $samples.Add([pscustomobject]@{elapsedMs=$clock.Elapsed.TotalMilliseconds;cpuMs=$owned.TotalProcessorTime.TotalMilliseconds;cycles=$cycles;privateBytes=$owned.PrivateMemorySize64;workingSetBytes=$owned.WorkingSet64;handles=$owned.HandleCount;threads=$owned.Threads.Count})
        if ($i -lt $Seconds) { Start-Sleep -Seconds 1 }
    }
    if (-not (Stop-QaOwnedProcess $owned)) { throw 'Application did not close normally.' }
    $report = Get-Content -LiteralPath (Join-Path $output ('process-'+$owned.Id+'.json')) -Raw | ConvertFrom-Json
    if (-not $report.readOnly -or -not $report.desktop.isolated -or $report.build.hasTestPanel) { throw 'Expected an isolated read-only production executable.' }
    if ((Get-FileHash -LiteralPath $exe).Hash -ne $digest) { throw 'Executable changed during measurement.' }
    $first=$samples[0]; $last=$samples[$samples.Count-1]
    $wallMs=$last.elapsedMs-$first.elapsedMs
    [ordered]@{
        status='PASS';executableSha256=$digest;appBuildId=$report.build.appBuildId
        logicalProcessors=[Environment]::ProcessorCount;elapsedMs=$wallMs
        cpuTimeMs=$last.cpuMs-$first.cpuMs
        cpuPercentOneCore=100*($last.cpuMs-$first.cpuMs)/$wallMs
        cpuPercentMachine=100*($last.cpuMs-$first.cpuMs)/$wallMs/[Environment]::ProcessorCount
        cyclesDelta=$last.cycles-$first.cycles
        privateBytesDelta=$last.privateBytes-$first.privateBytes
        workingSetBytesDelta=$last.workingSetBytes-$first.workingSetBytes
        samples=$samples.ToArray();blockedOperations=$report.blockedOperations
        scope='Exact production EXE after initial refresh and 3 s settling, invisible read-only desktop. No physical input or populated raw-input buffers are assumed. CPU percentages use observed process CPU time; cycles are not converted to time. A short observation cannot rule out a long-term leak.'
    } | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $output 'resources.json') -Encoding UTF8
    Write-Output ('PASS: '+$Seconds+' s production idle observation; '+$digest)
} finally {
    if ($null -ne $owned -and -not $owned.HasExited) { $null=Stop-QaOwnedProcess $owned }
}
