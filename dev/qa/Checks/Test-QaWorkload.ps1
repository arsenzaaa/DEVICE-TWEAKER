param([Parameter(Mandatory=$true)][string]$OutputDirectory)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\Support\Private-Desktop.ps1')
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$checks = [Collections.Generic.List[object]]::new()
function Check([string]$Name, [bool]$Passed) {
    $checks.Add(@{name=$Name;passed=$Passed})
    if (-not $Passed) { throw "QA workload contract failed: $Name" }
}
try {
    Check 'memory-boundary-accepted' ([DeviceTweakerQaDesktop]::HasMemoryHeadroom(2GB,2GB,2GB))
    Check 'physical-pressure-refused' (-not [DeviceTweakerQaDesktop]::HasMemoryHeadroom(1GB,3GB,2GB))
    Check 'commit-pressure-refused' (-not [DeviceTweakerQaDesktop]::HasMemoryHeadroom(3GB,1GB,2GB))
    $lease = [DeviceTweakerQaDesktop]::AcquireWorkload()
    if ($null -eq $lease) { throw 'Run this short contract from the normal desktop; it never creates visible windows.' }
    try {
        $probe = Join-Path $OutputDirectory 'concurrency-probe.ps1'
        $source = @'
param([string]$Helper)
$ErrorActionPreference = 'Stop'
. $Helper
try { $lease = [DeviceTweakerQaDesktop]::AcquireWorkload(); $lease.Dispose(); exit 2 }
catch { if ($_.Exception.ToString().Contains('Another QA workload')) { exit 0 }; throw }
'@
        [IO.File]::WriteAllText($probe,$source)
        $start = [Diagnostics.ProcessStartInfo]::new()
        $start.FileName = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
        $start.Arguments = '-NoProfile -NonInteractive -File "'+$probe+'" -Helper "'+(Join-Path $PSScriptRoot '..\Support\Private-Desktop.ps1')+'"'
        $start.UseShellExecute = $false; $start.CreateNoWindow = $true
        $child = [Diagnostics.Process]::Start($start)
        try {
            if (-not $child.WaitForExit(15000)) { $child.Kill(); throw 'Concurrency contract timed out' }
            Check 'independent-launch-refused' ($child.ExitCode -eq 0)
        } finally { $child.Dispose() }
        foreach ($entry in @('Invoke-Qa.ps1','Invoke-DevSession.ps1','Export-Gallery.ps1','Review-DevSession.ps1','Open-TestPanel.ps1')) {
            # Also use a forbidden output path: a broken lock must still never
            # start a real build or full QA as part of this negative contract.
            $blockedOutput = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot ('..\..\..\src\qa-guard-'+$entry)))
            $refusalLog = Join-Path $OutputDirectory ($entry+'-refusal.log')
            $start = [Diagnostics.ProcessStartInfo]::new()
            $start.FileName = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
            $entryArguments = if ($entry -eq 'Review-DevSession.ps1') { '-RunDirectory "'+$OutputDirectory+'"' }
                elseif ($entry -eq 'Open-TestPanel.ps1') { '-BuildOnly -PreparationDirectory "'+$blockedOutput+'"' }
                elseif ($entry -eq 'Export-Gallery.ps1') { '-OutputDirectory "'+$blockedOutput+'"' }
                else { '-OutputRoot "'+$blockedOutput+'"' }
            $start.Arguments = '-NoProfile -NonInteractive -File "'+(Join-Path (Split-Path $PSScriptRoot) $entry)+'" '+$entryArguments
            $start.UseShellExecute = $false; $start.CreateNoWindow = $true
            $start.RedirectStandardOutput = $true; $start.RedirectStandardError = $true
            $child = [Diagnostics.Process]::Start($start)
            try {
                $stdout = $child.StandardOutput.ReadToEndAsync(); $stderr = $child.StandardError.ReadToEndAsync()
                if (-not $child.WaitForExit(15000)) { $child.Kill(); $child.WaitForExit(); throw 'Entrypoint refusal timed out' }
                $text = $stdout.Result+$stderr.Result
                [IO.File]::WriteAllText($refusalLog,$text)
                Check ($entry+'-refuses-before-preparation') ($child.ExitCode -ne 0 -and $text.Contains('Another QA workload') -and -not (Test-Path -LiteralPath $blockedOutput))
            } finally { $child.Dispose() }
        }
        $nested = [DeviceTweakerQaDesktop]::AcquireWorkload()
        $nested.Dispose()
        Check 'same-launcher-reentrant' $true
    } finally { $lease.Dispose() }
    $lease = [DeviceTweakerQaDesktop]::AcquireWorkload()
    $lease.Dispose()
    Check 'lease-released' $true

    $nestedScript = Join-Path $OutputDirectory 'nested-lease.ps1'
    [IO.File]::WriteAllText($nestedScript,@'
param([string]$Helper)
$ErrorActionPreference = 'Stop'
. $Helper
$lease = [DeviceTweakerQaDesktop]::AcquireWorkload()
if ($null -ne $lease) { $lease.Dispose(); exit 1 }
exit 0
'@)
    $nestedCode = Invoke-QaPrivateDesktop -ScriptPath $nestedScript -Parameters @{Helper=(Join-Path $PSScriptRoot '..\Support\Private-Desktop.ps1')} -ReportDirectory (Join-Path $OutputDirectory 'nested') -TimeoutSeconds 15
    Check 'bounded-private-worker-reuses-lease' ($nestedCode -eq 0)

    $fixture = Join-Path $OutputDirectory 'pressure-worker.ps1'
    $pids = Join-Path $OutputDirectory 'pressure-pids.json'
    $source = @'
param([string]$Pids)
$start = [Diagnostics.ProcessStartInfo]::new()
$start.FileName = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
$start.Arguments = '-NoProfile -NonInteractive -Command "Start-Sleep -Seconds 60"'
$start.UseShellExecute = $false; $start.CreateNoWindow = $true
$child = [Diagnostics.Process]::Start($start)
@($PID,$child.Id) | ConvertTo-Json | Set-Content -LiteralPath $Pids
Start-Sleep -Seconds 60
'@
    [IO.File]::WriteAllText($fixture,$source)
    $hostExe = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $command = '"'+$hostExe+'" -NoProfile -NonInteractive -File "'+$fixture+'" -Pids "'+$pids+'"'
    $resources = Join-Path $OutputDirectory 'pressure-resources.json'
    $stopped = $false
    try {
        $null = [DeviceTweakerQaDesktop]::RunPressureProbe(('DeviceTweakerQa_'+[Guid]::NewGuid().ToString('N')), $hostExe, $command, $OutputDirectory, 15, $resources)
    } catch { $stopped = $_.Exception.ToString().Contains('preserve host memory headroom') }
    Check 'pressure-stops-owned-job' $stopped
    [int[]]$ids = Get-Content -LiteralPath $pids -Raw | ConvertFrom-Json
    $remaining = @($ids | ForEach-Object { Get-Process -Id $_ -ErrorAction SilentlyContinue })
    Check 'pressure-cleans-owned-descendants' ($ids.Count -eq 2 -and $remaining.Count -eq 0)
    $report = Get-Content -LiteralPath $resources -Raw | ConvertFrom-Json
    Check 'pressure-evidence-recorded' ($report.status -eq 'FAIL' -and $report.injectedPressure -and $report.cleanupConfirmed -and $report.samples -ge 8 -and $report.peakJobCommitBytes -gt 0 -and $report.jobMemoryLimitBytes -le 3GB)
} finally {
    [ordered]@{status=$(if ($checks.Count -eq 12 -and @($checks | Where-Object {-not $_.passed}).Count -eq 0) {'PASS'} else {'FAIL'});checks=$checks.ToArray();scope='Small child processes and injected low-memory observation. Does not allocate to exhaustion, start the product GUI, or modify host settings.'} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'workload.json') -Encoding UTF8
}
Write-Output 'PASS: QA serialization, pressure policy, owned-tree cleanup and telemetry.'
