param([Parameter(Mandatory = $true)][string]$OutputDirectory,
    [Parameter(Mandatory = $true)][string]$ExePath)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Private-Desktop.ps1')
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$results = [Collections.Generic.List[object]]::new()
$savedDesktop = $env:DEVICE_TWEAKER_QA_DESKTOP
$savedOutput = $env:DEVICE_TWEAKER_QA_OUTPUT
try {
    # All probes are read-only and create no product windows.
    $env:DEVICE_TWEAKER_QA_DESKTOP = 'DeviceTweakerQa_wrong_assignment'
    $env:DEVICE_TWEAKER_QA_OUTPUT = Join-Path $OutputDirectory 'refused-process'
    $null = Start-QaOwnedProcess $ExePath '--qa-read-only --qa-probe'
} catch {
    $results.Add(@{ case = 'worker-refuses-wrong-desktop'; expected = 'Refusal before child launch';
        observed = $_.Exception.Message; passed = $_.Exception.Message -like '*not on its assigned*' })
} finally { $env:DEVICE_TWEAKER_QA_DESKTOP = $savedDesktop }
try {
    # Bypass only the PS assertion to verify the independent EXE startup check.
    $env:DEVICE_TWEAKER_QA_DESKTOP = 'DeviceTweakerQa_wrong_assignment'
    $env:DEVICE_TWEAKER_QA_OUTPUT = Join-Path $OutputDirectory 'refused-process'
    $start = [Diagnostics.ProcessStartInfo]::new()
    $start.FileName = $ExePath; $start.Arguments = '--qa-read-only --qa-probe'; $start.UseShellExecute = $false
    $start.CreateNoWindow = $true
    $probe = [Diagnostics.Process]::Start($start)
    if (-not $probe.WaitForExit(15000)) { $probe.Kill(); throw 'Refusal probe timed out' }
    $ok = $probe.ExitCode -eq 3 -and -not (Test-Path -LiteralPath $env:DEVICE_TWEAKER_QA_OUTPUT)
    $results.Add(@{ case = 'exe-refuses-wrong-desktop'; expected = 'Exit 3 before reports/windows';
        observed = "exit=$($probe.ExitCode)"; passed = $ok })
    $env:DEVICE_TWEAKER_QA_DESKTOP = $null
    $exitScript = Join-Path $OutputDirectory 'exit-seven.ps1'
    [IO.File]::WriteAllText($exitScript, 'exit 7')
    $code = Invoke-QaPrivateDesktop $exitScript @{} (Join-Path $OutputDirectory 'exit-seven') 15
    $results.Add(@{ case = 'worker-exit-code'; expected = 7; observed = $code; passed = $code -eq 7 })
    $timeoutScript = Join-Path $OutputDirectory 'timeout.ps1'
    $helper = (Join-Path $PSScriptRoot 'Private-Desktop.ps1').Replace("'", "''")
    $pidFile = (Join-Path $OutputDirectory 'timeout-pids.json').Replace("'", "''")
    $body = @'
. '__HELPER__'
$hostExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$child = Start-QaOwnedProcess $hostExe '-NoProfile -NonInteractive -Command "Start-Sleep -Seconds 120"'
@($PID, $child.Id) | ConvertTo-Json | Set-Content -LiteralPath '__PIDFILE__'
Start-Sleep -Seconds 120
'@.Replace('__HELPER__', $helper).Replace('__PIDFILE__', $pidFile)
    [IO.File]::WriteAllText($timeoutScript, $body)
    $timedOut = $false
    try { Invoke-QaPrivateDesktop $timeoutScript @{} (Join-Path $OutputDirectory 'timeout') 4 | Out-Null }
    catch { $timedOut = $_.Exception.ToString() -like '*timed out*' }
    [int[]]$ids = Get-Content -LiteralPath (Join-Path $OutputDirectory 'timeout-pids.json') -Raw | ConvertFrom-Json
    $remaining = @($ids | ForEach-Object { Get-Process -Id $_ -ErrorAction SilentlyContinue })
    $results.Add(@{ case = 'timeout-cleans-owned-descendants'; expected = 'Timeout; worker and child stopped';
        observed = "timeout=$timedOut remaining=$($remaining.Count)"; ownedPids = $ids;
        passed = $timedOut -and $ids.Count -eq 2 -and $remaining.Count -eq 0 })
} finally {
    $env:DEVICE_TWEAKER_QA_DESKTOP = $savedDesktop
    $env:DEVICE_TWEAKER_QA_OUTPUT = $savedOutput
    $results.ToArray() | ConvertTo-Json -Depth 6 |
        Set-Content -LiteralPath (Join-Path $OutputDirectory 'contracts.json') -Encoding UTF8
}
if ($results.Count -ne 4 -or @($results | Where-Object { -not $_.passed }).Count) { throw 'Private desktop contracts failed' }
Write-Host 'PASS: independent desktop refusal, exact exit code and timeout cleanup.'
