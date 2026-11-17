[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$OutputDirectory)
$ErrorActionPreference = 'Stop'
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $OutputDirectory) { throw 'Use a new preparation contract directory' }
$fixture = Join-Path $OutputDirectory 'fixture\dev\qa'
New-Item -ItemType Directory -Path $fixture | Out-Null
# Exercise the real entry script with a tiny build-boundary fixture. No compiler
# or application is run, and the intentional failure stops before cache use.
foreach ($name in @('Invoke-DevSession.ps1','Support\Private-Desktop.ps1','Support\PrivateDesktop.cs','Support\Input-Snapshot.ps1','Support\Read-QaJson.ps1')) {
    $target = Join-Path $fixture $name
    New-Item -ItemType Directory -Force -Path (Split-Path $target) | Out-Null
    Copy-Item -LiteralPath (Join-Path (Split-Path $PSScriptRoot) $name) -Destination $target
}
[IO.File]::WriteAllText((Join-Path $fixture 'scenario.json'),'{}')
$mock = @'
function Get-QaDevBuild([string]$Workspace, [string]$BuildRoot, [switch]$FreshBuild) {
    $lease = [DeviceTweakerQaDesktop]::AcquireWorkload()
    $bounded = $null -eq $lease
    if ($lease) { $lease.Dispose() }
    @{bounded=$bounded;pid=$PID} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $BuildRoot 'preparation.json')
    if (-not $bounded) { throw 'Preparation is outside the bounded private job' }
    $child = Start-QaOwnedProcess (Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe') `
        ('-NoProfile -NonInteractive -File "'+(Join-Path $PSScriptRoot '..\child.ps1')+'" -Helper "'+(Join-Path $PSScriptRoot 'Private-Desktop.ps1')+'" -Report "'+(Join-Path $BuildRoot 'child.json')+'"')
    try {
        if (-not $child.WaitForExit(15000)) { $child.Kill(); $child.WaitForExit(); throw 'Preparation child timed out' }
        if ($child.ExitCode -ne 0) { throw 'Preparation child failed' }
    } finally { $child.Dispose() }
    throw 'Preparation contract intentional stop'
}
'@
[IO.File]::WriteAllText((Join-Path $fixture 'Support\Build-Cache.ps1'),$mock)
[IO.File]::WriteAllText((Join-Path $fixture 'child.ps1'),@'
param([string]$Helper,[string]$Report)
$ErrorActionPreference = 'Stop'
. $Helper
$lease = [DeviceTweakerQaDesktop]::AcquireWorkload()
$bounded = $null -eq $lease
if ($lease) { $lease.Dispose() }
@{bounded=$bounded;pid=$PID} | ConvertTo-Json | Set-Content -LiteralPath $Report
'@)
$start = New-Object Diagnostics.ProcessStartInfo
$start.FileName = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
$start.Arguments = '-NoProfile -NonInteractive -File "'+(Join-Path $fixture 'Invoke-DevSession.ps1')+'" -Scenario "'+(Join-Path $fixture 'scenario.json')+'" -OutputRoot "'+(Join-Path $OutputDirectory 'runs')+'" -BuildRoot "'+$OutputDirectory+'"'
$start.UseShellExecute = $false; $start.CreateNoWindow = $true
$start.RedirectStandardOutput = $true; $start.RedirectStandardError = $true
$child = [Diagnostics.Process]::Start($start)
try {
    $stdout = $child.StandardOutput.ReadToEndAsync(); $stderr = $child.StandardError.ReadToEndAsync()
    if (-not $child.WaitForExit(30000)) { $child.Kill(); $child.WaitForExit(); throw 'Preparation contract timed out' }
    $log = $stdout.Result + $stderr.Result
    [IO.File]::WriteAllText((Join-Path $OutputDirectory 'entry.log'),$log)
    $preparation = Get-Content -LiteralPath (Join-Path $OutputDirectory 'preparation.json') -Raw | ConvertFrom-Json
    if (-not $preparation.bounded) { throw 'Preparation must already be inside the bounded private job' }
    $descendant = Get-Content -LiteralPath (Join-Path $OutputDirectory 'child.json') -Raw | ConvertFrom-Json
    if (-not $preparation.bounded -or -not $descendant.bounded) { throw 'Preparation and its child must already be inside the bounded private job' }
    if ($child.ExitCode -eq 0 -or -not $log.Contains('Preparation contract intentional stop')) { throw 'Preparation failure was hidden' }
    $resources = @(Get-ChildItem -LiteralPath (Join-Path $OutputDirectory 'runs') -Recurse -File -Filter resources.json)
    if ($resources.Count -ne 1) { throw 'Expected one preparation resource report' }
    $resource = Get-Content -LiteralPath $resources[0].FullName -Raw | ConvertFrom-Json
    if (-not $resource.cleanupConfirmed -or $resource.jobMemoryLimitBytes -gt 3GB -or $resource.peakJobCommitBytes -le 0) { throw 'Missing bounded cleanup evidence' }
    [ordered]@{status='PASS';preparation=$preparation;descendant=$descendant;failurePropagated=$true;resources=$resource;
        scope='Real entry point, substituted build boundary and small child; no compilation or product GUI'} |
        ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'preparation-contract.json') -Encoding UTF8
} finally { $child.Dispose() }
Write-Output 'PASS: cache preparation and descendants are bounded before build; failures propagate.'
