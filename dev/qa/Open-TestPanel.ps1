[CmdletBinding()]
param([switch]$BuildOnly, [switch]$Worker, [string]$PreparationDirectory)
$ErrorActionPreference = 'Stop'
$workspace = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..')).Path
. (Join-Path $PSScriptRoot 'Support\Input-Snapshot.ps1')
. (Join-Path $PSScriptRoot 'Support\Private-Desktop.ps1')
$output = Join-Path $workspace 'local\tools\TestPanel'
$lease = if (-not $Worker) { [DeviceTweakerQaDesktop]::AcquireWorkload() } else { $null }
try {
if (-not $Worker) {
    if ($PreparationDirectory) { throw 'PreparationDirectory belongs to the bounded worker.' }
    $build = Join-Path $workspace ('local\work\TestPanel\build_' + [Guid]::NewGuid().ToString('N'))
    $code = Invoke-QaPrivateDesktop $PSCommandPath @{ Worker = $true; BuildOnly = $true; PreparationDirectory = $build } (Join-Path $build 'isolation')
    if ($null -eq $code) { & $PSCommandPath -Worker -BuildOnly -PreparationDirectory $build }
    elseif ($code -ne 0) { exit $code }
    $receipt = Get-Content -LiteralPath (Join-Path $build 'test-panel.json') -Raw | ConvertFrom-Json
    if ($receipt.status -ne 'PASS') { throw 'Developer panel preparation was not verified.' }
    Write-Output ('Verified developer panel: ' + $receipt.exe)
    if (-not $BuildOnly) {
        # This entry is explicitly interactive; build/probe windows stay on the private desktop.
        $saved = @{}
        foreach ($name in @('DEVICE_TWEAKER_QA_SESSION','DEVICE_TWEAKER_QA_OUTPUT','DEVICE_TWEAKER_QA_TEST_ADMIN')) { $saved[$name] = [Environment]::GetEnvironmentVariable($name) }
        try {
            $env:DEVICE_TWEAKER_QA_SESSION = [Guid]::NewGuid().ToString('N')
            $env:DEVICE_TWEAKER_QA_OUTPUT = Join-Path $output ('session_' + $env:DEVICE_TWEAKER_QA_SESSION)
            $env:DEVICE_TWEAKER_QA_TEST_ADMIN = '1'
            Start-Process -FilePath $receipt.exe -ArgumentList '--qa-read-only' | Out-Null
        } finally { foreach ($name in $saved.Keys) { [Environment]::SetEnvironmentVariable($name,$saved[$name]) } }
    }
    return
}
Assert-QaBoundedWorkload
if (-not $PreparationDirectory -or -not $BuildOnly) { throw 'Worker only prepares the test panel.' }
$build = [IO.Path]::GetFullPath($PreparationDirectory)
$snapshot = New-QaInputSnapshot $workspace $build
$project = Join-Path $snapshot 'src\DeviceTweakerCS.csproj'
& dotnet publish $project -c Release -r win-x64 -o $output `
    -p:DeveloperQa=true -p:BuildImodDriver=false -p:SelfContained=true `
    -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true `
    -p:DebugType=None -p:DebugSymbols=false -p:TreatWarningsAsErrors=true
if ($LASTEXITCODE -ne 0) { throw 'Developer test panel build failed' }
$exe = Join-Path $output 'DEVICE TWEAKER QA.exe'
$session = [Guid]::NewGuid().ToString('N')
$probeOutput = Join-Path $output ('probe_' + $session)
$variables = @('DEVICE_TWEAKER_QA_SESSION', 'DEVICE_TWEAKER_QA_OUTPUT', 'DEVICE_TWEAKER_QA_TEST_ADMIN')
$saved = @{}
foreach ($name in $variables) { $saved[$name] = [Environment]::GetEnvironmentVariable($name) }
try {
    $env:DEVICE_TWEAKER_QA_SESSION = $session
    $env:DEVICE_TWEAKER_QA_OUTPUT = $probeOutput
    $probe = Start-QaOwnedProcess $exe '--qa-probe'
    try {
        if (-not $probe.WaitForExit(20000)) { throw 'Developer panel probe timed out' }
        if ($probe.ExitCode -ne 0) { throw 'Developer panel probe failed' }
        $report = Get-Content -LiteralPath (Join-Path $probeOutput ('process-' + $probe.Id + '.json')) -Raw | ConvertFrom-Json
        if ($report.readOnly -ne $true -or $report.build.hasTestPanel -ne $true) {
            throw 'Expected a permanently read-only developer build with its test panel'
        }
    } finally {
        if (-not $probe.HasExited) { $probe.Kill(); $probe.WaitForExit(5000) | Out-Null }
        $probe.Dispose()
    }
    Write-Output ('Verified developer panel: ' + $exe)
    @{ status = 'PASS'; exe = $exe; readOnly = $true; probe = $probeOutput } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $build 'test-panel.json') -Encoding UTF8
} finally {
    foreach ($name in $variables) { [Environment]::SetEnvironmentVariable($name, $saved[$name]) }
}
} finally { if ($lease) { $lease.Dispose() } }
