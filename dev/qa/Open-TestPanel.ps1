[CmdletBinding()]
param([switch]$BuildOnly)
$ErrorActionPreference = 'Stop'
$workspace = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..')).Path
. (Join-Path $PSScriptRoot 'Input-Snapshot.ps1')
$output = Join-Path $workspace 'local\tools\TestPanel'
$build = Join-Path $workspace ('local\work\TestPanel\build_' + [Guid]::NewGuid().ToString('N'))
$snapshot = New-QaInputSnapshot $workspace $build
$project = Join-Path $snapshot 'DEVICE TWEAKER\DeviceTweakerCS.csproj'
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
    $probe = Start-Process -FilePath $exe -ArgumentList '--qa-probe' -WindowStyle Hidden -PassThru
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
    if (-not $BuildOnly) {
        # Interactive launch requested by the person running this script.
        $env:DEVICE_TWEAKER_QA_SESSION = [Guid]::NewGuid().ToString('N')
        $env:DEVICE_TWEAKER_QA_OUTPUT = Join-Path $output ('session_' + $env:DEVICE_TWEAKER_QA_SESSION)
        $env:DEVICE_TWEAKER_QA_TEST_ADMIN = '1'
        Start-Process -FilePath $exe -ArgumentList '--qa-read-only' | Out-Null
    }
} finally {
    foreach ($name in $variables) { [Environment]::SetEnvironmentVariable($name, $saved[$name]) }
}
