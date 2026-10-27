# Fast, private-desktop development loop. Never launches the production EXE.
[CmdletBinding()]
param(
    [string]$Scenario,
    [string]$OutputRoot,
    [string]$CompareWith,
    [switch]$Worker,
    [string]$RunDirectory,
    [string]$ExePath
)
$ErrorActionPreference = 'Stop'
if (-not $Scenario) { $Scenario = Join-Path $PSScriptRoot 'Scenarios\device-flow.json' }
. (Join-Path $PSScriptRoot 'Private-Desktop.ps1')
. (Join-Path $PSScriptRoot 'Input-Snapshot.ps1')
. (Join-Path $PSScriptRoot 'Read-QaJson.ps1')
if (-not $Worker) {
    $workspace = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
    if (-not $OutputRoot) { $OutputRoot = Join-Path $workspace 'local\dev-sessions' }
    $outputFull = [IO.Path]::GetFullPath($OutputRoot).TrimEnd('\')
    $appRoot = Join-Path $workspace 'DEVICE TWEAKER'
    if (($outputFull -eq $appRoot -or $outputFull.StartsWith($appRoot+'\',[StringComparison]::OrdinalIgnoreCase)) -and
        -not $outputFull.StartsWith($appRoot+'\tmp\',[StringComparison]::OrdinalIgnoreCase) -and
        -not $outputFull.StartsWith($appRoot+'\bin\',[StringComparison]::OrdinalIgnoreCase)) { throw 'Use local/dev-sessions or app tmp/bin for output' }
    $run = Join-Path ([IO.Path]::GetFullPath($OutputRoot)) ('dev_{0}_{1}' -f (Get-Date -Format 'yyyyMMdd_HHmmss'),[Guid]::NewGuid().ToString('N').Substring(0,8))
    New-Item -ItemType Directory -Path $run | Out-Null
    Copy-Item -LiteralPath (Resolve-Path -LiteralPath $Scenario).Path -Destination (Join-Path $run 'scenario.json')
    $snapshot = New-QaInputSnapshot $workspace $run
    & dotnet build (Join-Path $snapshot 'DEVICE TWEAKER\DeviceTweakerCS.csproj') -c Release -p:DeveloperQa=true -p:BuildImodDriver=false -p:TreatWarningsAsErrors=true --nologo *> (Join-Path $run 'build.log')
    if ($LASTEXITCODE -ne 0) { throw "Developer build failed; see $run\build.log" }
    $parameters = @{ Worker = $true; RunDirectory = $run; Scenario = (Join-Path $run 'scenario.json');
        ExePath = (Join-Path $snapshot 'DEVICE TWEAKER\bin\DeveloperQa\Release\net8.0-windows\win-x64\DEVICE TWEAKER QA.exe') }
    if ($CompareWith) { $parameters.CompareWith = [IO.Path]::GetFullPath($CompareWith) }
    $code = Invoke-QaPrivateDesktop -ScriptPath (Join-Path $snapshot 'dev\qa\Invoke-DevSession.ps1') -Parameters $parameters -ReportDirectory (Join-Path $run 'isolation') -TimeoutSeconds 240
    Test-QaInputSnapshot $run
    Write-Output $run
    exit $code
}
Assert-QaPrivateDesktop
if (-not $ExePath -or -not $RunDirectory) { throw 'Worker requires an isolated executable and run directory' }
$RunDirectory = [IO.Path]::GetFullPath($RunDirectory)
New-Item -ItemType Directory -Force -Path $RunDirectory | Out-Null
if ([IO.Path]::GetFullPath($Scenario) -ne (Join-Path $RunDirectory 'scenario.json')) {
    Copy-Item -LiteralPath $Scenario -Destination (Join-Path $RunDirectory 'scenario.json')
    $Scenario = Join-Path $RunDirectory 'scenario.json'
}
$scene = Get-Content -LiteralPath $Scenario -Raw | ConvertFrom-Json
$inspector = Join-Path $PSScriptRoot 'Inspect\Inspect.csproj'
& dotnet run --project $inspector -- --dev-validate $Scenario *> (Join-Path $RunDirectory 'scenario-validation.log')
if ($LASTEXITCODE -ne 0) { throw 'Invalid scenario; see scenario-validation.log' }
$vars = @('DEVICE_TWEAKER_QA_READ_ONLY','DEVICE_TWEAKER_QA_OUTPUT','DEVICE_TWEAKER_QA_SESSION','DEVICE_TWEAKER_DEV_SCENARIO',
    'DEVICE_TWEAKER_SHOWCASE','DEVICE_TWEAKER_SHOWCASE_FILTER','DEVICE_TWEAKER_QA_TEST_ADMIN','DEVICE_TWEAKER_QA_SANDBOX',
    'DEVICE_TWEAKER_SHOWCASE_DEVICE','DEVICE_TWEAKER_SHOWCASE_ISOLATE_TARGET','DEVICE_TWEAKER_QA_DPI','DEVICE_TWEAKER_QA_LAYOUT_DPI')
$saved = @{}; foreach ($name in $vars) { $saved[$name] = [Environment]::GetEnvironmentVariable($name) }
$process = $null
try {
    foreach ($name in $vars) { [Environment]::SetEnvironmentVariable($name,$null) }
    $env:DEVICE_TWEAKER_QA_READ_ONLY = '1'
    $env:DEVICE_TWEAKER_QA_OUTPUT = $RunDirectory
    $env:DEVICE_TWEAKER_QA_SESSION = [Guid]::NewGuid().ToString('N')
    $env:DEVICE_TWEAKER_DEV_SCENARIO = [IO.Path]::GetFullPath($Scenario)
    $env:DEVICE_TWEAKER_SHOWCASE = $scene.preset + '_' + $scene.language
    $env:DEVICE_TWEAKER_SHOWCASE_FILTER = $scene.filter
    $env:DEVICE_TWEAKER_QA_LAYOUT_DPI = if ($scene.layoutDpi) { [string]$scene.layoutDpi } else { '96' }
    $process = Start-QaOwnedProcess ([IO.Path]::GetFullPath($ExePath)) '--qa-read-only'
    $deadline = [DateTime]::UtcNow.AddSeconds(120)
    while (-not $process.WaitForExit(250)) {
        $sessionFile = Join-Path $RunDirectory 'session.json'
        if (Test-Path -LiteralPath $sessionFile) {
            $current = Read-QaJson $sessionFile
            if ($current.status -in @('PASS','FAIL')) {
                if (-not $process.WaitForExit(5000)) { throw 'Scenario finished but application did not exit normally; inspect crash reports' }
                break
            }
        }
        if ([DateTime]::UtcNow -gt $deadline) { throw 'Development scenario timed out' }
    }
    if ($process.ExitCode -ne 0) { throw 'Development scenario failed; see session.json / index.html' }
    if (@(Get-ChildItem -LiteralPath $RunDirectory -Recurse -File -Filter 'crash_*.txt').Count) { throw 'Application crash report found after scenario' }
    $logs = @(Get-ChildItem -LiteralPath (Join-Path $RunDirectory 'logs') -File -Filter 'DeviceTweaker_*.log')
    if ($logs.Count -ne 1) { throw 'Expected exactly one application session log' }
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Test-QaLogErrors.ps1') `
        -LogPath $logs[0].FullName -ReportPath (Join-Path $RunDirectory 'runtime-errors.json') *> (Join-Path $RunDirectory 'runtime-errors.log')
    if ($LASTEXITCODE -ne 0) { throw 'Unexpected ERROR / FATAL in development session log' }
    & dotnet run --project $inspector -- --dev-verify $RunDirectory *> (Join-Path $RunDirectory 'evidence.log')
    if ($LASTEXITCODE -ne 0) { throw 'Development session evidence is incomplete' }
    [ordered]@{ sessionSha256=(Get-FileHash -LiteralPath (Join-Path $RunDirectory 'session.json')).Hash;
        processReport=('process-'+$process.Id+'.json'); processId=$process.Id; exitCode=$process.ExitCode } |
        ConvertTo-Json | Set-Content -LiteralPath (Join-Path $RunDirectory 'execution.json') -Encoding UTF8
    $comparisonStatus = 'NOT_COMPARED'
    if ($CompareWith) {
        & dotnet run --project $inspector -- --dev-compare ([IO.Path]::GetFullPath($CompareWith)) $RunDirectory (Join-Path $RunDirectory 'comparison') *> (Join-Path $RunDirectory 'comparison.log')
        if ($LASTEXITCODE -ne 0) { throw 'Comparison failed / runs are incompatible' }
        $comparisonStatus = (Get-Content -LiteralPath (Join-Path $RunDirectory 'comparison\comparison.json') -Raw | ConvertFrom-Json).status
    }
    $reviewStatus = if ($comparisonStatus -eq 'MATCH') { 'MATCH' } else { 'REVIEW_REQUIRED' }
    [ordered]@{ status = $reviewStatus; functionalStatus = 'PASS'; comparison = $comparisonStatus;
        sessionSha256 = (Get-FileHash -LiteralPath (Join-Path $RunDirectory 'session.json')).Hash;
        report = 'index.html'; hardwareValidation = 'NOT_PERFORMED'; sourceReview = 'Inspect images before marking a visual change reviewed' } |
        ConvertTo-Json | Set-Content -LiteralPath (Join-Path $RunDirectory 'development.json') -Encoding UTF8
} catch {
    [ordered]@{ status = 'FAIL'; functionalStatus = 'FAIL'; error = $_.Exception.ToString(); hardwareValidation = 'NOT_PERFORMED' } |
        ConvertTo-Json | Set-Content -LiteralPath (Join-Path $RunDirectory 'development.json') -Encoding UTF8
    throw
} finally {
    if ($process) { if (-not $process.HasExited) { $null = Stop-QaOwnedProcess $process }; $process.Dispose() }
    foreach ($name in $vars) { [Environment]::SetEnvironmentVariable($name,$saved[$name]) }
}
