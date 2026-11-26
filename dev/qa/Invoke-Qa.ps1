# Local QA entry point. No settings writes, publication, git actions or driver loading.
[CmdletBinding()]
param(
    [ValidateSet('Fast', 'Full')]
    [string]$Mode = 'Fast',
    [string]$OutputRoot,
    [switch]$SkipGui,
    [switch]$VisibleGui,
    [switch]$WithNative,
    [switch]$Worker,
    [string]$RunDirectory,
    [string]$SourceWorkspace,
    [string]$BackupSamples,
    [ValidateSet('Secondary', 'Primary')]
    [string]$Monitor = 'Secondary'
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Support\Private-Desktop.ps1')
. (Join-Path $PSScriptRoot 'Support\Input-Snapshot.ps1')
. (Join-Path $PSScriptRoot 'Support\Evidence.ps1')

if ($SkipGui -and $Mode -eq 'Full') { $Mode = 'Fast' }
$qaWorkloadLease = if (-not $Worker) { [DeviceTweakerQaDesktop]::AcquireWorkload() } else { $null }
try {
$workspace = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$root = Join-Path $workspace 'src'
if ([string]::IsNullOrWhiteSpace($OutputRoot)) { $OutputRoot = Join-Path $workspace 'local\qa-runs' }
$outputFull = [IO.Path]::GetFullPath($OutputRoot).TrimEnd('\')
if (($outputFull.Equals($root, [StringComparison]::OrdinalIgnoreCase) -or
    $outputFull.StartsWith($root + '\', [StringComparison]::OrdinalIgnoreCase)) -and
    -not $outputFull.StartsWith((Join-Path $root 'tmp\'), [StringComparison]::OrdinalIgnoreCase) -and
    -not $outputFull.StartsWith((Join-Path $root 'bin\'), [StringComparison]::OrdinalIgnoreCase)) {
    throw 'QA output inside the project must be under tmp or bin'
}
$runDir = if ($RunDirectory) { [IO.Path]::GetFullPath($RunDirectory) } else {
    Join-Path ([IO.Path]::GetFullPath($OutputRoot)) ('run_{0}_{1}' -f
        (Get-Date -Format 'yyyyMMdd_HHmmss'), [Guid]::NewGuid().ToString('N').Substring(0, 8))
}
New-Item -ItemType Directory -Path $runDir -Force | Out-Null

if ($Mode -eq 'Full' -and -not $Worker) {
    if ($BackupSamples) { $BackupSamples = (Resolve-Path -LiteralPath $BackupSamples).Path; if (-not (Test-Path -LiteralPath $BackupSamples -PathType Container)) { throw 'BackupSamples must be a directory' } }
    try {
        $snapshot = New-QaInputSnapshot $workspace $runDir
        $parameters = @{ Mode = 'Full'; Worker = $true; RunDirectory = $runDir; SourceWorkspace = $workspace;
            SkipGui = [bool]$SkipGui; VisibleGui = [bool]$VisibleGui; WithNative = [bool]$WithNative; Monitor = $Monitor; BackupSamples = $BackupSamples }
        $script = Join-Path $snapshot 'dev\qa\Invoke-Qa.ps1'
        if ($VisibleGui) {
            & $script @parameters
            exit $LASTEXITCODE
        }
        $exitCode = Invoke-QaPrivateDesktop -ScriptPath $script -Parameters $parameters `
            -ReportDirectory (Join-Path $runDir 'isolation') -TimeoutSeconds 1800
        if ($null -eq $exitCode) { throw 'QA entry point cannot create a nested full run' }
        $report = Get-Content -LiteralPath (Join-Path $runDir 'qa.json') -Raw | ConvertFrom-Json
        if ($report.status -eq 'FAIL' -or $report.steps.Count -eq 0) { $exitCode = 1 }
        if ($exitCode -eq 0) {
            try { Test-QaEvidence $runDir }
            catch {
                $report.status = 'FAIL'
                $report.steps += [pscustomobject]@{ name = 'evidence'; status = 'FAIL'; error = $_.Exception.Message }
                $report | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $runDir 'qa.json') -Encoding UTF8
                throw
            }
        }
        Write-Output (Join-Path $runDir 'qa.json')
        exit $exitCode
    } catch {
        if (-not (Test-Path -LiteralPath (Join-Path $runDir 'qa.json'))) {
            @{ status = 'FAIL'; mode = $Mode; error = $_.Exception.Message; artifacts = $runDir } |
                ConvertTo-Json | Set-Content -LiteralPath (Join-Path $runDir 'qa.json') -Encoding UTF8
        }
        throw
    }
}

if ($Mode -eq 'Full') {
    if (-not $VisibleGui) { Assert-QaPrivateDesktop }
    Test-QaInputSnapshot $runDir
}

$savedReadOnly = $env:DEVICE_TWEAKER_QA_READ_ONLY
$savedOutput = $env:DEVICE_TWEAKER_QA_OUTPUT
$savedBackups = $env:DEVICE_TWEAKER_QA_BACKUPS
$env:DEVICE_TWEAKER_QA_BACKUPS = $null
if ($BackupSamples) {
    if (-not (Test-Path -LiteralPath $BackupSamples -PathType Container)) { throw 'BackupSamples must be an existing directory' }
    $env:DEVICE_TWEAKER_QA_BACKUPS = [IO.Path]::GetFullPath($BackupSamples)
}
$env:DEVICE_TWEAKER_QA_READ_ONLY = '1'
$env:DEVICE_TWEAKER_QA_OUTPUT = Join-Path $runDir 'unit-processes'
$steps = New-Object 'System.Collections.Generic.List[object]'
$status = 'FAIL'
$coverage = @()

function Invoke-DotnetStep([string]$Name, [string[]]$Arguments) {
    $log = Join-Path $runDir "$Name.log"
    $savedPreference = $ErrorActionPreference
    try {
        $ErrorActionPreference = 'Continue'
        & dotnet @Arguments 2>&1 | ForEach-Object { $_.ToString() } | Tee-Object -FilePath $log | Out-Host
        $code = $LASTEXITCODE
    } finally { $ErrorActionPreference = $savedPreference }
    $steps.Add([pscustomobject]@{ name = $Name; status = $(if ($code -eq 0) { 'PASS' } else { 'FAIL' }); exitCode = $code; log = $log })
    if ($code -ne 0) { throw "$Name failed with exit code $code; see $log" }
}

try {
    $activeWorkspace = if ($Mode -eq 'Full') { Split-Path $root } else { $workspace }
    Test-QaDriverArtifact $activeWorkspace
    $steps.Add([pscustomobject]@{ name = 'driver-artifact'; status = 'PASS' })

    Invoke-DotnetStep 'build' @('build', (Join-Path $root 'DeviceTweakerCS.csproj'), '-c', 'Release',
        '-p:TreatWarningsAsErrors=true', '-p:BuildImodDriver=false', '--nologo', '-m:1')
    Invoke-DotnetStep 'unit' @('test', (Join-Path $PSScriptRoot 'tests\DeviceTweaker.Tests\DeviceTweaker.Tests.csproj'),
        '-c', 'Release', '-p:BuildImodDriver=false', '--nologo', '-m:1', '--logger', 'trx;LogFileName=unit.trx', '--results-directory', $runDir)

    [xml]$trx = Get-Content -LiteralPath (Join-Path $runDir 'unit.trx') -Raw
    $unitResults = @($trx.TestRun.Results.UnitTestResult)
    if ($unitResults.Count -eq 0 -or @($unitResults | Where-Object outcome -ne 'Passed').Count -gt 0) {
        throw 'Required unit tests are missing, failed or skipped'
    }
    $definitions = @{}
    foreach ($test in $trx.TestRun.TestDefinitions.UnitTest) { $definitions[$test.id] = $test.TestMethod.className }
    $coverage = @($trx.TestRun.Results.UnitTestResult | Group-Object { $definitions[$_.testId] } | ForEach-Object {
        [pscustomobject]@{ testClass = $_.Name; passed = @($_.Group | Where-Object outcome -eq 'Passed').Count;
            failed = @($_.Group | Where-Object outcome -eq 'Failed').Count; total = $_.Count }
    })

    if ($WithNative) {
        $nativeCases = [ordered]@{ 'native-model' = $null; 'native-unload-mutant' = 'unload'; 'native-write-range-mutant' = 'write-past-device-end'; 'native-write-guard-mutant' = 'bypass-device-write-guard' }
        foreach ($nativeCase in $nativeCases.Keys) {
            $nativeArguments = @('-NoProfile','-ExecutionPolicy','Bypass','-File',(Join-Path $PSScriptRoot 'Checks\Test-NativeDriverModel.ps1'),'-OutputDirectory',(Join-Path $runDir $nativeCase))
            if ($nativeCases[$nativeCase]) { $nativeArguments += @('-InjectRegression', $nativeCases[$nativeCase]) }
            & powershell.exe @nativeArguments *> (Join-Path $runDir ($nativeCase+'.log'))
            if ($LASTEXITCODE -ne 0) { throw "Native model stage failed: $nativeCase" }
        }
        $steps.Add([pscustomobject]@{ name = 'native-driver'; status = 'PASS'; scope = 'Actual native functions with fake kernel APIs; unload, write-range and write-guard mutations detected; no hardware access' })
    }

    if ($Mode -ne 'Full') {
        $status = 'PASS'
    } else {
        Invoke-DotnetStep 'source-inventory' @('run', '--project', (Join-Path $PSScriptRoot 'Inspect\Inspect.csproj'), '--',
            $activeWorkspace, (Join-Path $root 'bin\Release\net10.0-windows\win-x64\DEVICE TWEAKER.dll'),
            (Join-Path $runDir 'source-ui-inventory.json'))
        Invoke-DotnetStep 'build-developer' @('build', (Join-Path $root 'DeviceTweakerCS.csproj'), '-c', 'Release',
            '-p:DeveloperQa=true', '-p:TreatWarningsAsErrors=true', '-p:BuildImodDriver=false', '--nologo', '-m:1')
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Checks\Test-UiControls.ps1') `
            -AssemblyPath (Join-Path $root 'bin\Release\net10.0-windows\win-x64\DEVICE TWEAKER.dll') `
            -OutputDirectory (Join-Path $runDir 'controls')
        if ($LASTEXITCODE -ne 0) { throw 'Production UI control contracts failed' }
        $steps.Add([pscustomobject]@{ name = 'ui-controls'; status = 'PASS'; report = (Join-Path $runDir 'controls\controls.json') })

        $exe = Join-Path $root 'bin\DeveloperQa\Release\net10.0-windows\win-x64\DEVICE TWEAKER QA.exe'
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Checks\Test-Workbench.ps1') `
            -ExePath $exe -OutputDirectory (Join-Path $runDir 'workbench')
        if ($LASTEXITCODE -ne 0) { throw 'Development workbench failed' }
        $steps.Add([pscustomobject]@{ name = 'workbench'; status = 'PASS'; report = (Join-Path $runDir 'workbench\workbench.json') })

        $guiDebugArgs = @()
        if ($VisibleGui) { $guiDebugArgs += '-VisibleGui' }
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Checks\Smoke-SafeGui.ps1') `
            -ExePath $exe -OutputRoot (Join-Path $runDir 'gui') -Monitor $Monitor @guiDebugArgs
        $guiCode = $LASTEXITCODE
        $steps.Add([pscustomobject]@{ name = 'gui'; status = $(if ($guiCode -eq 0) { 'PASS' } else { 'FAIL' }); exitCode = $guiCode })
        if ($guiCode -ne 0) { throw "GUI QA failed with exit code $guiCode" }

        $guiLogs = @(Get-ChildItem -LiteralPath (Join-Path $runDir 'gui'), (Join-Path $runDir 'controls') -Recurse -File -Filter 'DeviceTweaker_*.log' |
            Where-Object { $_.Directory.Name -eq 'logs' })
        if ($guiLogs.Count -lt 2) { throw 'Expected production and control QA session logs' }
        foreach ($sessionLog in $guiLogs) {
            & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Checks\Test-LogQuality.ps1') `
                -LogPath $sessionLog.FullName -RequireSessionSummary -RequireBuildIdentity
            if ($LASTEXITCODE -ne 0) { throw "Log contract failed: $($sessionLog.FullName)" }
            & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Checks\Test-QaLogErrors.ps1') `
                -LogPath $sessionLog.FullName -ReportPath (Join-Path $runDir ($sessionLog.BaseName + '-errors.json'))
            if ($LASTEXITCODE -ne 0) { throw "Unexpected errors: $($sessionLog.FullName)" }
        }
        $steps.Add([pscustomobject]@{ name = 'log-contract'; status = 'PASS'; sessions = $guiLogs.Count })
        $steps.Add([pscustomobject]@{ name = 'runtime-error-audit'; status = 'PASS'; sessions = $guiLogs.Count })

        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Checks\Test-GuiLayout.ps1') `
            -ExePath $exe -OutputDirectory (Join-Path $runDir 'layout')
        if ($LASTEXITCODE -ne 0) { throw 'GUI layout matrix failed' }
        $steps.Add([pscustomobject]@{ name = 'gui-layout'; status = 'PASS'; report = (Join-Path $runDir 'layout\layout.json') })
        foreach ($layoutLog in (Get-ChildItem -LiteralPath (Join-Path $runDir 'layout') -Recurse -File -Filter 'DeviceTweaker_*.log')) {
            & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Checks\Test-LogQuality.ps1') `
                -LogPath $layoutLog.FullName -RequireSessionSummary -RequireBuildIdentity
            if ($LASTEXITCODE -ne 0) { throw 'Layout log did not complete correctly' }
            & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Checks\Test-QaLogErrors.ps1') `
                -LogPath $layoutLog.FullName -ReportPath (Join-Path $layoutLog.Directory.FullName 'errors.json')
            if ($LASTEXITCODE -ne 0) { throw 'Unexpected layout scenario errors' }
        }

        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Checks\Test-GuiResourceStability.ps1') `
            -ExePath (Join-Path $root 'bin\Release\net10.0-windows\win-x64\DEVICE TWEAKER.exe') `
            -OutputDirectory (Join-Path $runDir 'resources') @guiDebugArgs
        $resourceCode = $LASTEXITCODE
        $steps.Add([pscustomobject]@{ name = 'gui-resources'; status = $(if ($resourceCode -eq 0) { 'PASS' } else { 'FAIL' });
            exitCode = $resourceCode; report = (Join-Path $runDir 'resources\resources.json') })
        if ($resourceCode -ne 0) { throw 'GUI resource stability failed' }
        $resourceLogs = @(Get-ChildItem -LiteralPath (Join-Path $runDir 'resources\processes\logs') -File -Filter 'DeviceTweaker_*.log')
        if ($resourceLogs.Count -ne 1) { throw 'Expected one production GUI resource session log' }
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Checks\Test-LogQuality.ps1') `
            -LogPath $resourceLogs[0].FullName -RequireSessionSummary -RequireBuildIdentity
        if ($LASTEXITCODE -ne 0) { throw 'Production GUI resource log contract failed' }
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Checks\Test-QaLogErrors.ps1') `
            -LogPath $resourceLogs[0].FullName -ReportPath (Join-Path $runDir 'resources\errors.json')
        if ($LASTEXITCODE -ne 0) { throw 'Unexpected production GUI resource errors' }

        Test-QaInputSnapshot $runDir
        $steps.Add([pscustomobject]@{ name = 'input-integrity'; status = 'PASS'; report = (Join-Path $runDir 'inputs.json') })
        $status = 'PASS'
    }
} catch {
    $steps.Add([pscustomobject]@{ name = 'runner'; status = 'FAIL'; error = $_.Exception.ToString();
        position = $_.InvocationInfo.PositionMessage })
    Write-Host ($_ | Out-String) -ForegroundColor Red
} finally {
    $report = [ordered]@{
        schemaVersion = 3
        mode = $Mode
        withNative = [bool]$WithNative
        status = $status
        inputs = $(if ($Mode -eq 'Full') { Join-Path $runDir 'inputs.json' } else { $null })
        inputManifestSha256 = $(if ($Mode -eq 'Full' -and (Test-Path -LiteralPath (Join-Path $runDir 'inputs.json'))) { (Get-FileHash -LiteralPath (Join-Path $runDir 'inputs.json')).Hash } else { $null })
        displayMode = $(if ($Mode -ne 'Full') { 'IN_MEMORY' } elseif ($VisibleGui) { 'VISIBLE_DEBUG' } else { 'PRIVATE_DESKTOP' })
        desktop = [DeviceTweakerQaDesktop]::CurrentDesktop()
        artifactsIdentity = @(Get-ChildItem -LiteralPath $root -Recurse -File |
            Where-Object { $_.FullName -match '\\bin\\(DeveloperQa\\)?Release\\' -and $_.Extension -in @('.exe', '.dll') } |
            ForEach-Object { @{ path = $_.FullName; sha256 = (Get-FileHash -LiteralPath $_.FullName).Hash } })
        hardwareValidation = 'SKIPPED'
        limitations = @('Registry/WMI/power settings are denied, not applied and verified.',
            'Live Raw Input frequency and real driver readback are not validated.')
        unitCoverage = $coverage
        steps = @($steps.ToArray())
        artifacts = $runDir
    }
    $report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $runDir 'qa.json') -Encoding UTF8
    if ($Mode -eq 'Fast' -and $status -eq 'PASS') { Test-QaEvidence $runDir }
    $env:DEVICE_TWEAKER_QA_READ_ONLY = $savedReadOnly
    $env:DEVICE_TWEAKER_QA_OUTPUT = $savedOutput
    $env:DEVICE_TWEAKER_QA_BACKUPS = $savedBackups
}
Write-Output (Join-Path $runDir 'qa.json')
if ($status -eq 'FAIL') { exit 1 }

} finally { if ($null -ne $qaWorkloadLease) { $qaWorkloadLease.Dispose() } }
