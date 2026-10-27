# Local QA entry point. No settings writes, publication, git actions or driver loading.
[CmdletBinding()]
param(
    [string]$OutputRoot,
    [switch]$SkipGui,
    [switch]$VisibleGui,
    [switch]$Worker,
    [string]$RunDirectory,
    [string]$SourceWorkspace,
    [ValidateSet('Secondary', 'Primary')]
    [string]$Monitor = 'Secondary'
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Private-Desktop.ps1')
. (Join-Path $PSScriptRoot 'Input-Snapshot.ps1')
. (Join-Path $PSScriptRoot 'Test-QaEvidence.ps1')
$root = (Resolve-Path (Join-Path $PSScriptRoot '..\..\DEVICE TWEAKER')).Path
if ([string]::IsNullOrWhiteSpace($OutputRoot)) { $OutputRoot = Join-Path $root 'tmp\qa' }
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
if (-not $Worker) {
    $workspace = Split-Path $root
    try {
        $snapshot = New-QaInputSnapshot $workspace $runDir
        $parameters = @{ Worker = $true; RunDirectory = $runDir; SourceWorkspace = $workspace;
            SkipGui = [bool]$SkipGui; VisibleGui = [bool]$VisibleGui; Monitor = $Monitor }
        $script = Join-Path $snapshot 'dev\qa\Invoke-Qa.ps1'
        if ($VisibleGui) {
            & $script @parameters
            exit $LASTEXITCODE
        }
        $exitCode = Invoke-QaPrivateDesktop -ScriptPath $script -Parameters $parameters `
            -ReportDirectory (Join-Path $runDir 'isolation') -TimeoutSeconds 1800
        if ($null -eq $exitCode) { throw 'QA entry point cannot create a nested full run' }
        # Inspect persisted reports, never trust a zero worker exit on its own.
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
            @{ status = 'FAIL'; error = $_.Exception.Message; artifacts = $runDir } |
                ConvertTo-Json | Set-Content -LiteralPath (Join-Path $runDir 'qa.json') -Encoding UTF8
        }
        throw
    }
}
if (-not $VisibleGui) { Assert-QaPrivateDesktop }
Test-QaInputSnapshot $runDir
$savedReadOnly = $env:DEVICE_TWEAKER_QA_READ_ONLY
$savedOutput = $env:DEVICE_TWEAKER_QA_OUTPUT
$savedBackups = $env:DEVICE_TWEAKER_QA_BACKUPS
$backupDirectory = Join-Path $SourceWorkspace 'DEVICE TWEAKER\bin\Publish\DEVICE TWEAKER (SELF-CONTAINED)\Backups'
if (Test-Path -LiteralPath $backupDirectory) { $env:DEVICE_TWEAKER_QA_BACKUPS = $backupDirectory }
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
    Invoke-DotnetStep 'build' @('build', (Join-Path $root 'DeviceTweakerCS.csproj'), '-c', 'Release',
        '-p:TreatWarningsAsErrors=true', '-p:BuildImodDriver=false', '--nologo')
    Invoke-DotnetStep 'unit' @('test', (Join-Path $PSScriptRoot 'tests\DeviceTweaker.Tests\DeviceTweaker.Tests.csproj'),
        '-c', 'Release', '-p:BuildImodDriver=false', '--nologo', '--logger', 'trx;LogFileName=unit.trx', '--results-directory', $runDir)
    [xml]$trx = Get-Content -LiteralPath (Join-Path $runDir 'unit.trx') -Raw
    $unitResults = @($trx.TestRun.Results.UnitTestResult)
    if ($unitResults.Count -eq 0 -or @($unitResults | Where-Object outcome -ne 'Passed').Count -gt 0) {
        throw 'Required unit tests are missing, failed or skipped'
    }
    $definitions = @{}
    foreach ($test in $trx.TestRun.TestDefinitions.UnitTest) { $definitions[$test.id] = $test.TestMethod.className }
    $coverage = @($trx.TestRun.Results.UnitTestResult | Group-Object { $definitions[$_.testId] } | ForEach-Object {
        $scope = switch -Regex ($_.Name) {
            'QaWriteGuardTests' { 'production write guards'; break }
            'PowerPlanRestoreTests' { 'production saved-scheme AC/DC restore and failures on a memory Power API'; break }
            'QaArtifactTests' { 'concurrent JSON readers/writers and absence of test panel in production assembly'; break }
            'BackupRestoreTests' { 'production restore transaction / rollback and version 1 snapshots on memory storage'; break }
            'ImodRegisterTests' { 'production IMOD register sequence on memory registers'; break }
            'AutoAffinityPlannerTests' { 'production CPU allocation on synthetic CPU topologies'; break }
            'NicItrProgrammingTests' { 'production NIC register sequence on an in-memory bus'; break }
            'RawPollingRoleTests' { 'production estimators on synthetic intervals and polling display rules; no live device measurement'; break }
            'RawPollingConfirmationTests' { 'production polling confirmation and downgrade state on synthetic events'; break }
            'DevicePowerDiagnosticsTests' { 'production WMI query/write/fresh-readback sequence on injected providers; no real WMI writes'; break }
            'NdisDiagnosticsTests' { 'production RSS comparison and query script with shadowed adapter cmdlets; no real adapter access'; break }
            'DiagnosticRegressionTests' { 'actual diagnostic level classifier and EXE build identity / file digest'; break }
            'DeviceSettingsWriterTests' { 'production MSI / IRQ / affinity value writes on a memory registry'; break }
            default { 'pure logic / parsing / profile lookup' }
        }
        [pscustomobject]@{ testClass = $_.Name; scope = $scope; passed = @($_.Group | Where-Object outcome -eq 'Passed').Count;
            failed = @($_.Group | Where-Object outcome -eq 'Failed').Count; total = $_.Count }
    })
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Test-QaSensitivity.ps1') `
        -OutputDirectory (Join-Path $runDir 'sensitivity')
    $sensitivityCode = $LASTEXITCODE
    $steps.Add([pscustomobject]@{ name = 'sensitivity'; status = $(if ($sensitivityCode -eq 0) { 'PASS' } else { 'FAIL' });
        exitCode = $sensitivityCode; report = (Join-Path $runDir 'sensitivity\sensitivity.json') })
    if ($sensitivityCode -ne 0) { throw 'QA sensitivity checks failed' }
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Test-QaScriptContracts.ps1') `
        -OutputDirectory (Join-Path $runDir 'script-contracts')
    $contractCode = $LASTEXITCODE
    $steps.Add([pscustomobject]@{ name = 'script-contracts'; status = $(if ($contractCode -eq 0) { 'PASS' } else { 'FAIL' }); exitCode = $contractCode })
    if ($contractCode -ne 0) { throw 'QA script contracts failed' }
    foreach ($nativeCase in @('native-model','native-unload-mutant')) {
        $nativeArguments = @('-NoProfile','-ExecutionPolicy','Bypass','-File',(Join-Path $PSScriptRoot 'Test-NativeDriverModel.ps1'),'-OutputDirectory',(Join-Path $runDir $nativeCase))
        if ($nativeCase -eq 'native-unload-mutant') { $nativeArguments += '-InjectUnloadRegression' }
        & powershell.exe @nativeArguments *> (Join-Path $runDir ($nativeCase+'.log'))
        if ($LASTEXITCODE -ne 0) { throw "Native model stage failed: $nativeCase" }
    }
    $steps.Add([pscustomobject]@{ name='native-driver'; status='PASS'; scope='Actual native functions with fake kernel APIs; unload mutation detected; no hardware access' })
    Invoke-DotnetStep 'build-developer' @('build', (Join-Path $root 'DeviceTweakerCS.csproj'), '-c', 'Release',
        '-p:DeveloperQa=true', '-p:TreatWarningsAsErrors=true', '-p:BuildImodDriver=false', '--nologo')
    Invoke-DotnetStep 'source-inventory' @('run', '--project', (Join-Path $PSScriptRoot 'Inspect\Inspect.csproj'), '--',
        (Split-Path $root), (Join-Path $root 'bin\Release\net8.0-windows\win-x64\DEVICE TWEAKER.dll'),
        (Join-Path $runDir 'source-ui-inventory.json'))
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Test-QaDesktopContracts.ps1') `
        -OutputDirectory (Join-Path $runDir 'desktop-contracts') `
        -ExePath (Join-Path $root 'bin\Release\net8.0-windows\win-x64\DEVICE TWEAKER.exe')
    if ($LASTEXITCODE -ne 0) { throw 'Private desktop launcher contracts failed' }
    $steps.Add([pscustomobject]@{ name = 'desktop-contracts'; status = 'PASS'; report = (Join-Path $runDir 'desktop-contracts\contracts.json') })
    if ($SkipGui) {
        $steps.Add([pscustomobject]@{ name = 'gui'; status = 'SKIPPED'; reason = '-SkipGui explicitly requested' })
        $status = 'PARTIAL'
        Test-QaInputSnapshot $runDir
    } else {
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Test-UiControls.ps1') `
            -AssemblyPath (Join-Path $root 'bin\Release\net8.0-windows\win-x64\DEVICE TWEAKER.dll') `
            -OutputDirectory (Join-Path $runDir 'controls')
        if ($LASTEXITCODE -ne 0) { throw 'Production UI control contracts failed' }
        $steps.Add([pscustomobject]@{ name = 'ui-controls'; status = 'PASS'; report = (Join-Path $runDir 'controls\controls.json') })
        $exe = Join-Path $root 'bin\DeveloperQa\Release\net8.0-windows\win-x64\DEVICE TWEAKER QA.exe'
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Test-Workbench.ps1') `
            -ExePath $exe -OutputDirectory (Join-Path $runDir 'workbench')
        if ($LASTEXITCODE -ne 0) { throw 'Development workbench failed' }
        $steps.Add([pscustomobject]@{ name = 'workbench'; status = 'PASS'; report = (Join-Path $runDir 'workbench\workbench.json');
            scope = 'Functional scenarios / same-build image determinism; visual inspection remains separate' })
        $guiDebugArgs = @()
        if ($VisibleGui) { $guiDebugArgs += '-VisibleGui' }
        # Separate PowerShell process keeps the UI helper definitions and environment local.
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Smoke-SafeGui.ps1') `
            -ExePath $exe -OutputRoot (Join-Path $runDir 'gui') -Monitor $Monitor @guiDebugArgs
        $guiCode = $LASTEXITCODE
        $steps.Add([pscustomobject]@{ name = 'gui'; status = $(if ($guiCode -eq 0) { 'PASS' } else { 'FAIL' }); exitCode = $guiCode })
        if ($guiCode -ne 0) { throw "GUI QA failed with exit code $guiCode" }
        $guiLogs = @(Get-ChildItem -LiteralPath (Join-Path $runDir 'gui'), (Join-Path $runDir 'controls') -Recurse -File -Filter 'DeviceTweaker_*.log' |
            Where-Object { $_.Directory.Name -eq 'logs' })
        if ($guiLogs.Count -lt 3) { throw 'Expected startup, sandbox and Russian GUI session logs' }
        foreach ($sessionLog in $guiLogs) {
            & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Test-LogQuality.ps1') `
                -LogPath $sessionLog.FullName -RequireSessionSummary -RequireBuildIdentity
            if ($LASTEXITCODE -ne 0) { throw "Log contract failed: $($sessionLog.FullName)" }
            & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Test-QaLogErrors.ps1') `
                -LogPath $sessionLog.FullName -ReportPath (Join-Path $runDir ($sessionLog.BaseName + '-errors.json'))
            if ($LASTEXITCODE -ne 0) { throw "Unexpected errors: $($sessionLog.FullName)" }

        }
        $steps.Add([pscustomobject]@{ name = 'log-contract'; status = 'PASS'; sessions = $guiLogs.Count })
        $steps.Add([pscustomobject]@{ name = 'runtime-error-audit'; status = 'PASS'; sessions = $guiLogs.Count })
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Test-GuiLayout.ps1') `
            -ExePath $exe -OutputDirectory (Join-Path $runDir 'layout')
        if ($LASTEXITCODE -ne 0) { throw 'GUI layout matrix failed' }
        $steps.Add([pscustomobject]@{ name = 'gui-layout'; status = 'PASS'; report = (Join-Path $runDir 'layout\layout.json') })
        foreach ($layoutLog in (Get-ChildItem -LiteralPath (Join-Path $runDir 'layout') -Recurse -File -Filter 'DeviceTweaker_*.log')) {
            & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Test-LogQuality.ps1') `
                -LogPath $layoutLog.FullName -RequireSessionSummary -RequireBuildIdentity
            if ($LASTEXITCODE -ne 0) { throw 'Layout log did not complete correctly' }
            & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Test-QaLogErrors.ps1') `
                -LogPath $layoutLog.FullName -ReportPath (Join-Path $layoutLog.Directory.FullName 'errors.json')
            if ($LASTEXITCODE -ne 0) { throw 'Unexpected layout scenario errors' }
        }
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Test-GuiResourceStability.ps1') `
            -ExePath (Join-Path $root 'bin\Release\net8.0-windows\win-x64\DEVICE TWEAKER.exe') `
            -OutputDirectory (Join-Path $runDir 'resources') @guiDebugArgs
        $resourceCode = $LASTEXITCODE
        $steps.Add([pscustomobject]@{ name = 'gui-resources'; status = $(if ($resourceCode -eq 0) { 'PASS' } else { 'FAIL' });
            exitCode = $resourceCode; report = (Join-Path $runDir 'resources\resources.json') })
        if ($resourceCode -ne 0) { throw 'GUI resource stability failed' }
        $resourceLogs = @(Get-ChildItem -LiteralPath (Join-Path $runDir 'resources\processes\logs') -File -Filter 'DeviceTweaker_*.log')
        if ($resourceLogs.Count -ne 1) { throw 'Expected one production GUI resource session log' }
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Test-LogQuality.ps1') `
            -LogPath $resourceLogs[0].FullName -RequireSessionSummary -RequireBuildIdentity
        if ($LASTEXITCODE -ne 0) { throw 'Production GUI resource log contract failed' }
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Test-QaLogErrors.ps1') `
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
        schemaVersion = 2
        status = $status
        inputs = (Join-Path $runDir 'inputs.json')
        inputManifestSha256 = (Get-FileHash -LiteralPath (Join-Path $runDir 'inputs.json')).Hash
        displayMode = $(if ($VisibleGui) { 'VISIBLE_DEBUG' } else { 'PRIVATE_DESKTOP' })
        desktop = [DeviceTweakerQaDesktop]::CurrentDesktop()
        artifactsIdentity = @(Get-ChildItem -LiteralPath $root -Recurse -File |
            Where-Object { $_.FullName -match '\\bin\\(DeveloperQa\\)?Release\\' -and $_.Extension -in @('.exe', '.dll') } |
            ForEach-Object { @{ path = $_.FullName; sha256 = (Get-FileHash -LiteralPath $_.FullName).Hash } })
        hardwareValidation = 'SKIPPED'
        limitations = @('Registry/WMI/power settings are denied, not applied and verified.',
            'NIC programming uses the production sequence on simulated registers.',
            'TEST ADMIN backup/restore and fault injection are virtual-model checks.',
            'Live Raw Input frequency and real driver readback are not validated.')
        unitCoverage = $coverage
        steps = @($steps.ToArray())
        artifacts = $runDir
    }
    $report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $runDir 'qa.json') -Encoding UTF8
    $env:DEVICE_TWEAKER_QA_READ_ONLY = $savedReadOnly
    $env:DEVICE_TWEAKER_QA_OUTPUT = $savedOutput
    $env:DEVICE_TWEAKER_QA_BACKUPS = $savedBackups
}
Write-Output (Join-Path $runDir 'qa.json')
if ($status -eq 'FAIL') { exit 1 }
