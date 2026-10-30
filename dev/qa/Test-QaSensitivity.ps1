# Deliberately corrupt isolated source copies; the expected tests must fail.
[CmdletBinding()]
param([Parameter(Mandatory = $true)][string]$OutputDirectory)

$ErrorActionPreference = 'Stop'
$root = (Resolve-Path (Join-Path $PSScriptRoot '..\..\src')).Path
$output = [IO.Path]::GetFullPath($OutputDirectory)
$rootPrefix = $root.TrimEnd('\') + '\'
if (($output.TrimEnd('\').Equals($root, [StringComparison]::OrdinalIgnoreCase) -or
    $output.StartsWith($rootPrefix, [StringComparison]::OrdinalIgnoreCase)) -and
    -not $output.StartsWith((Join-Path $root 'tmp\'), [StringComparison]::OrdinalIgnoreCase) -and
    -not $output.StartsWith((Join-Path $root 'bin\'), [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Source copies inside the project must be under tmp or bin, outside the compile graph'
}
New-Item -ItemType Directory -Path $output -Force | Out-Null
$results = New-Object 'System.Collections.Generic.List[object]'
$mutations = @(
    @{ name = 'shift-i225-register'; file = 'Core\NicItrService.cs'; filter = 'NicItrProgrammingTests';
        old = '"125D", "5503"], 0x1680, 0x4, 5, 32'; new = '"125D", "5503"], 0x1684, 0x4, 5, 32';
        expected = 'I225_WritesFiveExactAddressesAndVerifiesEachRegister' },
    @{ name = 'ignore-readback-mismatch'; file = 'Startup\DeviceTweakerImodRuntime.cs'; filter = 'NicItrProgrammingTests'; old = 'if (actual != expected)';
        new = 'if (actual == ulong.MaxValue && actual != expected)'; expected = 'Failure_StopsBeforeLaterQueues(failure: "mismatch"' },
    @{ name = 'erase-imod-counter'; file = 'Startup\DeviceTweakerImodRuntime.cs'; filter = 'ImodRegisterTests';
        old = '((uint)currentValue & 0xFFFF0000U) | (interval & 0xFFFFU)'; new = '(interval & 0xFFFFU)';
        expected = 'IntervalWritePreservesCounterAndVerifiesLowBits' },
    @{ name = 'ignore-backup-readback'; file = 'Core\BackupRestoreService.cs'; filter = 'BackupRestoreTests';
        old = 'Verify(targets[i], captureValue(targets[i]));'; new = '// readback deliberately removed';
        expected = 'SilentWriteFailureIsDetectedByReadback' },
    @{ name = 'wrong-gpu-physical-pair'; file = 'Core\AutoAffinityPlanner.cs'; filter = 'AutoAffinityPlannerTests';
        old = '.ThenByDescending(pair => Math.Max(pair.First.PrimaryLp, pair.Second.PrimaryLp))';
        new = '.ThenBy(pair => Math.Max(pair.First.PrimaryLp, pair.Second.PrimaryLp))';
        expected = 'GpuKeepsHighestAdjacentPhysicalPairDespiteCppcOrder' },
    @{ name = 'accept-polling-downgrade-too-early'; file = 'Core\RawPollingService.cs'; filter = 'RawPollingConfirmationTests';
        old = 'if (IsPollingDowngrade(hertz, state.LastDisplayedHertz) &&';
        new = 'if (IsPollingDowngrade(hertz, state.LastDisplayedHertz) && hertz < 0 &&';
        expected = 'MouseDowngradeNeedsRepeatedStableAndLiveConfirmation' },
    @{ name = 'restore-wrong-power-source'; file = 'Devices\PowerPlanRestoreService.cs'; filter = 'PowerPlanRestoreTests';
        old = 'backend.WriteIndex(scheme, ac, index);'; new = 'backend.WriteIndex(scheme, !ac, index);';
        expected = 'RestoreTouchesOnlySavedSchemeAndSelectedPowerSource' },
    @{ name = 'activate-inactive-power-plan'; file = 'Devices\PowerPlanRestoreService.cs'; filter = 'PowerPlanRestoreTests';
        old = 'if (active == scheme) backend.Activate(scheme);'; new = 'if (active != Guid.Empty) backend.Activate(scheme);';
        expected = 'RestoreTouchesOnlySavedSchemeAndSelectedPowerSource(active: False' },
    @{ name = 'release-read-only-lock-through-environment'; file = 'Core\QaRun.cs';
        filter = 'QaWriteGuardTests.ProcessLock_SurvivesEnvironmentChangesAndHasNoReleaseSetter';
        old = 'internal static readonly bool IsReadOnly ='; new = 'internal static bool IsReadOnly =>';
        expected = 'ProcessLock_SurvivesEnvironmentChangesAndHasNoReleaseSetter' },
    @{ name = 'confuse-msix-table-size-with-enable'; file = 'Core\PciInterruptMode.cs'; filter = 'PciInterruptModeTests';
        old = 'msiX = (U16(config, offset + 2) & 0x8000) != 0;';
        new = 'msiX = (U16(config, offset + 2) & 0x7FF) != 0;';
        expected = 'SingleInterruptDistinguishesMsiAndMsiXByEnableBits' },
    @{ name = 'ignore-single-pixel-regression'; file = '..\dev\qa\Workbench\VisualComparison.cs'; filter = 'WorkbenchTests';
        old = 'return new(count == 0 ? "MATCH" : "CHANGED"'; new = 'return new(count <= 1 ? "MATCH" : "CHANGED"';
        expected = 'ExactComparisonDetectsOneChangedPixelAndItsBounds' },
    @{ name = 'dispose-shared-font-before-controls'; file = 'GUI\MainForm.Theme.cs'; filter = 'WorkbenchLifecycleTests';
        old = 'base.Dispose(disposing);'; new = '_blockFont.Dispose(); base.Dispose(disposing);';
        expected = 'SharedFontsRemainAliveUntilChildControlsAreDisposed' },
    @{ name = 'reject-second-resource-list'; file = 'Core\PciInterruptMode.cs'; filter = 'WindowsResourceCaptureTests';
        old = 'line = rawLine; message = rawMessage;';
        new = 'if (consumed != data.Length) { error = "second list rejected"; return false; } line = rawLine; message = rawMessage;';
        expected = 'ActualWindowsRawAndTranslatedListsAreBothDecoded' },
    @{ name = 'parallel-driver-stage-without-lock'; file = 'Core\DriverStagingGate.cs'; filter = 'DriverStagingTests';
        old = 'lock (Sync) return stage();'; new = 'return stage();'; expected = 'ParallelStagesDoNotCompeteForTheSameFile' },
    @{ name = 'load-during-background-read'; file = 'Core\DriverRuntime.cs'; filter = 'DriverRuntimeTests';
        old = 'if (!allowLoad) return openCompatible() ? DriverReadyResult.Existing : DriverReadyResult.Unavailable;';
        new = 'if (!allowLoad) { loadKdu(); return DriverReadyResult.Unavailable; }'; expected = 'ReadbackNeverChecksJournalOrInvokesEitherLoader' },
    @{ name = 'fallback-after-loader-timeout'; file = 'Core\DriverRuntime.cs'; filter = 'DriverRuntimeTests';
        old = 'assertIdle(); // A timeout/unknown outcome must stop SCM fallback, even if a device appeared.';
        new = '// pending loader guard deliberately removed'; expected = 'PendingKduStopsServiceEvenIfDeviceAppeared' },
    @{ name = 'accept-incompatible-driver-protocol'; file = 'Core\DriverRuntime.cs'; filter = 'DriverRuntimeTests';
        old = 'protocol != ProtocolVersion'; new = 'protocol == 0'; expected = 'MalformedOrIncompatibleReplyIsRejected(offset: 4' },
    @{ name = 'drop-nic-values-from-startup-profile'; file = 'Core\MainForm.Startup.cs'; filter = 'StartupScriptTests';
        old = 'Values = entry.Values.Select(StartupProfile.Hex).ToList(),'; new = 'Values = [],';
        expected = 'ImportImodProfileGivesTheScriptTheSavedValues' },
    @{ name = 'route-by-slot-interrupter'; file = 'Startup\DeviceTweakerImodRuntime.cs'; filter = 'XhciTopologyTests';
        old = 'interrupter = endpointInterrupter;'; new = 'endpointInterrupter = interrupter;';
        expected = 'EndpointTargetWinsOverSlotTargetAndHubsAreSkipped' }
)

foreach ($mutation in $mutations) {
    Write-Host ("EXPECTED REGRESSION: " + $mutation.name + ' (isolated source copy)')
    # Keep the testhost path below Win32 MAX_PATH even in a nested QA snapshot.
    $caseDir = Join-Path $output ('m{0:D2}_{1}' -f $results.Count, [Guid]::NewGuid().ToString('N').Substring(0, 8))
    $copyRepo = Join-Path $caseDir 'repo'
    $copyRoot = Join-Path $copyRepo 'src'
    New-Item -ItemType Directory -Path $copyRoot -Force | Out-Null
    foreach ($folder in @('Affinity', 'Core', 'Devices', 'GUI', 'Interop', 'Localization', 'Models', 'Tweaks')) {
        foreach ($file in (Get-ChildItem -LiteralPath (Join-Path $root $folder) -Filter '*.cs' -Recurse -File)) {
            $relative = $file.FullName.Substring($rootPrefix.Length)
            $target = Join-Path $copyRoot $relative
            New-Item -ItemType Directory -Path (Split-Path $target) -Force | Out-Null
            Copy-Item -LiteralPath $file.FullName -Destination $target
        }
    }
    $extra = @('DeviceTweakerCS.csproj', '.editorconfig', 'assets\DEVICE TWEAKER.ico', 'assets\app.manifest',
        'Startup\ApplyIMOD.ps1', 'Startup\DeviceTweakerImodRuntime.cs',
        'IMOD\DTIMOD.sys', 'IMOD\DTIMOD.sys.sha256', 'IMOD\Loader\kdu.exe', 'IMOD\Loader\drv64.dll')
    $testCopy = Join-Path $copyRepo 'dev\qa\tests\DeviceTweaker.Tests'
    New-Item -ItemType Directory -Path $testCopy -Force | Out-Null
    Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot 'tests\DeviceTweaker.Tests') -File |
        Where-Object { $_.Extension -in @('.cs', '.csproj', '.json') } |
        ForEach-Object { Copy-Item -LiteralPath $_.FullName -Destination $testCopy }
    foreach ($relative in $extra) {
        $target = Join-Path $copyRoot $relative
        New-Item -ItemType Directory -Path (Split-Path $target) -Force | Out-Null
        Copy-Item -LiteralPath (Join-Path $root $relative) -Destination $target
    }
    $workbenchCopy = Join-Path $copyRepo 'dev\qa\Workbench'
    New-Item -ItemType Directory -Force -Path $workbenchCopy | Out-Null
    Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot 'Workbench') -File -Filter '*.cs' |
        ForEach-Object { Copy-Item -LiteralPath $_.FullName -Destination $workbenchCopy }
    $mutatedPath = Join-Path $copyRoot $mutation.file
    $source = [IO.File]::ReadAllText($mutatedPath)
    if ([regex]::Matches($source, [regex]::Escape($mutation.old)).Count -ne 1) {
        throw "Mutation anchor must occur exactly once: $($mutation.name)"
    }
    $index = $source.IndexOf($mutation.old, [StringComparison]::Ordinal)
    if ($index -lt 0) { throw "Mutation anchor missing: $($mutation.name)" }
    $mutated = $source.Substring(0, $index) + $mutation.new + $source.Substring($index + $mutation.old.Length)
    [IO.File]::WriteAllText($mutatedPath, $mutated, (New-Object Text.UTF8Encoding($false)))
    $resultDir = Join-Path $caseDir 'results'
    New-Item -ItemType Directory -Path $resultDir -Force | Out-Null
    $savedErrorPreference = $ErrorActionPreference
    try {
    # Windows PowerShell wraps native stderr in ErrorRecord; these failures are expected.
    $ErrorActionPreference = 'Continue'
    & dotnet test (Join-Path $testCopy 'DeviceTweaker.Tests.csproj') `
        -c Release -p:BuildImodDriver=false --nologo --filter ("FullyQualifiedName~" + $mutation.filter) `
        --logger 'trx;LogFileName=mutant.trx' --results-directory $resultDir 2>&1 |
        ForEach-Object { $_.ToString() } | Tee-Object -FilePath (Join-Path $resultDir 'test.log') | Out-Host
    $code = $LASTEXITCODE
    } finally { $ErrorActionPreference = $savedErrorPreference }
    $trxPath = Join-Path $resultDir 'mutant.trx'
    if (-not (Test-Path -LiteralPath $trxPath)) { throw "Mutation did not reach tests: $($mutation.name)" }
    [xml]$trx = Get-Content -LiteralPath $trxPath -Raw
    $failed = @($trx.TestRun.Results.UnitTestResult | Where-Object { $_.outcome -eq 'Failed' })
    $expectedFailures = @($failed | Where-Object { $_.testName.Contains($mutation.expected) })
    $detected = $code -ne 0 -and $expectedFailures.Count -gt 0
    $results.Add([pscustomobject]@{ mutation = $mutation.name; detected = $detected;
        expected = $mutation.expected; observed = @($expectedFailures | ForEach-Object { $_.testName });
        failedTests = @($failed | ForEach-Object { $_.testName }); exitCode = $code; trx = $trxPath;
        sourceSha256 = (Get-FileHash -LiteralPath $mutatedPath).Hash })
}
$results.ToArray() | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $output 'sensitivity.json') -Encoding UTF8
if (@($results | Where-Object { -not $_.detected }).Count -gt 0) { throw 'QA missed an intentional regression' }
Write-Output (Join-Path $output 'sensitivity.json')
