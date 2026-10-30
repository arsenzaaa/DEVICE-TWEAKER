. (Join-Path $PSScriptRoot 'Read-QaJson.ps1')
function Test-QaEvidence([string]$RunDirectory) {
    $facts = [Collections.Generic.List[object]]::new()
    $report = Read-QaJson (Join-Path $RunDirectory 'qa.json')
    $full = $report.status -eq 'PASS'
    if ($report.status -notin @('PASS', 'PARTIAL')) { throw 'Worker did not finish the required stages' }
    $required = @('build', 'unit', 'sensitivity', 'script-contracts', 'native-driver', 'build-developer', 'desktop-contracts', 'source-inventory')
    if ($full) { $required += @('gui', 'log-contract', 'runtime-error-audit', 'gui-layout', 'gui-resources', 'input-integrity', 'ui-controls', 'workbench') }
    foreach ($name in $required) {
        $step = @($report.steps | Where-Object { $_.name -eq $name -and $_.status -eq 'PASS' })
        if ($step.Count -ne 1) { throw "Mandatory QA stage is missing or ambiguous: $name" }
    }
    Test-QaInputSnapshot $RunDirectory
    if ((Get-FileHash -LiteralPath $report.inputs).Hash -ne $report.inputManifestSha256) { throw 'Input manifest digest changed' }
    [xml]$trx = Get-Content -LiteralPath (Join-Path $RunDirectory 'unit.trx') -Raw
    $tests = @($trx.TestRun.Results.UnitTestResult)
    if ($tests.Count -eq 0 -or @($tests | Where-Object outcome -ne 'Passed').Count) { throw 'Unit evidence contains absent/failed/skipped tests' }
    $manifest = Read-QaJson $report.inputs
    $classes = @($trx.TestRun.TestDefinitions.UnitTest | ForEach-Object { $_.TestMethod.className })
    foreach ($file in (Get-ChildItem -LiteralPath (Join-Path $manifest.snapshot 'dev\qa\tests\DeviceTweaker.Tests') -Filter '*Tests.cs' -File)) {
        $source = [IO.File]::ReadAllText($file.FullName)
        $namespace = [regex]::Match($source, 'namespace\s+([\w.]+)\s*[;{]')
        $testClasses = [regex]::Matches($source, 'public\s+(?:sealed\s+)?class\s+(\w+Tests)\b')
        if (-not $namespace.Success -or $testClasses.Count -eq 0) { throw "Could not discover test classes: $($file.Name)" }
        foreach ($class in $testClasses) {
            $expected = $namespace.Groups[1].Value + '.' + $class.Groups[1].Value
            if ($expected -notin $classes) { throw "An entire test class did not run: $expected" }
        }
    }
    $facts.Add(@{ area = 'unit'; observed = $tests.Count; evidence = (Join-Path $RunDirectory 'unit.trx') })
    $mutants = Read-QaJson (Join-Path $RunDirectory 'sensitivity\sensitivity.json')
    $expectedMutants = @('shift-i225-register', 'ignore-readback-mismatch', 'erase-imod-counter', 'ignore-backup-readback',
        'wrong-gpu-physical-pair', 'accept-polling-downgrade-too-early', 'restore-wrong-power-source',
        'activate-inactive-power-plan', 'release-read-only-lock-through-environment', 'confuse-msix-table-size-with-enable',
        'ignore-single-pixel-regression', 'dispose-shared-font-before-controls', 'reject-second-resource-list', 'parallel-driver-stage-without-lock',
        'load-during-background-read', 'fallback-after-loader-timeout', 'accept-incompatible-driver-protocol', 'drop-nic-values-from-startup-profile',
        'route-by-slot-interrupter')
    if ($mutants.Count -ne $expectedMutants.Count) { throw 'Intentional regression matrix is incomplete' }
    foreach ($name in $expectedMutants) {
        $entry = @($mutants | Where-Object mutation -eq $name)
        if ($entry.Count -ne 1 -or -not $entry[0].detected -or $entry[0].exitCode -eq 0) { throw "Regression went undetected: $name" }
        [xml]$result = Get-Content -LiteralPath $entry[0].trx -Raw
        $hits = @($result.TestRun.Results.UnitTestResult | Where-Object { $_.outcome -eq 'Failed' -and $_.testName.Contains($entry[0].expected) })
        if ($hits.Count -eq 0) { throw "No expected assertion failure in TRX: $name" }
    }
    $contracts = Read-QaJson (Join-Path $RunDirectory 'desktop-contracts\contracts.json')
    $driverContracts = Read-QaJson (Join-Path $RunDirectory 'script-contracts\driver-artifact-contracts.json')
    if ($driverContracts.Count -ne 5 -or @($driverContracts | Where-Object { -not $_.passed }).Count) { throw 'Driver / KDU source and binary rejection evidence is incomplete' }
    $nativeModel = Read-QaJson (Join-Path $RunDirectory 'native-model\result.json')
    $nativeMutations = @('native-unload-mutant', 'native-write-range-mutant' | ForEach-Object { Read-QaJson (Join-Path $RunDirectory ($_ + '\mutation.json')) })
    if ($nativeModel.status -ne 'PASS' -or $nativeModel.kernelExecuted -or $nativeModel.hardwareAccessed -or
        $nativeModel.driverSourceSha256 -ne (Get-FileHash (Join-Path $manifest.snapshot 'src\IMOD\Driver\imod_driver.c')).Hash -or
        $nativeModel.driverHeaderSha256 -ne (Get-FileHash (Join-Path $manifest.snapshot 'src\IMOD\Driver\imod_driver.h')).Hash -or
        $nativeModel.modelSourceSha256 -ne (Get-FileHash (Join-Path $RunDirectory 'native-model\driver-model.c')).Hash -or
        @($nativeMutations | Where-Object { $_.status -ne 'DETECTED' -or $_.kernelExecuted }).Count -or $nativeMutations.Count -ne 2) {
        throw 'Native model evidence is incomplete or changed' }
    if ($contracts.Count -ne 4 -or @($contracts | Where-Object { -not $_.passed }).Count) { throw 'Desktop refusal/exit/cleanup evidence is incomplete' }
    foreach ($identity in $report.artifactsIdentity) {
        if ((Get-FileHash -LiteralPath $identity.path).Hash -ne $identity.sha256) { throw "Tested binary changed: $($identity.path)" }
    }
    if ($full) {
        $bench = Read-QaJson (Join-Path $RunDirectory 'workbench\workbench.json')
        $expectedScenarios = @('device-flow','control-states','dtimod-status','imod-check','imod-editor',
            'text-fields-intel','text-fields-amd','text-fields-imod','text-fields-nic',
            'dialog-audit-en','dialog-audit-ru','dialog-audit-en-120','dialog-audit-en-144',
            'dialog-audit-ru-120','dialog-audit-ru-144','backup-menu-flow','navigation-audit','irq-check-loading')
        if ($bench.status -ne 'PASS' -or $bench.cases.Count -ne $expectedScenarios.Count -or
            @(Compare-Object $expectedScenarios @($bench.cases.scenario)).Count) { throw 'Workbench cases are incomplete' }
        foreach ($case in $bench.cases) {
            $comparison = Read-QaJson $case.comparison
            if ($case.status -ne 'PASS' -or $comparison.status -ne 'MATCH' -or $comparison.checkpoints.Count -lt 1) { throw 'Workbench comparison did not match' }
            foreach ($dir in @($case.before,$case.after)) {
                $inspector = Join-Path $manifest.snapshot 'dev\qa\Inspect\bin\Debug\net8.0-windows\Inspect.dll'
                & dotnet $inspector --dev-verify-complete $dir *> (Join-Path $dir 'final-evidence.log')
                if ($LASTEXITCODE -ne 0) { throw 'Workbench execution / scenario / state evidence is inconsistent' }
                $session = Read-QaJson (Join-Path $dir 'session.json')
                if ($session.status -ne 'PASS' -or -not $session.readOnly) { throw 'Workbench session failed / has real writes enabled' }
                if ($case.scenario -like 'text-fields-*') {
                    foreach ($auditId in @('editors-en','editors-ru-narrow')) {
                        $auditPath = 'state/'+$auditId+'/editors.json'
                        $audit = Read-QaJson (Join-Path $dir $auditPath)
                        if ($audit.status -ne 'PASS' -or $audit.cases.Count -lt 2 -or $audit.failures.Count -ne 0 -or
                            @($audit.cases | Where-Object { -not $_.completeGlyph -or -not $_.contained -or
                                -not $_.afterFocus -or -not $_.rendered -or ($_.enabled -and -not $_.focused) -or $_.family -eq 'unknown' }).Count) {
                            throw "Actual production editor evidence is incomplete: $($case.scenario) / $auditId"
                        }
                        $checkpoint = if ($auditId -eq 'editors-en') { 'audited-en' } else { 'audited-ru' }
                        $tree = Read-QaJson (Join-Path $dir ('trees/'+$checkpoint+'.json'))
                        $visibleEditors = @($tree.nodes | Where-Object { $_.visible -and $_.type -in @('TextBox','UpDownEdit') })
                        if ($visibleEditors.Count -ne $audit.cases.Count -or
                            @($audit.cases | Where-Object family -eq 'search').Count -ne 1 -or
                            @($audit.cases | Where-Object { $_.paths.Count -ne 4 -or ($_.enabled -and -not $_.focusPath) }).Count) {
                            throw "Editor inventory does not match the independent control tree: $($case.scenario) / $auditId"
                        }
                        $statePaths = @($session.stateFiles.path)
                        foreach ($path in @($auditPath)+@($audit.cases.paths)+@($audit.cases.focusPath | Where-Object { $_ })) {
                            if ($path -notin $statePaths) { throw "Missing editor screenshot/hash: $path" }
                        }
                    }
                }
                foreach ($capture in $session.captures) {
                    if ((Get-FileHash -LiteralPath (Join-Path $dir $capture.path)).Hash -ne $capture.sha256 -or
                        (Get-FileHash -LiteralPath (Join-Path $dir $capture.tree)).Hash -ne $capture.treeSha256) { throw 'Workbench evidence changed' }
                }
            }
            if ((Get-FileHash -LiteralPath (Join-Path $case.before 'session.json')).Hash -ne $comparison.beforeManifest -or
                (Get-FileHash -LiteralPath (Join-Path $case.after 'session.json')).Hash -ne $comparison.afterManifest) { throw 'Workbench comparison refers to different sessions' }
        }
        $facts.Add(@{ area='workbench'; observed=$bench.cases.Count; evidence=(Join-Path $RunDirectory 'workbench\workbench.json'); scope='Repeatability, not design approval' })
        $controls = Read-QaJson (Join-Path $RunDirectory 'controls\controls.json')
        if ($controls.status -ne 'PASS' -or $controls.checks.Count -lt 20 -or @($controls.checks | Where-Object status -ne 'PASS').Count) {
            throw 'Production control input/translation evidence is incomplete'
        }
        foreach ($name in @('Ctrl+Z stays with the active text editor', 'leaving picker closes popup',
            'right click leaves picker closed', 'RU to EN preserves the pending value',
            'throttle status update retains English source for other info rows')) {
            if (@($controls.checks | Where-Object name -eq $name).Count -ne 1) { throw "Missing control check: $name" }
        }
        if ($controls.images.Count -ne 8) { throw 'Production control state captures are incomplete' }
        $editors = Read-QaJson (Join-Path $RunDirectory 'controls\native-editors.json')
        if ($editors.status -ne 'PASS' -or $editors.cases.Count -ne 12 -or $editors.images.Count -ne 36 -or
            @($editors.cases | Where-Object { -not $_.completeGlyph -or -not $_.negativeDetected }).Count) {
            throw 'Native text glyph / deliberate clipping evidence is incomplete'
        }
        foreach ($shot in $editors.images) {
            if ((Get-FileHash -LiteralPath $shot.path).Hash -ne $shot.sha256) { throw 'Native text glyph evidence changed' }
        }
        foreach ($shot in $controls.images) {
            if ($shot.width -le 0 -or $shot.height -le 0 -or
                (Get-FileHash -LiteralPath $shot.path).Hash -ne $shot.sha256) { throw 'Control screenshot changed' }
            if (($shot.path.EndsWith('-normal.png') -and $shot.focused) -or
                ($shot.path.EndsWith('-focus.png') -and -not $shot.focused) -or
                ($shot.path.EndsWith('-disabled.png') -and $shot.enabled)) { throw 'Control capture state is incorrect' }
        }
        $facts.Add(@{ area = 'ui-controls'; observed = $controls.checks.Count; screenshots = $controls.images.Count;
            evidence = (Join-Path $RunDirectory 'controls\controls.json') })
        $guiFiles = @(Get-ChildItem -LiteralPath (Join-Path $RunDirectory 'gui') -Recurse -File -Filter 'smoke.json')
        if ($guiFiles.Count -ne 1) { throw 'GUI report is missing or ambiguous' }
        $gui = Read-QaJson $guiFiles[0].FullName
        if ($gui.status -ne 'PASS' -or @($gui.checks | Where-Object status -ne 'PASS').Count) { throw 'GUI checks are incomplete' }
        foreach ($description in @('QA process confirmed immutable read-only mode', 'runtime language switch EN -> RU',
            'runtime language switch RU -> EN', 'sandbox core matrix completed=', 'IMOD startup persistence contract completed=',
            'IMOD backup persistence contract completed=', 'Multi-controller affinity completed=',
            'test-mode disable action cleared the complete emulation state')) {
            if (@($gui.checks | Where-Object { $_.description.StartsWith($description) }).Count -ne 1) {
                throw "Mandatory GUI result is missing: $description"
            }
        }
        foreach ($name in @('success', 'warnings', 'partial', 'partial_3x', 'failed', 'stress')) {
            if (@($gui.checks | Where-Object description -eq "Russian result preview 25_result_${name}_ru.png").Count -ne 1) {
                throw "Missing Russian operation-result check: $name"
            }
        }
        if ($gui.processReports.Count -lt 4 -or $gui.screenshots.Count -lt 10) { throw 'GUI process/capture evidence is incomplete' }
        foreach ($processReport in $gui.processReports) {
            $process = Read-QaJson $processReport
            if (-not $process.readOnly -or ($report.displayMode -eq 'PRIVATE_DESKTOP' -and
                (-not $process.desktop.isolated -or $process.desktop.actual -ne $report.desktop))) { throw 'GUI process did not confirm read-only/desktop assignment' }
        }
        foreach ($shot in $gui.screenshots) {
            if ($shot.bytes -lt 1000 -or (Get-FileHash -LiteralPath $shot.path).Hash -ne $shot.sha256) { throw 'GUI screenshot evidence is invalid or changed' }
        }
        $layout = Read-QaJson (Join-Path $RunDirectory 'layout\layout.json')
        if ($layout.status -ne 'PASS' -or $layout.cases.Count -ne 6) { throw 'Layout matrix is incomplete' }
        foreach ($dpi in @(96, 120, 144)) {
            foreach ($language in @('en', 'ru')) {
                $case = @($layout.cases | Where-Object id -eq "layout-${dpi}-$language")
                if ($case.Count -ne 1 -or $case[0].status -ne 'PASS' -or $case[0].evidence.Count -ne 2) { throw 'Missing layout case/captures' }
                foreach ($shot in $case[0].evidence) {
                    if ((Get-FileHash -LiteralPath $shot.path).Hash -ne $shot.sha256) { throw 'Layout screenshot changed' }
                }
            }
        }
        $resources = Read-QaJson (Join-Path $RunDirectory 'resources\resources.json')
        if ($resources.status -ne 'PASS' -or $resources.cycles -ne 20 -or $resources.samples.Count -ne 21 -or
            $resources.refreshEvidence.Count -ne 21) { throw 'Resource cycles/completion evidence is incomplete' }
        foreach ($cycle in $resources.refreshEvidence) {
            if ($cycle.after.clicks -ne ($cycle.before.clicks + 1) -or $cycle.after.completed -ne ($cycle.before.completed + 1)) {
                throw 'Resource sample is not backed by exactly one completed REFRESH'
            }
        }
        $facts.Add(@{ area = 'gui'; observed = $gui.checks.Count; screenshots = $gui.screenshots.Count; evidence = $guiFiles[0].FullName })
        $facts.Add(@{ area = 'layout'; observed = 6; screenshots = 12; evidence = (Join-Path $RunDirectory 'layout\layout.json') })
    }
    $facts.Add(@{ area = 'sensitivity'; observed = $mutants.Count; evidence = (Join-Path $RunDirectory 'sensitivity\sensitivity.json') })
    @{ status = 'PASS'; scope = 'Evidence completeness, independent TRX failures and artifact digests; not protection against deliberate editing of tests and reports';
        facts = $facts.ToArray() } | ConvertTo-Json -Depth 7 |
        Set-Content -LiteralPath (Join-Path $RunDirectory 'evidence.json') -Encoding UTF8
}
