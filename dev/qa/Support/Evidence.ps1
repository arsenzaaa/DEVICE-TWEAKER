. (Join-Path $PSScriptRoot 'Read-QaJson.ps1')
function Test-QaEvidence([string]$RunDirectory) {
    $facts = [Collections.Generic.List[object]]::new()
    $report = Read-QaJson (Join-Path $RunDirectory 'qa.json')
    $full = $report.status -eq 'PASS' -and $report.mode -eq 'Full'
    if ($report.status -ne 'PASS') { throw 'Worker did not finish the required stages' }
    $required = @('build', 'unit', 'driver-artifact')
    if ($full) {
        $required += @('source-inventory', 'build-developer', 'ui-controls', 'workbench', 'gui', 'log-contract', 'runtime-error-audit', 'gui-layout', 'gui-resources', 'input-integrity')
    }
    foreach ($name in $required) {
        $step = @($report.steps | Where-Object { $_.name -eq $name -and $_.status -eq 'PASS' })
        if ($step.Count -ne 1) { throw "Mandatory QA stage is missing or ambiguous: $name" }
    }

    $unitTrxPath = Join-Path $RunDirectory 'unit.trx'
    [xml]$trx = Get-Content -LiteralPath $unitTrxPath -Raw
    $tests = @($trx.TestRun.Results.UnitTestResult)
    if ($tests.Count -eq 0 -or @($tests | Where-Object outcome -ne 'Passed').Count) {
        throw 'Unit evidence contains absent, failed or skipped tests'
    }
    $facts.Add(@{ area = 'unit'; observed = $tests.Count; evidence = $unitTrxPath })

    if ($report.withNative) {
        $nativeStep = @($report.steps | Where-Object { $_.name -eq 'native-driver' -and $_.status -eq 'PASS' })
        if ($nativeStep.Count -ne 1) { throw 'Native driver stage requested but did not PASS' }
        $nativeModel = Read-QaJson (Join-Path $RunDirectory 'native-model\result.json')
        if ($nativeModel.status -ne 'PASS' -or $nativeModel.kernelExecuted -or $nativeModel.hardwareAccessed) {
            throw 'Native model evidence is incomplete or invalid'
        }
        $facts.Add(@{ area = 'native-driver'; observed = $nativeModel.functions.Count; evidence = (Join-Path $RunDirectory 'native-model\result.json') })
    }

    if ($full) {
        Test-QaInputSnapshot $RunDirectory
        if ((Get-FileHash -LiteralPath $report.inputs).Hash -ne $report.inputManifestSha256) { throw 'Input manifest digest changed' }
        $manifest = Read-QaJson $report.inputs
        $classes = @($trx.TestRun.TestDefinitions.UnitTest | ForEach-Object { $_.TestMethod.className })
        foreach ($file in (Get-ChildItem -LiteralPath (Join-Path $manifest.snapshot 'dev\qa\tests\DeviceTweaker.Tests') -Filter '*Tests.cs' -File)) {
            $source = [IO.File]::ReadAllText($file.FullName)
            $namespace = [regex]::Match($source, 'namespace\s+([\w.]+)\s*[;{]')
            $testClasses = [regex]::Matches($source, 'public\s+(?:(?:sealed|abstract|partial)\s+)*class\s+(\w+Tests)\b')
            if (-not $namespace.Success -or $testClasses.Count -eq 0) { throw "Could not discover test classes: $($file.Name)" }
            foreach ($class in $testClasses) {
                $expected = $namespace.Groups[1].Value + '.' + $class.Groups[1].Value
                if ($expected -notin $classes) { throw "An entire test class did not run: $expected" }
            }
        }
        foreach ($identity in $report.artifactsIdentity) {
            if ((Get-FileHash -LiteralPath $identity.path).Hash -ne $identity.sha256) { throw "Tested binary changed: $($identity.path)" }
        }

        $bench = Read-QaJson (Join-Path $RunDirectory 'workbench\workbench.json')
        if ($bench.status -ne 'PASS' -or $bench.cases.Count -lt 10) { throw 'Workbench cases are incomplete' }
        $inspector = Join-Path $manifest.snapshot 'dev\qa\Inspect\bin\Release\net10.0-windows\Inspect.dll'
        if (-not (Test-Path -LiteralPath $inspector)) {
            $inspector = Join-Path $manifest.snapshot 'dev\qa\Inspect\bin\Debug\net10.0-windows\Inspect.dll'
        }
        foreach ($case in $bench.cases) {
            if ($case.status -ne 'PASS' -or -not $case.run) { throw "Workbench scenario did not pass: $($case.scenario)" }
            & dotnet $inspector --dev-verify-complete $case.run *> (Join-Path $case.run 'final-evidence.log')
            if ($LASTEXITCODE -ne 0) { throw "Workbench evidence is inconsistent: $($case.scenario)" }
            $session = Read-QaJson (Join-Path $case.run 'session.json')
            if ($session.status -ne 'PASS' -or -not $session.readOnly) { throw "Workbench session failed or leaked write mode: $($case.scenario)" }
            foreach ($capture in $session.captures) {
                if ((Get-FileHash -LiteralPath (Join-Path $case.run $capture.path)).Hash -ne $capture.sha256 -or
                    (Get-FileHash -LiteralPath (Join-Path $case.run $capture.tree)).Hash -ne $capture.treeSha256) { throw 'Workbench evidence changed' }
            }
        }
        $facts.Add(@{ area = 'workbench'; observed = $bench.cases.Count; evidence = (Join-Path $RunDirectory 'workbench\workbench.json') })

        $controls = Read-QaJson (Join-Path $RunDirectory 'controls\controls.json')
        if ($controls.status -ne 'PASS' -or $controls.checks.Count -lt 20 -or @($controls.checks | Where-Object status -ne 'PASS').Count) {
            throw 'Production control input/translation evidence is incomplete'
        }
        foreach ($shot in $controls.images) {
            if ($shot.width -le 0 -or $shot.height -le 0 -or (Get-FileHash -LiteralPath $shot.path).Hash -ne $shot.sha256) {
                throw 'Control screenshot changed'
            }
        }
        $facts.Add(@{ area = 'ui-controls'; observed = $controls.checks.Count; screenshots = $controls.images.Count; evidence = (Join-Path $RunDirectory 'controls\controls.json') })

        $guiFiles = @(Get-ChildItem -LiteralPath (Join-Path $RunDirectory 'gui') -Recurse -File -Filter 'smoke.json')
        if ($guiFiles.Count -ne 1) { throw 'GUI report is missing or ambiguous' }
        $gui = Read-QaJson $guiFiles[0].FullName
        if ($gui.status -ne 'PASS' -or @($gui.checks | Where-Object status -ne 'PASS').Count) { throw 'GUI checks are incomplete' }
        foreach ($description in @('QA process confirmed immutable read-only mode', 'runtime language switch EN -> RU', 'runtime language switch RU -> EN')) {
            if (@($gui.checks | Where-Object { $_.description.StartsWith($description) }).Count -ne 1) {
                throw "Mandatory GUI result is missing: $description"
            }
        }
        foreach ($processReport in $gui.processReports) {
            $process = Read-QaJson $processReport
            if (-not $process.readOnly -or ($report.displayMode -eq 'PRIVATE_DESKTOP' -and
                (-not $process.desktop.isolated -or $process.desktop.actual -ne $report.desktop))) { throw 'GUI process did not confirm read-only/desktop assignment' }
        }
        foreach ($shot in $gui.screenshots) {
            if ($shot.bytes -lt 1000 -or (Get-FileHash -LiteralPath $shot.path).Hash -ne $shot.sha256) { throw 'GUI screenshot evidence is invalid or changed' }
        }
        $facts.Add(@{ area = 'gui'; observed = $gui.checks.Count; screenshots = $gui.screenshots.Count; evidence = $guiFiles[0].FullName })

        $layout = Read-QaJson (Join-Path $RunDirectory 'layout\layout.json')
        if ($layout.status -ne 'PASS' -or $layout.cases.Count -ne 6) { throw 'Layout matrix is incomplete' }
        foreach ($case in $layout.cases) {
            if ($case.status -ne 'PASS' -or $case.evidence.Count -ne 2) { throw "Layout case failed: $($case.id)" }
            foreach ($shot in $case.evidence) {
                if ((Get-FileHash -LiteralPath $shot.path).Hash -ne $shot.sha256) { throw 'Layout screenshot changed' }
            }
        }
        $facts.Add(@{ area = 'layout'; observed = $layout.cases.Count; screenshots = 12; evidence = (Join-Path $RunDirectory 'layout\layout.json') })

        $resources = Read-QaJson (Join-Path $RunDirectory 'resources\resources.json')
        if ($resources.status -ne 'PASS' -or $resources.cycles -ne 20 -or $resources.samples.Count -ne 21) {
            throw 'Resource stability cycles evidence is incomplete'
        }
        $facts.Add(@{ area = 'resources'; observed = $resources.cycles; evidence = (Join-Path $RunDirectory 'resources\resources.json') })
    }

    @{ status = 'PASS'; mode = $report.mode; facts = $facts.ToArray() } | ConvertTo-Json -Depth 7 |
        Set-Content -LiteralPath (Join-Path $RunDirectory 'evidence.json') -Encoding UTF8
}
