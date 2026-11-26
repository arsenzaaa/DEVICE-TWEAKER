[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$ExePath, [Parameter(Mandatory=$true)][string]$OutputDirectory)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\Support\Private-Desktop.ps1')
Assert-QaPrivateDesktop
$ExePath = (Resolve-Path -LiteralPath $ExePath).Path
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$cases = @()
# Single-pass execution per core scenario. Use Invoke-DevSession.ps1 -CompareWith when comparing before/after edits.
foreach ($name in @(
    'device-flow', 'control-states', 'dtimod-status', 'imod-check', 'imod-editor', 'imod-invalid-role',
    'text-fields-imod', 'text-fields-nic', 'dialog-audit-en', 'dialog-audit-ru',
    'backup-menu-flow', 'navigation-audit', 'irq-check-loading', 'refresh-reuse-lifecycle')) {
    $scenarioPath = Join-Path $PSScriptRoot ('..\Scenarios\' + $name + '.json')
    $runDir = Join-Path $OutputDirectory $name
    $arguments = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File',
        (Join-Path $PSScriptRoot '..\Invoke-DevSession.ps1'),
        '-Worker', '-RunDirectory', $runDir, '-ExePath', $ExePath, '-Scenario', $scenarioPath)
    & powershell.exe @arguments *> (Join-Path $OutputDirectory ($name + '.log'))
    if ($LASTEXITCODE -ne 0) { throw "Workbench scenario failed: $name; see $runDir" }
    $development = Get-Content -LiteralPath (Join-Path $runDir 'development.json') -Raw | ConvertFrom-Json
    if ($development.functionalStatus -ne 'PASS') { throw "Workbench scenario failed verification: $name" }
    $cases += [pscustomobject]@{ scenario = $name; status = 'PASS'; run = $runDir }
}
[ordered]@{
    status = 'PASS'
    scope = 'Streamlined single-pass device, editor, IMOD, NIC, dialog (EN/RU), backup, navigation, IRQ and refresh-lifecycle scenarios in read-only mode. Use Invoke-DevSession.ps1 -CompareWith for before/after image diffs.'
    cases = $cases
    hardwareValidation = 'NOT_PERFORMED'
} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'workbench.json') -Encoding UTF8
