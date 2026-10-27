[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$ExePath, [Parameter(Mandatory=$true)][string]$OutputDirectory)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Private-Desktop.ps1')
Assert-QaPrivateDesktop
$ExePath = (Resolve-Path -LiteralPath $ExePath).Path
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$cases = @()
foreach ($name in @('device-flow','control-states','dtimod-status','imod-check','imod-editor',
    'text-fields-intel','text-fields-amd','text-fields-imod','text-fields-nic',
    'dialog-audit-en','dialog-audit-ru','dialog-audit-en-120','dialog-audit-en-144',
    'dialog-audit-ru-120','dialog-audit-ru-144','backup-menu-flow','navigation-audit','irq-check-loading')) {
    $first = Join-Path $OutputDirectory ($name+'-before')
    $second = Join-Path $OutputDirectory ($name+'-after')
    foreach ($entry in @(@{dir=$first;compare=$null},@{dir=$second;compare=$first})) {
        $arguments = @('-NoProfile','-ExecutionPolicy','Bypass','-File',(Join-Path $PSScriptRoot 'Invoke-DevSession.ps1'),
            '-Worker','-RunDirectory',$entry.dir,'-ExePath',$ExePath,'-Scenario',(Join-Path $PSScriptRoot ('Scenarios\'+$name+'.json')))
        if ($entry.compare) { $arguments += @('-CompareWith',$entry.compare) }
        & powershell.exe @arguments *> (Join-Path $OutputDirectory ($name+'-'+[IO.Path]::GetFileName($entry.dir)+'.log'))
        if ($LASTEXITCODE -ne 0) { throw "Workbench scenario failed: $name; see $($entry.dir)" }
    }
    $development = Get-Content -LiteralPath (Join-Path $second 'development.json') -Raw | ConvertFrom-Json
    if ($development.status -ne 'MATCH') { throw "Identical-source scenario is not deterministic: $name" }
    $cases += [pscustomobject]@{ scenario=$name; status='PASS'; before=$first; after=$second; comparison=(Join-Path $second 'comparison\comparison.json') }
}
[ordered]@{ status='PASS'; scope='Repeatable device/editor scenarios, production dialogs in RU/EN at 100/125/150% layout models, backup deletion/retention/failures, file locations and dropdown/filter navigation. Native snapshots, transitions, sandbox files and injected exceptions. Repeated same-build comparison is a determinism check; visual quality still requires inspection.';
    cases=$cases; hardwareValidation='NOT_PERFORMED' } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'workbench.json') -Encoding UTF8
