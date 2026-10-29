param([Parameter(Mandatory = $true)][string]$OutputDirectory)
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$cases = @(
    @{ name = 'fault-consumed'; exit = 0; text = @'
[INFO] TEST.SCENARIO.BEGIN: fault=RegistryWriteDenied
[ERROR] TEST.SANDBOX.WRITE.ERROR: id=memory error=denied
[INFO] TEST.SCENARIO.END: status=PASS assertions=7/0 faultConsumed=True
'@ },
    @{ name = 'preview'; exit = 0; text = '[ERROR] UI.RESULT.ITEM: severity=Error scope=qa-dialog-preview' },
    @{ name = 'unexpected'; exit = 1; text = '[ERROR] BACKUP.RESTORE: access denied' },
    @{ name = 'failed-scenario'; exit = 1; text = @'
[INFO] TEST.SCENARIO.BEGIN: fault=RegistryWriteDenied
[ERROR] TEST.SANDBOX.WRITE.ERROR: id=memory error=denied
[INFO] TEST.SCENARIO.END: status=FAIL assertions=6/1 faultConsumed=True
'@ },
    @{ name = 'unexpected-in-fault'; exit = 1; text = @'
[INFO] TEST.SCENARIO.BEGIN: fault=RegistryWriteDenied
[ERROR] BACKUP.RESTORE: unrelated failure
[INFO] TEST.SCENARIO.END: status=PASS assertions=7/0 faultConsumed=True
'@ },
    @{ name = 'fatal-preview'; exit = 1; text = '[FATAL] UI.RESULT.ITEM: severity=Error scope=qa-dialog-preview' },
    @{ name = 'backup-delete-consumed'; exit = 0; text = @'
[INFO] TEST.SCENARIO.BEGIN: dialog=backup-delete fault=BackupDelete
[ERROR] BACKUP.DELETE: failed path=fixture error=System.IO.IOException: Injected file deletion failure.
[INFO] TEST.SCENARIO.END: status=PASS assertions=2/0 faultConsumed=True
'@ },
    @{ name = 'backup-delete-unclosed'; exit = 1; text = @'
[INFO] TEST.SCENARIO.BEGIN: dialog=backup-delete fault=BackupDelete
[ERROR] BACKUP.DELETE: failed path=fixture error=System.IO.IOException: Injected file deletion failure.
'@ },
    @{ name = 'backup-delete-unrelated'; exit = 1; text = @'
[INFO] TEST.SCENARIO.BEGIN: dialog=backup-delete fault=BackupDelete
[ERROR] BACKUP.DELETE: failed path=fixture error=System.IO.IOException: Access denied.
[INFO] TEST.SCENARIO.END: status=PASS assertions=2/0 faultConsumed=True
'@ },
    @{ name = 'backup-delete-without-scenario'; exit = 1; text = '[ERROR] BACKUP.DELETE: failed path=fixture error=System.IO.IOException: Injected file deletion failure.' }
)
$results = @()
foreach ($case in $cases) {
    $file = Join-Path $OutputDirectory ($case.name + '.log')
    [IO.File]::WriteAllText($file, $case.text)
    $old = $ErrorActionPreference
    try {
        $ErrorActionPreference = 'Continue'
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Test-QaLogErrors.ps1') `
            -LogPath $file -ReportPath (Join-Path $OutputDirectory ($case.name + '.json')) *> (Join-Path $OutputDirectory ($case.name + '-output.txt'))
        $code = $LASTEXITCODE
    } finally { $ErrorActionPreference = $old }
    $correct = ($case.exit -eq 0 -and $code -eq 0) -or ($case.exit -ne 0 -and $code -ne 0)
    $results += [pscustomobject]@{ case = $case.name; expectedExit = $case.exit; actualExit = $code; passed = $correct }
}
$results | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $OutputDirectory 'contracts.json') -Encoding UTF8
. (Join-Path $PSScriptRoot 'Input-Snapshot.ps1')
$workspace = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$copy = Join-Path $OutputDirectory 'driver-artifact-copy'
$app = Join-Path $copy 'src'
foreach ($name in @('IMOD\Driver\imod_driver.c', 'IMOD\Driver\imod_driver.h', 'IMOD\Driver\ImodDriver.vcxproj',
    'publish-variants.ps1', 'IMOD\DTIMOD.sys', 'IMOD\DTIMOD.sys.sha256', 'IMOD\DTIMOD.build.json')) {
    $target = Join-Path $app $name
    New-Item -ItemType Directory -Force -Path (Split-Path $target) | Out-Null
    Copy-Item -LiteralPath (Join-Path (Join-Path $workspace 'src') $name) -Destination $target
}
foreach ($file in (Get-QaInputFiles $workspace | Where-Object { $_.FullName.Contains('\IMOD\KDU\') -or $_.FullName.Contains('\IMOD\Loader\') })) {
    $relative = $file.FullName.Substring((Join-Path $workspace 'src').Length + 1)
    $target = Join-Path $app $relative
    New-Item -ItemType Directory -Force -Path (Split-Path $target) | Out-Null
    Copy-Item -LiteralPath $file.FullName -Destination $target
}
Test-QaDriverArtifact $copy
$driverContracts = @()
foreach ($case in @(
    @{ name = 'stale-driver-source'; path = 'IMOD\Driver\imod_driver.c'; error = 'DTIMOD source changed since the binary was built:' },
    @{ name = 'altered-driver-binary'; path = 'IMOD\DTIMOD.sys'; error = 'DTIMOD binary differs from its build manifest' },
    @{ name = 'stale-kdu-source'; path = 'IMOD\KDU\Source\Hamakaze\main.cpp'; error = 'KDU source differs from pinned upstream:' },
    @{ name = 'altered-kdu-binary'; path = 'IMOD\Loader\kdu.exe'; error = 'KDU binary mismatch:' },
    @{ name = 'mixed-kdu-database'; path = 'IMOD\Loader\drv64.dll'; error = 'KDU binary mismatch:' })) {
    $target = Join-Path $app $case.path
    $original = [IO.File]::ReadAllBytes($target)
    try {
        $changed = [byte[]]$original.Clone(); $changed[$changed.Length - 1] = $changed[$changed.Length - 1] -bxor 1
        [IO.File]::WriteAllBytes($target, $changed)
        $rejected = $false
        try { Test-QaDriverArtifact $copy } catch { $rejected = $_.Exception.Message.StartsWith($case.error) }
        $driverContracts += [pscustomobject]@{ case = $case.name; passed = $rejected }
        if (-not $rejected) { throw "Driver artifact mismatch was not rejected: $($case.name)" }
    } finally { [IO.File]::WriteAllBytes($target, $original) }
}
$driverContracts | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $OutputDirectory 'driver-artifact-contracts.json') -Encoding UTF8
if (@($results | Where-Object { -not $_.passed }).Count) { throw 'QA error classifier failed its regression cases' }
. (Join-Path $PSScriptRoot 'Visual-Assertions.ps1')
Add-Type -AssemblyName System.Drawing
function Write-Report([string]$message) { Write-Host $message }
$origin = [pscustomobject]@{ Left=0; Top=0 }
$region = [pscustomobject]@{ Left=10; Top=10; Right=90; Bottom=30; Width=80; Height=20 }
$bitmap = New-Object System.Drawing.Bitmap 100,50
try {
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    try { $graphics.Clear([Drawing.Color]::Black); $graphics.FillRectangle([Drawing.Brushes]::White,20,15,50,5) }
    finally { $graphics.Dispose() }
    $positive = Join-Path $OutputDirectory 'subtitle-visible.png'
    $bitmap.Save($positive,[Drawing.Imaging.ImageFormat]::Png)
    Assert-RegionRendered $positive $origin $region 'subtitle positive control'
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    try { $graphics.Clear([Drawing.Color]::Black) } finally { $graphics.Dispose() }
    $negative = Join-Path $OutputDirectory 'subtitle-missing.png'
    $bitmap.Save($negative,[Drawing.Imaging.ImageFormat]::Png)
    $rejected = $false
    try { Assert-RegionRendered $negative $origin $region 'subtitle negative control' }
    catch { $rejected = $_.Exception.Message.StartsWith('Visual text disappeared from capture:') }
    if (-not $rejected) { throw 'A missing subtitle was not rejected by the live GUI pixel assertion' }
    @{ status='PASS'; visibleRegionAccepted=$true; missingRegionRejected=$true } | ConvertTo-Json |
        Set-Content -LiteralPath (Join-Path $OutputDirectory 'subtitle-contracts.json') -Encoding UTF8
} finally { $bitmap.Dispose() }
Write-Host 'PASS: QA error classifier accepts expected faults and rejects unexpected errors.'
$versionBase = 'LOG.VERSION: product=0.0.4 file=0.0.4.0 assembly=0.0.4.0'
$identity = ' build='+('a'*32)+' exeSha256='+('B'*64)
$identityContracts = @()
foreach ($case in @(
    @{name='legacy-log'; version=$versionBase; require=$false; expected=0},
    @{name='identified-log'; version=$versionBase+$identity; require=$true; expected=0},
    @{name='missing-build-identity'; version=$versionBase; require=$true; expected=1},
    @{name='damaged-build-digest'; version=$versionBase+$identity.Substring(0,$identity.Length-1); require=$true; expected=1}
)) {
    $events = @('LOG.SCHEMA: version=3 fixture=identity','LOG.SESSION.START: fixture=identity',$case.version,
        'LOG.SESSION.SUMMARY: entries=3 infoEntries=3 warningEntries=0 errorEntries=0 fatalEntries=0 continuationEntries=0')
    $fixture = Join-Path $OutputDirectory ($case.name+'.log')
    $sequence=0
    $fixtureLines = @($events | ForEach-Object { $sequence++; '[2026-10-24 00:00:00.000 +03:00] [#{0:D6}] [+0000000ms] [T01] [INFO] [LOG] {1}' -f $sequence,$_ })
    [IO.File]::WriteAllLines($fixture,$fixtureLines,[Text.UTF8Encoding]::new($true))
    $arguments=@('-NoProfile','-ExecutionPolicy','Bypass','-File',(Join-Path $PSScriptRoot 'Test-LogQuality.ps1'),'-LogPath',$fixture,'-RequireSessionSummary')
    if($case.require){$arguments += '-RequireBuildIdentity'}
    $previousPreference=$ErrorActionPreference
    try { $ErrorActionPreference='Continue'; & powershell.exe @arguments *> (Join-Path $OutputDirectory ($case.name+'-validation.txt')); $code=$LASTEXITCODE }
    finally { $ErrorActionPreference=$previousPreference }
    if(($case.expected -eq 0) -ne ($code -eq 0)){throw "Log build-identity contract failed: $($case.name)"}
    $identityContracts += @{case=$case.name;expectedExit=$case.expected;actualExit=$code;passed=$true}
}
$identityContracts | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $OutputDirectory 'identity-contracts.json') -Encoding UTF8
