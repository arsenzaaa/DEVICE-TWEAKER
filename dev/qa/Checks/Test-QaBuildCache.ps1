[CmdletBinding()]
param([string]$OutputDirectory)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\Support\Build-Cache.ps1')
$workspace = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..\..')).Path
if (-not $OutputDirectory) { $OutputDirectory = Join-Path $workspace ('local\qa-cache-contracts\test_' + [Guid]::NewGuid().ToString('N').Substring(0,8)) }
if (Test-Path -LiteralPath $OutputDirectory) { throw 'Use a new output directory for cache contract fixtures' }
New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
$results = New-Object 'System.Collections.Generic.List[object]'
function Confirm([string]$Name, [bool]$Condition) {
    if (-not $Condition) { throw "Cache contract failed: $Name" }
    $results.Add([pscustomobject]@{ name = $Name; status = 'PASS' })
}
function Reject([string]$Name, [scriptblock]$Action, [string]$Pattern) {
    $message = $null
    try { & $Action | Out-Null } catch { $message = $_.Exception.Message }
    Confirm $Name ($message -and $message -match $Pattern)
}
try {
    $a = [pscustomobject]@{ path = 'src\a.cs'; sha256 = ('A' * 64) }
    $b = [pscustomobject]@{ path = 'dev\qa\b.ps1'; sha256 = ('B' * 64) }
    $key = Get-QaBuildKey @($a,$b) '10.0.100'
    Confirm 'stable-order' ($key -eq (Get-QaBuildKey @($b,$a) '10.0.100'))
    Confirm 'sdk-change-invalidates' ($key -ne (Get-QaBuildKey @($a,$b) '10.0.101'))
    Confirm 'removed-file-invalidates' ($key -ne (Get-QaBuildKey @($a) '10.0.100'))
    Confirm 'added-file-invalidates' ($key -ne (Get-QaBuildKey @($a,$b,[pscustomobject]@{path='src\new.cs';sha256=('C'*64)}) '10.0.100'))
    Confirm 'content-change-invalidates' ($key -ne (Get-QaBuildKey @($a,[pscustomobject]@{path=$b.path;sha256=('C'*64)}) '10.0.100'))
    $entry = Join-Path $OutputDirectory $key
    New-Item -ItemType Directory -Path $entry | Out-Null
    $interruptedLog = Join-Path $entry 'build.log'
    [IO.File]::WriteAllText($interruptedLog, 'interrupted compilation')
    $retry = Get-QaBuildDirectory $OutputDirectory $key
    Confirm 'interrupted-cache-does-not-block-next-build' ($retry -ne $entry -and -not (Test-Path -LiteralPath $retry))
    Confirm 'interrupted-output-preserved' ([IO.File]::ReadAllText($interruptedLog) -eq 'interrupted compilation')
    New-Item -ItemType Directory -Path $retry | Out-Null
    [IO.File]::WriteAllText((Join-Path $retry 'build.json'), '{}')
    Confirm 'finished-retry-is-selected-for-validation' ((Get-QaBuildDirectory $OutputDirectory $key) -eq $retry)
    $fixture = Join-Path $OutputDirectory 'integrity-fixture'
    $snapshot = New-QaInputSnapshot $workspace $fixture
    $inputs = Get-Content -LiteralPath (Join-Path $fixture 'inputs.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $key = Get-QaBuildKey $inputs.files 'contract-fixture'
    # These files test inventory integrity only. They are never executed.
    $exe = Join-Path $snapshot 'src\bin\DeveloperQa\Release\net10.0-windows\win-x64\DEVICE TWEAKER QA.exe'
    $dll = Join-Path $snapshot 'dev\qa\Inspect\bin\Release\net10.0-windows\Inspect.dll'
    foreach ($path in @($exe,$dll)) {
        New-Item -ItemType Directory -Force -Path (Split-Path $path) | Out-Null
        [IO.File]::WriteAllText($path, 'Non-executable cache integrity fixture')
    }
    $manifest = [ordered]@{ schemaVersion=1; key=$key; sdkVersion='contract-fixture';
        inputsSha256=(Get-FileHash -LiteralPath (Join-Path $fixture 'inputs.json')).Hash;
        outputs=@(Get-QaBuildOutputInventory $fixture) }
    $manifest | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $fixture 'build.json') -Encoding UTF8
    $null = Test-QaDevBuild $fixture $key
    Confirm 'intact-inventory-accepted' $true
    Reject 'wrong-build-identity-rejected' { Test-QaDevBuild $fixture 'wrong-key' } 'Incompatible QA build manifest'
    $original = [IO.File]::ReadAllBytes($dll)
    try {
        [IO.File]::AppendAllText($dll, 'changed')
        Reject 'changed-binary-rejected' { Test-QaDevBuild $fixture $key } 'QA build outputs changed'
    } finally { [IO.File]::WriteAllBytes($dll, $original) }
    $injected = Join-Path (Split-Path $dll) 'unexpected.dll'
    [IO.File]::WriteAllText($injected,'unexpected dependency')
    try { Reject 'added-dependency-rejected' { Test-QaDevBuild $fixture $key } 'QA build output inventory changed' }
    finally { Move-Item -LiteralPath $injected -Destination (Join-Path $OutputDirectory 'quarantined-unexpected.dll') }
    $holding = Join-Path $OutputDirectory 'held-inspector.dll'
    Move-Item -LiteralPath $dll -Destination $holding
    try { Reject 'missing-binary-rejected' { Test-QaDevBuild $fixture $key } 'QA build output inventory changed' }
    finally { Move-Item -LiteralPath $holding -Destination $dll }
    $source = Join-Path $snapshot 'src\DeviceTweakerCS.csproj'
    $original = [IO.File]::ReadAllBytes($source)
    try {
        [IO.File]::AppendAllText($source,'changed')
        Reject 'changed-source-rejected' { Test-QaDevBuild $fixture $key } 'QA input was changed'
    } finally { [IO.File]::WriteAllBytes($source, $original) }
    $rootConfig = Join-Path $snapshot 'Directory.Build.props'
    [IO.File]::WriteAllText($rootConfig, '<Project />')
    try { Reject 'added-root-build-config-rejected' { Test-QaDevBuild $fixture $key } 'QA input file inventory changed' }
    finally { Move-Item -LiteralPath $rootConfig -Destination (Join-Path $OutputDirectory 'held-Directory.Build.props') }
    $inputManifest = Join-Path $fixture 'inputs.json'
    $original = [IO.File]::ReadAllBytes($inputManifest)
    try {
        [IO.File]::AppendAllText($inputManifest,' ')
        Reject 'changed-input-manifest-rejected' { Test-QaDevBuild $fixture $key } 'QA build input manifest changed'
    } finally { [IO.File]::WriteAllBytes($inputManifest,$original) }
    $null = Test-QaDevBuild $fixture $key
    Confirm 'restored-inventory-accepted' $true
    [ordered]@{ status='PASS'; cases=$results.ToArray(); scope='Cache identity and tamper rejection; fixture binaries were not executed' } |
        ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'contracts.json') -Encoding UTF8
    Write-Output (Join-Path $OutputDirectory 'contracts.json')
} catch {
    [ordered]@{ status='FAIL'; cases=$results.ToArray(); error=$_.Exception.ToString() } |
        ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'contracts.json') -Encoding UTF8
    throw
}
