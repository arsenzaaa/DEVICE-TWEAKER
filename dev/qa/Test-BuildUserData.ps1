# Exercises the actual build wrapper with owned fixture folders and a fake publisher.
# Does not compile, sign, start the application or touch real logs/backups.
[CmdletBinding()]
param([string]$OutputDirectory, [string]$SourceScript)
$ErrorActionPreference = 'Stop'
$workspace = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..')).Path
if (-not $SourceScript) { $SourceScript = Join-Path $workspace 'src\build.ps1' }
$SourceScript = (Resolve-Path -LiteralPath $SourceScript).Path
if (-not $OutputDirectory) { $OutputDirectory = Join-Path $workspace 'local\build-user-data' }
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$run = Join-Path $OutputDirectory ('run_' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $run | Out-Null
$harness = @'
$ErrorActionPreference = 'Stop'
$fixture = $env:DEVICE_TWEAKER_BUILD_FIXTURE
$mode = $env:DEVICE_TWEAKER_BUILD_FAULT
function Move-Item {
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline=$true)]$InputObject, [string]$LiteralPath,
          [Parameter(Mandatory=$true)][string]$Destination, [switch]$Force)
    process {
        $blocked = if ($mode -eq 'prepare-second-folder') {
            Join-Path $fixture 'bin\Publish\DEVICE TWEAKER (SELF-CONTAINED)\logs'
        } elseif ($mode -eq 'prepare-next-flavor') {
            Join-Path $fixture 'bin\Publish\DEVICE TWEAKER\Backups'
        } else { '' }
        if ($LiteralPath -and $LiteralPath -eq $blocked) { throw 'Injected preparation move failure' }
        if ($LiteralPath) { Microsoft.PowerShell.Management\Move-Item -LiteralPath $LiteralPath -Destination $Destination -Force:$Force }
        else { $InputObject | Microsoft.PowerShell.Management\Move-Item -Destination $Destination -Force:$Force }
    }
}
function powershell.exe {
    param([Parameter(ValueFromRemainingArguments=$true)]$Arguments)
    [IO.File]::WriteAllText((Join-Path $fixture 'publisher-called.txt'), 'fake publisher')
    if ($mode -eq 'publish-fails') { $global:LASTEXITCODE=17; return }
    foreach ($relative in @('bin\Publish\DEVICE TWEAKER\DEVICE TWEAKER.exe',
        'bin\Publish\DEVICE TWEAKER (SELF-CONTAINED)\DEVICE TWEAKER (SELF-CONTAINED).exe', 'IMOD\DTIMOD.sys')) {
        $path=Join-Path $fixture $relative
        New-Item -ItemType Directory -Force -Path (Split-Path -Parent $path) | Out-Null
        [IO.File]::WriteAllText($path, 'fixture only - never executed')
    }
    $global:LASTEXITCODE=0
}
function Get-AuthenticodeSignature { param([string]$LiteralPath) [pscustomobject]@{ Status='NotSigned'; SignerCertificate=$null } }
& (Join-Path $fixture 'build.ps1') -Flavor both -SkipImodDriverBuild
'@
$results = @()
foreach ($mode in @('prepare-second-folder','prepare-next-flavor','publish-fails','publish-success')) {
    $fixture = Join-Path $run $mode
    New-Item -ItemType Directory -Path $fixture | Out-Null
    Copy-Item -LiteralPath $SourceScript -Destination (Join-Path $fixture 'build.ps1')
    [IO.File]::WriteAllText((Join-Path $fixture 'publish-variants.ps1'), '# Not executed by this test')
    $expected = @()
    foreach ($flavor in @('DEVICE TWEAKER (SELF-CONTAINED)','DEVICE TWEAKER')) {
        foreach ($name in @('Backups','logs')) {
            $directory = Join-Path $fixture ('bin\Publish\' + $flavor + '\' + $name)
            New-Item -ItemType Directory -Path $directory -Force | Out-Null
            $path=Join-Path $directory 'preserve.txt'
            [IO.File]::WriteAllText($path, $flavor + '/' + $name)
            $expected += [pscustomobject]@{ path=$path; hash=(Get-FileHash -LiteralPath $path).Hash }
        }
    }
    $scriptPath=Join-Path $fixture 'test-harness.ps1'
    [IO.File]::WriteAllText($scriptPath, $harness, [Text.UTF8Encoding]::new($true))
    $savedFixture=$env:DEVICE_TWEAKER_BUILD_FIXTURE; $savedFault=$env:DEVICE_TWEAKER_BUILD_FAULT
    $env:DEVICE_TWEAKER_BUILD_FIXTURE=$fixture; $env:DEVICE_TWEAKER_BUILD_FAULT=$mode
    try {
        $preference=$ErrorActionPreference
        try {
            $ErrorActionPreference='Continue'
            & "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File $scriptPath *> (Join-Path $fixture 'execution.log')
            $exitCode=$LASTEXITCODE
        } finally { $ErrorActionPreference=$preference }
    } finally { $env:DEVICE_TWEAKER_BUILD_FIXTURE=$savedFixture; $env:DEVICE_TWEAKER_BUILD_FAULT=$savedFault }
    $missing=@($expected | Where-Object { -not (Test-Path -LiteralPath $_.path) -or (Get-FileHash -LiteralPath $_.path).Hash -ne $_.hash })
    $published=Test-Path -LiteralPath (Join-Path $fixture 'publisher-called.txt')
    $leftovers=@(Get-ChildItem -LiteralPath (Join-Path $fixture 'bin') -Directory -Filter '.kept-user-data-*')
    $passed=$missing.Count -eq 0 -and $leftovers.Count -eq 0 -and
        (($mode -eq 'publish-success' -and $exitCode -eq 0) -or ($mode -ne 'publish-success' -and $exitCode -ne 0)) -and
        ($published -eq (-not $mode.StartsWith('prepare-')))
    $results += [pscustomobject]@{ mode=$mode; passed=$passed; exitCode=$exitCode; publisherCalled=$published;
        missing=@($missing | ForEach-Object { $_.path }); keptDirectories=@($leftovers | ForEach-Object { $_.FullName }); fixture=$fixture }
}
$report=Join-Path $run 'contracts.json'
$results | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $report -Encoding UTF8
if (@($results | Where-Object { -not $_.passed }).Count) { throw "Build user-data preservation failed; see $report" }
Write-Output ('PASS: four build user-data preservation cases; source=' + (Get-FileHash -LiteralPath $SourceScript).Hash)
Write-Output $report
