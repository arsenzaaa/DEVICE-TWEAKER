param([Parameter(Mandatory = $true)][string]$PackageDirectory, [string]$OutputDirectory)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Read-QaJson.ps1')
$package = (Resolve-Path -LiteralPath $PackageDirectory).Path
if (-not $OutputDirectory) { $OutputDirectory = Join-Path $package '..\validation' }
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$results = @()
$expectedNames = @('DEVICE.TWEAKER.exe', 'DEVICE.TWEAKER.SelfContained.exe')
$manifestLines = @(Get-Content -LiteralPath (Join-Path $package 'SHA256SUMS.txt'))
$manifestNames = @($manifestLines | ForEach-Object {
    if ($_ -notmatch '^([0-9A-Fa-f]{64})  (.+)$') { throw 'Invalid checksum manifest line' }
    $Matches[2]
})
if ($manifestNames.Count -ne 2 -or @($manifestNames | Select-Object -Unique).Count -ne 2 -or
    @(Compare-Object $expectedNames $manifestNames).Count) { throw 'Expected exactly one checksum for each production EXE' }
$names = @('DEVICE_TWEAKER_QA_READ_ONLY', 'DEVICE_TWEAKER_QA_OUTPUT', 'DEVICE_TWEAKER_QA_SESSION')
$saved = @{}
foreach ($name in $names) { $saved[$name] = [Environment]::GetEnvironmentVariable($name) }
try {
    foreach ($line in $manifestLines) {
        if ($line -notmatch '^([0-9A-Fa-f]{64})  (.+)$') { throw 'Invalid checksum manifest line' }
        $expected = $Matches[1]; $fileName = $Matches[2]
        $exe = Join-Path $package $fileName
        $hash = (Get-FileHash -LiteralPath $exe -Algorithm SHA256).Hash
        if ($hash -ne $expected) { throw "Package checksum mismatch: $fileName" }
        $info = [Diagnostics.FileVersionInfo]::GetVersionInfo($exe)
        if ($info.ProductVersion -ne '0.0.4' -or $info.FileVersion -ne '0.0.4.0') { throw "Wrong release version: $fileName" }
        $env:DEVICE_TWEAKER_QA_READ_ONLY = '1'
        $env:DEVICE_TWEAKER_QA_SESSION = [Guid]::NewGuid().ToString('N')
        $env:DEVICE_TWEAKER_QA_OUTPUT = Join-Path $OutputDirectory ([IO.Path]::GetFileNameWithoutExtension($fileName))
        $process = Start-Process -FilePath $exe -ArgumentList @('--qa-read-only', '--qa-probe') -WindowStyle Hidden -PassThru
        try {
            if (-not $process.WaitForExit(20000)) { throw "Packaged EXE probe timed out: $fileName" }
            if ($process.ExitCode -ne 0) { throw "Packaged EXE probe failed: $fileName" }
            $report = Read-QaJson (Join-Path $env:DEVICE_TWEAKER_QA_OUTPUT "process-$($process.Id).json")
            if ($report.readOnly -ne $true -or $report.build.version -ne '0.0.4' -or $report.build.hasTestPanel -ne $false) {
                throw "Packaged EXE is not the verified production build: $fileName"
            }
            if ($report.build.appBuildId -notmatch '^[0-9a-f]{32}$' -or $report.build.executableSha256 -ne $hash) {
                throw "Packaged EXE build identity does not match its actual file: $fileName"
            }
            $results += [pscustomobject]@{ file = $fileName; sha256 = $hash; version = $info.ProductVersion;
                bytes = (Get-Item -LiteralPath $exe).Length; probe = 'PASS'; hasTestPanel = $report.build.hasTestPanel;
                driverSha256 = $report.build.driverSha256; resources = $report.build.embeddedResources;
                appBuildId = $report.build.appBuildId; executableSha256 = $report.build.executableSha256 }
        } finally {
            if (-not $process.HasExited) { $process.Kill(); $process.WaitForExit(5000) | Out-Null }
            $process.Dispose()
        }
    }
    if ($results.Count -ne 2 -or $results[0].driverSha256 -ne $results[1].driverSha256) { throw 'Expected two production EXEs with identical driver payloads' }
    $sourceDriver = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..\DEVICE TWEAKER\IMOD\DTIMOD.sys')).Path
    if ($results[0].driverSha256 -ne (Get-FileHash -LiteralPath $sourceDriver -Algorithm SHA256).Hash) { throw 'Embedded driver differs from the source build artifact' }
    foreach ($file in $results) {
        foreach ($name in @('DeviceTweakerCS.IMOD.DTIMOD.sys', 'DeviceTweakerCS.IMOD.DTIMOD.sys.sha256', 'DeviceTweakerCS.IMOD.Loader.kdu.exe', 'DeviceTweakerCS.IMOD.Loader.drv64.dll')) {
            if ($file.resources -notcontains $name) { throw "Embedded resource missing: $name" }
        }
    }
    [ordered]@{ status = 'PASS'; hardwareValidation = 'SKIPPED'; files = $results } |
        ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'release.json') -Encoding UTF8
    Write-Host 'PASS: both packaged 0.0.4 EXEs start in probe mode, contain the same driver, and contain no test panel.'
} finally {
    foreach ($name in $names) { [Environment]::SetEnvironmentVariable($name, $saved[$name]) }
}
