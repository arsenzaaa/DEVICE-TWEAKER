param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release",

    [ValidateSet("x64")]
    [string]$Platform = "x64",

    [string]$PlatformToolset = "v145",

    [switch]$CopyToLoader
)

$ErrorActionPreference = "Stop"
if ($CopyToLoader) { throw "The release loader pair is pinned. Build outputs are kept under bin/KduBuild; update the payload manifest and validation explicitly before replacing it." }

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Resolve-Path (Join-Path $root "..\..")
$upstream = Get-Content -LiteralPath (Join-Path $root 'upstream.json') -Raw | ConvertFrom-Json
# Build a copy. Upstream contains a prebuilt post-build helper; never replace vendor inputs in place.
$buildRoot = Join-Path $projectRoot ('bin\KduBuild\' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '_' + [Guid]::NewGuid().ToString('N').Substring(0,8))
foreach ($file in $upstream.files) {
    $original = Join-Path $root $file.path
    if ((Get-FileHash -LiteralPath $original -Algorithm SHA256).Hash -ne $file.sha256) { throw "Changed upstream input: $($file.path)" }
    $copy = Join-Path $buildRoot $file.path
    New-Item -ItemType Directory -Force -Path (Split-Path $copy) | Out-Null
    Copy-Item -LiteralPath $original -Destination $copy
}
$sourceRoot = Join-Path $buildRoot "Source"
$solutionDir = $sourceRoot

function Find-MSBuild {
    $candidates = @(
        "${env:ProgramFiles}\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles}\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles}\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019\Professional\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019\Enterprise\MSBuild\Current\Bin\MSBuild.exe"
    )

    foreach ($candidate in $candidates) {
        if ($candidate -and (Test-Path -LiteralPath $candidate)) {
            return $candidate
        }
    }

    $vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
    if (Test-Path -LiteralPath $vswhere) {
        $path = & $vswhere -latest -requires Microsoft.Component.MSBuild -find "MSBuild\Current\Bin\MSBuild.exe" | Select-Object -First 1
        if ($path -and (Test-Path -LiteralPath $path)) {
            return $path
        }
    }

    throw "MSBuild.exe was not found. Install Visual Studio C++ build tools."
}

function Invoke-MSBuildProject {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProjectPath
    )

    Write-Host "Building $ProjectPath"
    & $script:MSBuild $ProjectPath /m /t:Build /p:Configuration=$Configuration /p:Platform=$Platform /p:PlatformToolset=$PlatformToolset /p:SolutionDir="$solutionDir" /nologo
    if ($LASTEXITCODE -ne 0) {
        throw "MSBuild failed for $ProjectPath with exit code $LASTEXITCODE"
    }
}

$script:MSBuild = Find-MSBuild
Write-Host "MSBuild: $script:MSBuild"
Write-Host "Configuration: $Configuration|$Platform"
Write-Host "PlatformToolset: $PlatformToolset"

$genProject = Join-Path $sourceRoot "Utils\GenAsIo2Unlock\GenAsIo2Unlock.vcxproj"
$tanikazeProject = Join-Path $sourceRoot "Tanikaze\Tanikaze.vcxproj"
$hamakazeProject = Join-Path $sourceRoot "Hamakaze\KDU.vcxproj"

Invoke-MSBuildProject $genProject

$genExe = Join-Path $sourceRoot "Utils\GenAsIo2Unlock\output\$Platform\$Configuration\GenAsIo2Unlock.exe"
if (-not (Test-Path -LiteralPath $genExe)) {
    throw "GenAsIo2Unlock.exe was not produced: $genExe"
}

$hamakazeUtils = Join-Path $sourceRoot "Hamakaze\Utils"
New-Item -ItemType Directory -Force -Path $hamakazeUtils | Out-Null
Copy-Item -LiteralPath $genExe -Destination (Join-Path $hamakazeUtils "GenAsIo2Unlock.exe") -Force

Invoke-MSBuildProject $tanikazeProject
Invoke-MSBuildProject $hamakazeProject

$builtKdu = Join-Path $sourceRoot "Hamakaze\output\$Platform\$Configuration\kdu.exe"
$builtDrv = Join-Path $sourceRoot "Tanikaze\output\$Platform\$Configuration\drv64.dll"

if (-not (Test-Path -LiteralPath $builtKdu)) {
    throw "kdu.exe was not produced: $builtKdu"
}
if (-not (Test-Path -LiteralPath $builtDrv)) {
    throw "drv64.dll was not produced: $builtDrv"
}

Write-Host "Built KDU:"
Get-Item -LiteralPath $builtKdu, $builtDrv | Select-Object FullName, Length, LastWriteTime | Format-Table -AutoSize
Get-FileHash -Algorithm SHA256 $builtKdu, $builtDrv | Format-Table -AutoSize

if ($CopyToLoader) {
    if ($Configuration -ne 'Release') { throw 'Only Release KDU may enter the application payload' }
    $loaderDir = Join-Path $projectRoot "IMOD\Loader"
    if (-not (Test-Path -LiteralPath $loaderDir)) {
        throw "Loader directory does not exist: $loaderDir"
    }

    Copy-Item -LiteralPath $builtKdu -Destination (Join-Path $loaderDir "kdu.exe") -Force
    Copy-Item -LiteralPath $builtDrv -Destination (Join-Path $loaderDir "drv64.dll") -Force
    $binaryRecords = @($builtKdu, $builtDrv | ForEach-Object {
        [ordered]@{ name = [IO.Path]::GetFileName($_); sha256 = (Get-FileHash -LiteralPath $_ -Algorithm SHA256).Hash;
            version = [Diagnostics.FileVersionInfo]::GetVersionInfo($_).FileVersion }
    })
    [ordered]@{ schemaVersion = 1; tag = $upstream.tag; commit = $upstream.commit; configuration = $Configuration;
        toolset = $PlatformToolset; sourceManifestSha256 = (Get-FileHash -LiteralPath (Join-Path $root 'upstream.json')).Hash;
        buildScriptSha256 = (Get-FileHash -LiteralPath $PSCommandPath).Hash; binaries = $binaryRecords } |
        ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $loaderDir 'KDU.build.json') -Encoding UTF8
    Write-Host "Copied rebuilt KDU binaries to $loaderDir"
}
