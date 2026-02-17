param(
    [ValidateSet("both", "with-net", "without-net")]
    [string]$Flavor = "both",
    [ValidateSet("Release", "Debug")]
    [string]$Configuration = "Release",
    [string]$Runtime = "win-x64",
    [switch]$Clean
)

$ErrorActionPreference = "Stop"

$projectPath = Join-Path $PSScriptRoot "DeviceTweakerCS.csproj"
$publishRoot = Join-Path $PSScriptRoot "bin\\Publish"
$withNetOut = Join-Path $publishRoot "self-contained"
$withoutNetOut = Join-Path $publishRoot "framework-dependent"

if (-not (Test-Path -LiteralPath $projectPath)) {
    throw "Project file not found: $projectPath"
}

if (-not (Get-Command dotnet -ErrorAction SilentlyContinue)) {
    throw "dotnet CLI is not available in PATH."
}

function Invoke-PublishFlavor {
    param(
        [string]$Label,
        [bool]$SelfContained,
        [bool]$PublishSingleFile,
        [bool]$IncludeNativeLibrariesForSelfExtract,
        [string]$OutputDir
    )

    if ($Clean -and (Test-Path -LiteralPath $OutputDir)) {
        Remove-Item -LiteralPath $OutputDir -Recurse -Force
    }

    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null

    Write-Host ""
    Write-Host "=== Publishing: $Label ==="
    Write-Host "Output: $OutputDir"

    $selfContainedText = $SelfContained.ToString().ToLowerInvariant()
    $singleFileText = $PublishSingleFile.ToString().ToLowerInvariant()
    $nativeExtractText = $IncludeNativeLibrariesForSelfExtract.ToString().ToLowerInvariant()

    dotnet publish $projectPath `
        -c $Configuration `
        -r $Runtime `
        -p:SelfContained=$selfContainedText `
        -p:PublishSingleFile=$singleFileText `
        -p:IncludeNativeLibrariesForSelfExtract=$nativeExtractText `
        -p:DebugType=None `
        -p:DebugSymbols=false `
        -p:UseAppHost=true `
        -o $OutputDir
}

Write-Host "Restoring project..."
dotnet restore $projectPath

switch ($Flavor) {
    "both" {
        Invoke-PublishFlavor -Label "with .NET (self-contained)" -SelfContained $true -PublishSingleFile $true -IncludeNativeLibrariesForSelfExtract $true -OutputDir $withNetOut
        Invoke-PublishFlavor -Label "without .NET (framework-dependent)" -SelfContained $false -PublishSingleFile $true -IncludeNativeLibrariesForSelfExtract $false -OutputDir $withoutNetOut
    }
    "with-net" {
        Invoke-PublishFlavor -Label "with .NET (self-contained)" -SelfContained $true -PublishSingleFile $true -IncludeNativeLibrariesForSelfExtract $true -OutputDir $withNetOut
    }
    "without-net" {
        Invoke-PublishFlavor -Label "without .NET (framework-dependent)" -SelfContained $false -PublishSingleFile $true -IncludeNativeLibrariesForSelfExtract $false -OutputDir $withoutNetOut
    }
}

Write-Host ""
Write-Host "Done."
Write-Host "with .NET:    $withNetOut"
Write-Host "without .NET: $withoutNetOut"
