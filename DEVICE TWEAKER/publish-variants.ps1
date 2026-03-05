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
$publishRoot = Join-Path $PSScriptRoot "bin\Publish"
$selfContainedDisplayName = "DEVICE TWEAKER (NET FRAMEWORK)"
$frameworkDependentDisplayName = "DEVICE TWEAKER"
$withNetOut = Join-Path $publishRoot $selfContainedDisplayName
$withoutNetOut = Join-Path $publishRoot $frameworkDependentDisplayName

if (-not (Test-Path -LiteralPath $projectPath)) {
    throw "Project file not found: $projectPath"
}

if (-not (Get-Command dotnet -ErrorAction SilentlyContinue)) {
    throw "dotnet CLI is not available in PATH."
}

function Rename-PublishedExe {
    param(
        [string]$OutputDir,
        [string]$TargetName,
        [string]$SourceName = "DEVICE TWEAKER.exe"
    )

    $allExe = Get-ChildItem -LiteralPath $OutputDir -File -Filter *.exe

    $publishedExe = $allExe | Where-Object { $_.Name -ieq $SourceName } | Select-Object -First 1
    if (-not $publishedExe) {
        $publishedExe = $allExe |
            Where-Object { $_.Name -ine $TargetName } |
            Sort-Object Length -Descending |
            Select-Object -First 1
    }
    if (-not $publishedExe) {
        $publishedExe = $allExe | Where-Object { $_.Name -ieq $TargetName } | Select-Object -First 1
    }

    if (-not $publishedExe) {
        throw "No EXE found in publish output: $OutputDir"
    }

    if ($publishedExe.Name -ieq $TargetName) {
        return
    }

    $targetPath = Join-Path $OutputDir $TargetName
    if (Test-Path -LiteralPath $targetPath) {
        Remove-Item -LiteralPath $targetPath -Force
    }

    Rename-Item -LiteralPath $publishedExe.FullName -NewName $TargetName
    Write-Host "Renamed EXE: $TargetName"
}

function Invoke-PublishFlavor {
    param(
        [string]$Label,
        [bool]$SelfContained,
        [bool]$PublishSingleFile,
        [bool]$IncludeNativeLibrariesForSelfExtract,
        [string]$OutputDir,
        [string]$TargetExeName
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

    if ($LASTEXITCODE -ne 0) {
        throw "dotnet publish failed for flavor '$Label' (exit code $LASTEXITCODE)"
    }

    Rename-PublishedExe -OutputDir $OutputDir -TargetName $TargetExeName
}

Write-Host "Restoring project..."
dotnet restore $projectPath
if ($LASTEXITCODE -ne 0) {
    throw "dotnet restore failed (exit code $LASTEXITCODE)"
}

switch ($Flavor) {
    "both" {
        Invoke-PublishFlavor -Label "self-contained (runtime included)" -SelfContained $true -PublishSingleFile $true -IncludeNativeLibrariesForSelfExtract $true -OutputDir $withNetOut -TargetExeName "$selfContainedDisplayName.exe"
        Invoke-PublishFlavor -Label "framework-dependent (requires .NET runtime)" -SelfContained $false -PublishSingleFile $true -IncludeNativeLibrariesForSelfExtract $false -OutputDir $withoutNetOut -TargetExeName "$frameworkDependentDisplayName.exe"
    }
    "with-net" {
        Invoke-PublishFlavor -Label "self-contained (runtime included)" -SelfContained $true -PublishSingleFile $true -IncludeNativeLibrariesForSelfExtract $true -OutputDir $withNetOut -TargetExeName "$selfContainedDisplayName.exe"
    }
    "without-net" {
        Invoke-PublishFlavor -Label "framework-dependent (requires .NET runtime)" -SelfContained $false -PublishSingleFile $true -IncludeNativeLibrariesForSelfExtract $false -OutputDir $withoutNetOut -TargetExeName "$frameworkDependentDisplayName.exe"
    }
}

Write-Host ""
Write-Host "Done."
Write-Host "Integrated .NET build:      $withNetOut"
Write-Host "Requires installed .NET:    $withoutNetOut"
