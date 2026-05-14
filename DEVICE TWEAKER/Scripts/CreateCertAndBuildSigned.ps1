param(
    [ValidateSet("both", "with-net", "without-net")]
    [string]$Flavor = "both",
    [ValidateSet("Release", "Debug")]
    [string]$Configuration = "Release",
    [switch]$NoClean,
    [switch]$SkipImodDriverBuild,
    [switch]$InstallToTrustedStores,
    [string]$MsBuildPath,
    [string]$Subject = "CN=MADE BY ARSENZA",
    [string]$OutputDir = $(Join-Path (Split-Path -Parent $PSScriptRoot) "build\codesign"),
    [int]$Years = 3,
    [string]$PfxPassword = ""
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$createScript = Join-Path $PSScriptRoot "Create-DevCodeSignCert.ps1"
$buildScript = Join-Path (Split-Path -Parent $PSScriptRoot) "build.ps1"

if (-not (Test-Path -LiteralPath $createScript)) {
    throw "Script not found: $createScript"
}
if (-not (Test-Path -LiteralPath $buildScript)) {
    throw "Script not found: $buildScript"
}

$certResult = & $createScript `
    -Subject $Subject `
    -OutputDir $OutputDir `
    -PfxPassword $PfxPassword `
    -Years $Years `
    -InstallToTrustedStores:$InstallToTrustedStores

if ($null -eq $certResult -or [string]::IsNullOrWhiteSpace($certResult.Thumbprint)) {
    throw "Certificate creation did not return a thumbprint."
}

Write-Host ""
Write-Host "Starting signed build..."
Write-Host "Subject:    $($certResult.Subject)"
Write-Host "Thumbprint: $($certResult.Thumbprint)"

$buildArgs = @(
    "-Flavor", $Flavor,
    "-Configuration", $Configuration,
    "-TrustImodDriverCert",
    "-ImodDriverCertThumbprint", $certResult.Thumbprint,
    "-ImodDriverCertSubject", $certResult.Subject
)

if ($NoClean) {
    $buildArgs += "-NoClean"
}

if ($SkipImodDriverBuild) {
    $buildArgs += "-SkipImodDriverBuild"
}

if (-not [string]::IsNullOrWhiteSpace($MsBuildPath)) {
    $buildArgs += @("-MsBuildPath", $MsBuildPath)
}

& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $buildScript @buildArgs
if ($LASTEXITCODE -ne 0) {
    throw "build.ps1 failed with exit code $LASTEXITCODE"
}
