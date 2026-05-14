param(
    [ValidateSet("both", "with-net", "without-net")]
    [string]$Flavor = "both",
    [ValidateSet("Release", "Debug")]
    [string]$Configuration = "Release",
    [switch]$NoClean,
    [switch]$SkipImodDriverBuild,
    [switch]$TrustImodDriverCert,
    [string]$MsBuildPath,
    [string]$ImodDriverCertThumbprint = "9CE4C30CD75905786774B1DDFAC126329ACAEA8D",
    [string]$ImodDriverCertSubject = "MADE BY ARSENZA"
)

$ErrorActionPreference = "Stop"

$root = $PSScriptRoot
$publishScript = Join-Path $root "publish-variants.ps1"
$withNetExe = Join-Path $root "bin\Publish\DEVICE TWEAKER (NET FRAMEWORK)\DEVICE TWEAKER (NET FRAMEWORK).exe"
$withoutNetExe = Join-Path $root "bin\Publish\DEVICE TWEAKER\DEVICE TWEAKER.exe"
$driverPath = Join-Path $root "IMOD\DTIMOD.sys"

if (-not (Test-Path -LiteralPath $publishScript)) {
    throw "Build script was not found: $publishScript"
}

$argsList = @(
    "-NoProfile",
    "-ExecutionPolicy",
    "Bypass",
    "-File",
    $publishScript,
    "-Flavor",
    $Flavor,
    "-Configuration",
    $Configuration
)

if (-not $NoClean) {
    $argsList += "-Clean"
}

if ($SkipImodDriverBuild) {
    $argsList += "-SkipImodDriverBuild"
}

if ($TrustImodDriverCert) {
    $argsList += "-TrustImodDriverCert"
}

if (-not [string]::IsNullOrWhiteSpace($MsBuildPath)) {
    $argsList += @("-MsBuildPath", $MsBuildPath)
}

if (-not [string]::IsNullOrWhiteSpace($ImodDriverCertThumbprint)) {
    $argsList += @("-ImodDriverCertThumbprint", $ImodDriverCertThumbprint)
}

if (-not [string]::IsNullOrWhiteSpace($ImodDriverCertSubject)) {
    $argsList += @("-ImodDriverCertSubject", $ImodDriverCertSubject)
}

Write-Host "DEVICE TWEAKER build"
Write-Host "Flavor:        $Flavor"
Write-Host "Configuration: $Configuration"
Write-Host "Clean:         $(-not $NoClean)"
Write-Host "Driver cert:   $ImodDriverCertSubject"
Write-Host "Thumbprint:    $ImodDriverCertThumbprint"
Write-Host ""

& powershell.exe @argsList
if ($LASTEXITCODE -ne 0) {
    throw "publish-variants.ps1 failed with exit code $LASTEXITCODE"
}

$artifacts = @()
if ($Flavor -in @("both", "with-net")) {
    $artifacts += $withNetExe
}
if ($Flavor -in @("both", "without-net")) {
    $artifacts += $withoutNetExe
}
$artifacts += $driverPath

foreach ($artifact in $artifacts) {
    if (-not (Test-Path -LiteralPath $artifact)) {
        throw "Expected build artifact was not found: $artifact"
    }
}

Write-Host ""
Write-Host "Artifacts:"
foreach ($artifact in $artifacts) {
    $item = Get-Item -LiteralPath $artifact
    $signature = Get-AuthenticodeSignature -LiteralPath $artifact
    $subject = if ($signature.SignerCertificate) { $signature.SignerCertificate.Subject } else { "not signed" }
    Write-Host ("{0} | {1:N0} bytes | signature={2} | {3}" -f $item.FullName, $item.Length, $signature.Status, $subject)
}

Write-Host ""
Write-Host "Build completed."
