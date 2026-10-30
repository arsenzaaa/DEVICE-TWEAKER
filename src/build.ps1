param(
    [ValidateSet("both", "with-net", "without-net")]
    [string]$Flavor = "both",
    [ValidateSet("Release", "Debug")]
    [string]$Configuration = "Release",
    [switch]$NoClean,
    [switch]$SkipImodDriverBuild,
    [switch]$TrustImodDriverCert,
    [string]$MsBuildPath,
    [string]$ImodDriverCertThumbprint = "B263CA124676681D028FFD7A3D605F5733E991B5",
    [string]$ImodDriverCertSubject = "MADE BY ARSENZA"
)

$ErrorActionPreference = "Stop"

$root = $PSScriptRoot
$publishScript = Join-Path $root "publish-variants.ps1"
$withNetExe = Join-Path $root "bin\Publish\DEVICE TWEAKER (SELF-CONTAINED)\DEVICE TWEAKER (SELF-CONTAINED).exe"
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

# A clean publish deletes the output folders. Backups and logs written next to the EXE
# are the user's restore points and history, so they are moved aside and put back.
$keptUserData = @()
$keepRoot = Join-Path $root ("bin\.kept-user-data-" + [Guid]::NewGuid().ToString("N"))
if (-not $NoClean) {
    foreach ($exe in @($withNetExe, $withoutNetExe)) {
        foreach ($name in @("Backups", "logs")) {
            $source = Join-Path (Split-Path -Parent $exe) $name
            if (Test-Path -LiteralPath $source -PathType Container) {
                $kept = Join-Path $keepRoot ((Split-Path -Leaf (Split-Path -Parent $exe)) + "\" + $name)
                New-Item -ItemType Directory -Path (Split-Path -Parent $kept) -Force | Out-Null
                Move-Item -LiteralPath $source -Destination $kept
                $keptUserData += [pscustomobject]@{ Source = $source; Kept = $kept }
            }
        }
    }
}

try {
    & powershell.exe @argsList
    $publishExitCode = $LASTEXITCODE
} finally {
    foreach ($item in $keptUserData) {
        New-Item -ItemType Directory -Path (Split-Path -Parent $item.Source) -Force | Out-Null
        if (Test-Path -LiteralPath $item.Source) {
            # The new build created the folder; merge the kept files into it.
            Get-ChildItem -LiteralPath $item.Kept -Force | Move-Item -Destination $item.Source -Force
            Remove-Item -LiteralPath $item.Kept -Recurse -Force
        } else {
            Move-Item -LiteralPath $item.Kept -Destination $item.Source
        }
        Write-Host "Kept: $($item.Source)"
    }
    if (Test-Path -LiteralPath $keepRoot) {
        Remove-Item -LiteralPath $keepRoot -Recurse -Force
    }
}
if ($publishExitCode -ne 0) {
    throw "publish-variants.ps1 failed with exit code $publishExitCode"
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
