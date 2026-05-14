param(
    [string]$Subject = "CN=MADE BY ARSENZA",
    [string]$OutputDir = $(Join-Path (Split-Path -Parent $PSScriptRoot) "build\codesign"),
    [string]$PfxPassword = "",
    [int]$Years = 3,
    [switch]$InstallToTrustedStores
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

if ($Years -lt 1 -or $Years -gt 10) {
    throw "Years must be between 1 and 10."
}

function New-RandomPassword {
    param([int]$Length = 24)

    $chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    $bytes = New-Object byte[] $Length
    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    try {
        $rng.GetBytes($bytes)
    }
    finally {
        $rng.Dispose()
    }

    $builder = New-Object System.Text.StringBuilder
    foreach ($b in $bytes) {
        [void]$builder.Append($chars[$b % $chars.Length])
    }

    return $builder.ToString()
}

if (-not (Test-Path -LiteralPath $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

if ([string]::IsNullOrWhiteSpace($PfxPassword)) {
    $PfxPassword = New-RandomPassword -Length 24
}

$notAfter = (Get-Date).AddYears($Years)
$cert = New-SelfSignedCertificate `
    -Type CodeSigningCert `
    -Subject $Subject `
    -KeyAlgorithm RSA `
    -KeyLength 3072 `
    -HashAlgorithm SHA256 `
    -NotAfter $notAfter `
    -KeyExportPolicy Exportable `
    -CertStoreLocation "Cert:\CurrentUser\My"

$safeSubject = ($Subject -replace "[^A-Za-z0-9._-]", "_").Trim("_")
if ([string]::IsNullOrWhiteSpace($safeSubject)) {
    $safeSubject = "device_tweaker_dev"
}

$baseName = "{0}_{1}" -f $safeSubject, $cert.Thumbprint
$cerPath = Join-Path $OutputDir ($baseName + ".cer")
$pfxPath = Join-Path $OutputDir ($baseName + ".pfx")

Export-Certificate -Cert $cert -FilePath $cerPath -Force | Out-Null

$securePassword = ConvertTo-SecureString -String $PfxPassword -AsPlainText -Force
Export-PfxCertificate `
    -Cert $cert `
    -FilePath $pfxPath `
    -Password $securePassword `
    -ChainOption BuildChain `
    -Force | Out-Null

if ($InstallToTrustedStores) {
    $installScript = Join-Path $PSScriptRoot "Install-CodeSignCert.ps1"
    if (-not (Test-Path -LiteralPath $installScript)) {
        throw "Install script not found: $installScript"
    }

    & $installScript -CerPath $cerPath
    if ($LASTEXITCODE -ne 0) {
        throw "Install-CodeSignCert.ps1 failed with exit code $LASTEXITCODE."
    }
}

Write-Host ""
Write-Host "Done. Dev code-sign certificate created."
Write-Host "Subject:   $Subject"
Write-Host "Thumbprint:$($cert.Thumbprint)"
Write-Host "CER: $cerPath"
Write-Host "PFX: $pfxPath"
Write-Host ""
Write-Host "PFX password: $PfxPassword"

[PSCustomObject]@{
    Subject = $Subject
    Thumbprint = $cert.Thumbprint
    CerPath = $cerPath
    PfxPath = $pfxPath
    PfxPassword = $PfxPassword
}
