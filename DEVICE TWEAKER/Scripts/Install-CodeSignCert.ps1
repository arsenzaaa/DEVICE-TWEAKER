param(
    [Parameter(Mandatory = $true)]
    [string]$CerPath
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

if (-not (Test-Path -LiteralPath $CerPath)) {
    throw "Certificate file not found: $CerPath"
}

$resolvedCerPath = (Resolve-Path -LiteralPath $CerPath).Path
$certutil = Get-Command certutil.exe -ErrorAction SilentlyContinue
if ($null -eq $certutil -or [string]::IsNullOrWhiteSpace($certutil.Source)) {
    throw "certutil.exe not found in PATH."
}

function Test-CertificateInStore {
    param(
        [Parameter(Mandatory = $true)]
        [string]$StoreName,
        [Parameter(Mandatory = $true)]
        [string]$Thumbprint
    )

    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store($StoreName, "CurrentUser")
    try {
        $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)
        $found = $store.Certificates.Find(
            [System.Security.Cryptography.X509Certificates.X509FindType]::FindByThumbprint,
            $Thumbprint,
            $false)
        return ($null -ne $found -and $found.Count -gt 0)
    }
    finally {
        $store.Close()
    }
}

function Add-CertificateToStore {
    param(
        [Parameter(Mandatory = $true)]
        [string]$StoreName,
        [Parameter(Mandatory = $true)]
        [string]$Thumbprint
    )

    if (Test-CertificateInStore -StoreName $StoreName -Thumbprint $Thumbprint) {
        Write-Host "Certificate already present in '$StoreName'. Skipping install."
        return
    }

    & $certutil.Source -user -f -addstore $StoreName $resolvedCerPath | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw "certutil failed while adding certificate to store '$StoreName' (exit code $LASTEXITCODE)."
    }
}

$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($resolvedCerPath)
$thumbprint = $cert.Thumbprint
if ([string]::IsNullOrWhiteSpace($thumbprint)) {
    throw "Unable to read certificate thumbprint from: $resolvedCerPath"
}

Add-CertificateToStore -StoreName "TrustedPublisher" -Thumbprint $thumbprint
Add-CertificateToStore -StoreName "Root" -Thumbprint $thumbprint
