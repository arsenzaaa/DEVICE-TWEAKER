param(
    [ValidateSet('GpuOnly', 'ReadmeOnly')][string]$Mode,
    [string]$OutputDirectory,
    [string]$StatusPath
)

$ErrorActionPreference = 'Stop'
$captureScript = Join-Path $PSScriptRoot 'Capture-Showcase.ps1'
try {
    if ($Mode -eq 'GpuOnly') {
        & $captureScript -GpuOnly -OutputDirectory $OutputDirectory *> $StatusPath
    } else {
        & $captureScript -ReadmeOnly -OutputDirectory $OutputDirectory *> $StatusPath
    }
    if (-not $?) { exit 1 }
    exit $LASTEXITCODE
} catch {
    $_ | Out-String | Add-Content -LiteralPath $StatusPath
    exit 1
}
