param(
    [ValidateSet('GpuOnly', 'ReadmeOnly')][string]$Mode,
    [string]$OutputDirectory,
    [string]$StatusPath
)

$ErrorActionPreference = 'Stop'
$captureScript = Join-Path $PSScriptRoot 'Capture-Showcase.ps1'
try {
    if ($Mode -eq 'GpuOnly') {
        & $captureScript -GpuOnly -AllowForeground -OutputDirectory $OutputDirectory *> $StatusPath
    } else {
        & $captureScript -ReadmeOnly -AllowForeground -OutputDirectory $OutputDirectory *> $StatusPath
    }
    exit 0
} catch {
    $_ | Out-String | Add-Content -LiteralPath $StatusPath
    exit 1
}
