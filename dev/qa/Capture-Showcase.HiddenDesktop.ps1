param(
    [ValidateSet('GpuOnly', 'ReadmeOnly')][string]$Mode = 'GpuOnly',
    [string]$OutputDirectory = (Join-Path $env:TEMP 'DeviceTweakerShowcaseHidden')
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Private-Desktop.ps1')
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
$parameters = @{ OutputDirectory = $OutputDirectory }
$parameters[$Mode] = $true
$script = Join-Path $PSScriptRoot 'Capture-Showcase.ps1'
$code = Invoke-QaPrivateDesktop $script $parameters (Join-Path $OutputDirectory 'isolation')
if ($null -ne $code) { exit $code }
& $script @parameters