[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$RunDirectory, [string]$ReviewFile, [switch]$CreateTemplate, [switch]$Worker)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Support\Private-Desktop.ps1')
$lease = if (-not $Worker) { [DeviceTweakerQaDesktop]::AcquireWorkload() } else { $null }
try {
$RunDirectory = (Resolve-Path -LiteralPath $RunDirectory).Path
if (-not $ReviewFile) { $ReviewFile = Join-Path $RunDirectory 'review-input.json' }
$ReviewFile = [IO.Path]::GetFullPath($ReviewFile)
if (-not (Test-Path -LiteralPath (Join-Path $RunDirectory 'session.json') -PathType Leaf)) { throw 'Review requires an existing development session.' }
if (-not $Worker) {
    $parameters = @{ Worker = $true; RunDirectory = $RunDirectory; ReviewFile = $ReviewFile; CreateTemplate = [bool]$CreateTemplate }
    $code = Invoke-QaPrivateDesktop $PSCommandPath $parameters (Join-Path $RunDirectory 'review-isolation') 300
    if ($null -ne $code) { exit $code }
}
Assert-QaBoundedWorkload
$mode = if ($CreateTemplate) { '--dev-review-template' } else { '--dev-review' }
. (Join-Path $PSScriptRoot 'Support\Build-Cache.ps1')
$inspector = Get-QaSessionInspector $RunDirectory
if ($inspector) { & dotnet $inspector $mode $RunDirectory ([IO.Path]::GetFullPath($ReviewFile)) }
else { & dotnet run --project (Join-Path $PSScriptRoot 'Inspect\Inspect.csproj') -- $mode $RunDirectory ([IO.Path]::GetFullPath($ReviewFile)) }
exit $LASTEXITCODE
} finally { if ($lease) { $lease.Dispose() } }
