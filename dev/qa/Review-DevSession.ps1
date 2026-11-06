[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$RunDirectory, [string]$ReviewFile, [switch]$CreateTemplate)
$ErrorActionPreference = 'Stop'
$RunDirectory = (Resolve-Path -LiteralPath $RunDirectory).Path
if (-not $ReviewFile) { $ReviewFile = Join-Path $RunDirectory 'review-input.json' }
$mode = if ($CreateTemplate) { '--dev-review-template' } else { '--dev-review' }
. (Join-Path $PSScriptRoot 'Build-Cache.ps1')
$inspector = Get-QaSessionInspector $RunDirectory
if ($inspector) { & dotnet $inspector $mode $RunDirectory ([IO.Path]::GetFullPath($ReviewFile)) }
else { & dotnet run --project (Join-Path $PSScriptRoot 'Inspect\Inspect.csproj') -- $mode $RunDirectory ([IO.Path]::GetFullPath($ReviewFile)) }
exit $LASTEXITCODE
