[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$RunDirectory, [string]$ReviewFile, [switch]$CreateTemplate)
$ErrorActionPreference = 'Stop'
$RunDirectory = (Resolve-Path -LiteralPath $RunDirectory).Path
if (-not $ReviewFile) { $ReviewFile = Join-Path $RunDirectory 'review-input.json' }
$mode = if ($CreateTemplate) { '--dev-review-template' } else { '--dev-review' }
& dotnet run --project (Join-Path $PSScriptRoot 'Inspect\Inspect.csproj') -- $mode $RunDirectory ([IO.Path]::GetFullPath($ReviewFile))
exit $LASTEXITCODE
