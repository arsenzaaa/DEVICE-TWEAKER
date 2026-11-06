[CmdletBinding()]
param([string]$OutputDirectory)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Package-Identity.ps1')
if (-not $OutputDirectory) { $OutputDirectory = Join-Path $PSScriptRoot ('..\..\local\package-contracts\run_'+[Guid]::NewGuid().ToString('N').Substring(0,8)) }
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$same = 'a' * 32
$cases = @(
    @{ name='matching-flavors'; valid=$true; files=@(@{appBuildId=$same},@{appBuildId=$same}) },
    @{ name='mixed-flavors'; valid=$false; files=@(@{appBuildId=$same},@{appBuildId=('b'*32)}) },
    @{ name='missing-identity'; valid=$false; files=@(@{appBuildId=$same},@{}) },
    @{ name='missing-flavor'; valid=$false; files=@(@{appBuildId=$same}) }
)
$results = @()
foreach ($case in $cases) {
    $accepted = $true
    try { Assert-QaPackageBuildIdentity $case.files } catch { $accepted = $false }
    if ($accepted -ne $case.valid) { throw "Package identity contract failed: $($case.name)" }
    $results += [pscustomobject]@{name=$case.name;status='PASS'}
}
[ordered]@{ status='PASS'; cases=$results } | ConvertTo-Json -Depth 4 |
    Set-Content -LiteralPath (Join-Path $OutputDirectory 'identity.json') -Encoding UTF8
Write-Output 'PASS: matching EXEs accepted; mixed, missing and incomplete identities rejected.'
