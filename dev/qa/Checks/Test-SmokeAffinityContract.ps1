param([Parameter(Mandatory=$true)][string]$ReportPath)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot '..\Support\Smoke-Contracts.ps1')
$cases=@(
    @{name='[ ALL ]';id='FILTER_ALL';expected=$false},
    @{name='[ ALL ]';id='';expected=$true},
    @{name='[ NONE ]';id='';expected=$true},
    @{name='[ P-CORES ]';id='';expected=$true},
    @{name='[ CCD0 ]';id='';expected=$true},
    @{name='[ NONE ]';id='FILTER_ALL';expected=$true},
    @{name='REFRESH';id='REFRESH';expected=$false})
$checks=@($cases | ForEach-Object {
    $actual=Test-QaObsoleteAffinityButton $_.name $_.id
    @{name=$_.name;automationId=$_.id;expected=$_.expected;actual=$actual;passed=($actual -eq $_.expected)}
})
$passed=@($checks | Where-Object {-not $_.passed}).Count -eq 0
@{status=$(if($passed){'PASS'}else{'FAIL'});checks=$checks} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $ReportPath -Encoding UTF8
if (-not $passed) { throw 'Affinity button smoke contract failed' }
