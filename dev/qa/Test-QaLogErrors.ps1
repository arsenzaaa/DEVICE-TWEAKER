[CmdletBinding()]
param([Parameter(Mandatory = $true)][string]$LogPath, [string]$ReportPath)
$ErrorActionPreference = 'Stop'
$unexpected = New-Object 'System.Collections.Generic.List[string]'
$expected = New-Object 'System.Collections.Generic.List[object]'
$scenario = $null
$pending = New-Object 'System.Collections.Generic.List[string]'
foreach ($line in [IO.File]::ReadAllLines($LogPath)) {
    if ($line -match 'TEST.SCENARIO.BEGIN: .*fault=(\w+)') {
        if ($scenario) { $unexpected.Add('Unclosed fault scenario: ' + $scenario) }
        $scenario = $Matches[1]; $pending.Clear()
    }
    if ($line -match '\[(ERROR|FATAL)\]') {
        if ($line -match '\[FATAL\]') { $unexpected.Add($line) }
        elseif ($line -match 'UI.RESULT.(ITEM|SUMMARY):' -and $line -match 'scope=qa-dialog-preview') {
            $expected.Add([pscustomobject]@{ kind = 'dialog-preview'; line = $line })
        } elseif ($scenario -and $scenario -ne 'None' -and
            ($line -match 'TEST.SANDBOX.(WRITE|IMOD).ERROR:' -or
             ($scenario -eq 'BackupDelete' -and $line -match 'BACKUP.DELETE: failed .*System.IO.IOException: Injected file deletion failure\.'))) { $pending.Add($line) }
        else { $unexpected.Add($line) }
    }
    if ($line -match 'TEST.SCENARIO.END:') {
        if ($line -match 'status=PASS\b.*assertions=\d+/0\b' -and $line -match 'faultConsumed=True\b') {
            foreach ($faultLine in $pending) { $expected.Add([pscustomobject]@{ kind = 'injected-fault'; fault = $scenario; line = $faultLine }) }
        } else { foreach ($faultLine in $pending) { $unexpected.Add($faultLine) } }
        $pending.Clear(); $scenario = $null
    }
}
if ($scenario) { $unexpected.Add('Unclosed fault scenario: ' + $scenario) }
$report = [ordered]@{ status = $(if ($unexpected.Count -eq 0) { 'PASS' } else { 'FAIL' });
    expectedErrors = @($expected.ToArray()); unexpectedErrors = @($unexpected.ToArray()); logPath = $LogPath }
if ($ReportPath) { $report | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $ReportPath -Encoding UTF8 }
Write-Host "Error audit: expected=$($expected.Count) unexpected=$($unexpected.Count) log=$LogPath"
if ($unexpected.Count -gt 0) { $unexpected | Out-Host; throw 'Unexpected runtime errors in QA log' }
