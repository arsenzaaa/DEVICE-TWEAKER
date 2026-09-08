param(
    [string]$LogPath,
    [switch]$RequireSessionSummary
)

$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($LogPath)) {
    $candidates = @(Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot '..\bin') -Recurse -File -Filter 'DeviceTweaker_*.log' -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTimeUtc -Descending)
    if ($candidates.Count -eq 0) {
        throw 'No DEVICE TWEAKER session log was found.'
    }
    $LogPath = $candidates[0].FullName
}

$resolved = (Resolve-Path -LiteralPath $LogPath).Path
$bytes = [IO.File]::ReadAllBytes($resolved)
if ($bytes.Length -lt 3 -or $bytes[0] -ne 0xEF -or $bytes[1] -ne 0xBB -or $bytes[2] -ne 0xBF) {
    throw 'Log is not UTF-8 with BOM.'
}

$lines = [IO.File]::ReadAllLines($resolved, [Text.Encoding]::UTF8)
if ($lines.Count -eq 0) {
    throw 'Log is empty.'
}

$linePattern = '^\[(?<timestamp>\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\.\d{3} [+-]\d{2}:\d{2})\] \[#(?<sequence>\d+)\] \[\+(?<elapsed>\d+)ms\] \[T(?<thread>\d+)\] \[(?<level>INFO|WARN|ERROR|FATAL)\] \[(?<category>[A-Z0-9_-]{1,10})\] (?<event>.+)$'
$records = New-Object System.Collections.Generic.List[object]
$problems = New-Object System.Collections.Generic.List[string]

for ($index = 0; $index -lt $lines.Count; $index++) {
    $line = $lines[$index]
    if ($line -notmatch $linePattern) {
        $problems.Add("line $($index + 1) is not structured: $line")
        continue
    }

    $records.Add([pscustomobject]@{
        Line = $index + 1
        Sequence = [int]$Matches.sequence
        Elapsed = [long]$Matches.elapsed
        Level = $Matches.level.Trim()
        Category = $Matches.category.Trim()
        Event = $Matches.event
    })
}

for ($index = 0; $index -lt $records.Count; $index++) {
    $expectedSequence = $index + 1
    if ($records[$index].Sequence -ne $expectedSequence) {
        $problems.Add("sequence mismatch at line $($records[$index].Line): expected $expectedSequence, got $($records[$index].Sequence)")
    }
    if ($index -gt 0 -and $records[$index].Elapsed -lt $records[$index - 1].Elapsed) {
        $problems.Add("elapsed time moved backwards at line $($records[$index].Line)")
    }
}

$schema = @($records | Where-Object { $_.Event -like 'LOG.SCHEMA: version=3 *' })
if ($schema.Count -ne 1) {
    $problems.Add("expected one LOG.SCHEMA version=3 entry, found $($schema.Count)")
} elseif ($schema[0].Line -ne 1) {
    $problems.Add('LOG.SCHEMA is not the first log entry')
}

$sessionStarts = @($records | Where-Object { $_.Event -like 'LOG.SESSION.START:*' })
if ($sessionStarts.Count -ne 1) {
    $problems.Add("expected one LOG.SESSION.START entry, found $($sessionStarts.Count)")
}

$versions = @($records | Where-Object { $_.Event -like 'LOG.VERSION:*' })
if ($versions.Count -ne 1) {
    $problems.Add("expected one LOG.VERSION entry, found $($versions.Count)")
} elseif ($versions[0].Event -notmatch '^LOG\.VERSION: product=\S+ file=\d+(?:\.\d+){3} assembly=\d+(?:\.\d+){3}$') {
    $problems.Add('LOG.VERSION does not contain complete product, file, and assembly versions')
}

$legacyResult = @($records | Where-Object { $_.Event -match '^UI\.RESULT:.*\bissues=' })
if ($legacyResult.Count -gt 0) {
    $problems.Add("legacy ambiguous UI.RESULT issues field found $($legacyResult.Count) time(s)")
}

$nestedAutoResultQuotes = @($records | Where-Object {
    $_.Event -match '^AUTO\.RESULT\.(?:APPLIED|SKIPPED|PRESERVED|CONFIGURED):.*\|\s*".*\broles="'
})
if ($nestedAutoResultQuotes.Count -gt 0) {
    $problems.Add("AUTO result entries with nested quoted roles found $($nestedAutoResultQuotes.Count) time(s)")
}

$legacyImodUiPayloads = @($records | Where-Object {
    $_.Event -like 'IMOD.UI:*' -and $_.Event -match '\s(?:map|detail)="'
})
if ($legacyImodUiPayloads.Count -gt 0) {
    $problems.Add("legacy oversized IMOD.UI payload found $($legacyImodUiPayloads.Count) time(s)")
}

$summaries = @($records | Where-Object { $_.Event -like 'LOG.SESSION.SUMMARY:*' })
if ($RequireSessionSummary -and $summaries.Count -ne 1) {
    $problems.Add("expected one session summary, found $($summaries.Count)")
}
if ($summaries.Count -gt 1) {
    $problems.Add("multiple session summaries found: $($summaries.Count)")
}
if ($summaries.Count -eq 1) {
    $summary = $summaries[0]
    if ($summary.Line -ne $lines.Count) {
        $problems.Add('session summary is not the final log entry')
    }
    if ($summary.Event -notmatch 'entries=(?<entries>\d+) infoEntries=(?<info>\d+) warningEntries=(?<warnings>\d+) errorEntries=(?<errors>\d+) fatalEntries=(?<fatal>\d+) continuationEntries=(?<continuations>\d+)') {
        $problems.Add('session summary counters could not be parsed')
    } else {
        $before = @($records | Where-Object { $_.Line -lt $summary.Line })
        $actualInfo = @($before | Where-Object Level -eq 'INFO').Count
        $actualWarnings = @($before | Where-Object Level -eq 'WARN').Count
        $actualErrors = @($before | Where-Object Level -eq 'ERROR').Count
        $actualFatal = @($before | Where-Object Level -eq 'FATAL').Count
        $actualContinuations = @($before | Where-Object { $_.Event.StartsWith('CONT | ') }).Count
        $expected = @{
            entries = $before.Count
            info = $actualInfo
            warnings = $actualWarnings
            errors = $actualErrors
            fatal = $actualFatal
            continuations = $actualContinuations
        }
        foreach ($name in $expected.Keys) {
            if ([int64]$Matches[$name] -ne [int64]$expected[$name]) {
                $problems.Add("session summary $name mismatch: expected $($expected[$name]), got $($Matches[$name])")
            }
        }
    }
}

if ($problems.Count -gt 0) {
    $problems | ForEach-Object { Write-Error $_ }
    throw "Log quality validation failed with $($problems.Count) problem(s): $resolved"
}

$warningCount = @($records | Where-Object Level -eq 'WARN').Count
$errorCount = @($records | Where-Object { $_.Level -in @('ERROR', 'FATAL') }).Count
Write-Output "LOG QUALITY: PASS"
Write-Output "path=$resolved"
Write-Output "lines=$($lines.Count) warnings=$warningCount errors=$errorCount summary=$($summaries.Count)"
