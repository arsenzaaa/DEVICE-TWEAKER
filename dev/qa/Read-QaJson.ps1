# Readers permit atomic replacement. Retry only Windows handoff/sharing errors;
# malformed JSON is an actual failure and is never retried or ignored.
function Read-QaText([string]$Path) {
    $stream = [IO.FileStream]::new($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read,
        [IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete)
    $reader = [IO.StreamReader]::new($stream)
    try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
}

function Read-QaJson([string]$Path) {
    for ($attempt = 0; $attempt -le 10; $attempt++) {
        $stream = $null
        $reader = $null
        try {
            $stream = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read,
                [IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete)
            $reader = New-Object IO.StreamReader($stream)
            $text = $reader.ReadToEnd()
        } catch {
            if ($attempt -eq 10 -or ($_.Exception.InnerException -isnot [IO.IOException] -and
                $_.Exception.InnerException -isnot [UnauthorizedAccessException] -and
                $_.Exception -isnot [IO.IOException])) { throw }
            Start-Sleep -Milliseconds 10
            continue
        } finally {
            if ($reader) { $reader.Dispose() }
            elseif ($stream) { $stream.Dispose() }
        }
        return $text | ConvertFrom-Json
    }
}
