# Capture documentation images through the same bounded, read-only workbench.
[CmdletBinding()]
param(
    [string]$OutputDirectory,
    [string[]]$Only,
    [switch]$Worker
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Support\Private-Desktop.ps1')
. (Join-Path $PSScriptRoot 'Support\Read-QaJson.ps1')
$lease = if (-not $Worker) { [DeviceTweakerQaDesktop]::AcquireWorkload() } else { $null }
try {
    if (-not $OutputDirectory) {
        $OutputDirectory = Join-Path $PSScriptRoot ('..\..\local\gallery\' + (Get-Date -Format 'yyyyMMdd-HHmmss'))
    }
    $OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
    if (-not $Worker) {
        if (Test-Path -LiteralPath $OutputDirectory) { throw 'Use a new gallery output directory.' }
        $parameters = @{ Worker = $true; OutputDirectory = $OutputDirectory }
        if ($Only) { $parameters.Only = $Only }
        $code = Invoke-QaPrivateDesktop $PSCommandPath $parameters (Join-Path $OutputDirectory 'isolation') 3600
        if ($null -ne $code) {
            if ($code -ne 0) { exit $code }
            $result = Read-QaJson (Join-Path $OutputDirectory 'gallery.json')
            if ($result.status -ne 'CAPTURED_REVIEW_REQUIRED' -or @($result.images).Count -eq 0 -or
                -not (Test-Path -LiteralPath (Join-Path $OutputDirectory 'index.html'))) { throw 'Gallery worker produced no complete review package.' }
            Write-Output $OutputDirectory
            return
        }
    }
    Assert-QaBoundedWorkload
    $catalog = Read-QaJson (Join-Path $PSScriptRoot 'Scenarios\Gallery\catalog.json')
    if ($Only) {
        foreach ($name in $Only) {
            if ($name -notin $catalog.scenario) { throw "Unknown gallery scenario: $name" }
        }
        $catalog = @($catalog | Where-Object { $_.scenario -in $Only })
    }
    if (@($catalog).Count -eq 0) { throw 'Gallery selection is empty.' }
    $images = Join-Path $OutputDirectory 'images'
    New-Item -ItemType Directory -Force -Path $images | Out-Null
    $results = [Collections.Generic.List[object]]::new()
    foreach ($entry in $catalog) {
        $scenario = Join-Path $PSScriptRoot ('Scenarios\Gallery\' + $entry.scenario + '.json')
        $sessions = Join-Path $OutputDirectory $entry.scenario
        & (Join-Path $PSScriptRoot 'Invoke-DevSession.ps1') -Scenario $scenario -OutputRoot $sessions | Out-Host
        $runs = @(Get-ChildItem -LiteralPath $sessions -Directory -Filter 'dev_*')
        if ($runs.Count -ne 1) { throw "Expected one gallery session: $($entry.scenario)" }
        $run = $runs[0].FullName
        $session = Read-QaJson (Join-Path $run 'session.json')
        if ($session.status -ne 'PASS' -or -not $session.readOnly) { throw 'Gallery session did not pass in read-only mode.' }
        foreach ($item in $entry.captures.PSObject.Properties) {
            $capture = @($session.captures | Where-Object { $_.id -eq $item.Name })
            if ($capture.Count -ne 1) { throw "Missing gallery checkpoint: $($entry.scenario)/$($item.Name)" }
            $source = Join-Path $run $capture[0].path
            $hash = (Get-FileHash -LiteralPath $source).Hash
            if ($hash -ne $capture[0].sha256) { throw 'Gallery capture changed after verification.' }
            Copy-Item -LiteralPath $source -Destination (Join-Path $images $item.Value)
            $results.Add([ordered]@{ file = $item.Value; scenario = $entry.scenario; checkpoint = $item.Name;
                sha256 = $hash; width = $capture[0].width; height = $capture[0].height; run = $run })
        }
        [ordered]@{ status = 'CAPTURED_REVIEW_REQUIRED'; images = @($results.ToArray());
            scope = 'Native read-only model captures. Layout canvas sizes are not physical monitor/DPI tests. Review before copying into assets/screenshots.' } |
            ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'gallery.json') -Encoding UTF8
    }
    $html = '<!doctype html><meta charset="utf-8"><title>DEVICE TWEAKER gallery review</title><style>body{background:#0d1117;color:#e6edf3;font:16px system-ui;margin:32px}figure{margin:36px 0}img{display:block;max-width:100%;height:auto;border:1px solid #30363d}a{color:#58a6ff}</style><h1>Gallery review</h1><p>Native captures, synthetic data. Open an image at its original size.</p>'
    foreach ($item in $results) {
        $name = [Net.WebUtility]::HtmlEncode($item.file)
        $html += '<figure><figcaption>' + $name + ' (' + $item.width + ' x ' + $item.height + ')</figcaption><a href="images/' + $name + '"><img src="images/' + $name + '" alt="' + $name + '"></a></figure>'
    }
    [IO.File]::WriteAllText((Join-Path $OutputDirectory 'index.html'), $html)
    Write-Output $OutputDirectory
} finally { if ($lease) { $lease.Dispose() } }
