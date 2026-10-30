function Get-QaInputFiles([string]$Workspace) {
    $app = Join-Path $Workspace 'src'
    $files = @()
    foreach ($folder in @('Affinity', 'Core', 'Devices', 'GUI', 'Interop', 'Localization', 'Models', 'Tweaks')) {
        $files += @(Get-ChildItem -LiteralPath (Join-Path $app $folder) -Recurse -File -Filter '*.cs')
    }
    $files += @(Get-ChildItem -LiteralPath (Join-Path $app 'Startup') -File)
    foreach ($file in @('DeviceTweakerCS.csproj', '.editorconfig', 'assets\DEVICE TWEAKER.ico', 'assets\app.manifest',
        'IMOD\DTIMOD.sys', 'IMOD\DTIMOD.sys.sha256', 'IMOD\Loader\kdu.exe', 'IMOD\Loader\drv64.dll',
        'IMOD\Driver\imod_driver.c', 'IMOD\Driver\imod_driver.h', 'IMOD\Driver\ImodDriver.vcxproj',
        'IMOD\DTIMOD.build.json', 'publish-variants.ps1')) {
        $files += Get-Item -LiteralPath (Join-Path $app $file)
    }
    foreach ($file in @('IMOD\Loader\KDU.build.json', 'IMOD\KDU\upstream.json', 'IMOD\KDU\build-kdu.ps1', 'IMOD\KDU\README.md')) {
        $files += Get-Item -LiteralPath (Join-Path $app $file)
    }
    $vendor = Get-Content -LiteralPath (Join-Path $app 'IMOD\KDU\upstream.json') -Raw | ConvertFrom-Json
    foreach ($file in $vendor.files) { $files += Get-Item -LiteralPath (Join-Path $app ('IMOD\KDU\' + $file.path)) }
    $qa = Join-Path $Workspace 'dev\qa'
    $files += @(Get-ChildItem -LiteralPath $qa -File)
    foreach ($folder in @('Panel', 'tests\DeviceTweaker.Tests', 'Inspect', 'Workbench', 'Scenarios')) {
        $files += @(Get-ChildItem -LiteralPath (Join-Path $qa $folder) -File |
            Where-Object { $_.Extension -in @('.cs', '.csproj', '.json') })
    }
    return @($files | Sort-Object FullName -Unique)
}

function New-QaInputSnapshot([string]$Workspace, [string]$RunDirectory) {
    Test-QaDriverArtifact $Workspace
    $Workspace = [IO.Path]::GetFullPath($Workspace).TrimEnd('\')
    $copyRoot = Join-Path $RunDirectory 'input\repo'
    $inventory = @()
    foreach ($file in (Get-QaInputFiles $Workspace)) {
        $relative = $file.FullName.Substring($Workspace.Length + 1)
        $target = Join-Path $copyRoot $relative
        New-Item -ItemType Directory -Force -Path (Split-Path $target) | Out-Null
        $before = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
        Copy-Item -LiteralPath $file.FullName -Destination $target
        $copied = (Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash
        $after = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
        if ($before -ne $copied -or $after -ne $copied) { throw "Source changed while making QA snapshot: $relative" }
        $inventory += [pscustomobject]@{ path = $relative; sha256 = $copied; bytes = $file.Length }
    }
    $manifest = Join-Path $RunDirectory 'inputs.json'
    [ordered]@{ schemaVersion = 1; workspace = $Workspace; snapshot = $copyRoot; files = $inventory } |
        ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $manifest -Encoding UTF8
    return $copyRoot
}

function Test-QaInputSnapshot([string]$RunDirectory) {
    $manifest = Get-Content -LiteralPath (Join-Path $RunDirectory 'inputs.json') -Raw | ConvertFrom-Json
    foreach ($entry in $manifest.files) {
        $path = Join-Path $manifest.snapshot $entry.path
        if (-not (Test-Path -LiteralPath $path) -or (Get-FileHash -LiteralPath $path).Hash -ne $entry.sha256) {
            throw "QA input was changed during the run: $($entry.path)"
        }
    }
    # Detect added source files as well as changed/deleted files.
    $actual = @(Get-QaInputFiles $manifest.snapshot | ForEach-Object {
        $_.FullName.Substring($manifest.snapshot.Length + 1)
    })
    if (@(Compare-Object @($manifest.files.path) $actual).Count) { throw 'QA input file inventory changed during the run' }
    Test-QaDriverArtifact $manifest.snapshot
}

function Test-QaDriverArtifact([string]$Workspace) {
    $app = Join-Path $Workspace 'src'
    $manifest = Get-Content -LiteralPath (Join-Path $app 'IMOD\DTIMOD.build.json') -Raw | ConvertFrom-Json
    $expected = @('IMOD\Driver\imod_driver.c', 'IMOD\Driver\imod_driver.h', 'IMOD\Driver\ImodDriver.vcxproj', 'publish-variants.ps1')
    if ($manifest.schemaVersion -ne 1 -or @(Compare-Object $expected @($manifest.sources.path)).Count) {
        throw 'DTIMOD build manifest is incomplete'
    }
    $hash = (Get-FileHash -LiteralPath (Join-Path $app 'IMOD\DTIMOD.sys') -Algorithm SHA256).Hash
    $declared = (Get-Content -LiteralPath (Join-Path $app 'IMOD\DTIMOD.sys.sha256') -Raw).Trim()
    if ($hash -ne $manifest.binarySha256 -or $hash -ne $declared) { throw 'DTIMOD binary differs from its build manifest' }
    foreach ($source in $manifest.sources) {
        if ((Get-FileHash -LiteralPath (Join-Path $app $source.path) -Algorithm SHA256).Hash -ne $source.sha256) {
            throw "DTIMOD source changed since the binary was built: $($source.path). Rebuild the driver before QA."
        }
    }
    Test-QaLoaderArtifact $Workspace
}

function Test-QaLoaderArtifact([string]$Workspace) {
    $app = Join-Path $Workspace 'src'
    $vendorRoot = Join-Path $app 'IMOD\KDU'
    $sourcePath = Join-Path $vendorRoot 'upstream.json'
    $source = Get-Content -LiteralPath $sourcePath -Raw | ConvertFrom-Json
    $build = Get-Content -LiteralPath (Join-Path $app 'IMOD\Loader\KDU.build.json') -Raw | ConvertFrom-Json
    if ($build.schemaVersion -ne 1 -or $build.configuration -ne 'Release' -or $build.commit -ne $source.commit -or $build.tag -ne $source.tag -or
        (Get-FileHash -LiteralPath $sourcePath).Hash -ne $build.sourceManifestSha256 -or
        (Get-FileHash -LiteralPath (Join-Path $vendorRoot 'build-kdu.ps1')).Hash -ne $build.buildScriptSha256) {
        throw 'KDU binary build provenance does not match its vendor source manifest'
    }
    if (@(Compare-Object @('kdu.exe','drv64.dll') @($build.binaries.name)).Count) { throw 'KDU binary pair is incomplete' }
    foreach ($binary in $build.binaries) {
        $path = Join-Path $app ('IMOD\Loader\' + $binary.name)
        if ((Get-FileHash -LiteralPath $path).Hash -ne $binary.sha256 -or
            [Diagnostics.FileVersionInfo]::GetVersionInfo($path).FileVersion -ne $binary.version -or $binary.version -ne '1.4.5.2512') {
            throw "KDU binary mismatch: $($binary.name)"
        }
    }
    foreach ($file in $source.files) {
        if ((Get-FileHash -LiteralPath (Join-Path $vendorRoot $file.path)).Hash -ne $file.sha256) { throw "KDU source differs from pinned upstream: $($file.path)" }
    }
}
