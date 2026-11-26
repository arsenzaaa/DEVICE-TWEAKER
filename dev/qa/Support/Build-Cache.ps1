# One source snapshot and build per input/SDK identity; retains the newest 3 completed builds.
. (Join-Path $PSScriptRoot 'Input-Snapshot.ps1')

function Get-QaBuildKey($Files, [string]$SdkVersion) {
    $lines = @('developer-release-v1', $SdkVersion)
    $lines += @($Files | Sort-Object path | ForEach-Object { $_.path.Replace('\','/') + ':' + $_.sha256.ToUpperInvariant() })
    $algorithm = [Security.Cryptography.SHA256]::Create()
    try { return [BitConverter]::ToString($algorithm.ComputeHash([Text.Encoding]::UTF8.GetBytes(($lines -join "`n")))).Replace('-','').ToLowerInvariant() }
    finally { $algorithm.Dispose() }
}

function Get-QaBuildOutputInventory([string]$BuildDirectory) {
    $snapshot = Join-Path $BuildDirectory 'input\repo'
    foreach ($relativeRoot in @('src\bin\DeveloperQa\Release\net10.0-windows\win-x64', 'dev\qa\Inspect\bin\Release\net10.0-windows')) {
        $directory = Join-Path $snapshot $relativeRoot
        if (-not (Test-Path -LiteralPath $directory -PathType Container)) { throw "Missing QA build output: $relativeRoot" }
        foreach ($file in (Get-ChildItem -LiteralPath $directory -File -Recurse | Sort-Object FullName)) {
            [pscustomobject]@{ path = $file.FullName.Substring($snapshot.Length + 1); bytes = $file.Length;
                sha256 = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash }
        }
    }
}

function Test-QaDevBuild([string]$BuildDirectory, [string]$ExpectedKey) {
    $BuildDirectory = [IO.Path]::GetFullPath($BuildDirectory)
    $manifest = Get-Content -LiteralPath (Join-Path $BuildDirectory 'build.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($manifest.schemaVersion -ne 1 -or ($ExpectedKey -and $manifest.key -ne $ExpectedKey)) { throw 'Incompatible QA build manifest' }
    if ((Get-FileHash -LiteralPath (Join-Path $BuildDirectory 'inputs.json')).Hash -ne $manifest.inputsSha256) { throw 'QA build input manifest changed' }
    $inputs = Get-Content -LiteralPath (Join-Path $BuildDirectory 'inputs.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    if ([IO.Path]::GetFullPath($inputs.snapshot) -ne (Join-Path $BuildDirectory 'input\repo')) { throw 'QA build snapshot escaped its directory' }
    if ((Get-QaBuildKey $inputs.files $manifest.sdkVersion) -ne $manifest.key) { throw 'QA build input identity changed' }
    Test-QaInputSnapshot $BuildDirectory
    $actual = @(Get-QaBuildOutputInventory $BuildDirectory)
    $expected = @($manifest.outputs)
    if ($actual.Count -eq 0 -or $actual.Count -ne $expected.Count) { throw 'QA build output inventory changed' }
    $expectedLines = @($expected | ForEach-Object { $_.path + ':' + $_.bytes + ':' + $_.sha256 })
    $actualLines = @($actual | ForEach-Object { $_.path + ':' + $_.bytes + ':' + $_.sha256 })
    if (@(Compare-Object $expectedLines $actualLines).Count) { throw 'QA build outputs changed; preserve this directory and use -FreshBuild' }
    foreach ($entry in @('src\bin\DeveloperQa\Release\net10.0-windows\win-x64\DEVICE TWEAKER QA.exe',
        'dev\qa\Inspect\bin\Release\net10.0-windows\Inspect.dll')) {
        if ($entry -notin $actual.path) { throw "Missing QA executable: $entry" }
    }
    return $manifest
}

function Get-QaBuildDirectory([string]$BuildRoot, [string]$Name) {
    $attempt = 0
    $directory = Join-Path $BuildRoot $Name
    # The caller holds the lock for this build key. A missing manifest means
    # preparation never finished. Keep that directory and try another slot.
    while ((Test-Path -LiteralPath $directory -PathType Container) -and
        -not (Test-Path -LiteralPath (Join-Path $directory 'build.json'))) {
        $attempt++
        $directory = Join-Path $BuildRoot ($Name + '-retry-' + $attempt)
    }
    return $directory
}

function Get-QaDevBuild([string]$Workspace, [string]$BuildRoot, [switch]$FreshBuild) {
    $Workspace = [IO.Path]::GetFullPath($Workspace).TrimEnd('\')
    if (-not $BuildRoot) { $BuildRoot = Join-Path $Workspace 'local\qa-builds' }
    $BuildRoot = [IO.Path]::GetFullPath($BuildRoot)
    Test-QaDriverArtifact $Workspace
    Push-Location -LiteralPath $Workspace
    try {
        $sdkVersion = (& dotnet --version | Out-String).Trim()
        if ($LASTEXITCODE -ne 0 -or -not $sdkVersion) { throw 'Cannot determine the selected .NET SDK' }
    } finally { Pop-Location }
    $files = @(Get-QaInputFiles $Workspace | ForEach-Object {
        [pscustomobject]@{ path = $_.FullName.Substring($Workspace.TrimEnd('\').Length + 1);
            sha256 = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash }
    })
    $key = Get-QaBuildKey $files $sdkVersion
    $name = if ($FreshBuild) { $key + '_' + [Guid]::NewGuid().ToString('N').Substring(0,8) } else { $key }
    New-Item -ItemType Directory -Force -Path $BuildRoot | Out-Null
    # A second writer must fail rather than modify a build in use.
    $lock = [IO.File]::Open((Join-Path $BuildRoot ($name+'.lock')), [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
    try {
        $directory = Get-QaBuildDirectory $BuildRoot $name
        $reused = Test-Path -LiteralPath $directory
        if (-not $reused) {
            New-Item -ItemType Directory -Path $directory | Out-Null
            $snapshot = New-QaInputSnapshot $Workspace $directory
            $inputs = Get-Content -LiteralPath (Join-Path $directory 'inputs.json') -Raw -Encoding UTF8 | ConvertFrom-Json
            if ((Get-QaBuildKey $inputs.files $sdkVersion) -ne $key) { throw 'Sources changed while preparing the QA build; rerun with stable inputs' }
            Push-Location -LiteralPath $snapshot
            try {
                & dotnet build 'src\DeviceTweakerCS.csproj' -c Release -p:DeveloperQa=true -p:BuildImodDriver=false -p:TreatWarningsAsErrors=true -p:UseSharedCompilation=false -m:1 -nr:false --nologo *> (Join-Path $directory 'build.log')
                if ($LASTEXITCODE -ne 0) { throw "Developer build failed; see $directory\build.log" }
                & dotnet build 'dev\qa\Inspect\Inspect.csproj' -c Release -p:TreatWarningsAsErrors=true -p:UseSharedCompilation=false -m:1 -nr:false --nologo *> (Join-Path $directory 'inspect-build.log')
                if ($LASTEXITCODE -ne 0) { throw "Inspector build failed; see $directory\inspect-build.log" }
            } finally { Pop-Location }
            Test-QaInputSnapshot $directory
            $manifestTemp = Join-Path $directory 'build.json.tmp'
            [ordered]@{ schemaVersion = 1; key = $key; sdkVersion = $sdkVersion;
                inputsSha256 = (Get-FileHash -LiteralPath (Join-Path $directory 'inputs.json')).Hash;
                outputs = @(Get-QaBuildOutputInventory $directory) } |
                ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $manifestTemp -Encoding UTF8
            [IO.File]::Move($manifestTemp, (Join-Path $directory 'build.json'))
        }
        $null = Test-QaDevBuild $directory $key
        if (-not $FreshBuild) {
            # Retain only the newest 3 completed QA builds so local/qa-builds stays compact.
            $completed = @(Get-ChildItem -LiteralPath $BuildRoot -Directory -ErrorAction SilentlyContinue |
                Where-Object { $_.FullName -ne $directory -and (Test-Path -LiteralPath (Join-Path $_.FullName 'build.json') -PathType Leaf) } |
                Sort-Object LastWriteTimeUtc -Descending)
            if ($completed.Count -gt 2) {
                foreach ($stale in ($completed | Select-Object -Skip 2)) {
                    try {
                        Remove-Item -LiteralPath $stale.FullName -Recurse -Force -ErrorAction SilentlyContinue
                        $staleLock = Join-Path $BuildRoot ($stale.Name + '.lock')
                        if (Test-Path -LiteralPath $staleLock) { Remove-Item -LiteralPath $staleLock -Force -ErrorAction SilentlyContinue }
                    } catch { }
                }
            }
        }
        return [pscustomobject]@{ directory = $directory; key = $key; reused = [bool]$reused;
            manifestSha256 = (Get-FileHash -LiteralPath (Join-Path $directory 'build.json')).Hash;
            snapshot = (Join-Path $directory 'input\repo'); sdkVersion = $sdkVersion }
    } finally { $lock.Dispose() }
}

function Get-QaSessionInspector([string]$RunDirectory) {
    $reference = Join-Path $RunDirectory 'build-reference.json'
    if (-not (Test-Path -LiteralPath $reference)) { return $null }
    $build = Get-Content -LiteralPath $reference -Raw -Encoding UTF8 | ConvertFrom-Json
    $manifestPath = Join-Path $build.directory 'build.json'
    if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) { return $null }
    if ((Get-FileHash -LiteralPath $manifestPath).Hash -ne $build.manifestSha256) { throw 'Session build reference changed' }
    $null = Test-QaDevBuild $build.directory $build.key
    return Join-Path $build.directory 'input\repo\dev\qa\Inspect\bin\Release\net10.0-windows\Inspect.dll'
}
