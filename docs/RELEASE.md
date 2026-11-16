# Releasing DEVICE TWEAKER

[Documentation](./README.md) / Release · [Русский](./RELEASE.ru.md)

Build one reviewed source commit, accept its exact files and publish a new release. The package workflow only builds and validates; it has no publication permissions.

## Sources and release notes

Use a clean, isolated checkout. Align the product version, assembly/file versions, application manifest, links and screenshots before building. Preserve internal driver and profile protocol versions unless their formats actually change.

Write the changelog from implementation differences against the preceding published release. Commit subjects help locate changes but are not evidence by themselves. The first release establishes the baseline and describes the included implementation. Keep English and Russian notes consistent, retain historical dates and include the final source SHA in the release description.

## Local package

Run from the repository root in the clean checkout:

```powershell
[xml]$project = Get-Content .\src\DeviceTweakerCS.csproj
$version = [string]($project.Project.PropertyGroup.InformationalVersion | Where-Object { $_ } | Select-Object -First 1)
$package = ".\src\bin\ReleasePackages\v$version"

powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\src\build.ps1 -Flavor both -Configuration Release -SkipImodDriverBuild
if ($LASTEXITCODE -ne 0) { throw 'Build failed' }
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\dev\qa\Test-ReleasePackage.ps1 -PackageDirectory $package
if ($LASTEXITCODE -ne 0) { throw 'Package check failed' }

$sourceRoot = '.\local\release-sources'
New-Item -ItemType Directory -Force -Path $sourceRoot | Out-Null
$sourceCommit = (git rev-parse HEAD).Trim()
if ($LASTEXITCODE -ne 0) { throw 'Source revision lookup failed' }
$sourceName = "DEVICE.TWEAKER.Source.$sourceCommit.zip"
git archive --format=zip --output="$sourceRoot/$sourceName" HEAD
if ($LASTEXITCODE -ne 0) { throw 'Source archive failed' }
Get-FileHash -LiteralPath "$sourceRoot/$sourceName" -Algorithm SHA256
```

The release package contains `DEVICE.TWEAKER.exe`, `DEVICE.TWEAKER.SelfContained.exe` and `SHA256SUMS.txt`. Sources are available through Git and GitHub's standard archives for the tag. Keep the exact local source ZIP and its checksum separately from the published files.

The regular EXE requires .NET 10 Desktop Runtime x64. The self-contained EXE includes it. The [package checker](../dev/qa/Test-ReleasePackage.ps1) compares both files with the source project version, checks their shared application identity, embedded components and absence of the test panel, using a read-only startup probe. It does not load DTIMOD or validate hardware behavior. Verify source archive entries against the selected Git tree.

Complete checks for the changed behavior, inspect EN/RU captures and perform acceptance on the final EXEs. Record the tested Windows/device/driver versions and uncovered configurations. Loading drivers, writing registers and changing Windows settings require explicit authorization on the test system.

## Build without publishing

Run `Release package` with the exact 40-character `source_sha`. It restores, tests, builds both EXEs, checks the package and uploads it as an Actions artifact. The version and output path come from the selected source project. No tag or release is created or replaced.

Download the artifact, review its contents and record the workflow run, source SHA and file checksums. Acceptance and publication must use these same files. Rebuilding after acceptance creates a different candidate.

## Publish a new release

After publication approval:

1. Confirm that the selected source is the reviewed commit, CI succeeded and the downloaded package has the accepted checksums.
2. Check remote refs and releases. Stop if the intended tag or version already exists; do not overwrite it implicitly.
3. Create the new tag at the reviewed source SHA and publish the accepted files with the reviewed notes. Include source SHA and Windows/runtime requirements.
4. Download the public assets and compare checksums. Check the tag target, source archives, both README languages, gallery links and support link.

Changes to older tags, release pages or public history require a separate migration decision. Before such a migration, preserve all old refs, annotated tag objects, release descriptions and assets locally; verify the archive can restore them. Keep private history and recovery archives out of public release assets and Actions artifacts. An ordinary new release does not need to rewrite earlier releases.

## Maintain the project page

The README uses a native Markdown heading, text navigation and one application screenshot. Additional scenarios belong in the [gallery](./SCREENSHOTS.md). Capture them from the read-only developer build on a hidden desktop:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\dev\qa\Invoke-DevSession.ps1 -Scenario .\dev\qa\Scenarios\gallery-intel-gpu.json -OutputRoot .\local\screenshots
```

The `gallery-*` scenarios cover Intel/AMD affinity, USB roles/interrupters, networking, restore, operation reports and system checks. Each records EN/RU captures. Select the reviewed `images/en.png` and `images/ru.png` for the matching gallery asset; keep the session and build reference locally. `Capture-Showcase.ps1 -ReadmeOnly` produces older capture names and does not regenerate this gallery.

Inspect both languages, label demonstration data and check agreement between controls, summaries and readback. Use the original captures without decorative frames. Keep the GitHub About description aligned with the English overview.
