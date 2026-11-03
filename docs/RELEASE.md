# Releasing DEVICE TWEAKER

[Documentation](./README.md) / Release · [Русский](./RELEASE.ru.md)

A release contains two portable EXEs, their checksums and the exact source archive. The current preparation replaces **0.0.4** at its existing URL. Keep the former release recoverable throughout the replacement.

## Sources and release notes

Choose one final source commit after code, documentation and history review. Build from a clean, isolated checkout. Version fields in `src/DeviceTweakerCS.csproj` and `src/assets/app.manifest`, package checks and GUI fallback must agree on `0.0.4`.

Derive the 0.0.4 changelog from **0.0.3 → final source**, checking implementation changes rather than copying commit subjects. Keep an additional comparison of the former published 0.0.4 against its replacement. For older releases, compare adjacent source revisions; 0.0.1 is the baseline.

Keep English and Russian changelogs and release notes consistent. Preserve historical dates. The replacement script appends the actual update date and exact source SHA to the published release description.

## Local package

From the **repository root**, in the clean checkout:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\src\build.ps1 -Flavor both -Configuration Release -SkipImodDriverBuild
if ($LASTEXITCODE -ne 0) { throw 'Build failed' }
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\dev\qa\Test-ReleasePackage.ps1 -PackageDirectory .\src\bin\ReleasePackages\v0.0.4
if ($LASTEXITCODE -ne 0) { throw 'Package check failed' }

$package = '.\src\bin\ReleasePackages\v0.0.4'
$sourceName = 'DEVICE.TWEAKER.Source.v0.0.4.zip'
git archive --format=zip --output="$package/$sourceName" HEAD
if ($LASTEXITCODE -ne 0) { throw 'Source archive failed' }
$hash = (Get-FileHash -LiteralPath "$package/$sourceName" -Algorithm SHA256).Hash
"$hash  $sourceName" | Set-Content -LiteralPath "$package/SOURCE-SHA256SUMS.txt" -Encoding ascii
```

The five assets are `DEVICE.TWEAKER.exe`, `DEVICE.TWEAKER.SelfContained.exe`, `SHA256SUMS.txt`, `DEVICE.TWEAKER.Source.v0.0.4.zip` and `SOURCE-SHA256SUMS.txt`.

The normal EXE needs .NET 10 Desktop Runtime x64; the self-contained EXE includes it. The [package checker](../dev/qa/Test-ReleasePackage.ps1) verifies versions, shared application identity, embedded components and absence of the test panel, using a read-only startup probe. It does not load DTIMOD or validate hardware settings. Check the archive against the chosen source commit.

Complete relevant software checks, inspect EN/RU captures and then perform hardware acceptance on the final files. Record the tested Windows/device/driver versions and any uncovered configurations. Driver loading, register writes and Windows configuration changes require explicit authorization on the test system.

## Build without publishing

The `Release package` workflow takes an exact 40-character `source_sha`. Its default `action=build` restores, tests, builds both EXEs, checks the package and uploads the five files as an Actions artifact. It has read-only repository permissions and creates no release or tag.

Inspect that artifact, complete acceptance on its EXEs and record the build run ID and SHA-256 of `SHA256SUMS.txt`. Replacement downloads those exact files from the successful build run, verifies their source archive against Git blobs and checks the accepted EXE manifest hash. It does not rebuild them.

## Replace the existing 0.0.4

After the owner's publication approval:

1. Push the reviewed commits to `main` and wait for CI. The selected source must equal remote `main`.
2. Read the current `main` SHA, the **tag object** for `v0.0.4` (not just its dereferenced commit), and the release API `updated_at`. Review any difference from the archived state before continuing. Confirm in repository settings that release immutability is disabled for the replacement: a mutable older release alone does not establish the policy for republication.
3. Run `Release package` from `main` with `action=replace-0.0.4`, the exact `source_sha`, `expected_tag_object`, `expected_release_updated_at`, `verified_run_id` and `expected_exe_sums_sha256`.
4. The separate write-enabled job archives the release metadata, old tag/source bundle and all five old assets, verifies their SHA-256 values and saves a recovery artifact **before any remote write**. It then checks the remote state again.
5. The script makes the existing release a draft, conditionally replaces the tag using `--force-with-lease`, uploads the five files and downloads them for checksum verification. It updates the description, republishes the same release and verifies the public downloads and URL.

An immutable release, changed refs, changed metadata or a changed asset set stops preparation. The old main commits and tags 0.0.1-0.0.3 are preserved.

Replacement is implemented in [replace_release.py](../dev/release/replace_release.py). Running it without `--apply` only prepares and verifies a local backup. The write phase requires that backup. Python 3.11+, Git and authenticated GitHub CLI are required.

## Recovery

GitHub CLI [deletes an existing asset before uploading its replacement](https://cli.github.com/manual/gh_release_upload). A failed replacement therefore restores the archived tag, five files, title and description; downloaded hashes must match before the former release becomes public again. Asset IDs and download counts can change when files are replaced.

A cancelled job or terminated process cannot run automatic recovery. Download the saved recovery artifact and run `replace_release.py` with the original parameters and `--restore` instead of `--apply`; it imports the archived tag bundle and verifies the restored downloads.

If automatic recovery fails, inspect `rollback-error.json` and the recovery artifact. Keep the release draft while restoring consistent files and refs. A concurrent remote ref change stops automatic recovery rather than overwriting another update. Do not retry with new expected values without reviewing the cause.

Keep the recovery artifact outside its 90-day Actions retention window in the maintainer's local archive. After success, check both README languages, release links, runtime requirements, tag source and all downloaded checksums. Apply the reviewed About/topics and funding metadata separately.

## Maintain the project page

The README uses the theme-aware wordmark and text navigation. Full application windows belong in the [gallery](./SCREENSHOTS.md). Capture them from the read-only developer build on a hidden desktop:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\dev\qa\Capture-Showcase.ps1 -ReadmeOnly -ExePath ".\src\bin\DeveloperQa\Release\net10.0-windows\win-x64\DEVICE TWEAKER QA.exe" -OutputDirectory ".\local\screenshots"
```

Inspect both languages, label demonstration data and check agreement between controls, summaries and readback. Use the original captures without decorative frames. Keep the GitHub About description aligned with the English overview.
