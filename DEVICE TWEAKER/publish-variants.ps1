param(
    [ValidateSet("both", "with-net", "without-net")]
    [string]$Flavor = "both",
    [ValidateSet("Release", "Debug")]
    [string]$Configuration = "Release",
    [string]$Runtime = "win-x64",
    [switch]$Clean,
    [switch]$SkipImodDriverBuild,
    [switch]$BuildDriverOnly,
    [switch]$TrustImodDriverCert,
    [string]$MsBuildPath,
    [string]$ImodDriverCertThumbprint = "9CE4C30CD75905786774B1DDFAC126329ACAEA8D",
    [string]$ImodDriverCertSubject = "MADE BY ARSENZA",
    [string]$TimestampUrl = "http://timestamp.digicert.com"
)

$ErrorActionPreference = "Stop"

$projectPath = Join-Path $PSScriptRoot "DeviceTweakerCS.csproj"
$driverProject = Join-Path $PSScriptRoot "IMOD\Driver\ImodDriver.vcxproj"
$driverOutDir = Join-Path $PSScriptRoot "IMOD"
$driverOutPath = Join-Path $driverOutDir "DTIMOD.sys"
$driverHashPath = "$driverOutPath.sha256"
$driverCertPath = Join-Path $driverOutDir "DTIMOD.cer"
$publishRoot = Join-Path $PSScriptRoot "bin\Publish"
$selfContainedDisplayName = "DEVICE TWEAKER (NET FRAMEWORK)"
$frameworkDependentDisplayName = "DEVICE TWEAKER"
$withNetOut = Join-Path $publishRoot $selfContainedDisplayName
$withoutNetOut = Join-Path $publishRoot $frameworkDependentDisplayName

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = [Security.Principal.WindowsPrincipal]::new($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Resolve-MsBuild {
    param([string]$RequestedPath)

    if (-not [string]::IsNullOrWhiteSpace($RequestedPath)) {
        if (-not (Test-Path -LiteralPath $RequestedPath)) {
            throw "MSBuild path does not exist: $RequestedPath"
        }

        return (Resolve-Path -LiteralPath $RequestedPath).Path
    }

    $vswhere = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"
    if (Test-Path -LiteralPath $vswhere) {
        $installPath = & $vswhere -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath
        if ($LASTEXITCODE -eq 0 -and -not [string]::IsNullOrWhiteSpace($installPath)) {
            $candidate = Join-Path $installPath "MSBuild\Current\Bin\MSBuild.exe"
            if (Test-Path -LiteralPath $candidate) {
                return $candidate
            }
        }
    }

    $command = Get-Command MSBuild.exe -ErrorAction SilentlyContinue
    if ($command) {
        return $command.Source
    }

    throw "MSBuild.exe was not found. Install Visual Studio Build Tools with C++ and WDK."
}

function Resolve-VsMsvcToolchain {
    $vswhere = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"
    $installPath = $null
    if (Test-Path -LiteralPath $vswhere) {
        $installPath = & $vswhere -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath
        if ($LASTEXITCODE -ne 0) {
            $installPath = $null
        }
    }

    if ([string]::IsNullOrWhiteSpace($installPath)) {
        throw "Visual Studio installation was not found."
    }

    $vcToolsRoot = Join-Path $installPath "VC\Tools\MSVC"
    $toolset = Get-ChildItem -LiteralPath $vcToolsRoot -Directory -ErrorAction SilentlyContinue |
        Sort-Object Name -Descending |
        Select-Object -First 1
    if (-not $toolset) {
        throw "MSVC toolset was not found under $vcToolsRoot."
    }

    $cl = Join-Path $toolset.FullName "bin\Hostx64\x64\cl.exe"
    $link = Join-Path $toolset.FullName "bin\Hostx64\x64\link.exe"
    $include = Join-Path $toolset.FullName "include"
    $lib = Join-Path $toolset.FullName "lib\x64"
    if (-not (Test-Path -LiteralPath $cl)) {
        throw "cl.exe was not found: $cl"
    }
    if (-not (Test-Path -LiteralPath $link)) {
        throw "link.exe was not found: $link"
    }

    return [pscustomobject]@{
        Cl = (Resolve-Path -LiteralPath $cl).Path
        Link = (Resolve-Path -LiteralPath $link).Path
        Include = (Resolve-Path -LiteralPath $include).Path
        Lib = (Resolve-Path -LiteralPath $lib).Path
    }
}

function Resolve-WdkToolchain {
    $kitsRoot = Join-Path ${env:ProgramFiles(x86)} "Windows Kits\10"
    $includeRoot = Join-Path $kitsRoot "Include"
    $libRoot = Join-Path $kitsRoot "Lib"

    $includeVersion = Get-ChildItem -LiteralPath $includeRoot -Directory -ErrorAction SilentlyContinue |
        Where-Object {
            (Test-Path -LiteralPath (Join-Path $_.FullName "km\ntddk.h")) -and
            (Test-Path -LiteralPath (Join-Path $_.FullName "shared\initguid.h"))
        } |
        Sort-Object Name -Descending |
        Select-Object -First 1
    $libVersion = Get-ChildItem -LiteralPath $libRoot -Directory -ErrorAction SilentlyContinue |
        Where-Object {
            (Test-Path -LiteralPath (Join-Path $_.FullName "km\x64\ntoskrnl.lib")) -and
            (Test-Path -LiteralPath (Join-Path $_.FullName "km\x64\wdmsec.lib"))
        } |
        Sort-Object Name -Descending |
        Select-Object -First 1

    if (-not $includeVersion) {
        throw "Windows Kits headers were not found under $includeRoot."
    }
    if (-not $libVersion) {
        throw "Windows Kits kernel libraries were not found under $libRoot."
    }

    return [pscustomobject]@{
        Include = $includeVersion.FullName
        Lib = (Join-Path $libVersion.FullName "km\x64")
    }
}

function Invoke-ImodDriverManualBuild {
    param([string]$Configuration)

    $vs = Resolve-VsMsvcToolchain
    $wdk = Resolve-WdkToolchain

    $driverBuildDir = Join-Path $PSScriptRoot "IMOD\Driver\build\x64\$Configuration"
    $driverObjDir = Join-Path $PSScriptRoot "IMOD\Driver\obj\x64\$Configuration\manual"
    $driverSource = Join-Path $PSScriptRoot "IMOD\Driver\imod_driver.c"
    $driverObj = Join-Path $driverObjDir "imod_driver.obj"
    $driverOut = Join-Path $driverBuildDir "DTIMOD.sys"

    New-Item -ItemType Directory -Force -Path $driverBuildDir | Out-Null
    New-Item -ItemType Directory -Force -Path $driverObjDir | Out-Null

    Write-Host "Building IMOD driver (manual toolchain)..."
    Write-Host "CL:    $($vs.Cl)"
    Write-Host "Link:  $($vs.Link)"

    & $vs.Cl /nologo /c /TC /GS- /Zl /W3 /D_AMD64_ /DAMD64 /DWIN64 /D_KERNEL_MODE /DNTDDI_VERSION=0x0A00000A /D_WIN32_WINNT=0x0A00 /Fo"$driverObj" /I"$($vs.Include)" /I"$($wdk.Include)\km" /I"$($wdk.Include)\shared" /I"$($wdk.Include)\ucrt" /I"$($wdk.Include)\um" $driverSource
    if ($LASTEXITCODE -ne 0) {
        throw "IMOD driver manual compile failed with exit code $LASTEXITCODE"
    }

    & $vs.Link /nologo /driver /subsystem:native /entry:DriverEntry /out:"$driverOut" /nodefaultlib /machine:x64 /libpath:"$($wdk.Lib)" /libpath:"$($vs.Lib)" "$driverObj" ntoskrnl.lib hal.lib wdm.lib wdmsec.lib libcntpr.lib BufferOverflowK.lib
    if ($LASTEXITCODE -ne 0) {
        throw "IMOD driver manual link failed with exit code $LASTEXITCODE"
    }

    return $driverOut
}

function Resolve-SignTool {
    $command = Get-Command signtool.exe -ErrorAction SilentlyContinue
    if ($command) {
        return $command.Source
    }

    $kitsRoot = Join-Path ${env:ProgramFiles(x86)} "Windows Kits\10\bin"
    if (Test-Path -LiteralPath $kitsRoot) {
        $candidate = Get-ChildItem -LiteralPath $kitsRoot -Recurse -Filter signtool.exe -ErrorAction SilentlyContinue |
            Where-Object { $_.FullName -match "\\x64\\signtool\.exe$" } |
            Sort-Object FullName -Descending |
            Select-Object -First 1
        if ($candidate) {
            return $candidate.FullName
        }

        $candidate = Get-ChildItem -LiteralPath $kitsRoot -Recurse -Filter signtool.exe -ErrorAction SilentlyContinue |
            Sort-Object FullName -Descending |
            Select-Object -First 1
        if ($candidate) {
            return $candidate.FullName
        }
    }

    throw "signtool.exe was not found. Install Windows SDK/WDK."
}

function Resolve-ImodDriverSigningCertificate {
    $thumbprintSource = if ($null -eq $ImodDriverCertThumbprint) { "" } else { $ImodDriverCertThumbprint }
    $subjectText = if ($null -eq $ImodDriverCertSubject) { "" } else { $ImodDriverCertSubject }
    $thumbprint = $thumbprintSource.Replace(" ", "").ToUpperInvariant()
    $stores = @(
        @{ Path = "Cert:\LocalMachine\My"; MachineStore = $true },
        @{ Path = "Cert:\CurrentUser\My"; MachineStore = $false }
    )

    foreach ($store in $stores) {
        $certificates = @(Get-ChildItem -Path $store.Path -ErrorAction SilentlyContinue)
        if (-not [string]::IsNullOrWhiteSpace($thumbprint)) {
            $match = $certificates |
                Where-Object { $_.HasPrivateKey -and $_.Thumbprint.ToUpperInvariant() -eq $thumbprint } |
                Sort-Object NotAfter -Descending |
                Select-Object -First 1
            if ($match) {
                return [pscustomobject]@{
                    Certificate = $match
                    Thumbprint = $match.Thumbprint
                    Subject = $match.Subject
                    MachineStore = [bool]$store.MachineStore
                }
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($subjectText)) {
            $match = $certificates |
                Where-Object { $_.HasPrivateKey -and $_.Subject -like "*$subjectText*" } |
                Sort-Object NotAfter -Descending |
                Select-Object -First 1
            if ($match) {
                return [pscustomobject]@{
                    Certificate = $match
                    Thumbprint = $match.Thumbprint
                    Subject = $match.Subject
                    MachineStore = [bool]$store.MachineStore
                }
            }
        }
    }

    return $null
}

function Invoke-ImodDriverPostSign {
    param([string]$TargetPath)

    if (-not (Test-Path -LiteralPath $TargetPath)) {
        throw "IMOD driver to sign was not found: $TargetPath"
    }

    $certificate = Resolve-ImodDriverSigningCertificate
    if (-not $certificate) {
        throw "IMOD driver signing certificate was not found in LocalMachine\My or CurrentUser\My. Expected subject '$ImodDriverCertSubject' or thumbprint '$ImodDriverCertThumbprint'."
    }

    $signTool = Resolve-SignTool
    Write-Host "Signing IMOD driver..."
    Write-Host "SignTool:  $signTool"
    Write-Host "Subject:   $($certificate.Subject)"
    Write-Host "Thumbprint:$($certificate.Thumbprint)"
    if (-not [string]::IsNullOrWhiteSpace($TimestampUrl)) {
        Write-Host "Timestamp: $TimestampUrl"
    }

    $signArgs = @("sign", "/fd", "sha256", "/s", "My", "/sha1", $certificate.Thumbprint)
    if ($certificate.MachineStore) {
        $signArgs = @("sign", "/fd", "sha256", "/sm", "/s", "My", "/sha1", $certificate.Thumbprint)
    }
    if (-not [string]::IsNullOrWhiteSpace($TimestampUrl)) {
        $signArgs += @("/tr", $TimestampUrl, "/td", "sha256")
    }
    $signArgs += $TargetPath

    & $signTool @signArgs
    if ($LASTEXITCODE -ne 0) {
        throw "IMOD driver signing failed with exit code $LASTEXITCODE"
    }

    Export-Certificate -Cert $certificate.Certificate -FilePath $driverCertPath -Force | Out-Null

    $signature = Get-AuthenticodeSignature -LiteralPath $TargetPath
    if (-not $signature.SignerCertificate -or $signature.SignerCertificate.Thumbprint -ne $certificate.Thumbprint) {
        throw "IMOD driver was signed by an unexpected certificate."
    }
}

function Test-CertificateInStore {
    param(
        [System.Security.Cryptography.X509Certificates.StoreLocation]$Location,
        [System.Security.Cryptography.X509Certificates.StoreName]$StoreName,
        [string]$Thumbprint
    )

    $store = [System.Security.Cryptography.X509Certificates.X509Store]::new($StoreName, $Location)
    try {
        $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)
        $found = $store.Certificates.Find(
            [System.Security.Cryptography.X509Certificates.X509FindType]::FindByThumbprint,
            $Thumbprint,
            $false)
        return ($null -ne $found -and $found.Count -gt 0)
    }
    finally {
        $store.Close()
    }
}

function Add-CertificateToStore {
    param(
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,
        [System.Security.Cryptography.X509Certificates.StoreLocation]$Location,
        [System.Security.Cryptography.X509Certificates.StoreName]$StoreName
    )

    if (Test-CertificateInStore -Location $Location -StoreName $StoreName -Thumbprint $Certificate.Thumbprint) {
        Write-Host "Certificate already trusted: $Location\$StoreName"
        return
    }

    $store = [System.Security.Cryptography.X509Certificates.X509Store]::new($StoreName, $Location)
    try {
        $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
        $store.Add($Certificate)
        Write-Host "Trusted driver certificate: $Location\$StoreName"
    }
    finally {
        $store.Close()
    }
}

function Import-ImodDriverCertificate {
    param([string]$CertificatePath)

    if (-not (Test-Path -LiteralPath $CertificatePath)) {
        throw "IMOD driver certificate not found: $CertificatePath"
    }

    $resolvedCert = (Resolve-Path -LiteralPath $CertificatePath).Path
    $certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($resolvedCert)
    Write-Host "Trusting IMOD driver certificate:"
    Write-Host "Subject:    $($certificate.Subject)"
    Write-Host "Thumbprint: $($certificate.Thumbprint)"

    Add-CertificateToStore -Certificate $certificate -Location CurrentUser -StoreName Root
    Add-CertificateToStore -Certificate $certificate -Location CurrentUser -StoreName TrustedPublisher

    if (Test-IsAdministrator) {
        Add-CertificateToStore -Certificate $certificate -Location LocalMachine -StoreName Root
        Add-CertificateToStore -Certificate $certificate -Location LocalMachine -StoreName TrustedPublisher
    }
    else {
        Write-Warning "Not elevated: LocalMachine trust stores were not updated. Driver load may still be blocked by Windows."
    }
}

function Stop-PublishedProcessIfRunning {
    param([string]$TargetExePath)

    if ([string]::IsNullOrWhiteSpace($TargetExePath)) {
        return
    }

    $resolvedTarget = [System.IO.Path]::GetFullPath($TargetExePath)
    $running = @(
        Get-Process -ErrorAction SilentlyContinue |
            Where-Object {
                try {
                    return $_.Path -and [string]::Equals(
                        [System.IO.Path]::GetFullPath($_.Path),
                        $resolvedTarget,
                        [System.StringComparison]::OrdinalIgnoreCase)
                }
                catch {
                    return $false
                }
            }
    )

    foreach ($proc in $running) {
        Write-Host "Stopping locked publish target: PID=$($proc.Id) $($proc.ProcessName)"
        Stop-Process -Id $proc.Id -Force -ErrorAction Stop
    }

    if ($running.Count -gt 0) {
        Start-Sleep -Milliseconds 250
    }
}

function Invoke-ImodDriverBuild {
    if (-not (Test-Path -LiteralPath $driverProject)) {
        throw "IMOD driver project not found: $driverProject"
    }

    $builtDriver = $null
    $msbuildFailed = $false
    $msbuild = $null

    try {
        $msbuild = Resolve-MsBuild -RequestedPath $MsBuildPath
        Write-Host "Building IMOD driver..."
        Write-Host "Project: $driverProject"
        Write-Host "MSBuild: $msbuild"

        & $msbuild $driverProject `
            /p:Configuration=$Configuration `
            /p:Platform=x64 `
            /m `
            /nologo `
            /v:m

        if ($LASTEXITCODE -ne 0) {
            $msbuildFailed = $true
            Write-Warning "MSBuild driver build failed with exit code $LASTEXITCODE. Falling back to manual toolchain build."
        }
        else {
            $builtDriver = Join-Path $PSScriptRoot "IMOD\Driver\build\x64\$Configuration\DTIMOD.sys"
            if (-not (Test-Path -LiteralPath $builtDriver)) {
                $msbuildFailed = $true
                Write-Warning "Built IMOD driver was not found at $builtDriver. Falling back to manual toolchain build."
            }
        }
    }
    catch {
        $msbuildFailed = $true
        Write-Warning "MSBuild driver build could not be used: $($_.Exception.Message). Falling back to manual toolchain build."
    }

    if ($msbuildFailed) {
        $builtDriver = Invoke-ImodDriverManualBuild -Configuration $Configuration
    }

    New-Item -ItemType Directory -Force -Path $driverOutDir | Out-Null
    Invoke-ImodDriverPostSign -TargetPath $builtDriver
    Copy-Item -LiteralPath $builtDriver -Destination $driverOutPath -Force

    $hash = (Get-FileHash -Algorithm SHA256 -LiteralPath $driverOutPath).Hash.ToUpperInvariant()
    Set-Content -LiteralPath $driverHashPath -Value $hash -Encoding ASCII

    $signature = Get-AuthenticodeSignature -LiteralPath $driverOutPath

    Write-Host "Generated: $driverOutPath"
    Write-Host "SHA256:    $hash"
    Write-Host "Signature: $($signature.Status)"
    if ($signature.SignerCertificate) {
        Write-Host "Signer:    $($signature.SignerCertificate.Subject)"
    }

    if ($TrustImodDriverCert) {
        Import-ImodDriverCertificate -CertificatePath $driverCertPath
        $signature = Get-AuthenticodeSignature -LiteralPath $driverOutPath
        Write-Host "Signature after trust: $($signature.Status)"
    }

    if ($signature.Status -ne "Valid") {
        Write-Warning "The generated driver is not production-trusted on this machine. Use a proper Microsoft-signed driver for release builds."
    }
}

function Rename-PublishedExe {
    param(
        [string]$OutputDir,
        [string]$TargetName,
        [string]$SourceName = "DEVICE TWEAKER.exe"
    )

    $allExe = Get-ChildItem -LiteralPath $OutputDir -File -Filter *.exe

    $publishedExe = $allExe | Where-Object { $_.Name -ieq $SourceName } | Select-Object -First 1
    if (-not $publishedExe) {
        $publishedExe = $allExe |
            Where-Object { $_.Name -ine $TargetName } |
            Sort-Object Length -Descending |
            Select-Object -First 1
    }
    if (-not $publishedExe) {
        $publishedExe = $allExe | Where-Object { $_.Name -ieq $TargetName } | Select-Object -First 1
    }

    if (-not $publishedExe) {
        throw "No EXE found in publish output: $OutputDir"
    }

    if ($publishedExe.Name -ieq $TargetName) {
        return
    }

    $targetPath = Join-Path $OutputDir $TargetName
    if (Test-Path -LiteralPath $targetPath) {
        Remove-Item -LiteralPath $targetPath -Force
    }

    Rename-Item -LiteralPath $publishedExe.FullName -NewName $TargetName
    Write-Host "Renamed EXE: $TargetName"
}

function Invoke-PublishFlavor {
    param(
        [string]$Label,
        [bool]$SelfContained,
        [bool]$PublishSingleFile,
        [bool]$IncludeNativeLibrariesForSelfExtract,
        [string]$OutputDir,
        [string]$TargetExeName
    )

    if ($Clean -and (Test-Path -LiteralPath $OutputDir)) {
        Stop-PublishedProcessIfRunning -TargetExePath (Join-Path $OutputDir $TargetExeName)
        Remove-Item -LiteralPath $OutputDir -Recurse -Force
    }

    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null

    Write-Host ""
    Write-Host "=== Publishing: $Label ==="
    Write-Host "Output: $OutputDir"

    $selfContainedText = $SelfContained.ToString().ToLowerInvariant()
    $singleFileText = $PublishSingleFile.ToString().ToLowerInvariant()
    $nativeExtractText = $IncludeNativeLibrariesForSelfExtract.ToString().ToLowerInvariant()

    dotnet publish $projectPath `
        -c $Configuration `
        -r $Runtime `
        -p:SelfContained=$selfContainedText `
        -p:PublishSingleFile=$singleFileText `
        -p:IncludeNativeLibrariesForSelfExtract=$nativeExtractText `
        -p:DebugType=None `
        -p:DebugSymbols=false `
        -p:BuildImodDriver=false `
        -p:UseAppHost=true `
        -o $OutputDir

    if ($LASTEXITCODE -ne 0) {
        throw "dotnet publish failed for flavor '$Label' (exit code $LASTEXITCODE)"
    }

    Rename-PublishedExe -OutputDir $OutputDir -TargetName $TargetExeName
}

if ($BuildDriverOnly) {
    Invoke-ImodDriverBuild
    return
}

if (-not (Test-Path -LiteralPath $projectPath)) {
    throw "Project file not found: $projectPath"
}

if (-not (Get-Command dotnet -ErrorAction SilentlyContinue)) {
    throw "dotnet CLI is not available in PATH."
}

if (-not $SkipImodDriverBuild) {
    Invoke-ImodDriverBuild
}

Write-Host "Restoring project..."
dotnet restore $projectPath -p:BuildImodDriver=false
if ($LASTEXITCODE -ne 0) {
    throw "dotnet restore failed (exit code $LASTEXITCODE)"
}

switch ($Flavor) {
    "both" {
        Invoke-PublishFlavor -Label "self-contained (runtime included)" -SelfContained $true -PublishSingleFile $true -IncludeNativeLibrariesForSelfExtract $true -OutputDir $withNetOut -TargetExeName "$selfContainedDisplayName.exe"
        Invoke-PublishFlavor -Label "framework-dependent (requires .NET runtime)" -SelfContained $false -PublishSingleFile $true -IncludeNativeLibrariesForSelfExtract $false -OutputDir $withoutNetOut -TargetExeName "$frameworkDependentDisplayName.exe"
    }
    "with-net" {
        Invoke-PublishFlavor -Label "self-contained (runtime included)" -SelfContained $true -PublishSingleFile $true -IncludeNativeLibrariesForSelfExtract $true -OutputDir $withNetOut -TargetExeName "$selfContainedDisplayName.exe"
    }
    "without-net" {
        Invoke-PublishFlavor -Label "framework-dependent (requires .NET runtime)" -SelfContained $false -PublishSingleFile $true -IncludeNativeLibrariesForSelfExtract $false -OutputDir $withoutNetOut -TargetExeName "$frameworkDependentDisplayName.exe"
    }
}

Write-Host ""
Write-Host "Done."
Write-Host "Integrated .NET build:      $withNetOut"
Write-Host "Requires installed .NET:    $withoutNetOut"
