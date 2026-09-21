# Safe GUI smoke for DEVICE TWEAKER (second monitor + TEST ADMIN sandbox).
param(
    [string]$Root,
    [string]$ExePath,
    [string]$OutputRoot,
    [ValidateSet('Secondary', 'Primary')]
    [string]$Monitor = 'Secondary'
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class SmokeNative {
  // No SetCursorPos / mouse_event / keybd_event — never touch physical input.
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
  [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
  [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr hWnd, IntPtr hdcBlt, uint nFlags);
  public const int SW_RESTORE = 9;
  public const uint BM_CLICK = 0x00F5;
  public const uint WM_KEYDOWN = 0x0100;
  public const uint WM_KEYUP = 0x0101;
  public const uint PW_RENDERFULLCONTENT = 0x00000002;
}
'@

$root = if ([string]::IsNullOrWhiteSpace($Root)) {
    (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
} else {
    (Resolve-Path $Root).Path
}
$exe = if ([string]::IsNullOrWhiteSpace($ExePath)) {
    Join-Path $root 'bin\Release\net8.0-windows\win-x64\DEVICE TWEAKER.exe'
} else {
    (Resolve-Path $ExePath).Path
}
$outputBase = if ([string]::IsNullOrWhiteSpace($OutputRoot)) {
    Join-Path $root 'bin\SmokeSafe'
} else {
    $OutputRoot
}
$runId = '{0:yyyyMMdd_HHmmss}' -f (Get-Date)
$outDir = Join-Path $outputBase ("run_$runId")
New-Item -ItemType Directory -Force -Path $outDir | Out-Null
$reportPath = Join-Path $outDir "smoke_$runId.txt"
$smokeFailed = $false

$targetScreen = if ($Monitor -eq 'Secondary') {
    [System.Windows.Forms.Screen]::AllScreens | Where-Object { -not $_.Primary } | Select-Object -First 1
} else {
    [System.Windows.Forms.Screen]::PrimaryScreen
}
if ($null -eq $targetScreen) {
    $targetScreen = [System.Windows.Forms.Screen]::PrimaryScreen
}
$targetArea = $targetScreen.WorkingArea
$targetWidth = [Math]::Min(1280, $targetArea.Width - 40)
$targetHeight = [Math]::Min(1000, $targetArea.Height - 40)
$targetX = $targetArea.X + [Math]::Max(0, [int](($targetArea.Width - $targetWidth) / 2))
$targetY = $targetArea.Y + [Math]::Max(0, [int](($targetArea.Height - $targetHeight) / 2))
$env:DEVICE_TWEAKER_QA_WINDOW_BOUNDS = "$targetX,$targetY,$targetWidth,$targetHeight"

function Write-Report([string]$msg) {
    $line = '[{0:HH:mm:ss}] {1}' -f (Get-Date), $msg
    Write-Host $line
    Add-Content -LiteralPath $reportPath -Value $line -Encoding UTF8
}

function Get-MainWindow([int]$processId, [int]$timeoutSec = 30) {
    $deadline = (Get-Date).AddSeconds($timeoutSec)
    while ((Get-Date) -lt $deadline) {
        $rootEl = [System.Windows.Automation.AutomationElement]::RootElement
        $cond = New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::ProcessIdProperty, $processId)
        $wins = $rootEl.FindAll([System.Windows.Automation.TreeScope]::Children, $cond)
        foreach ($win in $wins) {
            if ($win -and [string]$win.Current.Name -like '*DEVICE TWEAKER*') { return $win }
        }
        Start-Sleep -Milliseconds 250
    }
    return $null
}

function Find-Desc(
    [System.Windows.Automation.AutomationElement]$scope,
    [string]$name,
    [string]$controlTypeName = $null,
    [int]$timeoutSec = 15) {
    $deadline = (Get-Date).AddSeconds([Math]::Max(0, $timeoutSec))
    do {
        if ($null -eq $scope) { return $null }
        $conds = New-Object System.Collections.Generic.List[System.Windows.Automation.Condition]
        $conds.Add((New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::NameProperty, $name)))
        if ($controlTypeName) {
            $ct = [System.Windows.Automation.ControlType]::$controlTypeName
            $conds.Add((New-Object System.Windows.Automation.PropertyCondition(
                [System.Windows.Automation.AutomationElement]::ControlTypeProperty, $ct)))
        }
        $and = New-Object System.Windows.Automation.AndCondition($conds.ToArray())
        $el = $scope.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $and)
        if ($el) { return $el }
        if ((Get-Date) -ge $deadline) { break }
        Start-Sleep -Milliseconds 150
    } while ((Get-Date) -le $deadline)
    return $null
}

function Invoke-Click(
    [System.Windows.Automation.AutomationElement]$el,
    [string]$label,
    [int]$settleMilliseconds = 700) {
    if ($null -eq $el) { throw "UI element not found: $label" }

    $hwnd = [IntPtr]$el.Current.NativeWindowHandle
    if ($hwnd -ne [IntPtr]::Zero) {
        # Async Win32 click — does not move the cursor or block on modal dialogs.
        [SmokeNative]::PostMessage($hwnd, [SmokeNative]::BM_CLICK, [IntPtr]::Zero, [IntPtr]::Zero) | Out-Null
        Write-Report "CLICK(BM): $label hwnd=$hwnd"
        if ($settleMilliseconds -gt 0) {
            Start-Sleep -Milliseconds $settleMilliseconds
        }
        return
    }

    # Prefer TogglePattern for checkboxes without an HWND.
    try {
        $toggle = $el.GetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern)
        $toggle.Toggle()
        Write-Report "CLICK(TOGGLE): $label"
        Start-Sleep -Milliseconds 700
        return
    } catch {}

    throw "No HWND/Toggle for '$label' - refusing mouse/keyboard fallback"
}

function Find-DescNamePattern(
    [System.Windows.Automation.AutomationElement]$scope,
    [string]$pattern,
    [int]$timeoutSec = 5) {
    $deadline = (Get-Date).AddSeconds($timeoutSec)
    do {
        if ($null -eq $scope) { return $null }
        try {
            $match = $scope.FindAll(
                [System.Windows.Automation.TreeScope]::Descendants,
                [System.Windows.Automation.Condition]::TrueCondition) |
                Where-Object { [string]$_.Current.Name -match $pattern } |
                Select-Object -First 1
            if ($match) { return $match }
        } catch {}
        Start-Sleep -Milliseconds 100
    } while ((Get-Date) -le $deadline)
    return $null
}

function Find-AutomationId(
    [System.Windows.Automation.AutomationElement]$scope,
    [string]$automationId,
    [int]$timeoutSec = 5) {
    $deadline = (Get-Date).AddSeconds($timeoutSec)
    do {
        $condition = New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::AutomationIdProperty, $automationId)
        $el = $scope.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $condition)
        if ($el) { return $el }
        Start-Sleep -Milliseconds 150
    } while ((Get-Date) -le $deadline)
    return $null
}

function Set-ToggleOn([System.Windows.Automation.AutomationElement]$el, [string]$label) {
    if ($null -eq $el) { throw "Toggle not found: $label" }
    $toggle = $el.GetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern)
    if ($toggle.Current.ToggleState -ne [System.Windows.Automation.ToggleState]::On) {
        $toggle.Toggle()
        Write-Report "TOGGLE ON: $label"
        Start-Sleep -Milliseconds 1000
    } else {
        Write-Report "TOGGLE already ON: $label"
    }
}

function Send-WindowKey(
    [System.Windows.Automation.AutomationElement]$window,
    [int]$virtualKey,
    [string]$label) {
    if ($null -eq $window) { throw "Window not found for key: $label" }
    $hwnd = [IntPtr]$window.Current.NativeWindowHandle
    if ($hwnd -eq [IntPtr]::Zero) { throw "Window has no native handle for key: $label" }
    [SmokeNative]::SetForegroundWindow($hwnd) | Out-Null
    if (-not [SmokeNative]::PostMessage($hwnd, [SmokeNative]::WM_KEYDOWN, [IntPtr]$virtualKey, [IntPtr]::Zero)) {
        throw "WM_KEYDOWN failed: $label"
    }
    if (-not [SmokeNative]::PostMessage($hwnd, [SmokeNative]::WM_KEYUP, [IntPtr]$virtualKey, [IntPtr]::Zero)) {
        throw "WM_KEYUP failed: $label"
    }
    Write-Report ("KEY: {0} vk=0x{1:X2}" -f $label, $virtualKey)
    Start-Sleep -Milliseconds 250
}

function Capture-Window([string]$path, [System.Windows.Automation.AutomationElement]$window) {
    $r = $window.Current.BoundingRectangle
    $width = [Math]::Max(1, [int][Math]::Ceiling($r.Width))
    $height = [Math]::Max(1, [int][Math]::Ceiling($r.Height))
    $hwnd = [IntPtr]$window.Current.NativeWindowHandle
    if ($hwnd -eq [IntPtr]::Zero) { throw "Window has no native handle: $path" }
    $center = New-Object System.Drawing.Point(
        [int]($r.Left + ($r.Width / 2)),
        [int]($r.Top + ($r.Height / 2)))
    $captureScreen = [System.Windows.Forms.Screen]::FromPoint($center)
    if ($captureScreen.DeviceName -ne $targetScreen.DeviceName) {
        throw "Capture target is on $($captureScreen.DeviceName), expected $($targetScreen.DeviceName): $path"
    }
    $bmp = New-Object System.Drawing.Bitmap $width, $height
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $hdc = $g.GetHdc()
    try {
        $captured = [SmokeNative]::PrintWindow($hwnd, $hdc, [SmokeNative]::PW_RENDERFULLCONTENT)
    }
    finally {
        $g.ReleaseHdc($hdc)
        $g.Dispose()
    }
    if (-not $captured) {
        $bmp.Dispose()
        throw "PrintWindow failed: $path"
    }

    $minBrightness = 765
    $maxBrightness = 0
    $visibleSamples = 0
    for ($y = 0; $y -lt $height; $y += 6) {
        for ($x = 0; $x -lt $width; $x += 6) {
            $pixel = $bmp.GetPixel($x, $y)
            $brightness = [int]$pixel.R + [int]$pixel.G + [int]$pixel.B
            if ($brightness -lt $minBrightness) { $minBrightness = $brightness }
            if ($brightness -gt $maxBrightness) { $maxBrightness = $brightness }
            if ($brightness -gt 90) { $visibleSamples++ }
        }
    }
    if (($maxBrightness - $minBrightness) -lt 45 -or $visibleSamples -lt 20) {
        $bmp.Dispose()
        throw "Captured window is blank or visually invalid: $path range=$($maxBrightness - $minBrightness) visibleSamples=$visibleSamples"
    }

    $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    Write-Report "SHOT: $path method=PrintWindow size=${width}x${height} range=$($maxBrightness - $minBrightness) visibleSamples=$visibleSamples"
}

function Assert-ElementRendered(
    [string]$path,
    [System.Windows.Automation.AutomationElement]$window,
    [System.Windows.Automation.AutomationElement]$element,
    [string]$label) {
    if ($null -eq $element) { throw "Visual element not found: $label" }
    $windowRect = $window.Current.BoundingRectangle
    $elementRect = $element.Current.BoundingRectangle
    if ($element.Current.IsOffscreen -or $elementRect.Width -lt 2 -or $elementRect.Height -lt 2) {
        throw "Visual element is hidden or collapsed: $label"
    }

    $bmp = [System.Drawing.Bitmap]::FromFile($path)
    try {
        $left = [Math]::Max(0, [int][Math]::Floor($elementRect.Left - $windowRect.Left))
        $top = [Math]::Max(0, [int][Math]::Floor($elementRect.Top - $windowRect.Top))
        $right = [Math]::Min($bmp.Width - 1, [int][Math]::Ceiling($elementRect.Right - $windowRect.Left))
        $bottom = [Math]::Min($bmp.Height - 1, [int][Math]::Ceiling($elementRect.Bottom - $windowRect.Top))
        $textPixels = 0
        for ($y = $top; $y -le $bottom; $y++) {
            for ($x = $left; $x -le $right; $x++) {
                $pixel = $bmp.GetPixel($x, $y)
                if (([int]$pixel.R + [int]$pixel.G + [int]$pixel.B) -gt 210) {
                    $textPixels++
                }
            }
        }
        if ($textPixels -lt 30) {
            throw "Visual text disappeared from capture: $label pixels=$textPixels path=$path"
        }
        Write-Report "PASS: rendered $label pixels=$textPixels"
    }
    finally {
        $bmp.Dispose()
    }
}

function Dismiss-Dialogs([System.Windows.Automation.AutomationElement]$scope, [int]$rounds = 8) {
    for ($i = 0; $i -lt $rounds; $i++) {
        $dismissed = $false
        foreach ($btnName in @('SKIP', 'NO', 'OK', 'YES')) {
            $btn = Find-Desc $scope $btnName 'Button' 0
            if ($btn) {
                try {
                    Invoke-Click $btn "dialog/$btnName"
                    $dismissed = $true
                    break
                } catch {}
            }
        }
        if (-not $dismissed) { break }
        Start-Sleep -Milliseconds 400
    }
}

Write-Report "ROOT=$root"
Write-Report "EXE=$exe"
Write-Report ("MONITOR={0} primary={1} workingArea={2},{3},{4}x{5} initialBounds={6}" -f
    $targetScreen.DeviceName, $targetScreen.Primary,
    $targetArea.X, $targetArea.Y, $targetArea.Width, $targetArea.Height,
    $env:DEVICE_TWEAKER_QA_WINDOW_BOUNDS)
Get-Process | Where-Object {
    $_.ProcessName -like 'DEVICE TWEAKER*' -and
    -not [string]::IsNullOrWhiteSpace($_.Path) -and
    $_.Path.StartsWith($root, [StringComparison]::OrdinalIgnoreCase)
} | ForEach-Object {
    Write-Report "CLOSE EXISTING pid=$($_.Id) path=$($_.Path)"
    $closed = $_.CloseMainWindow()
    if ($closed) {
        $_.WaitForExit(5000) | Out-Null
    }
    if (-not $_.HasExited) {
        Write-Report "STOP EXISTING pid=$($_.Id) after graceful-close timeout"
        Stop-Process -Id $_.Id -Force
    }
}
Start-Sleep -Seconds 1

Remove-Item Env:DEVICE_TWEAKER_QA_TEST_ADMIN -ErrorAction SilentlyContinue
Remove-Item Env:DEVICE_TWEAKER_QA_SANDBOX -ErrorAction SilentlyContinue
$env:DEVICE_TWEAKER_LANGUAGE = 'en'
$normalProc = Start-Process -FilePath $exe -WorkingDirectory (Split-Path $exe) -PassThru
Write-Report "START NORMAL pid=$($normalProc.Id)"
$normalMain = Get-MainWindow -processId $normalProc.Id -timeoutSec 30
if ($null -eq $normalMain) { throw 'Normal main window not found' }
$startupSubtitle = Find-DescNamePattern $normalMain '^alpha version .+ - developed by @arsenza$' 5
if ($null -eq $startupSubtitle) { throw 'Startup version subtitle not found' }
Start-Sleep -Milliseconds 60
$startupWindowRects = @()
for ($startupFrame = 0; $startupFrame -lt 8; $startupFrame++) {
    $startupWindowRects += $normalMain.Current.BoundingRectangle
    $framePath = Join-Path $outDir ('01_startup_{0:D2}.png' -f $startupFrame)
    Capture-Window $framePath $normalMain
    Assert-ElementRendered $framePath $normalMain $startupSubtitle "startup subtitle frame=$startupFrame"
    if ($startupFrame -eq 0) {
        Copy-Item -LiteralPath $framePath -Destination (Join-Path $outDir '09_startup_loading.png') -Force
    }
    Start-Sleep -Milliseconds 120
}
$startupWidthSpan = ($startupWindowRects | Measure-Object -Property Width -Maximum -Minimum)
$startupHeightSpan = ($startupWindowRects | Measure-Object -Property Height -Maximum -Minimum)
if (($startupWidthSpan.Maximum - $startupWidthSpan.Minimum) -gt 1 -or
    ($startupHeightSpan.Maximum - $startupHeightSpan.Minimum) -gt 1) {
    $smokeFailed = $true
    Write-Report ("FAIL: main window geometry moved during startup width={0}..{1} height={2}..{3}" -f
        $startupWidthSpan.Minimum, $startupWidthSpan.Maximum,
        $startupHeightSpan.Minimum, $startupHeightSpan.Maximum)
} else {
    Write-Report ("PASS: main window geometry is stable during startup size={0}x{1}" -f
        $startupWidthSpan.Maximum, $startupHeightSpan.Maximum)
}
$normalRect = $normalMain.Current.BoundingRectangle
$normalCenter = New-Object System.Drawing.Point(
    [int]($normalRect.Left + ($normalRect.Width / 2)),
    [int]($normalRect.Top + ($normalRect.Height / 2)))
$actualNormalScreen = [System.Windows.Forms.Screen]::FromPoint($normalCenter)
if ($actualNormalScreen.DeviceName -ne $targetScreen.DeviceName) {
    $smokeFailed = $true
    Write-Report "FAIL: startup window opened on $($actualNormalScreen.DeviceName), expected $($targetScreen.DeviceName)"
} else {
    Write-Report "PASS: startup and all startup captures are on $($targetScreen.DeviceName)"
}
$startupRu = Find-AutomationId $normalMain 'LANGUAGE_RU' 1
$startupEn = Find-AutomationId $normalMain 'LANGUAGE_EN' 1
$startupApply = Find-AutomationId $normalMain 'APPLY' 1
if ($null -eq $startupRu -or $null -eq $startupEn -or $null -eq $startupApply) {
    $smokeFailed = $true
    Write-Report 'FAIL: startup header/buttons were incomplete during initial device loading'
} else {
    $startupEnRect = $startupEn.Current.BoundingRectangle
    $startupRuRect = $startupRu.Current.BoundingRectangle
    $startupLanguageGeometryOk =
        $startupEnRect.Width -ge 40 -and $startupEnRect.Height -ge 26 -and
        [Math]::Abs($startupEnRect.Width - $startupRuRect.Width) -le 1 -and
        [Math]::Abs($startupEnRect.Height - $startupRuRect.Height) -le 1 -and
        $startupRuRect.Left - $startupEnRect.Right -ge 4
    if (-not $startupLanguageGeometryOk) {
        $smokeFailed = $true
        Write-Report ("FAIL: startup language buttons are compressed EN={0}x{1} RU={2}x{3} gap={4}" -f
            $startupEnRect.Width, $startupEnRect.Height,
            $startupRuRect.Width, $startupRuRect.Height,
            ($startupRuRect.Left - $startupEnRect.Right))
    } else {
        Write-Report ("PASS: startup header/buttons complete; language geometry EN={0}x{1} RU={2}x{3} gap={4}" -f
            $startupEnRect.Width, $startupEnRect.Height,
            $startupRuRect.Width, $startupRuRect.Height,
            ($startupRuRect.Left - $startupEnRect.Right))
    }
}
$readyDeadline = (Get-Date).AddSeconds(30)
do {
    $readyRefresh = Find-AutomationId $normalMain 'REFRESH' 1
    if ($readyRefresh -and $readyRefresh.Current.IsEnabled) { break }
    Start-Sleep -Milliseconds 100
} while ((Get-Date) -lt $readyDeadline)
if (-not $readyRefresh -or -not $readyRefresh.Current.IsEnabled) {
    throw 'Normal startup did not reach ready state in time'
}
$firstStartupRect = $startupWindowRects[0]
$readyWindowRect = $normalMain.Current.BoundingRectangle
if ([Math]::Abs($readyWindowRect.Width - $firstStartupRect.Width) -gt 1 -or
    [Math]::Abs($readyWindowRect.Height - $firstStartupRect.Height) -gt 1 -or
    [Math]::Abs($readyWindowRect.Left - $firstStartupRect.Left) -gt 1 -or
    [Math]::Abs($readyWindowRect.Top - $firstStartupRect.Top) -gt 1) {
    $smokeFailed = $true
    Write-Report ("FAIL: main window bounds changed from first loading frame to ready first={0},{1},{2}x{3} ready={4},{5},{6}x{7}" -f
        $firstStartupRect.Left, $firstStartupRect.Top, $firstStartupRect.Width, $firstStartupRect.Height,
        $readyWindowRect.Left, $readyWindowRect.Top, $readyWindowRect.Width, $readyWindowRect.Height)
} else {
    Write-Report 'PASS: main window bounds are stable from first loading frame to ready'
}
$readyEn = Find-AutomationId $normalMain 'LANGUAGE_EN' 1
$readyRu = Find-AutomationId $normalMain 'LANGUAGE_RU' 1
if ($readyEn -and $readyRu -and $startupEn -and $startupRu) {
    $readyEnRect = $readyEn.Current.BoundingRectangle
    $readyRuRect = $readyRu.Current.BoundingRectangle
    $languageGeometryStable =
        [Math]::Abs($readyEnRect.Width - $startupEnRect.Width) -le 1 -and
        [Math]::Abs($readyEnRect.Height - $startupEnRect.Height) -le 1 -and
        [Math]::Abs($readyRuRect.Width - $startupRuRect.Width) -le 1 -and
        [Math]::Abs($readyRuRect.Height - $startupRuRect.Height) -le 1
    if (-not $languageGeometryStable) {
        $smokeFailed = $true
        Write-Report ("FAIL: language button geometry changed after loading startup={0}x{1}/{2}x{3} ready={4}x{5}/{6}x{7}" -f
            $startupEnRect.Width, $startupEnRect.Height, $startupRuRect.Width, $startupRuRect.Height,
            $readyEnRect.Width, $readyEnRect.Height, $readyRuRect.Width, $readyRuRect.Height)
    } else {
        Write-Report 'PASS: language button geometry is stable from loading to ready'
    }
}
$internalUiLeaks = @(
    $normalMain.FindAll(
        [System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.Condition]::TrueCondition) |
        Where-Object {
            [string]$name = $_.Current.Name
            $name -match '(?i)Sandbox|TEST ADMIN|Test Mode|test devices|dry-run'
        } |
        ForEach-Object { [string]$_.Current.Name } |
        Sort-Object -Unique
)
if ($internalUiLeaks.Count -gt 0) {
    $smokeFailed = $true
    Write-Report "FAIL: internal test UI leaked into normal mode: $($internalUiLeaks -join ' | ')"
} else {
    Write-Report 'PASS: normal UI contains no TEST ADMIN/Sandbox labels'
}
$normalHwnd = [IntPtr]$normalMain.Current.NativeWindowHandle
[SmokeNative]::ShowWindow($normalHwnd, [SmokeNative]::SW_RESTORE) | Out-Null
[SmokeNative]::SetForegroundWindow($normalHwnd) | Out-Null
Start-Sleep -Milliseconds 300
Capture-Window (Join-Path $outDir '10_normal_main.png') $normalMain
$ruButton = Find-AutomationId $normalMain 'LANGUAGE_RU' 5
Invoke-Click $ruButton 'language/RU'
$russianWindowRect = $normalMain.Current.BoundingRectangle
if ([Math]::Abs($russianWindowRect.Width - $readyWindowRect.Width) -gt 1 -or
    [Math]::Abs($russianWindowRect.Height - $readyWindowRect.Height) -gt 1) {
    $smokeFailed = $true
    Write-Report 'FAIL: language switch changed the main window size'
} else {
    Write-Report 'PASS: language switch preserves the main window size'
}
$russianApply = Find-AutomationId $normalMain 'APPLY' 5
if ($null -eq $russianApply -or $russianApply.Current.Name -eq 'APPLY') {
    $smokeFailed = $true
    Write-Report 'FAIL: APPLY caption did not change in Russian UI'
} else {
    Write-Report 'PASS: runtime language switch EN -> RU'
}
Start-Sleep -Milliseconds 600
$ruVisibleNames = @(
    $normalMain.FindAll(
        [System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.Condition]::TrueCondition) |
        ForEach-Object { [string]$_.Current.Name }
)
$untranslatedUi = @($ruVisibleNames | Where-Object {
    $_ -match '(?i)affinity masks not supported|Building device list|Scanning devices'
} | Sort-Object -Unique)
if ($untranslatedUi.Count -gt 0) {
    $smokeFailed = $true
    Write-Report "FAIL: untranslated Russian UI text: $($untranslatedUi -join ' | ')"
} else {
    Write-Report 'PASS: no untranslated non-technical UI text remains after switching to Russian'
}
$canonicalUiTerms = @($ruVisibleNames | Where-Object {
    $_ -match '(?i)^CPU Affinity$|^Affinity Mask:|^MSI Mode:$|^MSI Limit:$|^IRQ Priority:$|^Power Saving:$'
} | Sort-Object -Unique)
if ($canonicalUiTerms.Count -lt 6) {
    $smokeFailed = $true
    Write-Report "FAIL: canonical technical labels missing from Russian UI: found=$($canonicalUiTerms -join ' | ')"
} else {
    Write-Report "PASS: canonical technical labels preserved in Russian UI count=$($canonicalUiTerms.Count)"
}
$storageNoteRu = @($ruVisibleNames | Where-Object {
    $_ -match '(?i)^Affinity Mask:\s*(\u041F\u043E\s+\u0443\u043C\u043E\u043B\u0447\u0430\u043D\u0438\u044E|По умолчанию)\s+Windows'
})
if ($storageNoteRu.Count -lt 1) {
    $smokeFailed = $true
    Write-Report 'FAIL: Storage affinity mask text is not exact in Russian'
} else {
    Write-Report 'PASS: Storage affinity mask preserves natural Russian typography'
}
Capture-Window (Join-Path $outDir '11_normal_ru.png') $normalMain
$refreshButtonRu = Find-AutomationId $normalMain 'REFRESH' 5
$ruSubtitle = Find-DescNamePattern $normalMain '0\.0\.4-alpha\.2.+@arsenza$' 5
if ($null -eq $ruSubtitle) { throw 'Russian version subtitle not found before REFRESH' }
Invoke-Click $refreshButtonRu 'language/RU refresh progress' 20
for ($refreshFrame = 0; $refreshFrame -lt 8; $refreshFrame++) {
    $refreshPath = Join-Path $outDir ('11_refresh_ru_{0:D2}.png' -f $refreshFrame)
    Capture-Window $refreshPath $normalMain
    # WinForms may recreate the LinkLabel accessibility handle while the
    # device tree is rebuilt. Reacquire it for every frame so the assertion
    # checks the visible subtitle rather than a stale UIA object.
    $liveRuSubtitle = Find-DescNamePattern $normalMain '0\.0\.4-alpha\.2.+@arsenza$' 1
    Assert-ElementRendered $refreshPath $normalMain $liveRuSubtitle "REFRESH subtitle frame=$refreshFrame"
    if ($refreshFrame -eq 0) {
        Copy-Item -LiteralPath $refreshPath -Destination (Join-Path $outDir '11_refresh_ru_loading.png') -Force
    }
    Start-Sleep -Milliseconds 100
}
$ruBusyEnglish = @(
    $normalMain.FindAll(
        [System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.Condition]::TrueCondition) |
        ForEach-Object { [string]$_.Current.Name } |
        Where-Object { $_ -match '(?i)Scanning devices|Clearing device list|Enumerating devices|Building device list|Building reserved CPU sets|Laying out devices|Updating (IMOD|IRQ)' } |
        Sort-Object -Unique
)
if ($ruBusyEnglish.Count -gt 0) {
    $smokeFailed = $true
    Write-Report "FAIL: untranslated Russian progress text: $($ruBusyEnglish -join ' | ')"
} else {
    Write-Report 'PASS: device-loading progress is translated in Russian'
}
$refreshReady = $false
for ($i = 0; $i -lt 80; $i++) {
    $refreshButtonRu = Find-AutomationId $normalMain 'REFRESH' 1
    if ($refreshButtonRu -and $refreshButtonRu.Current.IsEnabled) {
        $refreshReady = $true
        break
    }
    Start-Sleep -Milliseconds 125
}
if (-not $refreshReady) { throw 'Russian REFRESH did not finish in time' }
$afterRefreshRect = $normalMain.Current.BoundingRectangle
if ([Math]::Abs($afterRefreshRect.Width - $readyWindowRect.Width) -gt 1 -or
    [Math]::Abs($afterRefreshRect.Height - $readyWindowRect.Height) -gt 1) {
    $smokeFailed = $true
    Write-Report 'FAIL: REFRESH changed the main window size'
} else {
    Write-Report 'PASS: REFRESH preserves the main window size'
}
$restoreButtonRu = Find-AutomationId $normalMain 'RESTORE' 5
Invoke-Click $restoreButtonRu 'language/RU restore dialog'
$restoreDialogRu = Find-AutomationId $normalMain 'RESTORE_DIALOG' 8
if ($null -eq $restoreDialogRu) {
    $smokeFailed = $true
    Write-Report 'FAIL: Russian RESTORE dialog did not open'
} else {
    Write-Report 'PASS: Russian RESTORE dialog opened'
    Capture-Window (Join-Path $outDir '12_restore_ru.png') $restoreDialogRu
    Invoke-Click (Find-AutomationId $restoreDialogRu 'CANCEL' 5) 'language/RU restore cancel'
}
$enButton = Find-AutomationId $normalMain 'LANGUAGE_EN' 5
Invoke-Click $enButton 'language/EN'
$englishApply = Find-Desc $normalMain 'APPLY' 'Button' 5
if ($null -eq $englishApply) {
    $smokeFailed = $true
    Write-Report 'FAIL: English UI did not expose APPLY after switching back'
} else {
    Write-Report 'PASS: runtime language switch RU -> EN'
}
if (-not $normalProc.HasExited) {
    Stop-Process -Id $normalProc.Id -Force
    $normalProc.WaitForExit(5000)
    Write-Report "STOP NORMAL pid=$($normalProc.Id)"
}
Start-Sleep -Seconds 2

$env:DEVICE_TWEAKER_QA_TEST_ADMIN = '1'
$env:DEVICE_TWEAKER_QA_SANDBOX = '1'
$proc = Start-Process -FilePath $exe -WorkingDirectory (Split-Path $exe) -PassThru
Write-Report "START pid=$($proc.Id) QA_TEST_ADMIN=1 QA_SANDBOX=1"
Start-Sleep -Seconds 6

$main = Get-MainWindow -processId $proc.Id -timeoutSec 30
if ($null -eq $main) { throw 'Main window not found' }
$hwnd = [IntPtr]$main.Current.NativeWindowHandle
[SmokeNative]::ShowWindow($hwnd, [SmokeNative]::SW_RESTORE) | Out-Null

# A second launch must activate the first window and exit without creating a
# second long-running DEVICE TWEAKER process.
$secondProc = Start-Process -FilePath $exe -WorkingDirectory (Split-Path $exe) -PassThru
if (-not $secondProc.WaitForExit(8000)) {
    $smokeFailed = $true
    Write-Report "FAIL: second instance did not exit pid=$($secondProc.Id)"
    Stop-Process -Id $secondProc.Id -Force -ErrorAction SilentlyContinue
} else {
    Write-Report "PASS: second instance exited code=$($secondProc.ExitCode)"
}
$runningInstances = @(Get-Process | Where-Object { $_.ProcessName -eq $proc.ProcessName })
if ($runningInstances.Count -ne 1) {
    $smokeFailed = $true
    Write-Report "FAIL: expected one running instance, found=$($runningInstances.Count)"
} else {
    Write-Report 'PASS: exactly one DEVICE TWEAKER instance remains'
}

$mainRect = $main.Current.BoundingRectangle
$mainCenter = New-Object System.Drawing.Point(
    [int]($mainRect.Left + ($mainRect.Width / 2)),
    [int]($mainRect.Top + ($mainRect.Height / 2)))
$mainScreen = [System.Windows.Forms.Screen]::FromPoint($mainCenter)
if ($mainScreen.DeviceName -ne $targetScreen.DeviceName) {
    throw "QA window opened on $($mainScreen.DeviceName), expected $($targetScreen.DeviceName)"
}
Write-Report "PASS: QA main window remained on $($targetScreen.DeviceName)"
Start-Sleep -Milliseconds 300

# Owned modal dialog is a descendant of the main window, not a root child.
$admin = Find-Desc $main 'TEST ADMIN' 'Window' 25
if ($null -eq $admin) { throw 'TEST ADMIN dialog not found under main window' }
Write-Report 'OPEN: TEST ADMIN'
$adminRect = $admin.Current.BoundingRectangle
$adminPoint = New-Object System.Drawing.Point(
    [int]($adminRect.Left + ($adminRect.Width / 2)),
    [int]($adminRect.Top + ($adminRect.Height / 2)))
$adminScreen = [System.Windows.Forms.Screen]::FromPoint($adminPoint)
if ($adminScreen.DeviceName -ne $targetScreen.DeviceName) {
    $smokeFailed = $true
    Write-Report "FAIL: TEST ADMIN opened on $($adminScreen.DeviceName), expected $($targetScreen.DeviceName)"
} else {
    Write-Report "PASS: TEST ADMIN opened on $($targetScreen.DeviceName)"
}
# Wait for sandbox preset to populate fake devices.
$sandboxReady = $false
for ($i = 0; $i -lt 40; $i++) {
    $listCond = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
        [System.Windows.Automation.ControlType]::List)
    foreach ($el in $admin.FindAll([System.Windows.Automation.TreeScope]::Descendants, $listCond)) {
        $n = [string]$el.Current.Name
        if ($n -match 'Current test devices:\s*(\d+)') {
            $count = [int]$Matches[1]
            Write-Report "TEST DEVICES count=$count"
            if ($count -gt 0) { $sandboxReady = $true; break }
        }
    }
    if ($sandboxReady) { break }
    Start-Sleep -Milliseconds 500
    # refresh admin handle
    $admin = Find-Desc $main 'TEST ADMIN' 'Window' 2
    if ($null -eq $admin) { break }
}
if (-not $sandboxReady) { Write-Report 'WARN: sandbox test devices still 0 after wait' }

# QA startup runs the hidden core scenario matrix on the UI thread. Wait for
# its explicit completion marker before visual capture or button automation.
$matrixReady = $false
for ($i = 0; $i -lt 60; $i++) {
    $liveLog = Get-ChildItem -LiteralPath (Join-Path (Split-Path $exe) 'logs') -Filter 'DeviceTweaker_*.log' -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($liveLog -and (Select-String -LiteralPath $liveLog.FullName -Pattern 'TEST.QA.MATRIX: status=' -SimpleMatch -Quiet -ErrorAction SilentlyContinue)) {
        $matrixReady = $true
        break
    }
    Start-Sleep -Milliseconds 500
}
if (-not $matrixReady) {
    $smokeFailed = $true
    Write-Report 'FAIL: sandbox matrix did not complete before GUI capture'
} else {
    Write-Report 'PASS: sandbox matrix completed before GUI capture'
}

# Capture the populated device lists first. This verifies the custom scrollbar
# and long-item presentation instead of checking Scenario Lab alone.
$listsViewReady = $false
for ($i = 0; $i -lt 40; $i++) {
    $liveLog = Get-ChildItem -LiteralPath (Join-Path (Split-Path $exe) 'logs') -Filter 'DeviceTweaker_*.log' -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($liveLog -and (Select-String -LiteralPath $liveLog.FullName -Pattern 'TEST.QA.LISTS.VIEW:' -SimpleMatch -Quiet -ErrorAction SilentlyContinue)) {
        $listsViewReady = $true
        break
    }
    Start-Sleep -Milliseconds 250
}
if (-not $listsViewReady) {
    $smokeFailed = $true
    Write-Report 'FAIL: test-device lists were not positioned for visual capture'
} else {
    Write-Report 'PASS: test-device lists positioned for visual capture'
}
$admin = Find-Desc $main 'TEST ADMIN' 'Window' 5
Capture-Window (Join-Path $outDir '20_test_admin_lists.png') $admin

# Do not capture the top of the long admin form and assume the lower Scenario
# Lab is fine. Wait until the application deliberately scrolls that section.
$scenarioViewReady = $false
for ($i = 0; $i -lt 40; $i++) {
    $liveLog = Get-ChildItem -LiteralPath (Join-Path (Split-Path $exe) 'logs') -Filter 'DeviceTweaker_*.log' -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($liveLog -and (Select-String -LiteralPath $liveLog.FullName -Pattern 'TEST.QA.SCENARIO.VIEW:' -SimpleMatch -Quiet -ErrorAction SilentlyContinue)) {
        $scenarioViewReady = $true
        break
    }
    Start-Sleep -Milliseconds 250
}
if (-not $scenarioViewReady) {
    $smokeFailed = $true
    Write-Report 'FAIL: Scenario Lab was not positioned for visual capture'
} else {
    Write-Report 'PASS: Scenario Lab positioned for visual capture'
}
$admin = Find-Desc $main 'TEST ADMIN' 'Window' 5

Capture-Window (Join-Path $outDir '20_test_admin.png') $admin

# The panel is intentionally long, so its persistent section navigation is a
# release requirement rather than a cosmetic extra. Exercise every destination
# explicitly and capture it; this catches broken anchors, stale scroll offsets
# and navigation buttons that disappear after a layout rebuild.
$adminNavigation = @(
    @{ Id = 'TEST_ADMIN_NAV_CPU'; Name = 'CPU'; File = '20_test_admin_nav_cpu.png' },
    @{ Id = 'TEST_ADMIN_NAV_DEVICES'; Name = 'DEVICES'; File = '20_test_admin_nav_devices.png' },
    @{ Id = 'TEST_ADMIN_NAV_SCENARIO'; Name = 'SCENARIO'; File = '20_test_admin_nav_scenario.png' },
    @{ Id = 'TEST_ADMIN_NAV_RESULTS'; Name = 'RESULTS'; File = '20_test_admin_nav_results.png' }
)
foreach ($destination in $adminNavigation) {
    $admin = Find-Desc $main 'TEST ADMIN' 'Window' 5
    $navigationButton = Find-AutomationId $admin $destination.Id 5
    if ($null -eq $navigationButton) {
        $smokeFailed = $true
        Write-Report "FAIL: TEST ADMIN navigation button missing: $($destination.Name)"
        continue
    }

    Invoke-Click $navigationButton "TEST ADMIN/NAV/$($destination.Name)"
    Start-Sleep -Milliseconds 350
    $admin = Find-Desc $main 'TEST ADMIN' 'Window' 5
    Capture-Window (Join-Path $outDir $destination.File) $admin
}

# Repeat the complete tab cycle to catch stateful layout faults that only
# appear after controls have been hidden, shown and relaid out several times.
for ($navigationCycle = 1; $navigationCycle -le 3; $navigationCycle++) {
    foreach ($destination in $adminNavigation) {
        $admin = Find-Desc $main 'TEST ADMIN' 'Window' 5
        Invoke-Click (Find-AutomationId $admin $destination.Id 5) "TEST ADMIN/NAV STABILITY $navigationCycle/$($destination.Name)"
        Start-Sleep -Milliseconds 120
    }
}
Write-Report 'PASS: TEST ADMIN navigation stability cycles=3 destinations=4'

$setOnly = Find-Desc $admin 'Show test devices only' 'CheckBox' 3
if ($setOnly) { Set-ToggleOn $setOnly 'Show test devices only' }
$setDry = Find-Desc $admin 'Sandbox dry-run (no registry writes)' 'CheckBox' 3
if ($setDry) { Set-ToggleOn $setDry 'dry-run' }
Capture-Window (Join-Path $outDir '21_test_admin_armed.png') $admin

# Release UI gate: exercise every operation-result state without touching hardware.
$resultPreviews = @(
    @{ Button = 'RESULT: SUCCESS'; Title = 'AUTO-OPTIMIZATION'; File = '21_result_success.png'; Details = $false },
    @{ Button = 'RESULT: WARNINGS'; Title = 'APPLY'; File = '21_result_warnings.png'; Details = $false },
    @{ Button = 'RESULT: PARTIAL'; Title = 'AUTO-OPTIMIZATION'; File = '21_result_partial.png'; Details = $true },
    @{ Button = 'RESULT: PARTIAL (3x)'; Title = 'AUTO-OPTIMIZATION'; File = '21_result_partial_3x.png'; Details = $false },
    @{ Button = 'RESULT: FAILED'; Title = 'APPLY'; File = '21_result_failed.png'; Details = $true },
    @{ Button = 'RESULT: STRESS'; Title = 'GUI STRESS TEST'; File = '21_result_stress.png'; Details = $true },
    @{ Button = 'PROMPT: IMOD'; Title = 'USB IMOD TUNING'; File = '21_prompt_imod.png'; Details = $false; Dismiss = 'SKIP' },
    @{ Button = 'PROMPT: CONFIRM'; Title = 'RESET ADAPTER SETTINGS'; File = '21_prompt_confirm.png'; Details = $false; Dismiss = 'CANCEL' },
    @{ Button = 'PROMPT: BACKUP'; Title = 'AUTO BACKUP'; File = '21_prompt_backup.png'; Details = $false; Dismiss = 'SKIP' },
    @{ Button = 'PROMPT: RESTORE'; Title = 'RESTORE'; File = '21_prompt_restore.png'; Details = $false; Dismiss = 'CANCEL' },
    @{ Button = 'PROMPT: RESTORE (0)'; Title = 'RESTORE'; File = '21_prompt_restore_empty.png'; Details = $false; Dismiss = 'CANCEL' },
    @{ Button = 'PROMPT: INFO'; Title = 'TEST MODE INFO'; File = '21_prompt_info.png'; Details = $false; Dismiss = 'OK' }
)
foreach ($preview in $resultPreviews) {
    $admin = Find-Desc $main 'TEST ADMIN' 'Window' 5
    Invoke-Click (Find-Desc $admin $preview.Button 'Button' 5) "TEST ADMIN/$($preview.Button)"
    $resultWindow = Find-Desc $main $preview.Title 'Window' 8
    if ($null -eq $resultWindow) { throw "Result preview not found: $($preview.Title)" }
    $resultRect = $resultWindow.Current.BoundingRectangle
    $resultPoint = New-Object System.Drawing.Point(
        [int]($resultRect.Left + ($resultRect.Width / 2)),
        [int]($resultRect.Top + ($resultRect.Height / 2)))
    $resultScreen = [System.Windows.Forms.Screen]::FromPoint($resultPoint)
    Capture-Window (Join-Path $outDir $preview.File) $resultWindow
    if ($preview.Details) {
        Invoke-Click (Find-Desc $resultWindow 'DETAILS' 'Button' 3) "$($preview.Title)/DETAILS"
        Capture-Window (Join-Path $outDir ($preview.File -replace '\.png$', '_details.png')) $resultWindow
    }
    $dismissName = if ($preview.Dismiss) { $preview.Dismiss } else { 'OK' }
    Invoke-Click (Find-Desc $resultWindow $dismissName 'Button' 3) "$($preview.Title)/$dismissName"
    Start-Sleep -Milliseconds 500
}

Invoke-Click (Find-Desc $admin 'CLOSE' 'Button' 5) 'TEST ADMIN/CLOSE'
Start-Sleep -Seconds 3

$main = Get-MainWindow -processId $proc.Id -timeoutSec 10
Capture-Window (Join-Path $outDir '22_test_only_main.png') $main

# Walk the complete synthetic device list deterministically. These keys are
# handled only while the hidden QA sandbox is active, so physical input and the
# normal application path are untouched.
$VK_HOME = 0x24
$VK_NEXT = 0x22
Send-WindowKey $main $VK_HOME 'devices/EN/HOME'
$enTopPath = Join-Path $outDir '22_devices_en_top.png'
Capture-Window $enTopPath $main
$previousPageHash = (Get-FileHash -LiteralPath $enTopPath -Algorithm SHA256).Hash
$enReachedEnd = $false
for ($page = 1; $page -le 12; $page++) {
    Send-WindowKey $main $VK_NEXT "devices/EN/PAGEDOWN/$page"
    $pagePath = Join-Path $outDir ('22_devices_en_page_{0:D2}.png' -f $page)
    Capture-Window $pagePath $main
    $pageHash = (Get-FileHash -LiteralPath $pagePath -Algorithm SHA256).Hash
    if ($pageHash -eq $previousPageHash) {
        Remove-Item -LiteralPath $pagePath
        $enReachedEnd = $true
        Write-Report "PASS: EN device walk reached the end after $($page - 1) distinct pages"
        break
    }
    $previousPageHash = $pageHash
}
if (-not $enReachedEnd) {
    $smokeFailed = $true
    Write-Report 'FAIL: EN device walk did not reach the end within 12 pages'
}
Send-WindowKey $main $VK_HOME 'devices/EN/HOME-RESET'

# Switch the complete synthetic device matrix to Russian. This covers USB,
# storage, audio, GPU and both network paths instead of validating only the
# first real device visible at startup.
Invoke-Click (Find-AutomationId $main 'LANGUAGE_RU' 5) 'test-matrix language/RU'
$matrixEnglishLabels = @(
    $main.FindAll(
        [System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.Condition]::TrueCondition) |
        Where-Object {
            $type = [string]$_.Current.ControlType.ProgrammaticName
            $name = [string]$_.Current.Name
            $type -match 'ControlType\.(Text|Document)' -and
            $name -match '(?i)affinity masks not supported|Sandbox|^(current|default|devices|detail|raw|time|map|preview|queues)(?: |):'
        } |
        ForEach-Object { [string]$_.Current.Name } |
        Sort-Object -Unique
)
if ($matrixEnglishLabels.Count -gt 0) {
    $smokeFailed = $true
    Write-Report "FAIL: untranslated labels in Russian synthetic device matrix: $($matrixEnglishLabels -join ' | ')"
} else {
    Write-Report 'PASS: Russian labels cover the complete synthetic device matrix'
}
Capture-Window (Join-Path $outDir '22_test_only_main_ru.png') $main
Send-WindowKey $main $VK_HOME 'devices/RU/HOME'
$ruTopPath = Join-Path $outDir '22_devices_ru_top.png'
Capture-Window $ruTopPath $main
$previousPageHash = (Get-FileHash -LiteralPath $ruTopPath -Algorithm SHA256).Hash
$ruReachedEnd = $false
for ($page = 1; $page -le 12; $page++) {
    Send-WindowKey $main $VK_NEXT "devices/RU/PAGEDOWN/$page"
    $pagePath = Join-Path $outDir ('22_devices_ru_page_{0:D2}.png' -f $page)
    Capture-Window $pagePath $main
    $pageHash = (Get-FileHash -LiteralPath $pagePath -Algorithm SHA256).Hash
    if ($pageHash -eq $previousPageHash) {
        Remove-Item -LiteralPath $pagePath
        $ruReachedEnd = $true
        Write-Report "PASS: RU device walk reached the end after $($page - 1) distinct pages"
        break
    }
    $previousPageHash = $pageHash
}
if (-not $ruReachedEnd) {
    $smokeFailed = $true
    Write-Report 'FAIL: RU device walk did not reach the end within 12 pages'
}
Send-WindowKey $main $VK_HOME 'devices/RU/HOME-RESET'
Invoke-Click (Find-AutomationId $main 'LANGUAGE_EN' 5) 'test-matrix language/EN'

foreach ($btnName in @('REFRESH', 'AUTO-OPTIMIZATION', 'APPLY')) {
    $main = Get-MainWindow -processId $proc.Id -timeoutSec 5
    Start-Sleep -Milliseconds 300
    Invoke-Click (Find-Desc $main $btnName 'Button' 5) "main/$btnName"
    Start-Sleep -Seconds 1
    Dismiss-Dialogs $main 10
    Start-Sleep -Seconds 1
    Capture-Window (Join-Path $outDir ('23_{0}.png' -f ($btnName -replace '[^A-Za-z0-9]+','_'))) $main
}

$main = Get-MainWindow -processId $proc.Id -timeoutSec 5
$restoreButton = Find-Desc $main 'RESTORE' 'Button' 5
if ($null -eq $restoreButton) {
    $smokeFailed = $true
    Write-Report 'FAIL: main/RESTORE button not found'
} else {
    Write-Report 'PASS: main/RESTORE button found'
    Invoke-Click $restoreButton 'main/RESTORE'
    $restoreWindow = Find-Desc $main 'RESTORE' 'Window' 8
    if ($null -eq $restoreWindow) {
        $smokeFailed = $true
        Write-Report 'FAIL: RESTORE dialog not found'
    } else {
        Capture-Window (Join-Path $outDir '24_RESTORE_dialog.png') $restoreWindow
        Invoke-Click (Find-Desc $restoreWindow 'CANCEL' 'Button' 3) 'RESTORE/CANCEL'
        Write-Report 'PASS: RESTORE dialog opened and cancelled without changes'
    }
}

$logDir = Join-Path (Split-Path $exe) 'logs'
$latestLog = Get-ChildItem -LiteralPath $logDir -Filter 'DeviceTweaker_*.log' |
    Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($latestLog) {
    Copy-Item -LiteralPath $latestLog.FullName -Destination (Join-Path $outDir $latestLog.Name) -Force
    Write-Report "LOG: $($latestLog.FullName)"
    foreach ($p in @(
        'TEST.QA.SANDBOX',
        'TEST.QA.MATRIX',
        'TEST.QA.LOCALIZATION',
        'TEST.QA.IMOD.STARTUP',
        'TEST.QA.BACKUP.IMOD',
        'TEST.QA.SCROLL',
        'TEST.QA.AFFINITY.8940HX',
        'TEST.QA.MULTI_CONTROLLER',
        'TEST.SCENARIO.END',
        'autoDryRun=True',
        'testDevicesOnly=True',
        'AUTO.DRYRUN',
        'AUTO.POWER',
        'APPLY.PREVIEW',
        'APPLY.DRYRUN',
        'USB.SUSPEND.PLAN: skipped (no real USB blocks)',
        'UI: APPLY dry-run completed',
        'UI: AUTO-OPTIMIZATION',
        'UI: REFRESH',
        'NIC.ITR.CHECK',
        'IMOD.CHECK',
        'KDU',
        'APPLY.REG'
    )) {
        $hits = @(Select-String -LiteralPath $latestLog.FullName -Pattern $p -SimpleMatch -ErrorAction SilentlyContinue)
        Write-Report ("LOGHIT {0} => {1}" -f $p, $hits.Count)
    }

    $regHits = @(Select-String -LiteralPath $latestLog.FullName -Pattern 'APPLY.REG' -SimpleMatch -ErrorAction SilentlyContinue)
    $planWrite = @(Select-String -LiteralPath $latestLog.FullName -Pattern 'USB.SUSPEND.PLAN: enabled|USB.SUSPEND.PLAN: disabled|SAFE_RESET.SUSPEND.PLAN: USB selective suspend power plan restored' -ErrorAction SilentlyContinue)
    $planSkip = @(Select-String -LiteralPath $latestLog.FullName -Pattern 'USB.SUSPEND.PLAN: skipped \(no real USB blocks\)|SAFE_RESET.SUSPEND.PLAN: skipped \(no real USB blocks\)' -ErrorAction SilentlyContinue)
    if ($regHits.Count -gt 0) { Write-Report "FAIL: unexpected APPLY.REG writes=$($regHits.Count)" }
    if ($regHits.Count -gt 0) { $smokeFailed = $true }
    if ($planWrite.Count -gt 0) {
        $smokeFailed = $true
        Write-Report "FAIL: unexpected real USB power-plan writes=$($planWrite.Count)"
    }
    Write-Report ("LOGHIT power-plan-skip => {0}" -f $planSkip.Count)
    if ($regHits.Count -eq 0 -and $planWrite.Count -eq 0) { Write-Report 'PASS: no registry/power-plan writes in this smoke' }

    $policyPass = @(Select-String -LiteralPath $latestLog.FullName -Pattern 'TEST.AUTO.POLICY.FINAL: status=PASS' -SimpleMatch -ErrorAction SilentlyContinue)
    $policyFail = @(Select-String -LiteralPath $latestLog.FullName -Pattern 'TEST.AUTO.POLICY.' -SimpleMatch -ErrorAction SilentlyContinue | Where-Object { $_.Line -match 'status=FAIL' })
    $imodDeclined = @(Select-String -LiteralPath $latestLog.FullName -Pattern 'AUTO.IMOD.RESULT: status=skipped reason=user-declined' -SimpleMatch -ErrorAction SilentlyContinue)
    $imodMapClip = @(Select-String -LiteralPath $latestLog.FullName -Pattern 'imodMapClip=True' -SimpleMatch -ErrorAction SilentlyContinue)
    $layoutIssues = @(Select-String -LiteralPath $latestLog.FullName -Pattern 'GUI.LAYOUT.ISSUE:' -SimpleMatch -ErrorAction SilentlyContinue)
    $adminLayoutPass = @(Select-String -LiteralPath $latestLog.FullName -Pattern 'TEST.ADMIN.LAYOUT: status=PASS' -SimpleMatch -ErrorAction SilentlyContinue)
    $adminLayoutFail = @(Select-String -LiteralPath $latestLog.FullName -Pattern 'TEST.ADMIN.LAYOUT: status=FAIL' -SimpleMatch -ErrorAction SilentlyContinue)
    $matrixPass = @(Select-String -LiteralPath $latestLog.FullName -Pattern 'TEST.QA.MATRIX: status=PASS' -SimpleMatch -ErrorAction SilentlyContinue)
    $matrixFail = @(Select-String -LiteralPath $latestLog.FullName -Pattern 'TEST.MATRIX.FINAL: status=FAIL' -SimpleMatch -ErrorAction SilentlyContinue)
    $localizationPass = @(Select-String -LiteralPath $latestLog.FullName -Pattern 'TEST.QA.LOCALIZATION: status=PASS' -SimpleMatch -ErrorAction SilentlyContinue)
    $localizationFail = @(Select-String -LiteralPath $latestLog.FullName -Pattern 'TEST.QA.LOCALIZATION.FAIL:' -SimpleMatch -ErrorAction SilentlyContinue)
    $imodStartupPass = @(Select-String -LiteralPath $latestLog.FullName -Pattern 'TEST.QA.IMOD.STARTUP: status=PASS' -SimpleMatch -ErrorAction SilentlyContinue)
    $imodStartupFail = @(Select-String -LiteralPath $latestLog.FullName -Pattern 'TEST.QA.IMOD.STARTUP.FAIL:' -SimpleMatch -ErrorAction SilentlyContinue)
    $imodBackupPass = @(Select-String -LiteralPath $latestLog.FullName -Pattern 'TEST.QA.BACKUP.IMOD: status=PASS' -SimpleMatch -ErrorAction SilentlyContinue)
    $imodBackupFail = @(Select-String -LiteralPath $latestLog.FullName -Pattern 'TEST.QA.BACKUP.IMOD.FAIL:' -SimpleMatch -ErrorAction SilentlyContinue)
    $qaScroll = @(Select-String -LiteralPath $latestLog.FullName -Pattern 'TEST.QA.SCROLL:' -SimpleMatch -ErrorAction SilentlyContinue)
    $affinity8940Pass = @(Select-String -LiteralPath $latestLog.FullName -Pattern 'TEST.QA.AFFINITY.8940HX: status=PASS' -SimpleMatch -ErrorAction SilentlyContinue)
    $affinity8940Fail = @(Select-String -LiteralPath $latestLog.FullName -Pattern 'TEST.QA.AFFINITY.8940HX: status=FAIL' -SimpleMatch -ErrorAction SilentlyContinue)
    $multiControllerPass = @(Select-String -LiteralPath $latestLog.FullName -Pattern 'TEST.QA.MULTI_CONTROLLER: status=PASS' -SimpleMatch -ErrorAction SilentlyContinue)
    $multiControllerFail = @(Select-String -LiteralPath $latestLog.FullName -Pattern 'TEST.QA.MULTI_CONTROLLER: status=FAIL' -SimpleMatch -ErrorAction SilentlyContinue)
    $adminNavigationHits = @(Select-String -LiteralPath $latestLog.FullName -Pattern 'TEST.ADMIN.NAV:' -SimpleMatch -ErrorAction SilentlyContinue)
    if ($policyPass.Count -eq 0 -or $policyFail.Count -gt 0) {
        $smokeFailed = $true
        Write-Report "FAIL: AUTO policy regression pass=$($policyPass.Count) fail=$($policyFail.Count)"
    } else {
        Write-Report "PASS: AUTO policy regression checks=$($policyPass.Count)"
    }
    if ($imodDeclined.Count -eq 0) {
        $smokeFailed = $true
        Write-Report 'FAIL: exact IMOD user-declined result was not logged'
    } else {
        Write-Report "PASS: exact IMOD user-declined result logged=$($imodDeclined.Count)"
    }
    if ($imodMapClip.Count -gt 0) {
        $smokeFailed = $true
        Write-Report "FAIL: IMOD map clipped in GUI snapshots count=$($imodMapClip.Count)"
    } else {
        Write-Report 'PASS: IMOD map is not clipped in GUI snapshots'
    }
    if ($layoutIssues.Count -gt 0) {
        $smokeFailed = $true
        Write-Report "FAIL: GUI layout issues found=$($layoutIssues.Count)"
    } else {
        Write-Report 'PASS: no GUI layout issues in snapshots'
    }
    if ($adminLayoutFail.Count -gt 0 -or $adminLayoutPass.Count -lt 3) {
        $smokeFailed = $true
        Write-Report "FAIL: TEST ADMIN layout pass=$($adminLayoutPass.Count) fail=$($adminLayoutFail.Count)"
    } else {
        Write-Report "PASS: TEST ADMIN layout checkpoints=$($adminLayoutPass.Count)"
    }
    if ($matrixPass.Count -eq 0 -or $matrixFail.Count -gt 0) {
        $smokeFailed = $true
        Write-Report "FAIL: sandbox core matrix pass=$($matrixPass.Count) fail=$($matrixFail.Count)"
    } else {
        Write-Report "PASS: sandbox core matrix completed=$($matrixPass.Count)"
    }
    if ($localizationPass.Count -eq 0 -or $localizationFail.Count -gt 0) {
        $smokeFailed = $true
        Write-Report "FAIL: Russian structured localization contract pass=$($localizationPass.Count) fail=$($localizationFail.Count)"
    } else {
        Write-Report "PASS: Russian structured localization contract completed=$($localizationPass.Count)"
    }
    if ($imodStartupPass.Count -eq 0 -or $imodStartupFail.Count -gt 0) {
        $smokeFailed = $true
        Write-Report "FAIL: IMOD startup persistence contract pass=$($imodStartupPass.Count) fail=$($imodStartupFail.Count)"
    } else {
        Write-Report "PASS: IMOD startup persistence contract completed=$($imodStartupPass.Count)"
    }
    if ($imodBackupPass.Count -eq 0 -or $imodBackupFail.Count -gt 0) {
        $smokeFailed = $true
        Write-Report "FAIL: IMOD backup persistence contract pass=$($imodBackupPass.Count) fail=$($imodBackupFail.Count)"
    } else {
        Write-Report "PASS: IMOD backup persistence contract completed=$($imodBackupPass.Count)"
    }
    if ($affinity8940Pass.Count -eq 0 -or $affinity8940Fail.Count -gt 0) {
        $smokeFailed = $true
        Write-Report "FAIL: 8940HX Dual-CCD affinity layout pass=$($affinity8940Pass.Count) fail=$($affinity8940Fail.Count)"
    } else {
        Write-Report "PASS: 8940HX Dual-CCD affinity layout completed=$($affinity8940Pass.Count)"
    }
    if ($multiControllerPass.Count -eq 0 -or $multiControllerFail.Count -gt 0) {
        $smokeFailed = $true
        Write-Report "FAIL: Multi-controller affinity pass=$($multiControllerPass.Count) fail=$($multiControllerFail.Count)"
    } else {
        Write-Report "PASS: Multi-controller affinity completed=$($multiControllerPass.Count)"
    }
    if ($qaScroll.Count -lt 6) {
        $smokeFailed = $true
        Write-Report "FAIL: deterministic device-list navigation was not fully exercised entries=$($qaScroll.Count)"
    } else {
        Write-Report "PASS: deterministic device-list navigation entries=$($qaScroll.Count)"
    }
    if ($adminNavigationHits.Count -lt 16) {
        $smokeFailed = $true
        Write-Report "FAIL: TEST ADMIN section navigation entries=$($adminNavigationHits.Count)"
    } else {
        Write-Report "PASS: TEST ADMIN section navigation entries=$($adminNavigationHits.Count)"
    }
}

if (-not $proc.HasExited) {
    Stop-Process -Id $proc.Id -Force
    Write-Report "STOP pid=$($proc.Id)"
}

# Open the hidden QA panel once in Russian and audit its own labels. The main
# device matrix is checked above; this pass covers the administrative sandbox.
Start-Sleep -Milliseconds 800
$env:DEVICE_TWEAKER_LANGUAGE = 'ru'
$env:DEVICE_TWEAKER_QA_SANDBOX = '1'
$ruAdminStartedAt = Get-Date
$ruAdminProc = Start-Process -FilePath $exe -WorkingDirectory (Split-Path $exe) -PassThru
Write-Report "START RU ADMIN pid=$($ruAdminProc.Id)"
try {
    $ruAdminMain = Get-MainWindow -processId $ruAdminProc.Id -timeoutSec 30
    if ($null -eq $ruAdminMain) { throw 'Russian QA main window not found' }
    $ruAdmin = Find-AutomationId $ruAdminMain 'TEST_ADMIN_DIALOG' 30
    if ($null -eq $ruAdmin) {
        $smokeFailed = $true
        Write-Report 'FAIL: Russian TEST ADMIN dialog did not open'
    } else {
        $ruListsReady = $false
        for ($attempt = 0; $attempt -lt 80; $attempt++) {
            $ruLog = Get-ChildItem -LiteralPath (Join-Path (Split-Path $exe) 'logs') -Filter 'DeviceTweaker_*.log' -ErrorAction SilentlyContinue |
                Where-Object { $_.LastWriteTime -ge $ruAdminStartedAt.AddSeconds(-1) } |
                Sort-Object LastWriteTime -Descending | Select-Object -First 1
            if ($ruLog -and (Select-String -LiteralPath $ruLog.FullName -Pattern 'TEST.QA.LISTS.VIEW:' -SimpleMatch -Quiet -ErrorAction SilentlyContinue)) {
                $ruListsReady = $true
                break
            }
            Start-Sleep -Milliseconds 250
        }
        if (-not $ruListsReady) {
            $smokeFailed = $true
            Write-Report 'FAIL: Russian test-device lists were not positioned for visual capture'
        } else {
            Write-Report 'PASS: Russian test-device lists positioned for visual capture'
            $ruAdmin = Find-AutomationId $ruAdminMain 'TEST_ADMIN_DIALOG' 5
            Capture-Window (Join-Path $outDir '25_test_admin_ru_lists.png') $ruAdmin
        }

        $ruAdminEnglish = $null
        for ($attempt = 0; $attempt -lt 20 -and $null -eq $ruAdminEnglish; $attempt++) {
            # The startup QA matrix can rebuild the owned dialog. Always
            # reacquire it instead of retaining a stale AutomationElement.
            $ruAdmin = Find-AutomationId $ruAdminMain 'TEST_ADMIN_DIALOG' 2
            if ($null -eq $ruAdmin) {
                Start-Sleep -Milliseconds 250
                continue
            }
            try {
                $ruAdminEnglish = @(
                    $ruAdmin.FindAll(
                        [System.Windows.Automation.TreeScope]::Descendants,
                        [System.Windows.Automation.Condition]::TrueCondition) |
                        ForEach-Object { [string]$_.Current.Name } |
                        Where-Object {
                            $_ -match '(?i)^(Test CPU Topology|Test CPU mode|CPU preset|System preset|Test Devices|Test options|Add fake device|Enable test devices|Show test devices only|Sandbox dry-run|Device preset|Current test devices|Real device visibility|Visible real devices|Hidden real devices|Scenario Lab|Result Dialog Gallery|CPU TOPOLOGY|DEVICES|SCENARIO LAB|RESULT DIALOGS|APPLY CPU TOPOLOGY|Total logical processors|Group counts|How to use:|Note: UI and affinity masks|Tip: full PC presets)|individual fake devices|temporarily hiding real devices|^Group [0-9]+$'
                        } |
                        Sort-Object -Unique
                )
            } catch [System.Windows.Automation.ElementNotAvailableException] {
                $ruAdminEnglish = $null
                Start-Sleep -Milliseconds 250
            }
        }
        if ($null -eq $ruAdminEnglish) {
            throw 'Russian TEST ADMIN dialog did not remain stable for localization audit.'
        }
        if ($ruAdminEnglish.Count -gt 0) {
            $smokeFailed = $true
            Write-Report "FAIL: untranslated Russian TEST ADMIN labels: $($ruAdminEnglish -join ' | ')"
        } else {
            Write-Report 'PASS: Russian TEST ADMIN labels are localized'
        }
        $ruScenarioReady = $false
        for ($attempt = 0; $attempt -lt 40; $attempt++) {
            $ruLog = Get-ChildItem -LiteralPath (Join-Path (Split-Path $exe) 'logs') -Filter 'DeviceTweaker_*.log' -ErrorAction SilentlyContinue |
                Where-Object { $_.LastWriteTime -ge $ruAdminStartedAt.AddSeconds(-1) } |
                Sort-Object LastWriteTime -Descending | Select-Object -First 1
            if ($ruLog -and (Select-String -LiteralPath $ruLog.FullName -Pattern 'TEST.QA.SCENARIO.VIEW:' -SimpleMatch -Quiet -ErrorAction SilentlyContinue)) {
                $ruScenarioReady = $true
                break
            }
            Start-Sleep -Milliseconds 250
        }
        if (-not $ruScenarioReady) {
            $smokeFailed = $true
            Write-Report 'FAIL: Russian Scenario Lab was not positioned for visual capture'
        } else {
            Write-Report 'PASS: Russian Scenario Lab positioned for visual capture'
        }
        $ruAdmin = Find-AutomationId $ruAdminMain 'TEST_ADMIN_DIALOG' 5
        if ($null -eq $ruAdmin) { throw 'Russian TEST ADMIN dialog disappeared before capture.' }
        Capture-Window (Join-Path $outDir '25_test_admin_ru.png') $ruAdmin

        $ruAdminLayoutPass = @(Select-String -LiteralPath $ruLog.FullName -Pattern 'TEST.ADMIN.LAYOUT: status=PASS' -SimpleMatch -ErrorAction SilentlyContinue)
        $ruAdminLayoutFail = @(Select-String -LiteralPath $ruLog.FullName -Pattern 'TEST.ADMIN.LAYOUT: status=FAIL' -SimpleMatch -ErrorAction SilentlyContinue)
        if ($ruAdminLayoutFail.Count -gt 0 -or $ruAdminLayoutPass.Count -lt 3) {
            $smokeFailed = $true
            Write-Report "FAIL: Russian TEST ADMIN layout pass=$($ruAdminLayoutPass.Count) fail=$($ruAdminLayoutFail.Count)"
        } else {
            Write-Report "PASS: Russian TEST ADMIN layout checkpoints=$($ruAdminLayoutPass.Count)"
        }

        # Prove that emulation has an explicit, working exit path. This is
        # deliberately done only after every sandbox write-safety test.
        $disableTestButton = Find-AutomationId $ruAdmin 'TEST_ADMIN_DISABLE_MODE' 5
        if ($null -eq $disableTestButton) {
            $smokeFailed = $true
            Write-Report 'FAIL: disable-test-mode button not found in Russian panel'
        } else {
            Invoke-Click $disableTestButton 'TEST ADMIN/DISABLE TEST MODE'
            $resetLogged = $false
            for ($attempt = 0; $attempt -lt 40; $attempt++) {
                if (Select-String -LiteralPath $ruLog.FullName -Pattern 'TEST.RESET.REAL: full admin test state cleared' -SimpleMatch -Quiet -ErrorAction SilentlyContinue) {
                    $resetLogged = $true
                    break
                }
                Start-Sleep -Milliseconds 250
            }
            if (-not $resetLogged) {
                $smokeFailed = $true
                Write-Report 'FAIL: test-mode disable action did not complete'
            } else {
                Write-Report 'PASS: test-mode disable action cleared the complete emulation state'
                $ruAdmin = Find-AutomationId $ruAdminMain 'TEST_ADMIN_DIALOG' 5
                Invoke-Click (Find-AutomationId $ruAdmin 'TEST_ADMIN_CLOSE' 5) 'TEST ADMIN/CLOSE after disable'
                Start-Sleep -Seconds 2
                Capture-Window (Join-Path $outDir '26_main_after_test_mode_disabled_ru.png') $ruAdminMain
            }
        }
    }
} finally {
    if (-not $ruAdminProc.HasExited) {
        Stop-Process -Id $ruAdminProc.Id -Force
        Write-Report "STOP RU ADMIN pid=$($ruAdminProc.Id)"
    }
    $env:DEVICE_TWEAKER_LANGUAGE = 'en'
}
Write-Report "DONE report=$reportPath"
if ($smokeFailed) { throw "Smoke checks failed. Report: $reportPath" }
Write-Output $reportPath
