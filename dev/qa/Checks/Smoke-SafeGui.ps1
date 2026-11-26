# Safe GUI smoke for DEVICE TWEAKER (read-only main window, runtime EN/RU, single-instance).
param(
    [string]$Root,
    [string]$ExePath,
    [string]$OutputRoot,
    [switch]$VisibleGui,
    [ValidateSet('Secondary', 'Primary')]
    [string]$Monitor = 'Secondary'
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\Support\Private-Desktop.ps1')
if (-not $VisibleGui) {
    if (-not $OutputRoot) { $OutputRoot = Join-Path $PSScriptRoot '..\..\..\src\tmp\qa-gui' }
    $OutputRoot = [IO.Path]::GetFullPath($OutputRoot)
    $workerParameters = @{} + $PSBoundParameters
    $workerParameters.OutputRoot = $OutputRoot
    if ($ExePath) { $workerParameters.ExePath = [IO.Path]::GetFullPath($ExePath) }
    if ($Root) { $workerParameters.Root = [IO.Path]::GetFullPath($Root) }
    $workerExit = Invoke-QaPrivateDesktop -ScriptPath $PSCommandPath -Parameters $workerParameters `
        -ReportDirectory (Join-Path $OutputRoot 'isolation')
    if ($null -ne $workerExit) { exit $workerExit }
}
. (Join-Path $PSScriptRoot '..\Support\Read-QaJson.ps1')
. (Join-Path $PSScriptRoot '..\Support\Visual-Assertions.ps1')
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
  private delegate bool EnumWindow(IntPtr hwnd, IntPtr data);
  [DllImport("user32.dll")] private static extern bool EnumWindows(EnumWindow callback, IntPtr data);
  [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] private static extern int GetWindowText(IntPtr hwnd, System.Text.StringBuilder text, int count);
  [DllImport("user32.dll")] private static extern IntPtr GetWindow(IntPtr hwnd, uint command);
  public static IntPtr MainWindow(int id) {
    IntPtr found = IntPtr.Zero;
    EnumWindows((hwnd, data) => {
      uint pid; GetWindowThreadProcessId(hwnd, out pid);
      if (pid != id || GetWindow(hwnd, 4) != IntPtr.Zero) return true;
      var text = new System.Text.StringBuilder(256); GetWindowText(hwnd, text, text.Capacity);
      if (text.ToString() != "DEVICE TWEAKER") return true;
      found = hwnd; return false;
    }, IntPtr.Zero);
    return found;
  }
  [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
  [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
  [DllImport("user32.dll", SetLastError=true)] public static extern IntPtr SendMessageTimeout(IntPtr hwnd, uint msg, IntPtr w, IntPtr l, uint flags, uint timeout, out IntPtr result);
  [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr hWnd, IntPtr hdcBlt, uint nFlags);
  [DllImport("user32.dll")] public static extern bool RedrawWindow(IntPtr hwnd, IntPtr rect, IntPtr region, uint flags);
  public const int SW_RESTORE = 9;
  public const uint BM_CLICK = 0x00F5;
  public const uint WM_KEYDOWN = 0x0100;
  public const uint WM_KEYUP = 0x0101;
  public const uint PW_RENDERFULLCONTENT = 0x00000002;
}
'@

$root = if ([string]::IsNullOrWhiteSpace($Root)) {
    (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\src')).Path
} else {
    (Resolve-Path $Root).Path
}
$exe = if ([string]::IsNullOrWhiteSpace($ExePath)) {
    Join-Path $root 'bin\DeveloperQa\Release\net10.0-windows\win-x64\DEVICE TWEAKER QA.exe'
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
$ownedProcesses = New-Object 'System.Collections.Generic.List[System.Diagnostics.Process]'
$checks = New-Object 'System.Collections.Generic.List[object]'
$caseOrdinal = 0
$environmentNames = @('DEVICE_TWEAKER_QA_READ_ONLY', 'DEVICE_TWEAKER_QA_OUTPUT',
    'DEVICE_TWEAKER_QA_SESSION', 'DEVICE_TWEAKER_QA_WINDOW_BOUNDS',
    'DEVICE_TWEAKER_QA_TEST_ADMIN', 'DEVICE_TWEAKER_QA_SANDBOX', 'DEVICE_TWEAKER_LANGUAGE')
$savedEnvironment = @{}
foreach ($name in $environmentNames) { $savedEnvironment[$name] = [Environment]::GetEnvironmentVariable($name) }
$env:DEVICE_TWEAKER_QA_READ_ONLY = '1'
$env:DEVICE_TWEAKER_QA_SESSION = [Guid]::NewGuid().ToString('N')
$env:DEVICE_TWEAKER_QA_OUTPUT = Join-Path $outDir 'processes'
$logDir = Join-Path $env:DEVICE_TWEAKER_QA_OUTPUT 'logs'

function Start-OwnedProcess([switch]$Probe) {
    $arguments = @('--qa-read-only')
    if ($Probe) { $arguments += '--qa-probe' }
    $started = Start-QaOwnedProcess $exe ($arguments -join ' ')
    $ownedProcesses.Add($started)
    return $started
}

function Get-OwnedProcessLog([System.Diagnostics.Process]$owned) {
    $metadata = Join-Path $env:DEVICE_TWEAKER_QA_OUTPUT "process-$($owned.Id).json"
    if (-not (Test-Path -LiteralPath $metadata)) { return $null }
    $path = (Read-QaJson $metadata).logPath
    if ([string]::IsNullOrWhiteSpace($path)) { return $null }
    return Get-Item -LiteralPath $path -ErrorAction SilentlyContinue
}

function Stop-OwnedProcess([System.Diagnostics.Process]$owned) {
    if (-not $ownedProcesses.Contains($owned)) { throw 'Refusing to stop an unowned process' }
    if (-not $owned.HasExited) {
        if (-not (Stop-QaOwnedProcess $owned)) {
            $script:smokeFailed = $true
            Write-Report "FAIL: QA process did not shut down gracefully pid=$($owned.Id)"
        }
    }
}

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

function Write-Report([string]$msg, [string]$Expected = '', [string]$Observed = '') {
    $line = '[{0:HH:mm:ss}] {1}' -f (Get-Date), $msg
    if ($msg -match '^(PASS|FAIL|SKIP): (.+)$') {
        $script:caseOrdinal++
        $checks.Add([pscustomobject]@{ id = ('gui-{0:D3}' -f $script:caseOrdinal);
            status = $Matches[1]; description = $Matches[2]; expected = $(if ($Expected) { $Expected } else { $Matches[2] });
            observed = $(if ($Observed) { $Observed } else { $Matches[2] });
            evidence = $reportPath; time = (Get-Date).ToUniversalTime().ToString('o') })
    }
    Write-Host $line
    Add-Content -LiteralPath $reportPath -Value $line -Encoding UTF8
}

function Get-MainWindow([int]$processId, [int]$timeoutSec = 30) {
    $deadline = (Get-Date).AddSeconds($timeoutSec)
    while ((Get-Date) -lt $deadline) {
        $handle = [SmokeNative]::MainWindow($processId)
        if ($handle -ne [IntPtr]::Zero) { return [System.Windows.Automation.AutomationElement]::FromHandle($handle) }
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
    if ($null -eq $scope) { return $null }
    $scopeHandle = [IntPtr]$scope.Current.NativeWindowHandle
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
        try {
            $el = $scope.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $and)
        } catch {
            $cause = $_.Exception
            while ($cause.InnerException) { $cause = $cause.InnerException }
            if (($cause -isnot [Runtime.InteropServices.COMException] -and
                $cause -isnot [System.Windows.Automation.ElementNotAvailableException]) -or
                (Get-Date) -ge $deadline -or $scopeHandle -eq [IntPtr]::Zero) { throw }
            Write-Report ("UIA RETRY: name=" + $name + ' hresult=' + $cause.HResult)
            Start-Sleep -Milliseconds 150
            $scope = [System.Windows.Automation.AutomationElement]::FromHandle($scopeHandle)
            continue
        }
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
    if (-not [SmokeNative]::PostMessage($hwnd, [SmokeNative]::BM_CLICK, [IntPtr]::Zero, [IntPtr]::Zero)) {
        throw "Button message was rejected: $label"
    }
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
                Where-Object {
                    # A sibling can disappear while REFRESH rebuilds the device tree.
                    # Do not discard a valid header match because another element went stale.
                    try { [string]$_.Current.Name -match $pattern }
                    catch [System.Windows.Automation.ElementNotAvailableException] { $false }
                } |
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
    if (-not $VisibleGui) {
        $handled = [IntPtr]::Zero
        $response = [SmokeNative]::SendMessageTimeout($hwnd, 0x8476, [IntPtr]$virtualKey,
            [IntPtr]::Zero, 2, 10000, [ref]$handled)
        if ($response -eq [IntPtr]::Zero -or $handled -ne [IntPtr]1) { throw "Real navigation handler did not accept: $label" }
        Write-Report ("KEY(HANDLER): {0} vk=0x{1:X2} modifiers=None" -f $label, $virtualKey)
        return
    }
    [SmokeNative]::SetForegroundWindow($hwnd) | Out-Null
    # Synchronous delivery makes the screenshot/hash assertion deterministic.
    # Posting DOWN and UP back-to-back can let capture race the form message
    # queue, falsely reporting the end of a long device list.
    [SmokeNative]::SendMessage($hwnd, [SmokeNative]::WM_KEYDOWN, [IntPtr]$virtualKey, [IntPtr]::Zero) | Out-Null
    [SmokeNative]::SendMessage($hwnd, [SmokeNative]::WM_KEYUP, [IntPtr]$virtualKey, [IntPtr]::Zero) | Out-Null
    Write-Report ("KEY: {0} vk=0x{1:X2}" -f $label, $virtualKey)
    Start-Sleep -Milliseconds 250
}

function Capture-Window([string]$path, [System.Windows.Automation.AutomationElement]$window) {
    $r = $window.Current.BoundingRectangle
    $width = [Math]::Max(1, [int][Math]::Ceiling($r.Width))
    $height = [Math]::Max(1, [int][Math]::Ceiling($r.Height))
    $hwnd = [IntPtr]$window.Current.NativeWindowHandle
    if ($hwnd -eq [IntPtr]::Zero) { throw "Window has no native handle: $path" }
    if (-not [SmokeNative]::RedrawWindow($hwnd, [IntPtr]::Zero, [IntPtr]::Zero, 0x585)) {
        throw "Owned window redraw failed: $path"
    }
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
    $maxWhiteRowFraction = 0.0
    for ($y = 0; $y -lt $height; $y += 6) {
        $white = 0; $samples = 0
        for ($x = 0; $x -lt $width; $x += 6) {
            $pixel = $bmp.GetPixel($x, $y)
            $brightness = [int]$pixel.R + [int]$pixel.G + [int]$pixel.B
            if ($brightness -lt $minBrightness) { $minBrightness = $brightness }
            if ($brightness -gt $maxBrightness) { $maxBrightness = $brightness }
            if ($brightness -gt 90) { $visibleSamples++ }
            if ($pixel.R -gt 240 -and $pixel.G -gt 240 -and $pixel.B -gt 240) { $white++ }
            $samples++
        }
        $maxWhiteRowFraction = [Math]::Max($maxWhiteRowFraction, $white / [double]$samples)
    }
    if (($maxBrightness - $minBrightness) -lt 45 -or $visibleSamples -lt 20 -or $maxWhiteRowFraction -gt 0.35) {
        $bmp.Dispose()
        throw "Captured window is blank or visually invalid: $path range=$($maxBrightness - $minBrightness) visibleSamples=$visibleSamples whiteRow=$maxWhiteRowFraction"
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

    Assert-RegionRendered $path $windowRect $elementRect $label
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

try {
Write-Report "ROOT=$root"
Write-Report "EXE=$exe"
Write-Report ("MONITOR={0} primary={1} workingArea={2},{3},{4}x{5} initialBounds={6}" -f
    $targetScreen.DeviceName, $targetScreen.Primary,
    $targetArea.X, $targetArea.Y, $targetArea.Width, $targetArea.Height,
    $env:DEVICE_TWEAKER_QA_WINDOW_BOUNDS)
# Existing user processes are neither closed nor activated. QA uses its own mutex.
$probe = Start-OwnedProcess -Probe
if (-not $probe.WaitForExit(15000)) { throw 'QA probe did not exit; rebuild with read-only QA support' }
$probePath = Join-Path $env:DEVICE_TWEAKER_QA_OUTPUT "process-$($probe.Id).json"
if ($probe.ExitCode -ne 0 -or -not (Test-Path -LiteralPath $probePath)) { throw "QA read-only handshake missing: exit=$($probe.ExitCode) metadata=$probePath" }
if ($env:DEVICE_TWEAKER_QA_DESKTOP -and (Read-QaJson $probePath).desktop.isolated -ne $true) { throw 'Probe escaped the private desktop' }
$probeReport = Read-QaJson $probePath
if ($probeReport.readOnly -ne $true) { throw 'QA process did not confirm the immutable lock' }
Write-Report 'PASS: QA process confirmed immutable read-only mode' 'readOnly=True; assigned private desktop' `
    ("readOnly=$($probeReport.readOnly) desktop=$($probeReport.desktop.actual)")

Remove-Item Env:DEVICE_TWEAKER_QA_TEST_ADMIN -ErrorAction SilentlyContinue
Remove-Item Env:DEVICE_TWEAKER_QA_SANDBOX -ErrorAction SilentlyContinue
$env:DEVICE_TWEAKER_LANGUAGE = 'en'
$normalProc = Start-OwnedProcess
Write-Report "START NORMAL pid=$($normalProc.Id)"
$normalMain = Get-MainWindow -processId $normalProc.Id -timeoutSec 30
if ($null -eq $normalMain) { throw 'Normal main window not found' }
$englishSubtitlePattern = '^Version .+ · @arsenza$'
$russianSubtitlePattern = '^Версия .+ · @arsenza$'
$startupSubtitle = Find-DescNamePattern $normalMain $englishSubtitlePattern 5
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
    Write-Report 'PASS: real-device view contains no emulation controls; QA read-only indicator is allowed'
}
$normalHwnd = [IntPtr]$normalMain.Current.NativeWindowHandle
[SmokeNative]::ShowWindow($normalHwnd, [SmokeNative]::SW_RESTORE) | Out-Null
if ($VisibleGui) { [SmokeNative]::SetForegroundWindow($normalHwnd) | Out-Null }
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
    Write-Report 'PASS: three known English fallback phrases are absent from the Russian device view'
}
$canonicalUiTerms = @($ruVisibleNames | Where-Object {
    $_ -match '(?i)^CPU Affinity$|^Affinity Mask:|^MSI Mode:$|^MSI Limit:$|^IRQ Priority:$|^Power Saving:$|^\u042D\u043D\u0435\u0440\u0433\u043E\u0441\u0431\u0435\u0440\u0435\u0436\u0435\u043D\u0438\u0435:$'
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
$ruSubtitle = Find-DescNamePattern $normalMain $russianSubtitlePattern 5
if ($null -eq $ruSubtitle) { throw 'Russian version subtitle not found before REFRESH' }
$refreshSubtitleRect = $ruSubtitle.Current.BoundingRectangle
$refreshWindowRect = $normalMain.Current.BoundingRectangle
Invoke-Click $refreshButtonRu 'language/RU refresh progress' 20
for ($refreshFrame = 0; $refreshFrame -lt 8; $refreshFrame++) {
    $refreshPath = Join-Path $outDir ('11_refresh_ru_{0:D2}.png' -f $refreshFrame)
    Capture-Window $refreshPath $normalMain
    # Check every captured frame at the pre-refresh header location. UIA enumeration
    # is transient during device-tree replacement; it is not a rendering verdict.
    Assert-RegionRendered $refreshPath $refreshWindowRect $refreshSubtitleRect "REFRESH subtitle frame=$refreshFrame"
    if ($refreshFrame -eq 0) {
        Copy-Item -LiteralPath $refreshPath -Destination (Join-Path $outDir '11_refresh_ru_loading.png') -Force
    }
    Start-Sleep -Milliseconds 100
}
# Cards are being replaced while REFRESH runs; an element can vanish mid-walk, so the
# enumeration itself is retried. A real untranslated label is seen on every attempt.
$ruBusyEnglish = $null
for ($attempt = 1; $null -eq $ruBusyEnglish; $attempt++) {
    try {
        $ruBusyEnglish = @(
            $normalMain.FindAll(
                [System.Windows.Automation.TreeScope]::Descendants,
                [System.Windows.Automation.Condition]::TrueCondition) |
                ForEach-Object { [string]$_.Current.Name } |
                Where-Object { $_ -match '(?i)Scanning devices|Clearing device list|Enumerating devices|Building device list|Building reserved CPU sets|Laying out devices|Updating (IMOD|IRQ)' } |
                Sort-Object -Unique
        )
    } catch {
        if ($attempt -ge 5) { throw }
        Start-Sleep -Milliseconds 100
    }
}
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
$subtitleAfterRefresh = Find-DescNamePattern $normalMain $russianSubtitlePattern 5
if ($null -eq $subtitleAfterRefresh) { throw 'Russian version subtitle not found after REFRESH' }
$subtitleAfterRect = $subtitleAfterRefresh.Current.BoundingRectangle
foreach ($coordinate in @('Left','Top','Width','Height')) {
    if ([Math]::Abs($subtitleAfterRect.$coordinate - $refreshSubtitleRect.$coordinate) -gt 1) {
        throw "REFRESH changed the subtitle position: $coordinate"
    }
}
$refreshFinalPath = Join-Path $outDir '11_refresh_ru_ready.png'
Capture-Window $refreshFinalPath $normalMain
Assert-ElementRendered $refreshFinalPath $normalMain $subtitleAfterRefresh 'REFRESH subtitle ready'
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
$normalModeCheck = Find-AutomationId $normalMain 'IRQ_MODE_CHECK' 5
$normalModeValue = Find-AutomationId $normalMain 'IRQ_MODE_VALUE' 5
if (-not $normalModeCheck -or -not $normalModeValue) { throw 'Production IRQ Mode controls missing' }
$beforeModeCheck = [string]$normalModeValue.Current.Name
Invoke-Click $normalModeCheck 'read-only IRQ Mode CHECK'
$normalModeValue = Find-AutomationId $normalMain 'IRQ_MODE_VALUE' 5
if ([string]$normalModeValue.Current.Name -ne $beforeModeCheck) { throw 'Read-only mode query changed the hardware result' }
Write-Report 'PASS: production IRQ Mode CHECK is blocked in read-only GUI'


# Verify single-instance enforcement while the production read-only window is active.
$secondProc = Start-OwnedProcess
if (-not $secondProc.WaitForExit(20000)) {
    $smokeFailed = $true
    Write-Report "FAIL: second instance did not exit pid=$($secondProc.Id)"
    Stop-OwnedProcess $secondProc
} else {
    Write-Report "PASS: second instance exited code=$($secondProc.ExitCode)"
}
$runningInstances = @($ownedProcesses | Where-Object { -not $_.HasExited })
if ($runningInstances.Count -ne 1) {
    $smokeFailed = $true
    Write-Report "FAIL: expected one running instance, found=$($runningInstances.Count)"
} else {
    Write-Report 'PASS: exactly one DEVICE TWEAKER instance remains'
}

if (-not $normalProc.HasExited) {
    Stop-OwnedProcess $normalProc
    $normalProc.WaitForExit(5000) | Out-Null
    Write-Report "STOP NORMAL pid=$($normalProc.Id)"
}
Write-Report "DONE report=$reportPath"
if ($smokeFailed) { throw "Smoke checks failed. Report: $reportPath" }
Write-Output $reportPath
} catch {
    $smokeFailed = $true
    Write-Report ("FAIL: automation exception: " + $_.Exception.Message)
    throw
} finally {
    foreach ($owned in $ownedProcesses) { Stop-OwnedProcess $owned }
    $jsonReport = [ordered]@{
        schemaVersion = 2
        status = $(if ($smokeFailed) { 'FAIL' } else { 'PASS' })
        scope = 'Read-only main window smoke (startup geometry, EN/RU switch, REFRESH, RESTORE dialog, IRQ guard, single-instance)'
        hardwareValidation = 'SKIPPED'
        checks = @($checks.ToArray())
        ownedProcessIds = @($ownedProcesses | ForEach-Object { $_.Id })
        processReports = @(Get-ChildItem -LiteralPath $env:DEVICE_TWEAKER_QA_OUTPUT -Filter 'process-*.json' -ErrorAction SilentlyContinue | ForEach-Object { $_.FullName })
        displayMode = $(if ($VisibleGui) { 'VISIBLE_DEBUG' } else { 'PRIVATE_DESKTOP' })
        desktop = [DeviceTweakerQaDesktop]::CurrentDesktop()
        screenshots = @(Get-ChildItem -LiteralPath $outDir -Filter '*.png' -File | ForEach-Object {
            @{ path = $_.FullName; sha256 = (Get-FileHash -LiteralPath $_.FullName).Hash; bytes = $_.Length }
        })
    }
    $jsonReport | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $outDir 'smoke.json') -Encoding UTF8
    foreach ($name in $environmentNames) { [Environment]::SetEnvironmentVariable($name, $savedEnvironment[$name]) }
    if ($smokeFailed) { exit 1 }
}

