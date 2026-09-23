param(
    [string]$Monitor = 'Secondary'
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

if (-not ([System.Management.Automation.PSTypeName]'SmokeNative').Type) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class SmokeNative {
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
}

$root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$exe = Join-Path $root 'bin\Release\net8.0-windows\win-x64\DEVICE TWEAKER.exe'
if (-not (Test-Path $exe)) {
    throw "Executable not found at $exe. Please build first."
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

$assetsScreenshots = Join-Path $root 'assets\screenshots'
New-Item -ItemType Directory -Force -Path $assetsScreenshots | Out-Null

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

function Invoke-Click(
    [System.Windows.Automation.AutomationElement]$el,
    [string]$label,
    [int]$settleMilliseconds = 700) {
    if ($null -eq $el) { throw "UI element not found: $label" }
    try {
        $inv = $el.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
        $inv.Invoke()
        Start-Sleep -Milliseconds $settleMilliseconds
        return
    } catch {}

    $hwnd = [IntPtr]$el.Current.NativeWindowHandle
    if ($hwnd -ne [IntPtr]::Zero) {
        [SmokeNative]::PostMessage($hwnd, [SmokeNative]::BM_CLICK, [IntPtr]::Zero, [IntPtr]::Zero) | Out-Null
        Start-Sleep -Milliseconds $settleMilliseconds
        return
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
    [SmokeNative]::PostMessage($hwnd, [SmokeNative]::WM_KEYDOWN, [IntPtr]$virtualKey, [IntPtr]::Zero) | Out-Null
    [SmokeNative]::PostMessage($hwnd, [SmokeNative]::WM_KEYUP, [IntPtr]$virtualKey, [IntPtr]::Zero) | Out-Null
    Start-Sleep -Milliseconds 250
}

function Capture-Window([string]$path, [System.Windows.Automation.AutomationElement]$window) {
    $r = $window.Current.BoundingRectangle
    $width = [Math]::Max(1, [int][Math]::Ceiling($r.Width))
    $height = [Math]::Max(1, [int][Math]::Ceiling($r.Height))
    $hwnd = [IntPtr]$window.Current.NativeWindowHandle
    if ($hwnd -eq [IntPtr]::Zero) { throw "Window has no native handle: $path" }

    $bmp = New-Object System.Drawing.Bitmap($width, $height)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $hdc = $g.GetHdc()
    try {
        [SmokeNative]::PrintWindow($hwnd, $hdc, [SmokeNative]::PW_RENDERFULLCONTENT) | Out-Null
    } finally {
        $g.ReleaseHdc($hdc)
        $g.Dispose()
    }
    $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    Write-Host "Captured: $path"
}

function Dismiss-Dialogs([System.Windows.Automation.AutomationElement]$window, [int]$attempts = 5) {
    for ($i = 0; $i -lt $attempts; $i++) {
        $rootEl = [System.Windows.Automation.AutomationElement]::RootElement
        $cond = New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::ProcessIdProperty, $window.Current.ProcessId)
        $popups = $rootEl.FindAll([System.Windows.Automation.TreeScope]::Children, $cond)
        foreach ($p in $popups) {
            if ($p.Current.NativeWindowHandle -ne $window.Current.NativeWindowHandle) {
                # Look for OK / CANCEL / SKIP buttons
                foreach ($btnText in @('OK', 'ОК', 'CANCEL', 'ОТМЕНА', 'SKIP', 'ПРОПУСТИТЬ')) {
                    $b = Find-Desc $p $btnText 'Button' 1
                    if ($b) {
                        try {
                            Invoke-Click $b "Dismiss popup $btnText" 500
                            break
                        } catch {}
                    }
                }
            }
        }
        Start-Sleep -Milliseconds 200
    }
}

# --- CAPTURE 1: INTEL CORE i9-14900K (HYBRID CPU, P-CORE + E-CORE + HT) ---
Write-Host "Starting Intel 14900K capture..."
$env:DEVICE_TWEAKER_QA_TEST_ADMIN = '1'
$env:DEVICE_TWEAKER_QA_SANDBOX = '1'
$env:DEVICE_TWEAKER_QA_HIDE_SANDBOX_HEADER = '1'
$env:DEVICE_TWEAKER_QA_FINAL_PRESET = 'Intel14900K'

$proc1 = Start-Process -FilePath $exe -PassThru
try {
    $main1 = Get-MainWindow -processId $proc1.Id -timeoutSec 20
    if (-not $main1) { throw "Could not find main window for Intel run" }

    # Wait for ready
    Start-Sleep -Seconds 4
    # Switch to RU
    $ruBtn = Find-AutomationId $main1 'LANGUAGE_RU' 5
    if ($ruBtn) { Invoke-Click $ruBtn 'Switch to RU' 1000 }

    # Wait for test devices to finish loading
    Start-Sleep -Seconds 3

    # Capture top main window
    Capture-Window (Join-Path $assetsScreenshots 'showcase_intel_14900k_hybrid_ru.png') $main1

    # Scroll down to capture GPU + NIC
    Send-WindowKey $main1 0x22 'PageDown' # VK_NEXT
    Start-Sleep -Milliseconds 500
    Capture-Window (Join-Path $assetsScreenshots 'showcase_intel_14900k_gpu_nic_ru.png') $main1

} finally {
    Stop-Process -Id $proc1.Id -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 1
}

# --- CAPTURE 2: AMD RYZEN 9 9950X3D (DUAL-CCD, V-CACHE CCD0 + FREQUENCY CCD1 + SMT) ---
Write-Host "Starting AMD 9950X3D capture..."
$env:DEVICE_TWEAKER_QA_FINAL_PRESET = 'Ryzen9950X3D'

$proc2 = Start-Process -FilePath $exe -PassThru
try {
    $main2 = Get-MainWindow -processId $proc2.Id -timeoutSec 20
    if (-not $main2) { throw "Could not find main window for AMD run" }

    # Wait for ready
    Start-Sleep -Seconds 4
    # Switch to RU
    $ruBtn = Find-AutomationId $main2 'LANGUAGE_RU' 5
    if ($ruBtn) { Invoke-Click $ruBtn 'Switch to RU' 1000 }

    # Wait for test devices to finish loading
    Start-Sleep -Seconds 3

    # Capture top main window (AMD Dual-CCD)
    Capture-Window (Join-Path $assetsScreenshots 'showcase_amd_9950x3d_dual_ccd_ru.png') $main2

    # Scroll down to capture GPU + NIC + Storage
    Send-WindowKey $main2 0x22 'PageDown'
    Start-Sleep -Milliseconds 500
    Capture-Window (Join-Path $assetsScreenshots 'showcase_amd_9950x3d_devices_ru.png') $main2

} finally {
    Stop-Process -Id $proc2.Id -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 1
}

# Copy Russian restore dialog
Copy-Item (Join-Path $root 'bin\SmokeSafe\run_20260921_234551\12_restore_ru.png') (Join-Path $assetsScreenshots 'showcase_restore_ru.png') -Force

Write-Host "All showcase screenshots generated successfully!"
