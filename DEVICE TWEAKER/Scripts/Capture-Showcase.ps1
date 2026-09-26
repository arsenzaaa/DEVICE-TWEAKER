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
$targetWidth = [Math]::Min(1360, $targetArea.Width - 40)
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

function Click-CategoryButton($scope, [string[]]$textCandidates, [string]$label) {
    foreach ($cand in $textCandidates) {
        $btn = Find-Desc $scope $cand 'Button' 1
        if ($btn) {
            Invoke-Click $btn "Click category $label ($cand)" 600
            return $true
        }
    }
    Write-Warning "Category button not found for: $($textCandidates -join ', ')"
    return $false
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

function Capture-Scenario([string]$preset, [string]$lang, [string]$filter, [string]$outputFilename, [bool]$pageDown = $false) {
    Write-Host "Capturing $outputFilename (preset=$preset, lang=$lang, filter=$filter)..."
    Remove-Item Env:\DEVICE_TWEAKER_QA_TEST_ADMIN -ErrorAction SilentlyContinue
    $env:DEVICE_TWEAKER_QA_SANDBOX = '1'
    $env:DEVICE_TWEAKER_QA_HIDE_SANDBOX_HEADER = '1'
    $env:DEVICE_TWEAKER_SHOWCASE = "${preset}_${lang}"
    $env:DEVICE_TWEAKER_SHOWCASE_FILTER = $filter
    $env:DEVICE_TWEAKER_LANGUAGE = $lang
    $env:DEVICE_TWEAKER_CATEGORY_FILTER = $filter

    $proc = Start-Process -FilePath $exe -PassThru
    try {
        $main = Get-MainWindow -processId $proc.Id -timeoutSec 20
        if (-not $main) { throw "Could not find main window for $outputFilename" }

        # Wait for showcase setup to apply and render
        Start-Sleep -Milliseconds 3200

        if ($pageDown) {
            Send-WindowKey $main 0x22 'PageDown'
            Start-Sleep -Milliseconds 600
        }

        $outPath = Join-Path $assetsScreenshots $outputFilename
        Capture-Window $outPath $main
    } finally {
        Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
        Start-Sleep -Milliseconds 400
    }
}

# --- CAPTURES: INTEL 14900K ---
Capture-Scenario 'Intel14900K' 'RU' 'USB' 'showcase_intel_14900k_hybrid_ru.png'
Capture-Scenario 'Intel14900K' 'RU' 'ALL' 'showcase_intel_14900k_gpu_nic_ru.png' $true
Capture-Scenario 'Intel14900K' 'EN' 'USB' 'showcase_intel_14900k_hybrid_en.png'

# --- CAPTURES: AMD RYZEN 9 9950X3D (RU) ---
Capture-Scenario 'Ryzen9950X3D' 'RU' 'USB' 'showcase_amd_9950x3d_dual_ccd_ru.png'
Capture-Scenario 'Imod' 'RU' 'USB' 'showcase_amd_imod_table_ru.png'
Capture-Scenario 'Ryzen9950X3D' 'RU' 'GPU' 'showcase_gpu_affinity_ru.png'
Capture-Scenario 'NicItr' 'RU' 'NETWORK' 'showcase_nic_itr_ru.png'
Capture-Scenario 'Ryzen9950X3D' 'RU' 'ALL' 'main_interface_ru.png'
Capture-Scenario 'Ryzen9950X3D' 'RU' 'ALL' 'showcase_amd_9950x3d_devices_ru.png' $true

# --- CAPTURES: AMD RYZEN 9 9950X3D (EN) ---
Capture-Scenario 'Ryzen9950X3D' 'EN' 'USB' 'showcase_amd_9950x3d_dual_ccd_en.png'
Capture-Scenario 'Imod' 'EN' 'USB' 'showcase_amd_imod_table_en.png'
Capture-Scenario 'Ryzen9950X3D' 'EN' 'GPU' 'showcase_gpu_affinity_en.png'
Capture-Scenario 'NicItr' 'EN' 'NETWORK' 'showcase_nic_itr_en.png'
Capture-Scenario 'Ryzen9950X3D' 'EN' 'ALL' 'main_interface_en.png'

$restoreSample = Join-Path $root 'bin\SmokeSafe\run_20260921_234551\12_restore_ru.png'
if (Test-Path $restoreSample) {
    Copy-Item $restoreSample (Join-Path $assetsScreenshots 'showcase_restore_ru.png') -Force
}

Write-Host "All showcase screenshots generated successfully!"
