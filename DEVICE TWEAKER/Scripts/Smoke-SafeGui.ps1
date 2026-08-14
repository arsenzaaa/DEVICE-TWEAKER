# Safe GUI smoke for DEVICE TWEAKER (second monitor + TEST ADMIN sandbox).
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
  [DllImport("user32.dll")] public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);
  [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
  public const int SW_RESTORE = 9;
  public const uint BM_CLICK = 0x00F5;
}
'@

$root = 'C:\Users\Administrator\Desktop\DEVICE TWEAKER'
$exe = Join-Path $root 'bin\Release\net8.0-windows\win-x64\DEVICE TWEAKER.exe'
$outDir = Join-Path $root 'bin\SmokeSafe'
New-Item -ItemType Directory -Force -Path $outDir | Out-Null
$reportPath = Join-Path $outDir ('smoke_{0:yyyyMMdd_HHmmss}.txt' -f (Get-Date))

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
        $win = $rootEl.FindFirst([System.Windows.Automation.TreeScope]::Children, $cond)
        if ($win -and [string]$win.Current.Name -like '*DEVICE TWEAKER*') { return $win }
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

function Invoke-Click([System.Windows.Automation.AutomationElement]$el, [string]$label) {
    if ($null -eq $el) { throw "UI element not found: $label" }

    $hwnd = [IntPtr]$el.Current.NativeWindowHandle
    if ($hwnd -ne [IntPtr]::Zero) {
        # Async Win32 click — does not move the cursor or block on modal dialogs.
        [SmokeNative]::PostMessage($hwnd, [SmokeNative]::BM_CLICK, [IntPtr]::Zero, [IntPtr]::Zero) | Out-Null
        Write-Report "CLICK(BM): $label hwnd=$hwnd"
        Start-Sleep -Milliseconds 700
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

function Capture-Screen([string]$path, [System.Windows.Forms.Screen]$screen) {
    $b = $screen.Bounds
    $bmp = New-Object System.Drawing.Bitmap $b.Width, $b.Height
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($b.Location, [System.Drawing.Point]::Empty, $b.Size)
    $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose(); $bmp.Dispose()
    Write-Report "SHOT: $path"
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
Get-Process | Where-Object { $_.ProcessName -eq 'DEVICE TWEAKER' } | ForEach-Object {
    Write-Report "KILL pid=$($_.Id)"
    Stop-Process -Id $_.Id -Force
}
Start-Sleep -Seconds 1

$env:DEVICE_TWEAKER_QA_TEST_ADMIN = '1'
$env:DEVICE_TWEAKER_QA_SANDBOX = '1'
$proc = Start-Process -FilePath $exe -WorkingDirectory (Split-Path $exe) -PassThru
Write-Report "START pid=$($proc.Id) QA_TEST_ADMIN=1 QA_SANDBOX=1"
Start-Sleep -Seconds 6

$main = Get-MainWindow -processId $proc.Id -timeoutSec 30
if ($null -eq $main) { throw 'Main window not found' }
$hwnd = [IntPtr]$main.Current.NativeWindowHandle
[SmokeNative]::ShowWindow($hwnd, [SmokeNative]::SW_RESTORE) | Out-Null

$secondary = [System.Windows.Forms.Screen]::AllScreens | Where-Object { -not $_.Primary } | Select-Object -First 1
if ($null -eq $secondary) { $secondary = [System.Windows.Forms.Screen]::PrimaryScreen }
$wa = $secondary.WorkingArea
$width = [Math]::Min(1280, $wa.Width - 40)
$height = [Math]::Min(900, $wa.Height - 40)
$x = $wa.X + 40
$y = $wa.Y + 40
[SmokeNative]::MoveWindow($hwnd, [int]$x, [int]$y, [int]$width, [int]$height, $true) | Out-Null
Write-Report ("MOVE secondary={0} x={1} y={2}" -f $secondary.DeviceName, $x, $y)
Start-Sleep -Seconds 1

# Owned modal dialog is a descendant of the main window, not a root child.
$admin = Find-Desc $main 'TEST ADMIN' 'Window' 25
if ($null -eq $admin) { throw 'TEST ADMIN dialog not found under main window' }
Write-Report 'OPEN: TEST ADMIN'
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

Capture-Screen (Join-Path $outDir '20_test_admin.png') $secondary

$setOnly = Find-Desc $admin 'Show test devices only' 'CheckBox' 3
if ($setOnly) { Set-ToggleOn $setOnly 'Show test devices only' }
$setDry = Find-Desc $admin 'Sandbox dry-run (no registry writes)' 'CheckBox' 3
if ($setDry) { Set-ToggleOn $setDry 'dry-run' }
Capture-Screen (Join-Path $outDir '21_test_admin_armed.png') $secondary

Invoke-Click (Find-Desc $admin 'CLOSE' 'Button' 5) 'TEST ADMIN/CLOSE'
Start-Sleep -Seconds 3

$main = Get-MainWindow -processId $proc.Id -timeoutSec 10
Capture-Screen (Join-Path $outDir '22_test_only_main.png') $secondary

foreach ($btnName in @('REFRESH', 'AUTO-OPTIMIZATION', 'RESET ALL', 'APPLY')) {
    $main = Get-MainWindow -processId $proc.Id -timeoutSec 5
    Start-Sleep -Milliseconds 300
    Invoke-Click (Find-Desc $main $btnName 'Button' 5) "main/$btnName"
    Start-Sleep -Seconds 1
    Dismiss-Dialogs $main 10
    Start-Sleep -Seconds 1
    Capture-Screen (Join-Path $outDir ('23_{0}.png' -f ($btnName -replace '[^A-Za-z0-9]+','_'))) $secondary
}

$logDir = Join-Path (Split-Path $exe) 'logs'
$latestLog = Get-ChildItem -LiteralPath $logDir -Filter 'DeviceTweaker_*.log' |
    Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($latestLog) {
    Copy-Item -LiteralPath $latestLog.FullName -Destination (Join-Path $outDir $latestLog.Name) -Force
    Write-Report "LOG: $($latestLog.FullName)"
    foreach ($p in @(
        'TEST.QA.SANDBOX',
        'autoDryRun=True',
        'testDevicesOnly=True',
        'AUTO.DRYRUN',
        'AUTO.POWER',
        'APPLY.PREVIEW',
        'APPLY.DRYRUN',
        'USB.SUSPEND.PLAN: skipped (no real USB blocks)',
        'RESET.DONE: mode=dry-run',
        'UI: APPLY dry-run completed',
        'UI: AUTO-OPTIMIZATION',
        'UI: RESET ALL',
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
    $planWrite = @(Select-String -LiteralPath $latestLog.FullName -Pattern 'USB.SUSPEND.PLAN: enabled|USB.SUSPEND.PLAN: disabled|RESET.SUSPEND.PLAN: USB selective suspend power plan restored' -ErrorAction SilentlyContinue)
    $planSkip = @(Select-String -LiteralPath $latestLog.FullName -Pattern 'USB.SUSPEND.PLAN: skipped \(no real USB blocks\)|RESET.SUSPEND.PLAN: skipped \(no real USB blocks\)' -ErrorAction SilentlyContinue)
    if ($regHits.Count -gt 0) { Write-Report "FAIL: unexpected APPLY.REG writes=$($regHits.Count)" }
    if ($planWrite.Count -gt 0) { Write-Report "FAIL: unexpected real USB power-plan writes=$($planWrite.Count)" }
    Write-Report ("LOGHIT power-plan-skip => {0}" -f $planSkip.Count)
    if ($regHits.Count -eq 0 -and $planWrite.Count -eq 0) { Write-Report 'PASS: no registry/power-plan writes in this smoke' }
}

if (-not $proc.HasExited) {
    Stop-Process -Id $proc.Id -Force
    Write-Report "STOP pid=$($proc.Id)"
}
Write-Report "DONE report=$reportPath"
Write-Output $reportPath
