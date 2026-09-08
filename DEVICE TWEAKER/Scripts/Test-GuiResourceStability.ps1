param(
    [string]$ExePath,
    [ValidateRange(5, 100)]
    [int]$RefreshCycles = 20
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.Windows.Forms
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class DeviceTweakerResourceQaNative {
    [DllImport("user32.dll")]
    public static extern bool PostMessage(IntPtr hWnd, uint message, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")]
    public static extern uint GetGuiResources(IntPtr process, uint flags);
    public const uint BM_CLICK = 0x00F5;
}
'@

$root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$exe = if ([string]::IsNullOrWhiteSpace($ExePath)) {
    Join-Path $root 'bin\Release\net8.0-windows\win-x64\DEVICE TWEAKER.exe'
} else {
    (Resolve-Path $ExePath).Path
}

function Get-MainWindow([int]$ProcessId, [int]$TimeoutSeconds = 30) {
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        $rootElement = [System.Windows.Automation.AutomationElement]::RootElement
        $condition = New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::ProcessIdProperty, $ProcessId)
        $window = $rootElement.FindFirst([System.Windows.Automation.TreeScope]::Children, $condition)
        if ($window -and [string]$window.Current.Name -like '*DEVICE TWEAKER*') {
            return $window
        }
        Start-Sleep -Milliseconds 150
    }
    return $null
}

function Get-RefreshButton([System.Windows.Automation.AutomationElement]$Window) {
    $condition = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::AutomationIdProperty, 'REFRESH')
    return $Window.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $condition)
}

function Wait-Ready([System.Windows.Automation.AutomationElement]$Window, [int]$TimeoutSeconds = 45) {
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        $button = Get-RefreshButton $Window
        if ($button -and $button.Current.IsEnabled) {
            return $button
        }
        Start-Sleep -Milliseconds 125
    }
    throw 'REFRESH did not return to the ready state.'
}

function Get-ResourceSample([System.Diagnostics.Process]$Process, [int]$Cycle) {
    $Process.Refresh()
    [pscustomobject]@{
        Cycle = $Cycle
        Handles = $Process.HandleCount
        Gdi = [DeviceTweakerResourceQaNative]::GetGuiResources($Process.Handle, 0)
        User = [DeviceTweakerResourceQaNative]::GetGuiResources($Process.Handle, 1)
        PrivateMb = [Math]::Round($Process.PrivateMemorySize64 / 1MB, 1)
        WorkingMb = [Math]::Round($Process.WorkingSet64 / 1MB, 1)
    }
}

$process = $null
try {
    $env:DEVICE_TWEAKER_LANGUAGE = 'en'
    Remove-Item Env:DEVICE_TWEAKER_QA_TEST_ADMIN -ErrorAction SilentlyContinue
    Remove-Item Env:DEVICE_TWEAKER_QA_SANDBOX -ErrorAction SilentlyContinue
    $process = Start-Process -FilePath $exe -WorkingDirectory (Split-Path $exe) -PassThru
    $window = Get-MainWindow $process.Id
    if ($null -eq $window) { throw 'Main window was not found.' }
    $null = Wait-Ready $window

    # One warm-up cycle absorbs one-time device, font and UI Automation caches.
    $warmup = Wait-Ready $window
    $warmupHandle = [IntPtr]$warmup.Current.NativeWindowHandle
    if ($warmupHandle -eq [IntPtr]::Zero -or
        -not [DeviceTweakerResourceQaNative]::PostMessage($warmupHandle, [DeviceTweakerResourceQaNative]::BM_CLICK, [IntPtr]::Zero, [IntPtr]::Zero)) {
        throw 'Could not invoke the warm-up REFRESH button.'
    }
    Start-Sleep -Milliseconds 250
    $null = Wait-Ready $window
    Start-Sleep -Milliseconds 1500

    $samples = [System.Collections.Generic.List[object]]::new()
    $samples.Add((Get-ResourceSample $process 0))
    for ($cycle = 1; $cycle -le $RefreshCycles; $cycle++) {
        $refresh = Wait-Ready $window
        $handle = [IntPtr]$refresh.Current.NativeWindowHandle
        if ($handle -eq [IntPtr]::Zero -or
            -not [DeviceTweakerResourceQaNative]::PostMessage($handle, [DeviceTweakerResourceQaNative]::BM_CLICK, [IntPtr]::Zero, [IntPtr]::Zero)) {
            throw "Could not invoke REFRESH at cycle $cycle."
        }
        Start-Sleep -Milliseconds 250
        $null = Wait-Ready $window
        # REFRESH enables the UI before every asynchronous IRQ/readback task has
        # necessarily released its short-lived process handles. Sample only
        # after that work has had time to settle.
        Start-Sleep -Milliseconds 1500
        $sample = Get-ResourceSample $process $cycle
        $samples.Add($sample)
        Write-Host ("cycle={0} handles={1} gdi={2} user={3} privateMb={4} workingMb={5}" -f
            $sample.Cycle, $sample.Handles, $sample.Gdi, $sample.User, $sample.PrivateMb, $sample.WorkingMb)
    }

    $first = $samples[0]
    $last = $samples[$samples.Count - 1]
    $deltaHandles = $last.Handles - $first.Handles
    $deltaGdi = [int]$last.Gdi - [int]$first.Gdi
    $deltaUser = [int]$last.User - [int]$first.User
    $deltaPrivate = $last.PrivateMb - $first.PrivateMb
    Write-Host ("delta handles={0} gdi={1} user={2} privateMb={3}" -f
        $deltaHandles, $deltaGdi, $deltaUser, $deltaPrivate)

    # Managed WinForms drawing objects are reclaimed in GC waves, so comparing
    # two arbitrary endpoints is flaky. Compare the lower envelope of the first
    # and second halves instead: a real per-card leak raises that floor, while a
    # normal GC wave repeatedly returns it to the same range.
    $split = [Math]::Max(1, [int][Math]::Floor($samples.Count / 2))
    $firstHalf = @($samples | Select-Object -First $split)
    $secondHalf = @($samples | Select-Object -Skip $split)
    $floorHandles = ($secondHalf | Measure-Object Handles -Minimum).Minimum - ($firstHalf | Measure-Object Handles -Minimum).Minimum
    $floorGdi = ($secondHalf | Measure-Object Gdi -Minimum).Minimum - ($firstHalf | Measure-Object Gdi -Minimum).Minimum
    $floorUser = ($secondHalf | Measure-Object User -Minimum).Minimum - ($firstHalf | Measure-Object User -Minimum).Minimum
    $floorPrivate = ($secondHalf | Measure-Object PrivateMb -Minimum).Minimum - ($firstHalf | Measure-Object PrivateMb -Minimum).Minimum
    Write-Host ("floor-delta handles={0} gdi={1} user={2} privateMb={3}" -f
        $floorHandles, $floorGdi, $floorUser, $floorPrivate)

    # The original defect increased GDI by hundreds and USER by thousands in
    # only 12 cycles, so these limits remain deliberately strict.
    if ($floorHandles -gt 60 -or $floorGdi -gt 32 -or $floorUser -gt 20 -or $floorPrivate -gt 64) {
        throw 'GUI resources keep growing across REFRESH cycles.'
    }

    Write-Host "PASS: GUI resources are stable across $RefreshCycles REFRESH cycles."
}
finally {
    Remove-Item Env:DEVICE_TWEAKER_LANGUAGE -ErrorAction SilentlyContinue
    if ($process -and -not $process.HasExited) {
        if (-not $process.CloseMainWindow() -or -not $process.WaitForExit(5000)) {
            Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
        }
    }
}
