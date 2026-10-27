param(
    [string]$ExePath,
    [string]$OutputDirectory,
    [switch]$VisibleGui,
    [ValidateRange(5, 100)]
    [int]$RefreshCycles = 20
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Private-Desktop.ps1')
if (-not $VisibleGui) {
    if (-not $OutputDirectory) { $OutputDirectory = Join-Path $PSScriptRoot ('..\..\DEVICE TWEAKER\tmp\qa-resources\' + [Guid]::NewGuid().ToString('N')) }
    $OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
    $workerParameters = @{} + $PSBoundParameters
    $workerParameters.OutputDirectory = $OutputDirectory
    if ($ExePath) { $workerParameters.ExePath = [IO.Path]::GetFullPath($ExePath) }
    $workerExit = Invoke-QaPrivateDesktop -ScriptPath $PSCommandPath -Parameters $workerParameters `
        -ReportDirectory (Join-Path $OutputDirectory 'isolation')
    if ($null -ne $workerExit) { exit $workerExit }
}
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.Windows.Forms
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class DeviceTweakerResourceQaNative {
    private delegate bool EnumWindow(IntPtr hwnd, IntPtr data);
    [DllImport("user32.dll")] private static extern bool EnumWindows(EnumWindow callback, IntPtr data);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] private static extern int GetWindowText(IntPtr hwnd, System.Text.StringBuilder text, int count);
    [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr hwnd, int command);
    [DllImport("user32.dll")] private static extern IntPtr GetWindow(IntPtr hwnd, uint command);
    public static IntPtr GetOwnedWindow(int processId) {
        IntPtr found = IntPtr.Zero;
        EnumWindows((hwnd, data) => {
            uint pid; GetWindowThreadProcessId(hwnd, out pid);
            if (pid != processId) return true;
            var title = new System.Text.StringBuilder(256); GetWindowText(hwnd, title, 256);
            var kind = new System.Text.StringBuilder(256); GetClassName(hwnd, kind, 256);
            if (title.ToString() != "DEVICE TWEAKER" || !kind.ToString().StartsWith("WindowsForms10.Window") ||
                GetWindow(hwnd, 4) != IntPtr.Zero) return true;
            found = hwnd;
            ShowWindow(hwnd, 4); // Show without activating or moving physical input.
            return false;
        }, IntPtr.Zero);
        return found;
    }
    [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, EnumWindow callback, IntPtr data);
    [DllImport("user32.dll")] public static extern bool IsWindowEnabled(IntPtr hwnd);
    [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr hwnd);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] private static extern int GetClassName(IntPtr hwnd, System.Text.StringBuilder text, int count);
    public static string DescribeChildren(IntPtr parent) {
        var output = new System.Text.StringBuilder(); int count = 0;
        EnumChildWindows(parent, (hwnd, data) => {
            var text = new System.Text.StringBuilder(256); GetWindowText(hwnd, text, 256);
            var kind = new System.Text.StringBuilder(256); GetClassName(hwnd, kind, 256);
            if (count++ < 16) output.AppendLine(hwnd + " " + kind + " [" + text + "]");
            return true;
        }, IntPtr.Zero);
        output.AppendLine("children=" + count); return output.ToString();
    }
    public static IntPtr FindRefreshButton(IntPtr parent) {
        IntPtr found = IntPtr.Zero;
        EnumChildWindows(parent, (hwnd, data) => {
            var text = new System.Text.StringBuilder(128); GetWindowText(hwnd, text, 128);
            if (text.ToString() != "REFRESH") return true;
            found = hwnd; return false;
        }, IntPtr.Zero);
        return found;
    }
    [DllImport("user32.dll")]
    public static extern bool PostMessage(IntPtr hWnd, uint message, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")]
    public static extern uint GetGuiResources(IntPtr process, uint flags);
    public const uint BM_CLICK = 0x00F5;
}
'@

$root = (Resolve-Path (Join-Path $PSScriptRoot '..\..\DEVICE TWEAKER')).Path
$exe = if ([string]::IsNullOrWhiteSpace($ExePath)) {
    Join-Path $root 'bin\Release\net8.0-windows\win-x64\DEVICE TWEAKER.exe'
} else {
    (Resolve-Path $ExePath).Path
}

function Get-MainWindow([int]$ProcessId, [int]$TimeoutSeconds = 30) {
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        $window = [DeviceTweakerResourceQaNative]::GetOwnedWindow($ProcessId)
        if ($window -ne [IntPtr]::Zero) { return $window }
        Start-Sleep -Milliseconds 150
    }
    return [IntPtr]::Zero
}

function Wait-Ready([IntPtr]$Window, [IntPtr]$Button = [IntPtr]::Zero, [int]$TimeoutSeconds = 45) {
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        if ($Button -eq [IntPtr]::Zero -or -not [DeviceTweakerResourceQaNative]::IsWindow($Button)) {
            $Button = [DeviceTweakerResourceQaNative]::FindRefreshButton($Window)
        }
        if ($Button -ne [IntPtr]::Zero -and [DeviceTweakerResourceQaNative]::IsWindowEnabled($Button)) { return $Button }
        Start-Sleep -Milliseconds 125
    }
    Write-Host ([DeviceTweakerResourceQaNative]::DescribeChildren($Window))
    throw "REFRESH did not return to the ready state: window=$Window button=$Button"
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

function Get-RefreshEvidence {
    $text = Read-QaText $script:resourceLog
    return @{ clicks = ([regex]::Matches($text, 'UI: REFRESH button clicked')).Count;
        completed = ([regex]::Matches($text, 'REFRESH.DONE:')).Count }
}

function Invoke-VerifiedRefresh([IntPtr]$Window, [IntPtr]$Button, [int]$Cycle) {
    $before = Get-RefreshEvidence
    if (-not [DeviceTweakerResourceQaNative]::PostMessage($Button, [DeviceTweakerResourceQaNative]::BM_CLICK,
        [IntPtr]::Zero, [IntPtr]::Zero)) { throw "REFRESH message rejected: cycle=$Cycle" }
    $deadline = (Get-Date).AddSeconds(45)
    do {
        $after = Get-RefreshEvidence
        if ($after.clicks -gt $before.clicks -and $after.completed -gt $before.completed -and
            [DeviceTweakerResourceQaNative]::IsWindowEnabled($Button)) {
            if ($after.clicks -ne $before.clicks + 1 -or $after.completed -ne $before.completed + 1) {
                throw "REFRESH did not correspond to exactly one completed action: cycle=$Cycle"
            }
            $script:refreshEvidence.Add(@{ cycle = $Cycle; expected = 'One button action and completed refresh';
                before = $before; after = $after; log = $script:resourceLog })
            return (Wait-Ready $Window $Button)
        }
        if ($process.HasExited) { throw 'Resource QA process exited during REFRESH' }
        Start-Sleep -Milliseconds 125
    } while ((Get-Date) -lt $deadline)
    throw "REFRESH completion was not observed: cycle=$Cycle before=$($before | ConvertTo-Json -Compress) after=$($after | ConvertTo-Json -Compress)"
}

function Get-SettledResourceSample([System.Diagnostics.Process]$Process, [int]$Cycle) {
    # Await a quiet resource interval after the actual completion checkpoint.
    # This is bounded stabilization, not an unconditional sleep counted as a test.
    $deadline = (Get-Date).AddSeconds(8)
    $previous = Get-ResourceSample $Process $Cycle
    $stable = 0
    do {
        Start-Sleep -Milliseconds 200
        $current = Get-ResourceSample $Process $Cycle
        if ([Math]::Abs($current.Handles - $previous.Handles) -le 3 -and
            [Math]::Abs([int]$current.Gdi - [int]$previous.Gdi) -le 2 -and
            [Math]::Abs([int]$current.User - [int]$previous.User) -le 2) { $stable++ } else { $stable = 0 }
        if ($stable -ge 3) { return $current }
        $previous = $current
    } while ((Get-Date) -lt $deadline)
    throw "Resource sample never settled after completed REFRESH: cycle=$Cycle"
}

if (-not $OutputDirectory) { $OutputDirectory = Join-Path $root ('tmp\qa-resources\' + [Guid]::NewGuid().ToString('N')) }
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$names = @('DEVICE_TWEAKER_QA_READ_ONLY', 'DEVICE_TWEAKER_QA_OUTPUT', 'DEVICE_TWEAKER_QA_SESSION',
    'DEVICE_TWEAKER_LANGUAGE', 'DEVICE_TWEAKER_QA_TEST_ADMIN', 'DEVICE_TWEAKER_QA_SANDBOX')
$saved = @{}
foreach ($name in $names) { $saved[$name] = [Environment]::GetEnvironmentVariable($name) }
$resourceStatus = 'FAIL'
$samples = [System.Collections.Generic.List[object]]::new()
$refreshEvidence = [System.Collections.Generic.List[object]]::new()
. (Join-Path $PSScriptRoot 'Read-QaJson.ps1')
$process = $null
try {
    $env:DEVICE_TWEAKER_QA_READ_ONLY = '1'
    $env:DEVICE_TWEAKER_QA_SESSION = [Guid]::NewGuid().ToString('N')
    $env:DEVICE_TWEAKER_QA_OUTPUT = Join-Path $OutputDirectory 'processes'
    $env:DEVICE_TWEAKER_LANGUAGE = 'en'
    Remove-Item Env:DEVICE_TWEAKER_QA_TEST_ADMIN -ErrorAction SilentlyContinue
    Remove-Item Env:DEVICE_TWEAKER_QA_SANDBOX -ErrorAction SilentlyContinue
    $process = Start-QaOwnedProcess $exe '--qa-read-only'
    $window = Get-MainWindow $process.Id
    if ($window -eq [IntPtr]::Zero) { throw 'Main window was not found.' }
    # Header buttons initially exist and are enabled before the 80 ms startup
    # timer begins enumeration. Enabled alone is not a startup-ready signal.
    $startupDeadline = (Get-Date).AddSeconds(45)
    $metadataPath = Join-Path $env:DEVICE_TWEAKER_QA_OUTPUT "process-$($process.Id).json"
    $startupComplete = $false
    do {
        if ($process.HasExited) { throw 'Resource process exited during startup' }
        if (Test-Path -LiteralPath $metadataPath) {
            $metadata = Read-QaJson $metadataPath
            if ($metadata.logPath -and (Test-Path -LiteralPath $metadata.logPath) -and
                (Read-QaText $metadata.logPath) -match 'REFRESH.DONE:') { $startupComplete = $true; break }
        }
        Start-Sleep -Milliseconds 125
    } while ((Get-Date) -lt $startupDeadline)
    if (-not $startupComplete) { throw 'Initial enumeration completion was not observed' }
    $refreshButton = Wait-Ready $window
    if ($metadata.readOnly -ne $true -or $metadata.processId -ne $process.Id) { throw 'Resource QA read-only handshake failed' }
    if ($env:DEVICE_TWEAKER_QA_DESKTOP -and $metadata.desktop.isolated -ne $true) { throw 'Resource process escaped the private desktop' }
    $resourceLog = $metadata.logPath


    # One warm-up cycle absorbs one-time device, font and UI Automation caches.
    $warmup = Wait-Ready $window $refreshButton
    $refreshButton = Invoke-VerifiedRefresh $window $warmup 0
    $samples.Add((Get-SettledResourceSample $process 0))
    for ($cycle = 1; $cycle -le $RefreshCycles; $cycle++) {
        $refresh = Wait-Ready $window $refreshButton
        $refreshButton = Invoke-VerifiedRefresh $window $refresh $cycle
        $sample = Get-SettledResourceSample $process $cycle
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

    $resourceStatus = 'PASS'
    Write-Host "PASS: GUI resources are stable across $RefreshCycles REFRESH cycles."
} catch {
    $resourceError = $_.Exception.ToString()
    Write-Host ("FAIL: " + ($_ | Out-String))
}
finally {
    if ($process -and -not $process.HasExited) {
        if (-not (Stop-QaOwnedProcess $process)) { $resourceStatus = 'FAIL' }
    }
    [ordered]@{ status = $resourceStatus; readOnly = $true; cycles = $RefreshCycles;
        error = $resourceError;
        startupComplete = $startupComplete;
        samples = @($samples.ToArray()); hardwareValidation = 'SKIPPED';
        refreshEvidence = @($refreshEvidence.ToArray());
        desktop = [DeviceTweakerQaDesktop]::CurrentDesktop();
        floorDelta = @{ handles = $floorHandles; gdi = $floorGdi; user = $floorUser; privateMb = $floorPrivate } } |
        ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'resources.json') -Encoding UTF8
    foreach ($name in $names) { [Environment]::SetEnvironmentVariable($name, $saved[$name]) }
    if ($resourceStatus -ne 'PASS') { exit 1 }
}
