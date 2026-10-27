param([Parameter(Mandatory = $true)][string]$ExePath,
    [Parameter(Mandatory = $true)][string]$OutputDirectory)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Private-Desktop.ps1')
. (Join-Path $PSScriptRoot 'Read-QaJson.ps1')
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
$ExePath = [IO.Path]::GetFullPath($ExePath)
$code = Invoke-QaPrivateDesktop $PSCommandPath @{ ExePath = $ExePath; OutputDirectory = $OutputDirectory } `
    (Join-Path $OutputDirectory 'isolation')
if ($null -ne $code) { exit $code }
Assert-QaPrivateDesktop
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes, System.Drawing, System.Windows.Forms
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class LayoutQaNative {
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
    [DllImport("user32.dll")] public static extern bool MoveWindow(IntPtr hwnd, int x, int y, int width, int height, bool repaint);
    [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr hwnd, IntPtr dc, uint flags);
    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hwnd, uint msg, IntPtr w, IntPtr l);
    [DllImport("user32.dll")] public static extern bool RedrawWindow(IntPtr hwnd, IntPtr rect, IntPtr region, uint flags);
    [DllImport("user32.dll", SetLastError=true)] public static extern IntPtr SendMessageTimeout(IntPtr hwnd, uint msg, IntPtr w, IntPtr l, uint flags, uint timeout, out IntPtr result);
}
'@
function Wait-Element($window, [string]$Id) {
    $deadline = (Get-Date).AddSeconds(15)
    do {
        $condition = [System.Windows.Automation.PropertyCondition]::new(
            [System.Windows.Automation.AutomationElement]::AutomationIdProperty, $Id)
        $element = $window.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $condition)
        if ($element) { return $element }
        Start-Sleep -Milliseconds 100
    } while ((Get-Date) -lt $deadline)
    throw "Missing rendered control: $Id"
}
function Save-Frame($window, [string]$Path, [int]$Attempt = 0) {
    $rect = $window.Current.BoundingRectangle
    if ($rect.Width -lt 100 -or $rect.Height -lt 100) { throw 'Invalid capture geometry' }
    $result = [IntPtr]::Zero
    $response = [LayoutQaNative]::SendMessageTimeout($script:mainHandle, 0x8475,
        [IntPtr]$window.Current.NativeWindowHandle, [IntPtr]::Zero, 2, 10000, [ref]$result)
    if ($response -eq [IntPtr]::Zero -or $result -ne [IntPtr]1) { throw 'In-process WinForms rendering failed or timed out' }
    # DrawToBitmap supplies full layout evidence, but can omit native text/DWM.
    # Keep it explicitly diagnostic. The visual assertion uses native PrintWindow
    # only inside the physical display geometry where Windows paints reliably.
    $diagnostic = $Path + '.layout.png'
    Copy-Item -LiteralPath (Join-Path $env:DEVICE_TWEAKER_QA_OUTPUT 'visual-snapshot.png') -Destination $diagnostic -Force
    $fullImage = [Drawing.Bitmap]::new([int]$rect.Width, [int]$rect.Height)
    if (-not [LayoutQaNative]::RedrawWindow([IntPtr]$window.Current.NativeWindowHandle, [IntPtr]::Zero, [IntPtr]::Zero, 0x585)) {
        $fullImage.Dispose(); throw 'Native window did not repaint before capture'
    }
    $graphics = [Drawing.Graphics]::FromImage($fullImage); $dc = $graphics.GetHdc()
    try {
        if (-not [LayoutQaNative]::PrintWindow([IntPtr]$window.Current.NativeWindowHandle, $dc, 2)) { throw 'Native layout capture failed' }
    } finally { $graphics.ReleaseHdc($dc); $graphics.Dispose() }
    $screen = [System.Windows.Forms.Screen]::FromRectangle([Drawing.Rectangle]::new([int]$rect.Left, [int]$rect.Top, [int]$rect.Width, [int]$rect.Height))
    $captureWidth = [Math]::Min([int]$rect.Width, $screen.Bounds.Right - [int]$rect.Left)
    $captureHeight = [Math]::Min([int]$rect.Height, $screen.Bounds.Bottom - [int]$rect.Top)
    if ($captureWidth -lt 100 -or $captureHeight -lt 100) { $fullImage.Dispose(); throw 'Native capture viewport is unavailable' }
    try { $image = $fullImage.Clone([Drawing.Rectangle]::new(0, 0, $captureWidth, $captureHeight), $fullImage.PixelFormat) }
    finally { $fullImage.Dispose() }
    try {
        $bright = 0; $dark = 0; $maxWhiteRowFraction = 0.0
        for ($y = 0; $y -lt $image.Height; $y += 8) {
            $white = 0; $samples = 0
            for ($x = 0; $x -lt $image.Width; $x += 8) {
                $pixel = $image.GetPixel($x, $y); $sum = [int]$pixel.R + [int]$pixel.G + [int]$pixel.B
                if ($sum -gt 180) { $bright++ }; if ($sum -lt 90) { $dark++ }
                if ($pixel.R -gt 240 -and $pixel.G -gt 240 -and $pixel.B -gt 240) { $white++ }
                $samples++
            }
            $maxWhiteRowFraction = [Math]::Max($maxWhiteRowFraction, $white / [double]$samples)
        }
        if ($bright -lt 30 -or $dark -lt 30 -or $maxWhiteRowFraction -gt 0.35) {
            if ($Attempt -lt 20) {
                $image.Dispose()
                Start-Sleep -Milliseconds 100
                return (Save-Frame $window $Path ($Attempt + 1))
            }
            $image.Save($Path + '.invalid.png', [Drawing.Imaging.ImageFormat]::Png)
            throw "Blank or incomplete layout image: path=$Path bright=$bright dark=$dark whiteRow=$maxWhiteRowFraction"
        }
        $image.Save($Path, [Drawing.Imaging.ImageFormat]::Png)
        return @{ path = $Path; sha256 = (Get-FileHash -LiteralPath $Path).Hash;
            width = $image.Width; height = $image.Height; brightSamples = $bright;
            maxWhiteRowFraction = $maxWhiteRowFraction; method = 'PrintWindow';
            windowWidth = [int]$rect.Width; windowHeight = [int]$rect.Height;
            fullWindow = $captureWidth -eq [int]$rect.Width -and $captureHeight -eq [int]$rect.Height;
            diagnostic = @{ path = $diagnostic; sha256 = (Get-FileHash -LiteralPath $diagnostic).Hash;
                method = 'WinForms.DrawToBitmap'; scope = 'Layout only; native text and DWM can differ' } }
    } finally { $image.Dispose() }
}
function Assert-PaintedText($Window, $Element, [string]$Path) {
    $root = $Window.Current.BoundingRectangle; $rect = $Element.Current.BoundingRectangle
    $image = [Drawing.Bitmap]::FromFile($Path)
    try {
        $left = [int]($rect.Left - $root.Left) + 4; $top = [int]($rect.Top - $root.Top) + 4
        $right = [int]($rect.Right - $root.Left) - 4; $bottom = [int]($rect.Bottom - $root.Top) - 4
        if ($left -lt 0 -or $top -lt 0 -or $right -ge $image.Width -or $bottom -ge $image.Height) { throw 'Header text is outside the captured viewport' }
        $pixels = 0
        for ($y = $top; $y -lt $bottom; $y++) {
            for ($x = $left; $x -lt $right; $x++) {
                $color = $image.GetPixel($x, $y)
                if ([int]$color.R + [int]$color.G + [int]$color.B -gt 210) { $pixels++ }
            }
        }
        if ($pixels -lt 15) { throw "Header text did not paint: $($Element.Current.AutomationId) pixels=$pixels" }
    } finally { $image.Dispose() }
}
$names = @('DEVICE_TWEAKER_QA_READ_ONLY', 'DEVICE_TWEAKER_QA_OUTPUT', 'DEVICE_TWEAKER_QA_SESSION',
    'DEVICE_TWEAKER_LANGUAGE', 'DEVICE_TWEAKER_SHOWCASE', 'DEVICE_TWEAKER_SHOWCASE_FILTER',
    'DEVICE_TWEAKER_QA_LAYOUT_DPI', 'DEVICE_TWEAKER_QA_TEST_ADMIN', 'DEVICE_TWEAKER_QA_WINDOW_BOUNDS')
$saved = @{}; foreach ($name in $names) { $saved[$name] = [Environment]::GetEnvironmentVariable($name) }
$cases = [Collections.Generic.List[object]]::new()
$status = 'FAIL'; $process = $null
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
try {
    foreach ($dpi in @(96, 120, 144)) {
        foreach ($language in @('en', 'ru')) {
            $directory = Join-Path $OutputDirectory ("layout_${dpi}_$language")
            New-Item -ItemType Directory -Force -Path $directory | Out-Null
            $env:DEVICE_TWEAKER_QA_READ_ONLY = '1'
            $env:DEVICE_TWEAKER_QA_OUTPUT = Join-Path $directory 'process'
            $env:DEVICE_TWEAKER_QA_SESSION = [Guid]::NewGuid().ToString('N')
            $env:DEVICE_TWEAKER_QA_LAYOUT_DPI = [string]$dpi
            $env:DEVICE_TWEAKER_LANGUAGE = $language
            $env:DEVICE_TWEAKER_SHOWCASE = 'Intel14900K_' + $language
            $env:DEVICE_TWEAKER_SHOWCASE_FILTER = 'GPU'
            Remove-Item Env:DEVICE_TWEAKER_QA_TEST_ADMIN -ErrorAction SilentlyContinue
            Remove-Item Env:DEVICE_TWEAKER_QA_WINDOW_BOUNDS -ErrorAction SilentlyContinue
            $process = Start-QaOwnedProcess $ExePath '--qa-read-only'
            $deadline = (Get-Date).AddSeconds(40)
            $window = $null; $metadata = $null
            do {
                if ($process.HasExited) { throw "Layout process exited: $($process.ExitCode)" }
                $handle = [LayoutQaNative]::MainWindow($process.Id)
                $report = Join-Path $env:DEVICE_TWEAKER_QA_OUTPUT "process-$($process.Id).json"
                if ($handle -ne [IntPtr]::Zero -and (Test-Path -LiteralPath $report)) {
                    $metadata = Read-QaJson $report
                    if ($metadata.logPath -and (Test-Path -LiteralPath $metadata.logPath) -and
                        ((Read-QaText $metadata.logPath) -match 'SHOWCASE: initialized')) {
                        $window = [System.Windows.Automation.AutomationElement]::FromHandle($handle); break
                    }
                }
                Start-Sleep -Milliseconds 150
            } while ((Get-Date) -lt $deadline)
            if (-not $window -or -not $metadata.desktop.isolated -or -not $metadata.readOnly) { throw 'Layout startup handshake failed' }
            $factor = $dpi / 96.0
            $script:mainHandle = $handle
            $wantedWidth = [int](1280*$factor); $wantedHeight = [int](1000*$factor)
            $screenBounds = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
            if (-not [LayoutQaNative]::MoveWindow($handle, $screenBounds.X, $screenBounds.Y, $wantedWidth, $wantedHeight, $true)) { throw 'Layout resize failed' }
            # Resize calls the real synchronous LayoutBlocks handler, not REFRESH.
            # Observe resulting geometry and readiness instead of waiting for an unrelated log.
            $refresh = Wait-Element $window 'REFRESH'
            $deadline = (Get-Date).AddSeconds(30)
            do {
                $refresh = Wait-Element $window 'REFRESH'
                $actualRect = $window.Current.BoundingRectangle
                if ($refresh.Current.IsEnabled -and $actualRect.Width -eq $wantedWidth -and $actualRect.Height -eq $wantedHeight) { break }
                Start-Sleep -Milliseconds 100
            } while ((Get-Date) -lt $deadline)
            if (-not $refresh.Current.IsEnabled -or $actualRect.Width -ne $wantedWidth -or $actualRect.Height -ne $wantedHeight) {
                throw "Layout never reached requested geometry: wanted=${wantedWidth}x${wantedHeight} actual=$($actualRect.Width)x$($actualRect.Height)"
            }
            $apply = Wait-Element $window 'APPLY'; $en = Wait-Element $window 'LANGUAGE_EN'; $ru = Wait-Element $window 'LANGUAGE_RU'
            $expectedCaption = if ($language -eq 'en') { 'APPLY' } else { [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('0J/QoNCY0JzQldCd0JjQotCs')) }
            if ($apply.Current.Name -ne $expectedCaption) { throw "Wrong caption: expected=$expectedCaption actual=$($apply.Current.Name)" }
            $bounds = $window.Current.BoundingRectangle
            $geometry = @()
            foreach ($button in @($apply, $en, $ru, $refresh)) {
                $rect = $button.Current.BoundingRectangle
                if ($rect.Width -lt 35*$factor -or $rect.Height -lt 20*$factor -or
                    $rect.Left -lt $bounds.Left -or $rect.Right -gt $bounds.Right -or
                    $rect.Top -lt $bounds.Top -or $rect.Bottom -gt $bounds.Bottom) { throw "Clipped header control: $($button.Current.AutomationId)" }
                $geometry += @{ id = $button.Current.AutomationId; x = $rect.Left; y = $rect.Top; width = $rect.Width; height = $rect.Height }
            }
            if ($en.Current.BoundingRectangle.Right -gt $ru.Current.BoundingRectangle.Left) { throw 'Language buttons overlap' }
            $frame = Save-Frame $window (Join-Path $directory 'main.png')
            foreach ($button in @($apply, $en, $ru, $refresh)) { Assert-PaintedText $window $button $frame.path }
            $restore = Wait-Element $window 'RESTORE'
            if (-not [LayoutQaNative]::PostMessage([IntPtr]$restore.Current.NativeWindowHandle, 0xF5, [IntPtr]::Zero, [IntPtr]::Zero)) { throw 'RESTORE message rejected' }
            $dialog = Wait-Element $window 'RESTORE_DIALOG'
            $dialogFrame = Save-Frame $dialog (Join-Path $directory 'restore.png')
            $cancel = Wait-Element $dialog 'CANCEL'
            if (-not [LayoutQaNative]::PostMessage([IntPtr]$cancel.Current.NativeWindowHandle, 0xF5, [IntPtr]::Zero, [IntPtr]::Zero)) { throw 'Cancel message rejected' }
            $cases.Add(@{ id = "layout-${dpi}-$language"; status = 'PASS'; layoutDpi = $dpi; language = $language;
                expected = @{ caption = $expectedCaption; headerInsideWindow = $true; languageButtonsOverlap = $false };
                observed = @{ caption = $apply.Current.Name; geometry = $geometry; window = @{ width = $bounds.Width; height = $bounds.Height } };
                evidence = @($frame, $dialogFrame); processReport = $report; log = $metadata.logPath })
            if (-not (Stop-QaOwnedProcess $process)) { throw 'Layout process required forced termination' }
            $process = $null
        }
    }
    if ($cases.Count -ne 6) { throw 'Layout matrix is incomplete' }
    $status = 'PASS'
} catch {
    $cases.Add(@{ id = "layout-${dpi}-$language"; status = 'FAIL'; observed = $_.Exception.Message })
    throw
} finally {
    if ($process -and -not $process.HasExited) { $process.Kill(); $process.WaitForExit() }
    @{ status = $status; scope = 'Developer layout scaling model at 100/125/150%; native capture viewport and full diagnostic bitmap. Actual Windows DPI transitions are not exercised';
        hardwareValidation = 'SKIPPED'; desktop = [DeviceTweakerQaDesktop]::CurrentDesktop(); cases = $cases.ToArray() } |
        ConvertTo-Json -Depth 9 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'layout.json') -Encoding UTF8
    foreach ($name in $names) { [Environment]::SetEnvironmentVariable($name, $saved[$name]) }
}
