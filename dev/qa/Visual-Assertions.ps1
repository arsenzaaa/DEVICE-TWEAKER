# Pixel assertion shared by live GUI smoke and its positive/blank-image controls.
function Assert-RegionRendered([string]$path, $windowRect, $elementRect, [string]$label) {
    if ($null -eq $elementRect -or $elementRect.Width -lt 2 -or $elementRect.Height -lt 2) {
        throw "Visual region is missing or collapsed: $label"
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
                if (([int]$pixel.R + [int]$pixel.G + [int]$pixel.B) -gt 210) { $textPixels++ }
            }
        }
        if ($textPixels -lt 30) { throw "Visual text disappeared from capture: $label pixels=$textPixels path=$path" }
        Write-Report "PASS: rendered $label pixels=$textPixels"
    }
    finally { $bmp.Dispose() }
}
