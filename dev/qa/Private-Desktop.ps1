$ErrorActionPreference = 'Stop'
if (-not ('DeviceTweakerQaDesktop' -as [type])) {
    Add-Type -Path (Join-Path $PSScriptRoot 'PrivateDesktop.cs')
}

function Assert-QaPrivateDesktop {
    $actual = [DeviceTweakerQaDesktop]::CurrentDesktop()
    $expected = $env:DEVICE_TWEAKER_QA_DESKTOP
    if (-not $expected -or $actual -ne $expected -or $actual -notlike 'DeviceTweakerQa_*') {
        throw "QA worker is not on its assigned private desktop: actual=$actual expected=$expected"
    }
    if ([DeviceTweakerQaDesktop]::InputDesktop() -eq $actual) { throw 'QA desktop is visible.' }
}

function Start-QaOwnedProcess([string]$ExePath, [string]$Arguments) {
    if ($env:DEVICE_TWEAKER_QA_DESKTOP) { Assert-QaPrivateDesktop }
    $start = New-Object Diagnostics.ProcessStartInfo
    $start.FileName = $ExePath
    $start.Arguments = $Arguments
    $start.WorkingDirectory = Split-Path $ExePath
    # ShellExecute may route an app through Explorer on the user's desktop.
    $start.UseShellExecute = $false
    $start.CreateNoWindow = $true
    return [Diagnostics.Process]::Start($start)
}

function Stop-QaOwnedProcess([Diagnostics.Process]$Owned) {
    if ($Owned.HasExited) { return $Owned.ExitCode -eq 0 }
    if ($env:DEVICE_TWEAKER_QA_DESKTOP) { Assert-QaPrivateDesktop }
    # Process.CloseMainWindow may enumerate the input desktop and miss a private
    # window. Deliver WM_CLOSE to the HWNDs owned by this PID on our desktop.
    [DeviceTweakerQaDesktop]::CloseOwnedWindows($Owned.Id)
    if ($Owned.WaitForExit(5000)) { return $Owned.ExitCode -eq 0 }
    $Owned.Kill(); $Owned.WaitForExit()
    return $false # Forced cleanup is never a successful GUI shutdown.
}

function Invoke-QaPrivateDesktop {
    param([string]$ScriptPath, [System.Collections.IDictionary]$Parameters,
        [string]$ReportDirectory, [int]$TimeoutSeconds = 1800)
    if ($env:DEVICE_TWEAKER_QA_DESKTOP) {
        Assert-QaPrivateDesktop
        return $null # Already inside the worker; no recursive launch.
    }
    $destination = [IO.Path]::GetFullPath($ReportDirectory)
    New-Item -ItemType Directory -Force -Path $destination | Out-Null
    $log = Join-Path $destination 'worker.log'
    $plain = @{}
    foreach ($key in $Parameters.Keys) {
        $value = $Parameters[$key]
        if ($value -is [System.Management.Automation.SwitchParameter]) { $value = [bool]$value }
        $plain[$key] = $value
    }
    $payload = @{ script = [IO.Path]::GetFullPath($ScriptPath); parameters = $plain; log = $log } |
        ConvertTo-Json -Compress -Depth 8
    $encodedPayload = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($payload))
    $worker = @'
$ErrorActionPreference = 'Stop'
$payload = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('__PAYLOAD__')) | ConvertFrom-Json
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
$OutputEncoding = [Console]::OutputEncoding
$parameters = @{}
foreach ($property in $payload.parameters.PSObject.Properties) { $parameters[$property.Name] = $property.Value }
try {
    $global:LASTEXITCODE = 0
    & $payload.script @parameters *>> $payload.log
    $successful = $?
    $code = $LASTEXITCODE
    if (-not $successful -and $code -eq 0) { $code = 1 }
    exit $code
}
    catch { $_ | Out-String | Add-Content -LiteralPath $payload.log; exit 1 }
'@.Replace('__PAYLOAD__', $encodedPayload)
    $encodedCommand = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($worker))
    $hostExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $command = '"' + $hostExe + '" -NoProfile -NonInteractive -ExecutionPolicy Bypass -EncodedCommand ' + $encodedCommand
    $desktop = 'DeviceTweakerQa_' + [Guid]::NewGuid().ToString('N')
    $saved = $env:DEVICE_TWEAKER_QA_DESKTOP
    $result = $null
    try {
        $env:DEVICE_TWEAKER_QA_DESKTOP = $desktop
        $result = [DeviceTweakerQaDesktop]::Run($desktop, $hostExe, $command, (Split-Path $ScriptPath), $TimeoutSeconds)
        [ordered]@{ status = $(if ($result.exitCode -eq 0) { 'PASS' } else { 'FAIL' });
            desktop = $result.desktop; workerPid = $result.processId; exitCode = $result.exitCode;
            inputBefore = $result.inputBefore; inputAfter = $result.inputAfter;
            inputChecks = $result.inputChecks; cleanupConfirmed = $result.cleanupConfirmed;
            visibleFallback = $false } | ConvertTo-Json |
            Set-Content -LiteralPath (Join-Path $destination 'desktop.json') -Encoding UTF8
    } catch {
        @{ status = 'FAIL'; desktop = $desktop; error = $_.Exception.Message; visibleFallback = $false } |
            ConvertTo-Json | Set-Content -LiteralPath (Join-Path $destination 'desktop.json') -Encoding UTF8
        throw
    } finally { $env:DEVICE_TWEAKER_QA_DESKTOP = $saved }
    if (Test-Path -LiteralPath $log) { Get-Content -LiteralPath $log | Out-Host }
    return [int]$result.exitCode
}
