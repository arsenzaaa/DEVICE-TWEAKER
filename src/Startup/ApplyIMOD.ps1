param(
    [switch]$verbose
)

$imodScriptVersion = 30

{{IMOD_CONFIG_BLOCK}}

function Test-IsAdmin {
    try {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal $identity
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch {
        return $false
    }
}

function Start-ElevatedSelf {
    param(
        [string]$scriptPath,
        [string[]]$extraArgs
    )

    $psExe = (Get-Process -Id $PID).Path
    if (-not $psExe) {
        $psExe = 'powershell.exe'
    }

    $args = @(
        '-NoLogo',
        '-NoProfile',
        '-NonInteractive',
        '-WindowStyle', 'Hidden',
        '-ExecutionPolicy', 'Bypass',
        '-File', $scriptPath
    )
    if ($extraArgs) {
        $args += $extraArgs
    }

    Start-Process -FilePath $psExe -ArgumentList $args -Verb RunAs | Out-Null
}

function Write-ImodLog {
    param(
        [string]$message,
        [switch]$verboseOnly
    )
    try {
        if (-not $ImodStartupLogEnabled -and -not $verbose) {
            return
        }
        if ($verboseOnly -and -not $ImodStartupVerboseLogEnabled -and -not $verbose) {
            return
        }
        if ([string]::IsNullOrWhiteSpace($ImodLogPath)) {
            return
        }

        $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
        Add-Content -LiteralPath $ImodLogPath -Value "[$stamp] $message" -Encoding UTF8
    } catch {
    }
}

function Format-ImodHex {
    param([uint64]$value)
    return ('0x{0:X}' -f $value)
}

function Format-ImodVector {
    param($values)
    if ($null -eq $values) {
        return ''
    }

    $items = @($values | ForEach-Object { Format-ImodHex ([uint64]$_) })
    return ($items -join ',')
}

function Format-ImodText {
    param($value)
    if ($null -eq $value) {
        return '-'
    }

    $text = [string]$value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return '-'
    }

    return $text
}

function ConvertTo-ImodNormalizedDeviceId {
    param($value)
    if ($null -eq $value) {
        return ''
    }

    return (([string]$value) -replace '\\\\','\').Trim().ToUpperInvariant()
}

function Get-ImodVidPidKey {
    param([string]$deviceId)
    if ([string]::IsNullOrWhiteSpace($deviceId)) {
        return ''
    }

    if ($deviceId -match 'VID_([0-9A-Fa-f]{4})&PID_([0-9A-Fa-f]{4})') {
        return "$($Matches[1]):$($Matches[2])".ToUpperInvariant()
    }

    return ''
}

function Add-ImodRootPortRole {
    param(
        [hashtable]$rolesByPort,
        [int]$rootPort,
        [string]$role
    )

    if ($rootPort -le 0 -or [string]::IsNullOrWhiteSpace($role)) {
        return
    }

    if (-not $rolesByPort.ContainsKey($rootPort)) {
        $rolesByPort[$rootPort] = New-Object 'System.Collections.Generic.List[string]'
    }

    if (-not $rolesByPort[$rootPort].Contains($role)) {
        [void]$rolesByPort[$rootPort].Add($role)
    }
}

function Get-ImodUsbRootPort {
    param([string]$deviceId)

    if ([string]::IsNullOrWhiteSpace($deviceId)) {
        return $null
    }

    try {
        $properties = @(Get-PnpDeviceProperty -InstanceId $deviceId -KeyName @('DEVPKEY_Device_LocationPaths', 'DEVPKEY_Device_LocationInfo', 'DEVPKEY_Device_Address') -ErrorAction SilentlyContinue)
        $locPaths = $properties | Where-Object { $_.KeyName -eq 'DEVPKEY_Device_LocationPaths' } | Select-Object -First 1
        if ($locPaths -and $locPaths.Data) {
            foreach ($path in @($locPaths.Data)) {
                if ([string]$path -match 'USBROOT\(\d+\)#USB\((\d+)\)') {
                    return [int]$Matches[1]
                }
            }
        }

        $locInfo = $properties | Where-Object { $_.KeyName -eq 'DEVPKEY_Device_LocationInfo' } | Select-Object -First 1
        if ($locInfo -and $locInfo.Data -and [string]$locInfo.Data -match 'Port_#0*(\d+)\.Hub_#') {
            return [int]$Matches[1]
        }

        $address = $properties | Where-Object { $_.KeyName -eq 'DEVPKEY_Device_Address' } | Select-Object -First 1
        if ($address -and $null -ne $address.Data -and [int]$address.Data -gt 0) {
            return [int]$address.Data
        }
    } catch {
    }

    return $null
}

function Get-ImodUsbRootPortFromMap {
    param(
        [string]$deviceId,
        [hashtable]$propertyMap
    )

    $key = ConvertTo-ImodNormalizedDeviceId $deviceId
    if ([string]::IsNullOrWhiteSpace($key) -or $null -eq $propertyMap -or -not $propertyMap.ContainsKey($key)) {
        return (Get-ImodUsbRootPort $deviceId)
    }

    $props = $propertyMap[$key]
    if ($props.ContainsKey('DEVPKEY_Device_LocationPaths')) {
        foreach ($path in @($props['DEVPKEY_Device_LocationPaths'])) {
            if ([string]$path -match 'USBROOT\(\d+\)#USB\((\d+)\)') {
                return [int]$Matches[1]
            }
        }
    }

    if ($props.ContainsKey('DEVPKEY_Device_LocationInfo')) {
        $location = [string]$props['DEVPKEY_Device_LocationInfo']
        if ($location -match 'Port_#0*(\d+)\.Hub_#') {
            return [int]$Matches[1]
        }
    }

    if ($props.ContainsKey('DEVPKEY_Device_Address')) {
        $address = $props['DEVPKEY_Device_Address']
        if ($null -ne $address -and [int]$address -gt 0) {
            return [int]$address
        }
    }

    return $null
}

function Get-ImodParentDeviceId {
    param([string]$deviceId)
    try {
        $parent = Get-PnpDeviceProperty -InstanceId $deviceId -KeyName 'DEVPKEY_Device_Parent' -ErrorAction SilentlyContinue
        if ($parent -and $parent.Data) {
            return [string]$parent.Data
        }
    } catch {
    }

    return $null
}

function Add-ImodRoleByParentWalk {
    param(
        [hashtable]$rootPortByDeviceId,
        [hashtable]$rolesByPort,
        [string]$deviceId,
        [string]$role
    )

    $current = $deviceId
    for ($depth = 0; $depth -lt 12 -and -not [string]::IsNullOrWhiteSpace($current); $depth++) {
        $key = ConvertTo-ImodNormalizedDeviceId $current
        if ($rootPortByDeviceId.ContainsKey($key)) {
            Add-ImodRootPortRole -rolesByPort $rolesByPort -rootPort ([int]$rootPortByDeviceId[$key]) -role $role
            return
        }

        $current = Get-ImodParentDeviceId $current
    }
}

function Add-ImodRoleByControllerParentWalk {
    param(
        [System.Collections.Generic.HashSet[string]]$controllerDeviceIds,
        [hashtable]$rolesByPort,
        [string]$deviceId,
        [string]$role
    )

    $current = $deviceId
    for ($depth = 0; $depth -lt 12 -and -not [string]::IsNullOrWhiteSpace($current); $depth++) {
        $key = ConvertTo-ImodNormalizedDeviceId $current
        if ($controllerDeviceIds.Contains($key)) {
            $rootPort = Get-ImodUsbRootPort $current
            if ($null -ne $rootPort) {
                Add-ImodRootPortRole -rolesByPort $rolesByPort -rootPort ([int]$rootPort) -role $role
            }
            return
        }

        $current = Get-ImodParentDeviceId $current
    }
}

function ConvertTo-ImodRootPortRoleText {
    param([hashtable]$rolesByPort)
    if ($null -eq $rolesByPort -or $rolesByPort.Count -eq 0) {
        return ''
    }

    $parts = @()
    foreach ($port in @($rolesByPort.Keys | Sort-Object {[int]$_})) {
        $roles = @($rolesByPort[$port] | Sort-Object -Unique)
        if ($roles.Count -gt 0) {
            $parts += ("{0}={1}" -f ([int]$port), ($roles -join '+'))
        }
    }

    return ($parts -join ';')
}

function Resolve-ImodStartupRootPortRoles {
    param(
        [string]$controllerDeviceId,
        [string]$roleIntervalsText
    )

    $controllerKey = ConvertTo-ImodNormalizedDeviceId $controllerDeviceId
    if ([string]::IsNullOrWhiteSpace($controllerKey)) {
        return ''
    }

    $wantsAudio = $roleIntervalsText -match '(?i)(Audio|Microphone|Speaker)'
    $wantsGamepad = $roleIntervalsText -match '(?i)(Gamepad|Controller|Joystick)'
    $wantsWebcam = $roleIntervalsText -match '(?i)(Webcam|Camera)'
    $keyboardVidPid = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $mouseVidPid = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $keyboardDeviceIds = New-Object 'System.Collections.Generic.List[string]'
    $mouseDeviceIds = New-Object 'System.Collections.Generic.List[string]'
    $controllerDeviceIds = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $dependentDeviceIds = New-Object 'System.Collections.Generic.List[string]'
    $rolesByPort = @{}
    $foundKeyboardDirect = $false
    $foundMouseDirect = $false

    try {
        foreach ($keyboard in @(Get-CimInstance Win32_Keyboard -ErrorAction SilentlyContinue)) {
            $deviceId = [string]$keyboard.PNPDeviceID
            if (-not [string]::IsNullOrWhiteSpace($deviceId)) {
                [void]$keyboardDeviceIds.Add($deviceId)
            }
            $key = Get-ImodVidPidKey $deviceId
            if (-not [string]::IsNullOrWhiteSpace($key)) {
                [void]$keyboardVidPid.Add($key)
            }
        }
    } catch {
    }

    try {
        foreach ($mouse in @(Get-CimInstance Win32_PointingDevice -ErrorAction SilentlyContinue)) {
            $deviceId = [string]$mouse.PNPDeviceID
            if (-not [string]::IsNullOrWhiteSpace($deviceId)) {
                [void]$mouseDeviceIds.Add($deviceId)
            }
            $key = Get-ImodVidPidKey $deviceId
            if (-not [string]::IsNullOrWhiteSpace($key)) {
                [void]$mouseVidPid.Add($key)
            }
        }
    } catch {
    }

    try {
        foreach ($assoc in @(Get-WmiObject Win32_USBControllerDevice -ErrorAction SilentlyContinue)) {
            $antecedent = [string]$assoc.Antecedent
            $dependent = [string]$assoc.Dependent
            if ($antecedent -notmatch 'DeviceID="([^"]+)"' -or (ConvertTo-ImodNormalizedDeviceId $Matches[1]) -ne $controllerKey) {
                continue
            }
            if ($dependent -notmatch 'DeviceID="([^"]+)"') {
                continue
            }

            $depId = ([string]$Matches[1]) -replace '\\\\','\'
            $depKey = ConvertTo-ImodNormalizedDeviceId $depId
            if (-not [string]::IsNullOrWhiteSpace($depKey)) {
                [void]$controllerDeviceIds.Add($depKey)
            }
            [void]$dependentDeviceIds.Add($depId)
        }
    } catch {
    }

    foreach ($keyboardId in $keyboardDeviceIds) {
        $key = ConvertTo-ImodNormalizedDeviceId $keyboardId
        if ([string]::IsNullOrWhiteSpace($key) -or -not $controllerDeviceIds.Contains($key)) {
            continue
        }

        $rootPort = Get-ImodUsbRootPort $keyboardId
        if ($null -ne $rootPort) {
            Add-ImodRootPortRole -rolesByPort $rolesByPort -rootPort ([int]$rootPort) -role 'Keyboard'
            $foundKeyboardDirect = $true
        }
    }

    foreach ($mouseId in $mouseDeviceIds) {
        $key = ConvertTo-ImodNormalizedDeviceId $mouseId
        if ([string]::IsNullOrWhiteSpace($key) -or -not $controllerDeviceIds.Contains($key)) {
            continue
        }

        $rootPort = Get-ImodUsbRootPort $mouseId
        if ($null -ne $rootPort) {
            Add-ImodRootPortRole -rolesByPort $rolesByPort -rootPort ([int]$rootPort) -role 'Mouse'
            $foundMouseDirect = $true
        }
    }

    if (-not $foundKeyboardDirect -or -not $foundMouseDirect -or $wantsAudio -or $wantsGamepad -or $wantsWebcam) {
        foreach ($depId in $dependentDeviceIds) {
            $vidPid = Get-ImodVidPidKey $depId
            if (-not [string]::IsNullOrWhiteSpace($vidPid)) {
                if (-not $foundKeyboardDirect -and $keyboardVidPid.Contains($vidPid)) {
                    $rootPort = Get-ImodUsbRootPort $depId
                    if ($null -ne $rootPort) {
                        Add-ImodRootPortRole -rolesByPort $rolesByPort -rootPort ([int]$rootPort) -role 'Keyboard'
                        $foundKeyboardDirect = $true
                    }
                } elseif (-not $foundMouseDirect -and $mouseVidPid.Contains($vidPid)) {
                    $rootPort = Get-ImodUsbRootPort $depId
                    if ($null -ne $rootPort) {
                        Add-ImodRootPortRole -rolesByPort $rolesByPort -rootPort ([int]$rootPort) -role 'Mouse'
                        $foundMouseDirect = $true
                    }
                } elseif ($wantsAudio -or $wantsGamepad -or $wantsWebcam) {
                    try {
                        $pnp = Get-PnpDevice -InstanceId $depId -ErrorAction SilentlyContinue
                        if ($pnp) {
                            if ($wantsAudio -and ($pnp.Class -eq 'MEDIA' -or $pnp.Class -eq 'AudioEndpoint')) {
                                $rootPort = Get-ImodUsbRootPort $depId
                                if ($null -ne $rootPort) {
                                    Add-ImodRootPortRole -rolesByPort $rolesByPort -rootPort ([int]$rootPort) -role 'Audio'
                                }
                            } elseif ($wantsWebcam -and ($pnp.Class -eq 'Camera' -or $pnp.Class -eq 'Image' -or ([string]$pnp.FriendlyName) -match '(?i)(Webcam|Camera)')) {
                                $rootPort = Get-ImodUsbRootPort $depId
                                if ($null -ne $rootPort) {
                                    Add-ImodRootPortRole -rolesByPort $rolesByPort -rootPort ([int]$rootPort) -role 'Webcam'
                                }
                            } elseif ($wantsGamepad -and $pnp.Class -eq 'HIDClass') {
                                $rootPort = Get-ImodUsbRootPort $depId
                                if ($null -ne $rootPort) {
                                    Add-ImodRootPortRole -rolesByPort $rolesByPort -rootPort ([int]$rootPort) -role 'Gamepad'
                                }
                            }
                        }
                    } catch {
                    }
                }
            } elseif ($wantsAudio -and $depId -match '^SWD\\MMDEVAPI\\') {
                Add-ImodRoleByControllerParentWalk -controllerDeviceIds $controllerDeviceIds -rolesByPort $rolesByPort -deviceId $depId -role 'Audio'
            }
        }
    }

    if (-not $foundKeyboardDirect) {
        try {
            foreach ($keyboard in @(Get-CimInstance Win32_Keyboard -ErrorAction SilentlyContinue)) {
                Add-ImodRoleByControllerParentWalk -controllerDeviceIds $controllerDeviceIds -rolesByPort $rolesByPort -deviceId ([string]$keyboard.PNPDeviceID) -role 'Keyboard'
            }
        } catch {
        }
    }

    if (-not $foundMouseDirect) {
        try {
            foreach ($mouse in @(Get-CimInstance Win32_PointingDevice -ErrorAction SilentlyContinue)) {
                Add-ImodRoleByControllerParentWalk -controllerDeviceIds $controllerDeviceIds -rolesByPort $rolesByPort -deviceId ([string]$mouse.PNPDeviceID) -role 'Mouse'
            }
        } catch {
        }
    }

    if ($wantsAudio) {
        try {
            $audioDevices = @()
            $mediaDevices = @(Get-PnpDevice -Class 'MEDIA' -ErrorAction SilentlyContinue)
            $audioEndpoints = @(Get-PnpDevice -Class 'AudioEndpoint' -ErrorAction SilentlyContinue)
            if ($mediaDevices) { $audioDevices += $mediaDevices }
            if ($audioEndpoints) { $audioDevices += $audioEndpoints }
            foreach ($audio in $audioDevices) {
                Add-ImodRoleByControllerParentWalk -controllerDeviceIds $controllerDeviceIds -rolesByPort $rolesByPort -deviceId ([string]$audio.InstanceId) -role 'Audio'
            }
        } catch {
        }
    }

    if ($wantsGamepad) {
        try {
            foreach ($hid in @(Get-PnpDevice -Class 'HIDClass' -ErrorAction SilentlyContinue)) {
                $hidId = [string]$hid.InstanceId
                $vidPid = Get-ImodVidPidKey $hidId
                if (-not [string]::IsNullOrWhiteSpace($vidPid) -and ($keyboardVidPid.Contains($vidPid) -or $mouseVidPid.Contains($vidPid))) {
                    continue
                }

                $hidRole = 'Gamepad'
                try {
                    $compatibleIds = @((Get-PnpDeviceProperty -InstanceId $hidId -KeyName 'DEVPKEY_Device_CompatibleIds' -ErrorAction SilentlyContinue).Data)
                    foreach ($compatibleId in $compatibleIds) {
                        $cid = [string]$compatibleId
                        if ($cid -match 'Class_03.*SubClass_01.*Prot_01') {
                            $hidRole = 'Keyboard'
                            break
                        }
                        if ($cid -match 'Class_03.*SubClass_01.*Prot_02') {
                            $hidRole = 'Mouse'
                            break
                        }
                    }
                } catch {
                }

                Add-ImodRoleByControllerParentWalk -controllerDeviceIds $controllerDeviceIds -rolesByPort $rolesByPort -deviceId $hidId -role $hidRole
            }
        } catch {
        }
    }

    if ($wantsWebcam) {
        try {
            $cameraDevices = @()
            $cameraDevices += @(Get-PnpDevice -Class 'Camera' -ErrorAction SilentlyContinue)
            $cameraDevices += @(Get-PnpDevice -Class 'Image' -ErrorAction SilentlyContinue)
            $cameraDevices += @(Get-PnpDevice -PresentOnly -ErrorAction SilentlyContinue | Where-Object { ([string]$_.FriendlyName) -match '(?i)(Webcam|Camera)' })
            foreach ($camera in @($cameraDevices | Sort-Object InstanceId -Unique)) {
                Add-ImodRoleByControllerParentWalk -controllerDeviceIds $controllerDeviceIds -rolesByPort $rolesByPort -deviceId ([string]$camera.InstanceId) -role 'Webcam'
            }
        } catch {
        }
    }

    return (ConvertTo-ImodRootPortRoleText $rolesByPort)
}

$scriptRoot = $PSCommandPath
if ($scriptRoot) {
    $scriptRoot = Split-Path -Parent $scriptRoot
}
if (-not $scriptRoot) {
    $scriptRoot = $PSScriptRoot
}
if (-not $scriptRoot) {
    $scriptRoot = $env:TEMP
}
if ([string]::IsNullOrWhiteSpace($ImodLogDirectory)) {
    $ImodLogDirectory = $scriptRoot
}
try {
    New-Item -ItemType Directory -Path $ImodLogDirectory -Force | Out-Null
} catch {
    $ImodLogDirectory = $scriptRoot
}
$imodLogStamp = Get-Date -Format 'yyyyMMdd'
$ImodLogPath = Join-Path $ImodLogDirectory ("ApplyIMOD_{0}.log" -f $imodLogStamp)
Write-ImodLog "startup context: version=$imodScriptVersion pid=$PID root=$scriptRoot driver=$ImodDriverPath kdu=$ImodKduPath db=$ImodKduDbPath" -verboseOnly
$usbEntryCount = if ($userDefinedData) { $userDefinedData.Count } else { 0 }
$nicEntryCount = if ($nicItrData) { @($nicItrData).Count } else { 0 }
Write-ImodLog ("startup config: applyUsb=$applyUsbImod usbEntries=$usbEntryCount nicEntries=$nicEntryCount globalInterval=$(Format-ImodHex ([uint64]$globalInterval)) hcsparamsOffset=$(Format-ImodHex ([uint64]$globalHCSPARAMSOffset)) rtsoffOffset=$(Format-ImodHex ([uint64]$globalRTSOFF))")
if ($userDefinedData) {
    foreach ($key in $userDefinedData.Keys) {
        $cfg = $userDefinedData[$key]
        $enabledText = if ($cfg.ContainsKey('ENABLED')) { [string][bool]$cfg['ENABLED'] } else { 'default' }
        $intervalText = if ($cfg.ContainsKey('INTERVAL')) { Format-ImodHex ([uint64]$cfg['INTERVAL']) } else { 'default' }
        $intervalsText = if ($cfg.ContainsKey('INTERVALS')) { Format-ImodVector $cfg['INTERVALS'] } else { '' }
        $rolesText = if ($cfg.ContainsKey('ROLE_INTERVALS')) { [string]$cfg['ROLE_INTERVALS'] } else { '' }
        $hcsText = if ($cfg.ContainsKey('HCSPARAMS_OFFSET')) { Format-ImodHex ([uint64]$cfg['HCSPARAMS_OFFSET']) } else { 'default' }
        $rtsText = if ($cfg.ContainsKey('RTSOFF')) { Format-ImodHex ([uint64]$cfg['RTSOFF']) } else { 'default' }
        Write-ImodLog ("usb config: key=$key enabled=$enabledText interval=$intervalText intervals=[$intervalsText] roles=`"$rolesText`" hcsparamsOffset=$hcsText rtsoffOffset=$rtsText")
    }
}
if ($nicEntryCount -gt 0) {
    foreach ($nic in @($nicItrData)) {
        Write-ImodLog ("nic itr config: hwid=$(Format-ImodText $nic['HWID']) family=$(Format-ImodText $nic['FAMILY']) baseOffset=$(Format-ImodHex ([uint64]$nic['BASE_OFFSET'])) stride=$(Format-ImodHex ([uint64]$nic['STRIDE'])) queues=$($nic['QUEUES']) width=$($nic['WIDTH']) mask=$(Format-ImodHex ([uint64]$nic['MASK'])) orBits=$(Format-ImodHex ([uint64]$nic['OR_BITS'])) values=[$(Format-ImodVector $nic['VALUES'])]")
    }
} else {
    Write-ImodLog "nic itr config: entries=0"
}

if (-not (Test-IsAdmin)) {
    $extraArgs = @()
    if ($verbose) {
        $extraArgs += "-verbose"
    }

    try {
        Start-ElevatedSelf -scriptPath $PSCommandPath -extraArgs $extraArgs
        exit 0
    } catch {
        Write-Host "error: administrator privileges required"
        exit 1
    }
}

if (-not $ImodDriverPath -or -not (Test-Path $ImodDriverPath -PathType Leaf)) {
    Write-ImodLog "error: DTIMOD.sys not found: $ImodDriverPath"
    exit 1
}

Add-Type -Language CSharp -TypeDefinition @'
{{IMOD_RUNTIME_SOURCE}}
'@

function Invoke-ImodKduLoader {
    if ([DeviceTweakerImodRuntime]::IsDriverDeviceAvailable()) {
        Write-ImodLog "startup loader: existing device available"
        return $true
    }

    if ([string]::IsNullOrWhiteSpace($ImodKduPath) -or [string]::IsNullOrWhiteSpace($ImodKduDbPath)) {
        Write-ImodLog "startup loader: KDU paths missing"
        return $false
    }

    if (-not (Test-Path -LiteralPath $ImodKduPath -PathType Leaf)) {
        Write-ImodLog "startup loader: kdu.exe missing: $ImodKduPath"
        return $false
    }

    if (-not (Test-Path -LiteralPath $ImodKduDbPath -PathType Leaf)) {
        Write-ImodLog "startup loader: drv64.dll missing: $ImodKduDbPath"
        return $false
    }

    foreach ($payload in @(
        @('DTIMOD.sys', $ImodDriverPath, $ImodDriverSha256),
        @('kdu.exe', $ImodKduPath, $ImodKduSha256),
        @('drv64.dll', $ImodKduDbPath, $ImodKduDbSha256)
    )) {
        if ([string]::IsNullOrWhiteSpace([string]$payload[2])) {
            Write-ImodLog ("startup loader: missing trusted hash for " + $payload[0])
            return $false
        }
        try {
            $actualHash = (Get-FileHash -LiteralPath $payload[1] -Algorithm SHA256 -ErrorAction Stop).Hash
            if (-not $actualHash.Equals([string]$payload[2], [System.StringComparison]::OrdinalIgnoreCase)) {
                Write-ImodLog ("startup loader: hash mismatch for " + $payload[0] + " actual=" + $actualHash)
                return $false
            }
        } catch {
            Write-ImodLog ("startup loader: hash check failed for " + $payload[0] + ": " + $_.Exception.Message)
            return $false
        }
    }

    try {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $ImodKduPath
        $psi.WorkingDirectory = Split-Path -Parent $ImodKduPath
        $escapedDriver = $ImodDriverPath.Replace('"', '\"')
        $psi.Arguments = "-scv 3 -drvn DeviceTweakerImod2 -drvr DeviceTweakerImod2 -map `"$escapedDriver`""
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true

        Write-ImodLog "startup loader: KDU map start"
        $proc = [System.Diagnostics.Process]::Start($psi)
        if ($null -eq $proc) {
            Write-ImodLog "startup loader: KDU process failed to start"
            return $false
        }

        $stdoutTask = $proc.StandardOutput.ReadToEndAsync()
        $stderrTask = $proc.StandardError.ReadToEndAsync()
        if (-not $proc.WaitForExit(60000)) {
            try { $proc.Kill() } catch {}
            Write-ImodLog "startup loader: KDU timeout"
            try { $proc.WaitForExit(2000) } catch {}
        }

        $stdout = if ($stdoutTask.IsCompleted) { $stdoutTask.GetAwaiter().GetResult() } else { "<capture incomplete>" }
        $stderr = if ($stderrTask.IsCompleted) { $stderrTask.GetAwaiter().GetResult() } else { "<capture incomplete>" }
        foreach ($stream in @(@("STDOUT", $stdout), @("STDERR", $stderr))) {
            $normalized = ([string]$stream[1]).Replace("`r`n", "`n").Replace("`r", "`n")
            $lines = $normalized.Split("`n")
            Write-ImodLog ("startup loader: KDU " + $stream[0] + " BEGIN chars=" + $normalized.Length + " lines=" + $lines.Length)
            if ($normalized.Length -eq 0) {
                Write-ImodLog ("startup loader: KDU " + $stream[0] + ": <empty>")
            } else {
                for ($lineIndex = 0; $lineIndex -lt $lines.Length; $lineIndex++) {
                    Write-ImodLog ("startup loader: KDU " + $stream[0] + " " + ($lineIndex + 1).ToString("D4") + ": " + $lines[$lineIndex])
                }
            }
            Write-ImodLog ("startup loader: KDU " + $stream[0] + " END")
        }

        if (-not $proc.HasExited) {
            return $false
        }

        Write-ImodLog ("startup loader: KDU exit=" + $proc.ExitCode)
        if ([DeviceTweakerImodRuntime]::IsDriverDeviceAvailable()) {
            Write-ImodLog "startup loader: KDU device available"
            return $true
        }
    } catch {
        Write-ImodLog ("startup loader: KDU failed: " + $_.Exception.Message)
    }

    return $false
}

try {
    Write-ImodLog "startup apply begin; driver=$ImodDriverPath loader=kdu-first"
    if (-not (Invoke-ImodKduLoader)) {
        Write-ImodLog "startup loader: fallback service"
        [DeviceTweakerImodRuntime]::EnsureDriver($ImodDriverPath)
    }

    $appliedUsb = 0
    if ($applyUsbImod) {
        $controllers = [DeviceTweakerImodRuntime]::EnumerateXhciControllers()
        Write-ImodLog ("usb controllers=" + $controllers.Count)
        foreach ($controller in $controllers) {
            $controllerBaseText = if ($controller.HasBase) { Format-ImodHex ([uint64]$controller.BaseAddress) } else { '-' }
            Write-ImodLog ("usb controller: id=$($controller.DeviceId) caption=$(Format-ImodText $controller.Caption) problem=$($controller.ProblemCode) hasBase=$($controller.HasBase) base=$controllerBaseText length=$(Format-ImodHex ([uint64]$controller.MemoryLength))") -verboseOnly
            if ([DeviceTweakerImodRuntime]::IsDisabledProblem([uint32]$controller.ProblemCode)) {
                Write-ImodLog ("skip disabled " + $controller.DeviceId)
                continue
            }
            if (-not $controller.HasBase) {
                Write-ImodLog ("skip missing base " + $controller.DeviceId + " " + $controller.BaseError)
                continue
            }

            $entry = $null
            $matchedKey = ''
            foreach ($key in $userDefinedData.Keys) {
                if ($controller.DeviceId.IndexOf([string]$key, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                    $entry = $userDefinedData[$key]
                    $matchedKey = [string]$key
                }
            }

            $enabled = $true
            $interval = [uint32]$globalInterval
            $hcsparamsOffset = [uint32]$globalHCSPARAMSOffset
            $rtsoff = [uint32]$globalRTSOFF
            $intervalList = @()

            if ($null -ne $entry) {
                if ($entry.ContainsKey('ENABLED')) {
                    $enabled = [bool]$entry['ENABLED']
                }
                if ($entry.ContainsKey('INTERVAL')) {
                    $interval = [uint32]$entry['INTERVAL']
                }
                if ($entry.ContainsKey('INTERVALS')) {
                    $intervalList = @($entry['INTERVALS'] | ForEach-Object { [uint32]$_ })
                }
                if ($entry.ContainsKey('HCSPARAMS_OFFSET')) {
                    $hcsparamsOffset = [uint32]$entry['HCSPARAMS_OFFSET']
                }
                if ($entry.ContainsKey('RTSOFF')) {
                    $rtsoff = [uint32]$entry['RTSOFF']
                }
            }

            $adaptiveBinding = $false
            if ($null -ne $entry -and $entry.ContainsKey('ADAPTIVE_ROLE_BINDING')) {
                $adaptiveBinding = [bool]$entry['ADAPTIVE_ROLE_BINDING']
            }
            $roleText = if ($null -ne $entry -and $entry.ContainsKey('ROLE_INTERVALS')) { [string]$entry['ROLE_INTERVALS'] } else { '' }
            Write-ImodLog ("usb apply plan: id=$($controller.DeviceId) matchedKey=$(Format-ImodText $matchedKey) enabled=$enabled adaptive=$adaptiveBinding interval=$(Format-ImodHex ([uint64]$interval)) intervals=[$(Format-ImodVector $intervalList)] roles=`"$roleText`" hcsparamsOffset=$(Format-ImodHex ([uint64]$hcsparamsOffset)) rtsoffOffset=$(Format-ImodHex ([uint64]$rtsoff)) base=$controllerBaseText")

            if (-not $enabled) {
                Write-ImodLog ("skip config-disabled " + $controller.DeviceId)
                continue
            }

            if ($intervalList.Count -gt 0) {
                $intervalArray = [uint32[]]$intervalList
            } else {
                $intervalArray = New-Object 'System.UInt32[]' 0
            }
            $adaptiveApplied = $false
            [uint32[]]$adaptiveIntervals = $null
            if ($adaptiveBinding -and -not [string]::IsNullOrWhiteSpace($roleText)) {
                try {
                    $rootPortRoles = Resolve-ImodStartupRootPortRoles -controllerDeviceId $controller.DeviceId -roleIntervalsText $roleText
                    Write-ImodLog ("usb adaptive role scan: id=$($controller.DeviceId) rootPorts=`"$rootPortRoles`"")
                    if (-not [string]::IsNullOrWhiteSpace($rootPortRoles)) {
                        $adaptiveResult = [DeviceTweakerImodRuntime]::TryBuildAdaptiveIntervals(
                            [uint64]$controller.BaseAddress,
                            [uint64]$controller.MemoryLength,
                            $hcsparamsOffset,
                            $interval,
                            $roleText,
                            $rootPortRoles,
                            $intervalArray,
                            [ref]$adaptiveIntervals)
                        Write-ImodLog ("usb adaptive result: id=$($controller.DeviceId) $adaptiveResult")
                        if ($adaptiveResult.StartsWith('ok:', [System.StringComparison]::OrdinalIgnoreCase) -and $adaptiveIntervals -and $adaptiveIntervals.Count -gt 0) {
                            $adaptiveApplied = $true
                        }
                    }
                } catch {
                    Write-ImodLog ("usb adaptive failed: id=$($controller.DeviceId) " + $_.Exception.Message)
                }
            }
            if ($adaptiveApplied) {
                $intervalArray = $adaptiveIntervals
            }

            $result = [DeviceTweakerImodRuntime]::ApplyController([uint64]$controller.BaseAddress, [uint64]$controller.MemoryLength, $hcsparamsOffset, $rtsoff, $interval, $intervalArray)
            Write-ImodLog ($controller.DeviceId + " " + $result)
            if (-not $result.StartsWith('error:', [System.StringComparison]::OrdinalIgnoreCase)) {
                $appliedUsb++
            }
        }
    } else {
        Write-ImodLog "usb imod skipped; no custom USB IMOD config"
    }

    $appliedNic = 0
    if ($nicItrData -and $nicItrData.Count -gt 0) {
        $pciDevices = [DeviceTweakerImodRuntime]::EnumeratePciDevices()
        Write-ImodLog ("nic itr entries=" + $nicItrData.Count + " pci=" + $pciDevices.Count)
        foreach ($device in $pciDevices) {
            $pciBaseText = if ($device.HasBase) { Format-ImodHex ([uint64]$device.BaseAddress) } else { '-' }
            Write-ImodLog ("nic pci: id=$($device.DeviceId) caption=$(Format-ImodText $device.Caption) problem=$($device.ProblemCode) hasBase=$($device.HasBase) base=$pciBaseText") -verboseOnly
        }
        foreach ($nic in $nicItrData) {
            $hwid = [string]$nic['HWID']
            if ([string]::IsNullOrWhiteSpace($hwid)) {
                continue
            }

            Write-ImodLog ("nic itr apply plan: hwid=$hwid family=$(Format-ImodText $nic['FAMILY']) baseOffset=$(Format-ImodHex ([uint64]$nic['BASE_OFFSET'])) stride=$(Format-ImodHex ([uint64]$nic['STRIDE'])) queues=$($nic['QUEUES']) width=$($nic['WIDTH']) mask=$(Format-ImodHex ([uint64]$nic['MASK'])) orBits=$(Format-ImodHex ([uint64]$nic['OR_BITS'])) values=[$(Format-ImodVector $nic['VALUES'])]")

            $target = $null
            foreach ($device in $pciDevices) {
                if ($device.DeviceId.IndexOf($hwid, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                    $target = $device
                }
            }

            if ($null -eq $target) {
                Write-ImodLog ("nic itr skip missing " + $hwid)
                continue
            }
            $targetBaseText = if ($target.HasBase) { Format-ImodHex ([uint64]$target.BaseAddress) } else { '-' }
            Write-ImodLog ("nic itr matched: hwid=$hwid id=$($target.DeviceId) caption=$(Format-ImodText $target.Caption) problem=$($target.ProblemCode) hasBase=$($target.HasBase) base=$targetBaseText")
            if ([DeviceTweakerImodRuntime]::IsDisabledProblem([uint32]$target.ProblemCode)) {
                Write-ImodLog ("nic itr skip disabled " + $target.DeviceId)
                continue
            }
            if (-not $target.HasBase) {
                Write-ImodLog ("nic itr skip missing base " + $target.DeviceId + " " + $target.BaseError)
                continue
            }

            $values = [uint64[]]@($nic['VALUES'] | ForEach-Object { [uint64]$_ })
            $result = [DeviceTweakerImodRuntime]::ApplyNicItr(
                [uint64]$target.BaseAddress,
                [uint64]$target.MemoryLength,
                [uint32]$nic['BASE_OFFSET'],
                [uint32]$nic['STRIDE'],
                [uint32]$nic['QUEUES'],
                [uint32]$nic['WIDTH'],
                [uint64]$nic['MASK'],
                [uint64]$nic['OR_BITS'],
                $values)
            Write-ImodLog ("nic itr " + $target.DeviceId + " " + $result)
            if (-not $result.StartsWith('error:', [System.StringComparison]::OrdinalIgnoreCase)) {
                $appliedNic++
            }
        }
    } else {
        Write-ImodLog "nic itr skipped; entries=0"
    }

    Write-ImodLog "startup apply done; usb=$appliedUsb nic=$appliedNic"
    exit 0
} catch {
    Write-ImodLog ("error: " + $_.Exception.Message)
    exit 1
}
