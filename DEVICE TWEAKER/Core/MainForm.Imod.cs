using System.Globalization;
using System.Text;

namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private const uint ImodDefaultInterval = 0xC8;
    private const uint ImodDefaultHcsparamsOffset = 0x4;
    private const uint ImodDefaultRtsoff = 0x18;
    private const string ImodScriptFileName = "ApplyIMOD.ps1";
    private const string ImodDriverName = "DTIMOD.sys";
    private const string ImodScriptMarkerStart = "$imodSettingsBegin = $true";
    private const string ImodScriptMarkerEnd = "$imodSettingsEnd = $true";
    private const string ImodScriptVersionMarker = "$imodScriptVersion = 26";
    private const string ImodScriptConfigToken = "{{IMOD_CONFIG_BLOCK}}";
    private const bool ImodStartupScriptLoggingEnabled = false;
    private const bool ImodStartupScriptVerboseLoggingEnabled = false;
    private static readonly string ImodScriptTemplate = """
    param(
        [switch]$verbose
    )
    
    $imodScriptVersion = 26
    
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
    
        $args = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $scriptPath)
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

        $wantsAudio = $false
        $wantsGamepad = $roleIntervalsText -match '(?i)(Gamepad|Controller|Joystick)'
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

        if (-not $foundKeyboardDirect -or -not $foundMouseDirect -or $wantsAudio -or $wantsGamepad) {
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
                    } elseif ($wantsAudio -or $wantsGamepad) {
                        try {
                            $pnp = Get-PnpDevice -InstanceId $depId -ErrorAction SilentlyContinue
                            if ($pnp) {
                                if ($wantsAudio -and ($pnp.Class -eq 'MEDIA' -or $pnp.Class -eq 'AudioEndpoint')) {
                                    $rootPort = Get-ImodUsbRootPort $depId
                                    if ($null -ne $rootPort) {
                                        Add-ImodRootPortRole -rolesByPort $rolesByPort -rootPort ([int]$rootPort) -role 'Audio'
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
    $ImodLogPath = Join-Path $scriptRoot 'ApplyIMOD.log'
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
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Globalization;
    using System.Runtime.InteropServices;
    using System.Text;

    public sealed class DeviceTweakerImodController
    {
        public string DeviceId;
        public string Caption;
        public uint ProblemCode;
        public ulong BaseAddress;
        public bool HasBase;
        public string BaseError;
    }

    public static class DeviceTweakerImodRuntime
    {
        private const uint CmProbDisabled = 0x00000016;
        private const uint DigcfPresent = 0x00000002;
        private const uint DigcfAllClasses = 0x00000004;
        private const uint SpdrpDeviceDesc = 0x00000000;
        private const uint SpdrpHardwareId = 0x00000001;
        private const uint SpdrpCompatibleIds = 0x00000002;
        private const uint SpdrpService = 0x00000004;
        private const uint SpdrpFriendlyName = 0x0000000C;
        private const uint RegSz = 1;
        private const uint RegMultiSz = 7;
        private const uint AllocLogConf = 0x00000002;
        private const uint BootLogConf = 0x00000003;
        private const uint ResTypeMem = 0x00000001;
        private const uint ResTypeMemLarge = 0x00000007;
        private const int CrSuccess = 0;
        private const int ErrorInsufficientBuffer = 122;
        private const int ErrorNoMoreItems = 259;
        private const int ErrorServiceDoesNotExist = 1060;
        private const int ErrorServiceAlreadyRunning = 1056;
        private const int ErrorServiceNotActive = 1062;
        private const uint ScManagerAllAccess = 0x000F003F;
        private const uint ServiceAllAccess = 0x000F01FF;
        private const uint ServiceKernelDriver = 0x00000001;
        private const uint ServiceDemandStart = 0x00000003;
        private const uint ServiceErrorNormal = 0x00000001;
        private const uint ServiceControlStop = 0x00000001;
        private const uint ServiceRunning = 0x00000004;
        private const uint ServiceStopped = 0x00000001;
        private const uint ServiceStopPending = 0x00000003;
        private const int ScStatusProcessInfo = 0;
        private const uint GenericRead = 0x80000000;
        private const uint GenericWrite = 0x40000000;
        private const uint FileShareRead = 0x00000001;
        private const uint FileShareWrite = 0x00000002;
        private const uint OpenExisting = 3;
        private const uint FileAttributeNormal = 0x00000080;
        private const uint FileDeviceImod = 0x00008010;
        private const uint ImodIoctlIndex = 0x810;
        private const uint MethodBuffered = 0;
        private const uint FileAnyAccess = 0;
        private const int OpenRetryCount = 10;
        private const int OpenRetryDelayMs = 100;
        private const string ImodDriverDevicePath = "\\\\.\\DeviceTweakerImod2";
        private const string ImodDriverServiceName = "DeviceTweakerImod2";
        private static readonly IntPtr InvalidHandleValue = new IntPtr(-1);
        private static readonly uint IoctlImodReadPhysicalMemory = CtlCode(FileDeviceImod, ImodIoctlIndex + 2, MethodBuffered, FileAnyAccess);
        private static readonly uint IoctlImodWritePhysicalMemory = CtlCode(FileDeviceImod, ImodIoctlIndex + 3, MethodBuffered, FileAnyAccess);

        [DllImport("setupapi.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr SetupDiGetClassDevsW(IntPtr classGuid, string enumerator, IntPtr hwndParent, uint flags);

        [DllImport("setupapi.dll", SetLastError = true)]
        private static extern bool SetupDiEnumDeviceInfo(IntPtr deviceInfoSet, uint memberIndex, ref SP_DEVINFO_DATA deviceInfoData);

        [DllImport("setupapi.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool SetupDiGetDeviceRegistryPropertyW(IntPtr deviceInfoSet, ref SP_DEVINFO_DATA deviceInfoData, uint property, out uint propertyRegDataType, byte[] propertyBuffer, uint propertyBufferSize, out uint requiredSize);

        [DllImport("setupapi.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool SetupDiGetDeviceInstanceIdW(IntPtr deviceInfoSet, ref SP_DEVINFO_DATA deviceInfoData, StringBuilder deviceInstanceId, int deviceInstanceIdSize, out int requiredSize);

        [DllImport("setupapi.dll", SetLastError = true)]
        private static extern bool SetupDiDestroyDeviceInfoList(IntPtr deviceInfoSet);

        [DllImport("cfgmgr32.dll", SetLastError = true)]
        private static extern int CM_Get_DevNode_Status(out uint status, out uint problem, uint devInst, uint flags);

        [DllImport("cfgmgr32.dll", SetLastError = true)]
        private static extern int CM_Get_First_Log_Conf(out IntPtr logConf, uint devInst, uint flags);

        [DllImport("cfgmgr32.dll", SetLastError = true)]
        private static extern int CM_Get_Next_Res_Des(out IntPtr resDes, IntPtr logConfOrResDes, uint forResource, IntPtr resourceId, uint flags);

        [DllImport("cfgmgr32.dll", SetLastError = true)]
        private static extern int CM_Get_Res_Des_Data_Size(out uint dataSize, IntPtr resDes, uint flags);

        [DllImport("cfgmgr32.dll", SetLastError = true)]
        private static extern int CM_Get_Res_Des_Data(IntPtr resDes, byte[] buffer, uint bufferLen, uint flags);

        [DllImport("cfgmgr32.dll", SetLastError = true)]
        private static extern int CM_Free_Res_Des_Handle(IntPtr resDes);

        [DllImport("cfgmgr32.dll", SetLastError = true)]
        private static extern int CM_Free_Log_Conf_Handle(IntPtr logConf);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr CreateFile(string fileName, uint desiredAccess, uint shareMode, IntPtr securityAttributes, uint creationDisposition, uint flagsAndAttributes, IntPtr templateFile);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool DeviceIoControl(IntPtr deviceHandle, uint ioControlCode, ref PhysAccessStruct inBuffer, int inBufferSize, ref PhysAccessStruct outBuffer, int outBufferSize, out int bytesReturned, IntPtr overlapped);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool CloseHandle(IntPtr handle);

        [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr OpenSCManager(string machineName, string databaseName, uint desiredAccess);

        [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr OpenService(IntPtr serviceManager, string serviceName, uint desiredAccess);

        [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr CreateService(
            IntPtr serviceManager,
            string serviceName,
            string displayName,
            uint desiredAccess,
            uint serviceType,
            uint startType,
            uint errorControl,
            string binaryPathName,
            string loadOrderGroup,
            IntPtr tagId,
            string dependencies,
            string serviceStartName,
            string password);

        [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool ChangeServiceConfig(
            IntPtr service,
            uint serviceType,
            uint startType,
            uint errorControl,
            string binaryPathName,
            string loadOrderGroup,
            IntPtr tagId,
            string dependencies,
            string serviceStartName,
            string password,
            string displayName);

        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern bool StartService(IntPtr service, uint numServiceArgs, IntPtr serviceArgVectors);

        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern bool QueryServiceStatusEx(IntPtr service, int infoLevel, ref SERVICE_STATUS_PROCESS status, uint bufferSize, out uint bytesNeeded);

        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern bool ControlService(IntPtr service, uint control, ref SERVICE_STATUS serviceStatus);

        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern bool CloseServiceHandle(IntPtr handle);

        public static void EnsureDriver(string driverPath)
        {
            IntPtr handle;
            string openError;
            if (TryOpenDriverDevice(1, out handle, out openError))
            {
                CloseHandle(handle);
                return;
            }

            EnsureDriverService(driverPath);

            if (!TryOpenDriverDevice(OpenRetryCount, out handle, out openError))
            {
                throw new InvalidOperationException(openError);
            }

            CloseHandle(handle);
        }

        public static bool IsDriverDeviceAvailable()
        {
            IntPtr handle;
            string openError;
            if (!TryOpenDriverDevice(1, out handle, out openError))
            {
                return false;
            }

            CloseHandle(handle);
            return true;
        }

        private static void EnsureDriverService(string driverPath)
        {
            IntPtr scm = OpenSCManager(null, null, ScManagerAllAccess);
            if (scm == IntPtr.Zero)
            {
                throw new InvalidOperationException("failed to open service manager: " + GetWin32ErrorMessage(Marshal.GetLastWin32Error()));
            }

            try
            {
                IntPtr service = OpenService(scm, ImodDriverServiceName, ServiceAllAccess);
                if (service == IntPtr.Zero)
                {
                    int openError = Marshal.GetLastWin32Error();
                    if (openError != ErrorServiceDoesNotExist)
                    {
                        throw new InvalidOperationException("failed to open IMOD driver service: " + GetWin32ErrorMessage(openError));
                    }

                    service = CreateService(
                        scm,
                        ImodDriverServiceName,
                        ImodDriverServiceName,
                        ServiceAllAccess,
                        ServiceKernelDriver,
                        ServiceDemandStart,
                        ServiceErrorNormal,
                        driverPath,
                        null,
                        IntPtr.Zero,
                        null,
                        null,
                        null);
                    if (service == IntPtr.Zero)
                    {
                        throw new InvalidOperationException("failed to create IMOD driver service: " + GetWin32ErrorMessage(Marshal.GetLastWin32Error()));
                    }
                }
                else if (!ChangeServiceConfig(
                             service,
                             ServiceKernelDriver,
                             ServiceDemandStart,
                             ServiceErrorNormal,
                             driverPath,
                             null,
                             IntPtr.Zero,
                             null,
                             null,
                             null,
                             ImodDriverServiceName))
                {
                    throw new InvalidOperationException("failed to configure IMOD driver service path: " + GetWin32ErrorMessage(Marshal.GetLastWin32Error()));
                }

                try
                {
                    SERVICE_STATUS_PROCESS status = new SERVICE_STATUS_PROCESS();
                    uint bytesNeeded;
                    if (QueryServiceStatusEx(service, ScStatusProcessInfo, ref status, (uint)Marshal.SizeOf(typeof(SERVICE_STATUS_PROCESS)), out bytesNeeded)
                        && status.dwCurrentState == ServiceRunning)
                    {
                        IntPtr runningHandle;
                        string runningOpenError;
                        if (TryOpenDriverDevice(1, out runningHandle, out runningOpenError))
                        {
                            CloseHandle(runningHandle);
                            return;
                        }

                        StopDriverService(service);
                    }

                    if (!StartService(service, 0, IntPtr.Zero))
                    {
                        int startError = Marshal.GetLastWin32Error();
                        if (startError != ErrorServiceAlreadyRunning)
                        {
                            throw new InvalidOperationException("failed to start IMOD driver service: " + GetWin32ErrorMessage(startError));
                        }
                    }
                }
                finally
                {
                    CloseServiceHandle(service);
                }
            }
            finally
            {
                CloseServiceHandle(scm);
            }
        }

        private static void StopDriverService(IntPtr service)
        {
            SERVICE_STATUS_PROCESS status = new SERVICE_STATUS_PROCESS();
            uint bytesNeeded;
            if (!QueryServiceStatusEx(service, ScStatusProcessInfo, ref status, (uint)Marshal.SizeOf(typeof(SERVICE_STATUS_PROCESS)), out bytesNeeded)
                || status.dwCurrentState == ServiceStopped)
            {
                return;
            }

            if (status.dwCurrentState != ServiceStopPending)
            {
                SERVICE_STATUS stopStatus = new SERVICE_STATUS();
                if (!ControlService(service, ServiceControlStop, ref stopStatus))
                {
                    int stopError = Marshal.GetLastWin32Error();
                    if (stopError != ErrorServiceNotActive)
                    {
                        throw new InvalidOperationException("failed to stop stale IMOD driver service: " + GetWin32ErrorMessage(stopError));
                    }
                }
            }

            for (int i = 0; i < 25; i++)
            {
                SERVICE_STATUS_PROCESS check = new SERVICE_STATUS_PROCESS();
                if (!QueryServiceStatusEx(service, ScStatusProcessInfo, ref check, (uint)Marshal.SizeOf(typeof(SERVICE_STATUS_PROCESS)), out bytesNeeded)
                    || check.dwCurrentState == ServiceStopped)
                {
                    return;
                }

                System.Threading.Thread.Sleep(200);
            }

            throw new InvalidOperationException("timed out while stopping stale IMOD driver service");
        }

        public static DeviceTweakerImodController[] EnumerateXhciControllers()
        {
            List<DeviceTweakerImodController> controllers = new List<DeviceTweakerImodController>();
            IntPtr devInfoSet = SetupDiGetClassDevsW(IntPtr.Zero, "PCI", IntPtr.Zero, DigcfPresent | DigcfAllClasses);
            if (devInfoSet == InvalidHandleValue)
            {
                throw new InvalidOperationException("failed to enumerate PCI devices: " + GetWin32ErrorMessage(Marshal.GetLastWin32Error()));
            }

            try
            {
                for (uint index = 0; ; index++)
                {
                    SP_DEVINFO_DATA devInfo = new SP_DEVINFO_DATA();
                    devInfo.cbSize = (uint)Marshal.SizeOf(typeof(SP_DEVINFO_DATA));

                    if (!SetupDiEnumDeviceInfo(devInfoSet, index, ref devInfo))
                    {
                        int lastError = Marshal.GetLastWin32Error();
                        if (lastError == ErrorNoMoreItems)
                        {
                            break;
                        }
                        throw new InvalidOperationException("failed to enumerate device info: " + GetWin32ErrorMessage(lastError));
                    }

                    if (!IsXhciDevice(devInfoSet, ref devInfo))
                    {
                        continue;
                    }

                    string instanceId;
                    if (!TryGetDeviceInstanceId(devInfoSet, ref devInfo, out instanceId))
                    {
                        continue;
                    }

                    uint problemCode;
                    TryGetDeviceProblemCode(devInfo.DevInst, out problemCode);

                    ulong baseAddress;
                    string baseError;
                    bool hasBase = TryGetDeviceMemoryBase(devInfo.DevInst, out baseAddress, out baseError);

                    DeviceTweakerImodController controller = new DeviceTweakerImodController();
                    controller.DeviceId = instanceId;
                    controller.Caption = GetDeviceCaption(devInfoSet, ref devInfo);
                    controller.ProblemCode = problemCode;
                    controller.BaseAddress = baseAddress;
                    controller.HasBase = hasBase;
                    controller.BaseError = baseError ?? string.Empty;
                    controllers.Add(controller);
                }
            }
            finally
            {
                SetupDiDestroyDeviceInfoList(devInfoSet);
            }

            return controllers.ToArray();
        }

        public static DeviceTweakerImodController[] EnumeratePciDevices()
        {
            List<DeviceTweakerImodController> devices = new List<DeviceTweakerImodController>();
            IntPtr devInfoSet = SetupDiGetClassDevsW(IntPtr.Zero, "PCI", IntPtr.Zero, DigcfPresent | DigcfAllClasses);
            if (devInfoSet == InvalidHandleValue)
            {
                throw new InvalidOperationException("failed to enumerate PCI devices: " + GetWin32ErrorMessage(Marshal.GetLastWin32Error()));
            }

            try
            {
                for (uint index = 0; ; index++)
                {
                    SP_DEVINFO_DATA devInfo = new SP_DEVINFO_DATA();
                    devInfo.cbSize = (uint)Marshal.SizeOf(typeof(SP_DEVINFO_DATA));

                    if (!SetupDiEnumDeviceInfo(devInfoSet, index, ref devInfo))
                    {
                        int lastError = Marshal.GetLastWin32Error();
                        if (lastError == ErrorNoMoreItems)
                        {
                            break;
                        }
                        throw new InvalidOperationException("failed to enumerate device info: " + GetWin32ErrorMessage(lastError));
                    }

                    string instanceId;
                    if (!TryGetDeviceInstanceId(devInfoSet, ref devInfo, out instanceId))
                    {
                        continue;
                    }

                    uint problemCode;
                    TryGetDeviceProblemCode(devInfo.DevInst, out problemCode);

                    ulong baseAddress;
                    string baseError;
                    bool hasBase = TryGetDeviceMemoryBase(devInfo.DevInst, out baseAddress, out baseError);

                    DeviceTweakerImodController device = new DeviceTweakerImodController();
                    device.DeviceId = instanceId;
                    device.Caption = GetDeviceCaption(devInfoSet, ref devInfo);
                    device.ProblemCode = problemCode;
                    device.BaseAddress = baseAddress;
                    device.HasBase = hasBase;
                    device.BaseError = baseError ?? string.Empty;
                    devices.Add(device);
                }
            }
            finally
            {
                SetupDiDestroyDeviceInfoList(devInfoSet);
            }

            return devices.ToArray();
        }

        public static string ApplyController(ulong capabilityAddress, uint hcsparamsOffset, uint rtsoff, uint interval, uint[] intervals)
        {
            IntPtr handle;
            string openError;
            if (!TryOpenDriverDevice(OpenRetryCount, out handle, out openError))
            {
                return "error: " + openError;
            }

            try
            {
                uint hcsparamsValue;
                string ioError;
                if (!TryReadPhys32(handle, capabilityAddress + hcsparamsOffset, out hcsparamsValue, out ioError))
                {
                    return "error: failed to read HCSPARAMS: " + ioError;
                }

                uint rtsoffValue;
                if (!TryReadPhys32(handle, capabilityAddress + rtsoff, out rtsoffValue, out ioError))
                {
                    return "error: failed to read RTSOFF: " + ioError;
                }

                uint maxIntrs = (hcsparamsValue >> 8) & 0x7FF;
                ulong runtimeAddress = capabilityAddress + rtsoffValue;
                uint writeCount = intervals != null && intervals.Length > 0
                    ? Math.Min(maxIntrs, (uint)intervals.Length)
                    : maxIntrs;
                uint failures = 0;

                for (uint i = 0; i < writeCount; i++)
                {
                    ulong interrupterAddress = runtimeAddress + 0x24 + (0x20 * i);
                    uint target = intervals != null && intervals.Length > 0 ? intervals[(int)i] : interval;
                    if (!TryWriteImodInterval(handle, interrupterAddress, target, out ioError))
                    {
                        failures++;
                    }
                }

                string mode = intervals != null && intervals.Length > 0 ? "vector=" + intervals.Length : "interval=0x" + interval.ToString("X");
                return "writes=" + writeCount + "/" + maxIntrs
                    + " hcsparams=0x" + hcsparamsValue.ToString("X")
                    + " rtsoff=0x" + rtsoffValue.ToString("X")
                    + " runtime=0x" + runtimeAddress.ToString("X")
                    + " " + mode
                    + " failures=" + failures;
            }
            finally
            {
                CloseHandle(handle);
            }
        }

        public static string TryBuildAdaptiveIntervals(
            ulong capabilityAddress,
            uint hcsparamsOffset,
            uint fallbackInterval,
            string roleIntervalsText,
            string rootPortRolesText,
            uint[] fallbackIntervals,
            out uint[] intervals)
        {
            intervals = new uint[0];

            Dictionary<string, uint> roleIntervals;
            string parseError;
            if (!TryParseRoleIntervals(roleIntervalsText, out roleIntervals, out parseError))
            {
                return "error: " + parseError;
            }

            Dictionary<uint, HashSet<string>> rolesByRootPort;
            if (!TryParseRootPortRoles(rootPortRolesText, out rolesByRootPort, out parseError))
            {
                return "error: " + parseError;
            }

            IntPtr handle;
            string openError;
            if (!TryOpenDriverDevice(OpenRetryCount, out handle, out openError))
            {
                return "error: " + openError;
            }

            try
            {
                uint maxIntrs;
                XhciInterrupterTopology topology;
                string topologyDetail;
                if (!TryReadXhciInterrupterTopology(handle, capabilityAddress, hcsparamsOffset, out maxIntrs, out topology, out topologyDetail))
                {
                    return "error: " + topologyDetail;
                }

                int count = (int)Math.Min(maxIntrs, 2048U);
                if (count <= 0)
                {
                    return "error: controller reports zero interrupters";
                }

                intervals = new uint[count];
                uint fallback = fallbackInterval & 0xFFFFU;
                for (int i = 0; i < intervals.Length; i++)
                {
                    intervals[i] = fallbackIntervals != null && i < fallbackIntervals.Length
                        ? fallbackIntervals[i] & 0xFFFFU
                        : fallback;
                }

                int matchedPorts = 0;
                int assignedIntrs = 0;
                List<string> shown = new List<string>();
                foreach (KeyValuePair<uint, HashSet<string>> pair in rolesByRootPort)
                {
                    List<uint> interrupters;
                    if (!topology.ByRootPort.TryGetValue(pair.Key, out interrupters) || interrupters.Count == 0)
                    {
                        continue;
                    }

                    string selectedRole;
                    uint selectedValue;
                    if (!TrySelectAdaptiveRoleInterval(pair.Value, roleIntervals, out selectedRole, out selectedValue))
                    {
                        continue;
                    }

                    matchedPorts++;
                    foreach (uint intr in interrupters)
                    {
                        if (intr >= intervals.Length)
                        {
                            continue;
                        }

                        intervals[(int)intr] = selectedValue & 0xFFFFU;
                        assignedIntrs++;
                        if (shown.Count < 8)
                        {
                            shown.Add("port" + pair.Key + "->I" + intr + "=" + selectedRole + ":0x" + selectedValue.ToString("X"));
                        }
                    }
                }

                if (assignedIntrs == 0)
                {
                    intervals = new uint[0];
                    return "error: no root-port roles matched active interrupters; roles=" + rootPortRolesText + "; " + topologyDetail;
                }

                string suffix = assignedIntrs > shown.Count ? " +" + (assignedIntrs - shown.Count) + " more" : string.Empty;
                return "ok: vector=" + intervals.Length
                    + " matchedPorts=" + matchedPorts
                    + " assignedIntrs=" + assignedIntrs
                    + " [" + string.Join(", ", shown.ToArray()) + suffix + "]; "
                    + topologyDetail;
            }
            finally
            {
                CloseHandle(handle);
            }
        }

        public static string ApplyNicItr(ulong baseAddress, uint baseOffset, uint stride, uint queues, uint width, ulong mask, ulong orBits, ulong[] values)
        {
            if (queues == 0)
            {
                return "error: queues=0";
            }
            if (width != 16 && width != 32)
            {
                return "error: unsupported width=" + width;
            }
            if (values == null || values.Length == 0)
            {
                return "error: no values";
            }

            IntPtr handle;
            string openError;
            if (!TryOpenDriverDevice(OpenRetryCount, out handle, out openError))
            {
                return "error: " + openError;
            }

            try
            {
                uint size = width == 16 ? 2U : 4U;
                uint failures = 0;
                StringBuilder applied = new StringBuilder();
                for (uint q = 0; q < queues; q++)
                {
                    ulong selected = q < values.Length ? values[q] : values[0];
                    ulong finalValue = (selected & mask) | orBits;
                    ulong address = baseAddress + baseOffset + (stride * q);
                    if (applied.Length > 0)
                    {
                        applied.Append(", ");
                    }
                    applied.Append("Q").Append(q).Append("@0x").Append(address.ToString("X")).Append("=0x").Append(finalValue.ToString("X"));
                    string ioError;
                    if (!TryWritePhysicalMemory(handle, address, size, finalValue, out ioError))
                    {
                        failures++;
                    }
                }

                return "writes=" + queues + " width=" + width + " values=[" + applied.ToString() + "] failures=" + failures;
            }
            finally
            {
                CloseHandle(handle);
            }
        }

        public static bool IsDisabledProblem(uint problemCode)
        {
            return problemCode == CmProbDisabled;
        }

        private static bool TryOpenDriverDevice(int attempts, out IntPtr handle, out string error)
        {
            handle = InvalidHandleValue;
            error = null;
            int lastError = 0;
            for (int i = 0; i < attempts; i++)
            {
                handle = CreateFile(ImodDriverDevicePath, GenericRead | GenericWrite, FileShareRead | FileShareWrite, IntPtr.Zero, OpenExisting, FileAttributeNormal, IntPtr.Zero);
                if (handle != InvalidHandleValue)
                {
                    return true;
                }
                lastError = Marshal.GetLastWin32Error();
                if (attempts > 1)
                {
                    System.Threading.Thread.Sleep(OpenRetryDelayMs);
                }
            }

            error = "failed to open " + ImodDriverDevicePath + ": " + GetWin32ErrorMessage(lastError);
            return false;
        }

        private static bool TryReadPhys32(IntPtr handle, ulong address, out uint value, out string error)
        {
            value = 0;
            ulong raw;
            if (!TryReadPhysicalMemory(handle, address, 4, out raw, out error))
            {
                return false;
            }
            value = unchecked((uint)raw);
            return true;
        }

        private static bool TryWriteImodInterval(IntPtr handle, ulong address, uint interval, out string error)
        {
            uint currentValue;
            if (!TryReadPhys32(handle, address, out currentValue, out error))
            {
                return false;
            }
            uint mergedValue = (currentValue & 0xFFFF0000U) | (interval & 0xFFFFU);
            return TryWritePhysicalMemory(handle, address, 4, mergedValue, out error);
        }

        private static bool TryReadPhysicalMemory(IntPtr handle, ulong address, uint size, out ulong value, out string error)
        {
            value = 0;
            error = null;
            PhysAccessStruct access = new PhysAccessStruct();
            access.physAddress = address;
            access.accessSizeInBytes = size;
            int bytesReturned;
            if (!DeviceIoControl(handle, IoctlImodReadPhysicalMemory, ref access, Marshal.SizeOf(typeof(PhysAccessStruct)), ref access, Marshal.SizeOf(typeof(PhysAccessStruct)), out bytesReturned, IntPtr.Zero))
            {
                error = GetWin32ErrorMessage(Marshal.GetLastWin32Error());
                return false;
            }
            value = access.value;
            return true;
        }

        private static bool TryWritePhysicalMemory(IntPtr handle, ulong address, uint size, ulong value, out string error)
        {
            error = null;
            PhysAccessStruct access = new PhysAccessStruct();
            access.physAddress = address;
            access.accessSizeInBytes = size;
            access.value = value;
            int bytesReturned;
            if (!DeviceIoControl(handle, IoctlImodWritePhysicalMemory, ref access, Marshal.SizeOf(typeof(PhysAccessStruct)), ref access, Marshal.SizeOf(typeof(PhysAccessStruct)), out bytesReturned, IntPtr.Zero))
            {
                error = GetWin32ErrorMessage(Marshal.GetLastWin32Error());
                return false;
            }
            return true;
        }

        private static bool TryParseRoleIntervals(string text, out Dictionary<string, uint> roleIntervals, out string error)
        {
            roleIntervals = new Dictionary<string, uint>(StringComparer.OrdinalIgnoreCase);
            error = null;
            if (string.IsNullOrWhiteSpace(text))
            {
                error = "ROLE_INTERVALS is empty";
                return false;
            }

            string[] parts = text.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string rawPart in parts)
            {
                string part = rawPart.Trim();
                int eq = part.IndexOf('=');
                if (eq <= 0 || eq >= part.Length - 1)
                {
                    continue;
                }

                string role = NormalizeAdaptiveRole(part.Substring(0, eq));
                if (role.Length == 0)
                {
                    continue;
                }

                uint value;
                if (!TryParseUInt32Flexible(part.Substring(eq + 1), out value))
                {
                    error = "invalid ROLE_INTERVALS value: " + part;
                    return false;
                }

                roleIntervals[role] = value & 0xFFFFU;
            }

            if (roleIntervals.Count == 0)
            {
                error = "ROLE_INTERVALS has no usable role values";
                return false;
            }

            return true;
        }

        private static bool TryParseRootPortRoles(string text, out Dictionary<uint, HashSet<string>> rolesByRootPort, out string error)
        {
            rolesByRootPort = new Dictionary<uint, HashSet<string>>();
            error = null;
            if (string.IsNullOrWhiteSpace(text))
            {
                error = "root-port role map is empty";
                return false;
            }

            string[] parts = text.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string rawPart in parts)
            {
                string part = rawPart.Trim();
                int eq = part.IndexOf('=');
                if (eq <= 0 || eq >= part.Length - 1)
                {
                    continue;
                }

                uint rootPort;
                if (!uint.TryParse(part.Substring(0, eq).Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out rootPort) || rootPort == 0)
                {
                    continue;
                }

                HashSet<string> roles;
                if (!rolesByRootPort.TryGetValue(rootPort, out roles))
                {
                    roles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    rolesByRootPort[rootPort] = roles;
                }

                string[] roleParts = part.Substring(eq + 1).Split(new[] { '+', ',', '|' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string rawRole in roleParts)
                {
                    string role = NormalizeAdaptiveRole(rawRole);
                    if (role.Length > 0)
                    {
                        roles.Add(role);
                    }
                }
            }

            if (rolesByRootPort.Count == 0)
            {
                error = "root-port role map has no usable entries";
                return false;
            }

            return true;
        }

        private static string NormalizeAdaptiveRole(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return string.Empty;
            }

            string value = text.Trim();
            if (value.IndexOf("Mouse", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "Mouse";
            }
            if (value.IndexOf("Keyboard", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "Keyboard";
            }
            if (value.IndexOf("Audio", StringComparison.OrdinalIgnoreCase) >= 0
                || value.IndexOf("Microphone", StringComparison.OrdinalIgnoreCase) >= 0
                || value.IndexOf("Speaker", StringComparison.OrdinalIgnoreCase) >= 0
                || value.IndexOf("Headphone", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "Audio";
            }
            if (value.IndexOf("Gamepad", StringComparison.OrdinalIgnoreCase) >= 0
                || value.IndexOf("Controller", StringComparison.OrdinalIgnoreCase) >= 0
                || value.IndexOf("Joystick", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "Gamepad";
            }
            if (value.IndexOf("Camera", StringComparison.OrdinalIgnoreCase) >= 0
                || value.IndexOf("Webcam", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "Webcam";
            }

            return value;
        }

        private static bool TryParseUInt32Flexible(string text, out uint value)
        {
            value = 0;
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            string trimmed = text.Trim();
            if (trimmed.StartsWith("0x", StringComparison.OrdinalIgnoreCase))
            {
                return uint.TryParse(trimmed.Substring(2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out value);
            }

            return uint.TryParse(trimmed, NumberStyles.Integer, CultureInfo.InvariantCulture, out value);
        }

        private static bool TrySelectAdaptiveRoleInterval(
            HashSet<string> roles,
            Dictionary<string, uint> roleIntervals,
            out string selectedRole,
            out uint selectedValue)
        {
            selectedRole = string.Empty;
            selectedValue = 0;
            for (int i = 0; i < AdaptiveRolePriority.Length; i++)
            {
                string role = AdaptiveRolePriority[i];
                if (roles.Contains(role) && roleIntervals.TryGetValue(role, out selectedValue))
                {
                    selectedRole = role;
                    return true;
                }
            }

            foreach (string role in roles)
            {
                if (roleIntervals.TryGetValue(role, out selectedValue))
                {
                    selectedRole = role;
                    return true;
                }
            }

            return false;
        }

        private static bool TryReadXhciInterrupterTopology(
            IntPtr handle,
            ulong capabilityAddress,
            uint hcsparamsOffset,
            out uint maxIntrs,
            out XhciInterrupterTopology topology,
            out string detail)
        {
            maxIntrs = 0;
            topology = new XhciInterrupterTopology();
            detail = null;

            uint capReg;
            string ioError;
            if (!TryReadPhys32(handle, capabilityAddress, out capReg, out ioError))
            {
                detail = "failed to read xHCI CAPLENGTH: " + ioError;
                return false;
            }

            uint capLength = capReg & 0xFFU;
            if (capLength == 0)
            {
                detail = "xHCI CAPLENGTH is zero";
                return false;
            }

            uint hcsparamsValue;
            if (!TryReadPhys32(handle, capabilityAddress + hcsparamsOffset, out hcsparamsValue, out ioError))
            {
                detail = "failed to read xHCI HCSPARAMS1: " + ioError;
                return false;
            }

            uint maxSlots = hcsparamsValue & 0xFFU;
            maxIntrs = (hcsparamsValue >> 8) & 0x7FFU;
            if (maxSlots == 0 || maxIntrs == 0)
            {
                detail = "xHCI reports maxSlots=" + maxSlots + " maxIntrs=" + maxIntrs;
                return false;
            }

            uint hccparamsValue;
            if (!TryReadPhys32(handle, capabilityAddress + 0x10, out hccparamsValue, out ioError))
            {
                detail = "failed to read xHCI HCCPARAMS1: " + ioError;
                return false;
            }

            uint contextSize = ((hccparamsValue >> 2) & 0x1U) != 0 ? 64U : 32U;
            ulong operationalAddress = capabilityAddress + capLength;
            ulong dcbaap;
            if (!TryReadPhys64(handle, operationalAddress + 0x30, out dcbaap, out ioError))
            {
                detail = "failed to read xHCI DCBAAP: " + ioError;
                return false;
            }

            dcbaap &= 0xFFFFFFFFFFFFFFC0UL;
            if (dcbaap == 0)
            {
                detail = "xHCI DCBAAP is zero";
                return false;
            }

            for (uint slot = 1; slot <= maxSlots; slot++)
            {
                ulong deviceContext;
                if (!TryReadPhys64(handle, dcbaap + ((ulong)slot * 8UL), out deviceContext, out ioError))
                {
                    continue;
                }

                deviceContext &= 0xFFFFFFFFFFFFFFC0UL;
                if (deviceContext == 0)
                {
                    continue;
                }

                uint slotDword0;
                uint slotDword1;
                uint slotDword2;
                uint slotDword3;
                if (!TryReadPhys32(handle, deviceContext, out slotDword0, out ioError)
                    || !TryReadPhys32(handle, deviceContext + 0x04, out slotDword1, out ioError)
                    || !TryReadPhys32(handle, deviceContext + 0x08, out slotDword2, out ioError)
                    || !TryReadPhys32(handle, deviceContext + 0x0C, out slotDword3, out ioError))
                {
                    continue;
                }

                bool isHub = ((slotDword0 >> 26) & 0x1U) != 0;
                uint contextEntries = (slotDword0 >> 27) & 0x1FU;
                uint slotState = (slotDword3 >> 27) & 0x1FU;
                uint rootPort = (slotDword1 >> 16) & 0xFFU;
                uint deviceAddress = slotDword3 & 0xFFU;
                uint interrupter = (slotDword2 >> 22) & 0x3FFU;

                if (slotState < 2 || isHub || rootPort == 0)
                {
                    continue;
                }

                topology.SlotTargetCount++;

                if (interrupter >= maxIntrs)
                {
                    continue;
                }

                AddUniqueInterrupter(topology.ByRootPort, rootPort, interrupter);
                if (deviceAddress > 0)
                {
                    AddUniqueInterrupter(topology.ByDeviceAddress, deviceAddress, interrupter);
                }
            }

            if (topology.ByRootPort.Count == 0 && topology.ByDeviceAddress.Count == 0)
            {
                detail = "no active xHCI device/interrupter topology was found";
                return false;
            }

            detail = "slots=" + maxSlots
                + " intrs=" + maxIntrs
                + " ctx=" + contextSize
                + " endpointTargets=" + topology.EndpointTargetCount
                + " slotTargets=" + topology.SlotTargetCount
                + " rootPortMap=[" + FormatInterrupterMap(topology.ByRootPort) + "]";
            return true;
        }

        private static bool TryReadEndpointInterrupterTarget(
            IntPtr handle,
            ulong deviceContext,
            uint contextSize,
            uint contextEntries,
            uint maxIntrs,
            out uint interrupter)
        {
            interrupter = 0;
            if (contextSize == 0 || contextEntries < 2)
            {
                return false;
            }

            for (uint contextIndex = 2; contextIndex <= contextEntries; contextIndex++)
            {
                ulong endpointContext = deviceContext + ((ulong)contextIndex * contextSize);
                uint epDword0;
                uint epDword1;
                string ioError;
                if (!TryReadPhys32(handle, endpointContext, out epDword0, out ioError)
                    || !TryReadPhys32(handle, endpointContext + 0x04, out epDword1, out ioError))
                {
                    continue;
                }

                uint endpointState = epDword0 & 0x7U;
                uint endpointType = (epDword1 >> 3) & 0x7U;
                if (endpointState == 0 || (endpointType != 3 && endpointType != 5 && endpointType != 7))
                {
                    continue;
                }

                ulong transferRing;
                if (!TryReadPhys64(handle, endpointContext + 0x08, out transferRing, out ioError))
                {
                    continue;
                }

                transferRing &= 0xFFFFFFFFFFFFFFF0UL;
                if (transferRing == 0)
                {
                    continue;
                }

                uint trbDword2;
                uint trbDword3;
                if (!TryReadPhys32(handle, transferRing + 0x08, out trbDword2, out ioError)
                    || !TryReadPhys32(handle, transferRing + 0x0C, out trbDword3, out ioError))
                {
                    continue;
                }

                uint trbType = (trbDword3 >> 10) & 0x3FU;
                if (trbType != 1 && trbType != 3 && trbType != 5)
                {
                    continue;
                }

                uint target = (trbDword2 >> 22) & 0x3FFU;
                if (target < maxIntrs)
                {
                    interrupter = target;
                    return true;
                }
            }

            return false;
        }

        private static bool TryReadPhys64(IntPtr handle, ulong address, out ulong value, out string error)
        {
            value = 0;
            uint low;
            if (!TryReadPhys32(handle, address, out low, out error))
            {
                return false;
            }

            uint high;
            if (!TryReadPhys32(handle, address + 4, out high, out error))
            {
                return false;
            }

            value = (ulong)low | ((ulong)high << 32);
            return true;
        }

        private static void AddUniqueInterrupter(Dictionary<uint, List<uint>> map, uint key, uint interrupter)
        {
            List<uint> interrupters;
            if (!map.TryGetValue(key, out interrupters))
            {
                interrupters = new List<uint>();
                map[key] = interrupters;
            }

            if (!interrupters.Contains(interrupter))
            {
                interrupters.Add(interrupter);
            }
        }

        private static string FormatInterrupterMap(Dictionary<uint, List<uint>> map)
        {
            List<string> parts = new List<string>();
            foreach (KeyValuePair<uint, List<uint>> pair in map)
            {
                if (parts.Count >= 16)
                {
                    break;
                }

                List<string> intrs = new List<string>();
                foreach (uint intr in pair.Value)
                {
                    intrs.Add("I" + intr);
                }

                parts.Add(pair.Key + ":" + string.Join("/", intrs.ToArray()));
            }

            if (map.Count > parts.Count)
            {
                parts.Add("+" + (map.Count - parts.Count) + " more");
            }

            return string.Join(", ", parts.ToArray());
        }

        private sealed class XhciInterrupterTopology
        {
            public readonly Dictionary<uint, List<uint>> ByRootPort = new Dictionary<uint, List<uint>>();
            public readonly Dictionary<uint, List<uint>> ByDeviceAddress = new Dictionary<uint, List<uint>>();
            public uint EndpointTargetCount;
            public uint SlotTargetCount;

            public XhciInterrupterTopology()
            {
                EndpointTargetCount = 0;
                SlotTargetCount = 0;
            }
        }

        private static readonly string[] AdaptiveRolePriority = new[] { "Mouse", "Keyboard", "Audio", "Gamepad", "Webcam" };

        private static bool IsXhciDevice(IntPtr devInfoSet, ref SP_DEVINFO_DATA devInfo)
        {
            string service;
            if (TryGetDeviceStringProperty(devInfoSet, ref devInfo, SpdrpService, out service)
                && string.Equals(service, "USBXHCI", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            List<string> ids;
            if (TryGetDeviceMultiSzProperty(devInfoSet, ref devInfo, SpdrpHardwareId, out ids) && HasXhciClassCode(ids))
            {
                return true;
            }
            if (TryGetDeviceMultiSzProperty(devInfoSet, ref devInfo, SpdrpCompatibleIds, out ids) && HasXhciClassCode(ids))
            {
                return true;
            }
            return false;
        }

        private static bool HasXhciClassCode(List<string> ids)
        {
            foreach (string id in ids)
            {
                if (id.IndexOf("CC_0C0330", StringComparison.OrdinalIgnoreCase) >= 0
                    || id.IndexOf("CLASS_0C0330", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }
            return false;
        }

        private static string GetDeviceCaption(IntPtr devInfoSet, ref SP_DEVINFO_DATA devInfo)
        {
            string caption;
            if (TryGetDeviceStringProperty(devInfoSet, ref devInfo, SpdrpFriendlyName, out caption))
            {
                return caption;
            }
            if (TryGetDeviceStringProperty(devInfoSet, ref devInfo, SpdrpDeviceDesc, out caption))
            {
                return caption;
            }
            return "Unknown USB Controller";
        }

        private static bool TryGetDeviceProblemCode(uint devInst, out uint problemCode)
        {
            uint status;
            uint problem;
            problemCode = 0;
            int cr = CM_Get_DevNode_Status(out status, out problem, devInst, 0);
            if (cr != CrSuccess)
            {
                return false;
            }
            problemCode = problem;
            return true;
        }

        private static bool TryGetDeviceMemoryBase(uint devInst, out ulong baseAddress, out string error)
        {
            baseAddress = 0;
            error = null;
            IntPtr logConf;
            int cr = CM_Get_First_Log_Conf(out logConf, devInst, AllocLogConf);
            if (cr != CrSuccess)
            {
                cr = CM_Get_First_Log_Conf(out logConf, devInst, BootLogConf);
            }
            if (cr != CrSuccess)
            {
                error = "failed to query logical config (CONFIGRET " + cr + ")";
                return false;
            }

            try
            {
                bool found = false;
                ulong minBase = 0;
                uint[] resTypes = new uint[] { ResTypeMem, ResTypeMemLarge };
                foreach (uint resType in resTypes)
                {
                    IntPtr resDes;
                    int resCr = CM_Get_Next_Res_Des(out resDes, logConf, resType, IntPtr.Zero, 0);
                    while (resCr == CrSuccess)
                    {
                        uint dataSize;
                        int sizeCr = CM_Get_Res_Des_Data_Size(out dataSize, resDes, 0);
                        if (sizeCr == CrSuccess && dataSize > 0)
                        {
                            byte[] buffer = new byte[dataSize];
                            if (CM_Get_Res_Des_Data(resDes, buffer, dataSize, 0) == CrSuccess)
                            {
                                ulong candidate;
                                if (TryExtractBaseFromResource(resType, buffer, out candidate))
                                {
                                    if (!found || candidate < minBase)
                                    {
                                        minBase = candidate;
                                        found = true;
                                    }
                                }
                            }
                        }

                        IntPtr nextResDes;
                        int nextCr = CM_Get_Next_Res_Des(out nextResDes, resDes, resType, IntPtr.Zero, 0);
                        CM_Free_Res_Des_Handle(resDes);
                        resDes = nextResDes;
                        resCr = nextCr;
                    }
                }

                if (!found)
                {
                    error = "no memory resource found";
                    return false;
                }

                baseAddress = minBase;
                return true;
            }
            finally
            {
                CM_Free_Log_Conf_Handle(logConf);
            }
        }

        private static bool TryExtractBaseFromResource(uint resType, byte[] data, out ulong baseAddress)
        {
            baseAddress = 0;
            if (resType == ResTypeMem)
            {
                if (data.Length < Marshal.SizeOf(typeof(MemDes)))
                {
                    return false;
                }
                MemDes mem = BytesToStruct<MemDes>(data, 0);
                ulong candidate = mem.MD_Alloc_Base;
                if (candidate == 0 && mem.MD_Count > 0)
                {
                    int offset = Marshal.SizeOf(typeof(MemDes));
                    if (data.Length >= offset + Marshal.SizeOf(typeof(MemRange)))
                    {
                        MemRange range = BytesToStruct<MemRange>(data, offset);
                        candidate = range.MR_Min;
                    }
                }
                if (candidate == 0)
                {
                    return false;
                }
                baseAddress = candidate;
                return true;
            }

            if (resType == ResTypeMemLarge)
            {
                if (data.Length < Marshal.SizeOf(typeof(MemLargeDes)))
                {
                    return false;
                }
                MemLargeDes mem = BytesToStruct<MemLargeDes>(data, 0);
                ulong candidate = mem.MLD_Alloc_Base;
                if (candidate == 0 && mem.MLD_Count > 0)
                {
                    int offset = Marshal.SizeOf(typeof(MemLargeDes));
                    if (data.Length >= offset + Marshal.SizeOf(typeof(MemLargeRange)))
                    {
                        MemLargeRange range = BytesToStruct<MemLargeRange>(data, offset);
                        candidate = range.MLR_Min;
                    }
                }
                if (candidate == 0)
                {
                    return false;
                }
                baseAddress = candidate;
                return true;
            }

            return false;
        }

        private static bool TryGetDeviceStringProperty(IntPtr devInfoSet, ref SP_DEVINFO_DATA devInfo, uint property, out string value)
        {
            value = string.Empty;
            List<string> values;
            if (!TryGetDeviceMultiSzProperty(devInfoSet, ref devInfo, property, out values) || values.Count == 0)
            {
                return false;
            }
            value = values[0];
            return true;
        }

        private static bool TryGetDeviceMultiSzProperty(IntPtr devInfoSet, ref SP_DEVINFO_DATA devInfo, uint property, out List<string> values)
        {
            values = new List<string>();
            byte[] data;
            uint regType;
            if (!TryGetDevicePropertyData(devInfoSet, ref devInfo, property, out data, out regType))
            {
                return false;
            }
            if (regType != RegMultiSz && regType != RegSz)
            {
                return false;
            }
            string text = Encoding.Unicode.GetString(data);
            string[] parts = text.Split(new char[] { '\0' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string part in parts)
            {
                string trimmed = part.Trim();
                if (trimmed.Length > 0)
                {
                    values.Add(trimmed);
                }
            }
            return values.Count > 0;
        }

        private static bool TryGetDevicePropertyData(IntPtr devInfoSet, ref SP_DEVINFO_DATA devInfo, uint property, out byte[] data, out uint regType)
        {
            data = null;
            regType = 0;
            uint requiredSize = 0;
            if (!SetupDiGetDeviceRegistryPropertyW(devInfoSet, ref devInfo, property, out regType, null, 0, out requiredSize))
            {
                int err = Marshal.GetLastWin32Error();
                if (err != ErrorInsufficientBuffer)
                {
                    return false;
                }
            }
            if (requiredSize == 0)
            {
                return false;
            }
            data = new byte[requiredSize];
            return SetupDiGetDeviceRegistryPropertyW(devInfoSet, ref devInfo, property, out regType, data, requiredSize, out requiredSize);
        }

        private static bool TryGetDeviceInstanceId(IntPtr devInfoSet, ref SP_DEVINFO_DATA devInfo, out string instanceId)
        {
            instanceId = string.Empty;
            int requiredSize;
            SetupDiGetDeviceInstanceIdW(devInfoSet, ref devInfo, null, 0, out requiredSize);
            int err = Marshal.GetLastWin32Error();
            if (err != ErrorInsufficientBuffer || requiredSize <= 0)
            {
                return false;
            }
            StringBuilder buffer = new StringBuilder(requiredSize);
            if (!SetupDiGetDeviceInstanceIdW(devInfoSet, ref devInfo, buffer, buffer.Capacity, out requiredSize))
            {
                return false;
            }
            instanceId = buffer.ToString();
            return instanceId.Length > 0;
        }

        private static T BytesToStruct<T>(byte[] data, int offset) where T : struct
        {
            GCHandle handle = GCHandle.Alloc(data, GCHandleType.Pinned);
            try
            {
                IntPtr ptr = IntPtr.Add(handle.AddrOfPinnedObject(), offset);
                return (T)Marshal.PtrToStructure(ptr, typeof(T));
            }
            finally
            {
                handle.Free();
            }
        }

        private static string GetWin32ErrorMessage(int error)
        {
            return new Win32Exception(error).Message;
        }

        private static uint CtlCode(uint deviceType, uint function, uint method, uint access)
        {
            return (deviceType << 16) | (access << 14) | (function << 2) | method;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct SP_DEVINFO_DATA
        {
            public uint cbSize;
            public Guid ClassGuid;
            public uint DevInst;
            public IntPtr Reserved;
        }

        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        private struct SERVICE_STATUS_PROCESS
        {
            public uint dwServiceType;
            public uint dwCurrentState;
            public uint dwControlsAccepted;
            public uint dwWin32ExitCode;
            public uint dwServiceSpecificExitCode;
            public uint dwCheckPoint;
            public uint dwWaitHint;
            public uint dwProcessId;
            public uint dwServiceFlags;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct SERVICE_STATUS
        {
            public uint dwServiceType;
            public uint dwCurrentState;
            public uint dwControlsAccepted;
            public uint dwWin32ExitCode;
            public uint dwServiceSpecificExitCode;
            public uint dwCheckPoint;
            public uint dwWaitHint;
        }

        private struct PhysAccessStruct
        {
            public ulong physAddress;
            public uint accessSizeInBytes;
            public uint reserved;
            public ulong value;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MemDes
        {
            public uint MD_Count;
            public uint MD_Type;
            public ulong MD_Alloc_Base;
            public ulong MD_Alloc_End;
            public uint MD_Flags;
            public uint MD_Reserved;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MemRange
        {
            public ulong MR_Align;
            public uint MR_nBytes;
            public ulong MR_Min;
            public ulong MR_Max;
            public uint MR_Flags;
            public uint MR_Reserved;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MemLargeDes
        {
            public uint MLD_Count;
            public uint MLD_Type;
            public ulong MLD_Alloc_Base;
            public ulong MLD_Alloc_End;
            public uint MLD_Flags;
            public uint MLD_Reserved;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MemLargeRange
        {
            public ulong MLR_Align;
            public ulong MLR_nBytes;
            public ulong MLR_Min;
            public ulong MLR_Max;
            public uint MLR_Flags;
            public uint MLR_Reserved;
        }
    }
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

            if (-not $proc.WaitForExit(60000)) {
                try { $proc.Kill() } catch {}
                Write-ImodLog "startup loader: KDU timeout"
                return $false
            }

            $stdout = $proc.StandardOutput.ReadToEnd()
            $stderr = $proc.StandardError.ReadToEnd()
            $combined = (($stdout + " " + $stderr) -replace '\s+', ' ').Trim()
            if ($combined.Length -gt 400) {
                $combined = $combined.Substring(0, 400)
            }

            Write-ImodLog ("startup loader: KDU exit=" + $proc.ExitCode + " output=" + $combined)
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
                Write-ImodLog ("usb controller: id=$($controller.DeviceId) caption=$(Format-ImodText $controller.Caption) problem=$($controller.ProblemCode) hasBase=$($controller.HasBase) base=$controllerBaseText") -verboseOnly
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

                $result = [DeviceTweakerImodRuntime]::ApplyController([uint64]$controller.BaseAddress, $hcsparamsOffset, $rtsoff, $interval, $intervalArray)
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

    """;

    private ImodConfig? _imodConfigCache;
    private string? _imodScriptPath;
    private bool _imodConfigLoaded;

    private sealed class ImodConfigEntry
    {
        public required string Hwid { get; set; }
        public uint? Interval { get; set; }
        public List<uint>? Intervals { get; set; }
        public bool? AdaptiveRoleBinding { get; set; }
        public Dictionary<string, uint>? RoleIntervals { get; set; }
        public uint? HcsparamsOffset { get; set; }
        public uint? Rtsoff { get; set; }
        public bool? Enabled { get; set; }
    }

    private sealed class NicItrConfigEntry
    {
        public required string Hwid { get; set; }
        public string FamilyName { get; set; } = string.Empty;
        public uint BaseOffset { get; set; }
        public uint Stride { get; set; }
        public int Queues { get; set; }
        public int Width { get; set; }
        public ulong Mask { get; set; }
        public ulong OrBits { get; set; }
        public List<ulong> Values { get; set; } = [];
    }

    private sealed class ImodConfig
    {
        public uint GlobalInterval { get; set; } = ImodDefaultInterval;
        public uint GlobalHcsparamsOffset { get; set; } = ImodDefaultHcsparamsOffset;
        public uint GlobalRtsoff { get; set; } = ImodDefaultRtsoff;
        public List<ImodConfigEntry> Overrides { get; } = [];
        public List<NicItrConfigEntry> NicItrEntries { get; } = [];
        public bool HasScript { get; set; }
    }

    private enum ImodApplyOutcome
    {
        Applied,
        SkippedNoUsb,
        SkippedNoController,
        SkippedNoConfig,
        Failed,
    }

    private sealed record ImodInput(uint? Interval, List<uint>? Intervals, Dictionary<string, uint>? RoleIntervals, bool IsDefault);

    private void InvalidateImodCache()
    {
        _imodConfigCache = null;
        _imodScriptPath = null;
        _imodConfigLoaded = false;
    }

    private void EnsureImodConfigLoaded()
    {
        if (_imodConfigLoaded)
        {
            return;
        }

        _imodConfigLoaded = true;
        _imodConfigCache = LoadImodConfig(out _imodScriptPath);
    }

    private ImodConfig LoadImodConfig(out string? scriptPath)
    {
        ResolveImodPaths(out scriptPath);
        if (string.IsNullOrWhiteSpace(scriptPath) || !File.Exists(scriptPath))
        {
            return new ImodConfig { HasScript = false };
        }

        try
        {
            ImodConfig config = ParseImodScriptFile(scriptPath);
            config.HasScript = true;
            return config;
        }
        catch (Exception ex)
        {
            WriteLog($"IMOD.CONFIG: failed to parse {scriptPath}: {ex.Message}");
            return new ImodConfig { HasScript = false };
        }
    }

    private string GetImodStartupPath()
    {
        string startup = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
        if (string.IsNullOrWhiteSpace(startup))
        {
            startup = GetScriptRoot();
        }

        return Path.Combine(startup, ImodScriptFileName);
    }

    private static string GetImodDriverSystemPath()
    {
        string windows = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
        if (string.IsNullOrWhiteSpace(windows))
        {
            windows = Path.GetPathRoot(Environment.SystemDirectory) ?? string.Empty;
            if (string.IsNullOrWhiteSpace(windows))
            {
                windows = AppContext.BaseDirectory;
                if (string.IsNullOrWhiteSpace(windows))
                {
                    windows = Environment.CurrentDirectory;
                }
            }
        }

        return Path.Combine(windows, ImodDriverName);
    }

    private static string GetImodStartupKduDirectory()
    {
        string root = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        if (string.IsNullOrWhiteSpace(root))
        {
            root = AppContext.BaseDirectory;
        }

        return Path.Combine(root, "DEVICE TWEAKER", "IMOD", "Loader");
    }

    private static bool EnsureImodStartupKduPayload(out string kduPath, out string dbPath, out string? error)
    {
        string directory = GetImodStartupKduDirectory();
        kduPath = Path.Combine(directory, ImodKduFileName);
        dbPath = Path.Combine(directory, ImodKduDatabaseFileName);

        if (!TrySecureImodDriverDirectory(directory, out error))
        {
            return false;
        }

        if (!TryWriteEmbeddedImodResource("DeviceTweakerCS.IMOD.Loader.kdu.exe", ".kdu.exe", kduPath, out error))
        {
            return false;
        }

        if (!TryWriteEmbeddedImodResource("DeviceTweakerCS.IMOD.Loader.drv64.dll", ".drv64.dll", dbPath, out error))
        {
            return false;
        }

        error = null;
        return true;
    }

    private static bool IsImodDriverSystemPath(string? path)
    {
        string systemPath = GetImodDriverSystemPath();
        if (string.IsNullOrWhiteSpace(path) || string.IsNullOrWhiteSpace(systemPath))
        {
            return false;
        }

        try
        {
            return string.Equals(Path.GetFullPath(path), Path.GetFullPath(systemPath), StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return string.Equals(path, systemPath, StringComparison.OrdinalIgnoreCase);
        }
    }

    private void ResolveImodPaths(out string? scriptPath)
    {
        string startupPath = GetImodStartupPath();
        scriptPath = File.Exists(startupPath) ? startupPath : null;
    }

    private void RemoveImodPersistenceFiles()
    {
        DeleteFileIfExists(GetImodStartupPath(), "IMOD.CONFIG");
        DeleteFileIfExists(Path.Combine(GetScriptRoot(), "dtimod.sys"), "IMOD.DRIVER.LEGACY");
        DeleteFileIfExists(Path.Combine(GetScriptRoot(), ImodDriverName), "IMOD.DRIVER.LEGACY");
        WriteLog($"IMOD.DRIVER: keep staged system driver {GetImodDriverSystemPath()}");
    }

    private void DeleteFileIfExists(string? path, string label)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return;
        }

        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
                WriteLog($"{label}: deleted {path}");
            }
        }
        catch (Exception ex)
        {
            WriteLog($"{label}: failed to delete {path}: {ex.Message}");
        }
    }

    private void DeleteDirectoryIfExists(string? path, string label)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return;
        }

        try
        {
            if (Directory.Exists(path))
            {
                Directory.Delete(path, recursive: true);
                WriteLog($"{label}: deleted {path}");
            }
        }
        catch (Exception ex)
        {
            WriteLog($"{label}: failed to delete {path}: {ex.Message}");
        }
    }

    private static ImodConfig ParseImodScriptFile(string path)
    {
        ImodConfig config = new();
        ImodConfigEntry? currentDevice = null;
        NicItrConfigEntry? currentNicItr = null;
        bool inOverrides = false;
        bool inNicItr = false;

        string[] lines = File.ReadAllLines(path, Encoding.UTF8);
        foreach (string raw in lines)
        {
            string line = StripInlineComment(raw).Trim();
            if (line.Length > 0 && line[0] == '\uFEFF')
            {
                line = line.TrimStart('\uFEFF');
            }
            if (line.Length == 0)
            {
                continue;
            }

            if (inNicItr)
            {
                if (currentNicItr is null)
                {
                    if (line.StartsWith(")", StringComparison.Ordinal))
                    {
                        inNicItr = false;
                        continue;
                    }

                    if (line.StartsWith("@{", StringComparison.Ordinal))
                    {
                        currentNicItr = new NicItrConfigEntry { Hwid = string.Empty };
                    }

                    continue;
                }

                if (line.StartsWith("}", StringComparison.Ordinal))
                {
                    if (!string.IsNullOrWhiteSpace(currentNicItr.Hwid)
                        && currentNicItr.Queues > 0
                        && (currentNicItr.Width == 16 || currentNicItr.Width == 32)
                        && currentNicItr.Values.Count > 0)
                    {
                        config.NicItrEntries.Add(currentNicItr);
                    }

                    currentNicItr = null;
                    continue;
                }

                if (!TryParseQuotedAssignment(line, out string nicKeyName, out string nicValueText))
                {
                    continue;
                }

                string nicKey = nicKeyName.Trim().ToUpperInvariant();
                if (nicKey == "HWID")
                {
                    currentNicItr.Hwid = UnquotePowerShellString(nicValueText);
                }
                else if (nicKey == "FAMILY")
                {
                    currentNicItr.FamilyName = UnquotePowerShellString(nicValueText);
                }
                else if (nicKey == "BASE_OFFSET" && TryParseUInt32Flexible(nicValueText, out uint baseOffset))
                {
                    currentNicItr.BaseOffset = baseOffset;
                }
                else if (nicKey == "STRIDE" && TryParseUInt32Flexible(nicValueText, out uint stride))
                {
                    currentNicItr.Stride = stride;
                }
                else if (nicKey == "QUEUES" && TryParseUInt32Flexible(nicValueText, out uint queues))
                {
                    currentNicItr.Queues = (int)Math.Min(queues, 1024);
                }
                else if (nicKey == "WIDTH" && TryParseUInt32Flexible(nicValueText, out uint width))
                {
                    currentNicItr.Width = (int)width;
                }
                else if (nicKey == "MASK" && TryParseUInt64Flexible(nicValueText, out ulong mask))
                {
                    currentNicItr.Mask = mask;
                }
                else if ((nicKey == "OR_BITS" || nicKey == "ORBITS") && TryParseUInt64Flexible(nicValueText, out ulong orBits))
                {
                    currentNicItr.OrBits = orBits;
                }
                else if (nicKey == "VALUES" && TryParseUInt64List(nicValueText, out List<ulong> nicValues))
                {
                    currentNicItr.Values = nicValues;
                }

                continue;
            }

            if (TryParseAssignment(line, "$globalInterval", out string valueText)
                && TryParseUInt32Flexible(valueText, out uint parsedGlobal))
            {
                config.GlobalInterval = parsedGlobal;
                continue;
            }

            if (TryParseAssignment(line, "$globalHCSPARAMSOffset", out valueText)
                && TryParseUInt32Flexible(valueText, out uint parsedHcsparams))
            {
                config.GlobalHcsparamsOffset = parsedHcsparams;
                continue;
            }

            if (TryParseAssignment(line, "$globalRTSOFF", out valueText)
                && TryParseUInt32Flexible(valueText, out uint parsedRtsoff))
            {
                config.GlobalRtsoff = parsedRtsoff;
                continue;
            }

            if (line.StartsWith("$nicItrData", StringComparison.OrdinalIgnoreCase))
            {
                inNicItr = true;
                currentNicItr = null;
                string compact = new(line.Where(ch => !char.IsWhiteSpace(ch)).ToArray());
                if (compact.Contains("@()"))
                {
                    inNicItr = false;
                }
                continue;
            }

            if (line.StartsWith("$userDefinedData", StringComparison.OrdinalIgnoreCase))
            {
                inOverrides = true;
                currentDevice = null;
                string compact = new(line.Where(ch => !char.IsWhiteSpace(ch)).ToArray());
                if (compact.Contains("@{}"))
                {
                    inOverrides = false;
                }
                continue;
            }

            if (!inOverrides)
            {
                continue;
            }

            if (currentDevice is null)
            {
                if (line.StartsWith("}", StringComparison.Ordinal))
                {
                    inOverrides = false;
                    continue;
                }

                if (TryParseQuotedKey(line, out string hwidKey))
                {
                    currentDevice = new ImodConfigEntry { Hwid = hwidKey };
                    config.Overrides.Add(currentDevice);
                }

                continue;
            }

            if (line.StartsWith("}", StringComparison.Ordinal))
            {
                currentDevice = null;
                continue;
            }

            if (!TryParseQuotedAssignment(line, out string keyName, out valueText))
            {
                continue;
            }

            string key = keyName.Trim().ToUpperInvariant();
            if (key == "INTERVALS")
            {
                if (TryParseImodIntervalList(valueText, out List<uint> parsedValues) && parsedValues.Count > 0)
                {
                    currentDevice.Intervals = parsedValues;
                }

                continue;
            }

            if (key == "ROLE_INTERVALS")
            {
                if (TryParseImodRoleIntervals(valueText, out Dictionary<string, uint> roleValues) && roleValues.Count > 0)
                {
                    currentDevice.RoleIntervals = roleValues;
                    currentDevice.AdaptiveRoleBinding = true;
                }

                continue;
            }

            if (key == "ADAPTIVE_ROLE_BINDING")
            {
                if (TryParseBoolFlexible(valueText, out bool adaptiveValue))
                {
                    currentDevice.AdaptiveRoleBinding = adaptiveValue;
                }

                continue;
            }

            if (!TryParseUInt32Flexible(valueText, out uint parsedValue))
            {
                if (key == "ENABLED" && TryParseBoolFlexible(valueText, out bool enabledValue))
                {
                    currentDevice.Enabled = enabledValue;
                }
                continue;
            }

            if (key == "INTERVAL")
            {
                currentDevice.Interval = parsedValue;
            }
            else if (key == "ENABLED")
            {
                currentDevice.Enabled = parsedValue != 0;
            }
            else if (key == "HCSPARAMS_OFFSET" || key == "HCSPARAPS_OFFSET")
            {
                currentDevice.HcsparamsOffset = parsedValue;
            }
            else if (key == "RTSOFF")
            {
                currentDevice.Rtsoff = parsedValue;
            }
        }

        if (config.GlobalInterval == 0)
        {
            config.GlobalInterval = ImodDefaultInterval;
        }

        return config;
    }

    private static bool TryParseAssignment(string line, string key, out string valueText)
    {
        valueText = string.Empty;
        if (!line.StartsWith(key, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        int eqPos = line.IndexOf('=');
        if (eqPos < 0)
        {
            return false;
        }

        valueText = line[(eqPos + 1)..].Trim();
        return valueText.Length > 0;
    }

    private static bool TryParseQuotedKey(string line, out string key)
    {
        key = string.Empty;
        int firstQuote = line.IndexOf('"');
        if (firstQuote < 0)
        {
            return false;
        }

        int secondQuote = line.IndexOf('"', firstQuote + 1);
        if (secondQuote <= firstQuote)
        {
            return false;
        }

        key = line.Substring(firstQuote + 1, secondQuote - firstQuote - 1).Trim();
        return !string.IsNullOrWhiteSpace(key);
    }

    private static bool TryParseQuotedAssignment(string line, out string key, out string valueText)
    {
        key = string.Empty;
        valueText = string.Empty;
        int firstQuote = line.IndexOf('"');
        if (firstQuote < 0)
        {
            return false;
        }

        int secondQuote = line.IndexOf('"', firstQuote + 1);
        if (secondQuote <= firstQuote)
        {
            return false;
        }

        key = line.Substring(firstQuote + 1, secondQuote - firstQuote - 1).Trim();
        int eqPos = line.IndexOf('=', secondQuote + 1);
        if (eqPos < 0)
        {
            return false;
        }

        valueText = line[(eqPos + 1)..].Trim();
        return !string.IsNullOrWhiteSpace(key) && valueText.Length > 0;
    }

    private static string StripInlineComment(string value)
    {
        int hashPos = value.IndexOf('#');
        int semiPos = value.IndexOf(';');
        int cut = -1;
        if (hashPos >= 0)
        {
            cut = hashPos;
        }
        if (semiPos >= 0)
        {
            cut = cut < 0 ? semiPos : Math.Min(cut, semiPos);
        }
        return cut >= 0 ? value[..cut] : value;
    }

    private static bool TryParseUInt32Flexible(string text, out uint value)
    {
        value = 0;
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        string trimmed = text.Trim();
        if (trimmed.StartsWith("0x", StringComparison.OrdinalIgnoreCase))
        {
            return uint.TryParse(trimmed[2..], NumberStyles.HexNumber, CultureInfo.InvariantCulture, out value);
        }

        return uint.TryParse(trimmed, NumberStyles.Integer, CultureInfo.InvariantCulture, out value);
    }

    private static bool TryParseImodInterval(string text, uint fallback, out uint value)
    {
        value = fallback;
        if (string.IsNullOrWhiteSpace(text))
        {
            return true;
        }

        string trimmed = text.Trim();
        if (trimmed.Equals("default", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return TryParseUInt32Flexible(trimmed, out value);
    }

    private static bool TryParseImodIntervalList(string text, out List<uint> values)
    {
        values = [];
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        string cleaned = text.Trim();
        if (cleaned.StartsWith("@(", StringComparison.Ordinal) && cleaned.EndsWith(")", StringComparison.Ordinal))
        {
            cleaned = cleaned[2..^1];
        }
        else if (cleaned.StartsWith("[", StringComparison.Ordinal) && cleaned.EndsWith("]", StringComparison.Ordinal))
        {
            cleaned = cleaned[1..^1];
        }

        string[] parts = cleaned.Split([',', ';', ' ', '\t'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length == 0)
        {
            return false;
        }

        foreach (string part in parts)
        {
            if (!TryParseUInt32Flexible(part, out uint parsed))
            {
                values.Clear();
                return false;
            }

            values.Add(parsed);
        }

        return values.Count > 0;
    }

    private static bool TryParseUInt64List(string text, out List<ulong> values)
    {
        values = [];
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        string cleaned = text.Trim();
        if (cleaned.StartsWith("@(", StringComparison.Ordinal) && cleaned.EndsWith(")", StringComparison.Ordinal))
        {
            cleaned = cleaned[2..^1];
        }
        else if (cleaned.StartsWith("[", StringComparison.Ordinal) && cleaned.EndsWith("]", StringComparison.Ordinal))
        {
            cleaned = cleaned[1..^1];
        }

        string[] parts = cleaned.Split([',', ';', ' ', '\t'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length == 0)
        {
            return false;
        }

        foreach (string part in parts)
        {
            if (!TryParseUInt64Flexible(part, out ulong parsed))
            {
                values.Clear();
                return false;
            }

            values.Add(parsed);
        }

        return values.Count > 0;
    }

    private static bool TryParseImodInput(string text, uint fallback, out ImodInput input)
    {
        input = new ImodInput(fallback, null, null, true);
        if (string.IsNullOrWhiteSpace(text))
        {
            return true;
        }

        string trimmed = text.Trim();
        if (trimmed.Equals("default", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        if (TryParseImodRoleIntervals(trimmed, out Dictionary<string, uint> roleValues) && roleValues.Count > 0)
        {
            input = new ImodInput(null, null, roleValues, false);
            return true;
        }

        if (TryParseImodIntervalList(trimmed, out List<uint> vector))
        {
            if (vector.Count > 1)
            {
                input = new ImodInput(null, vector, null, false);
            }
            else
            {
                uint single = vector[0];
                input = new ImodInput(single, null, null, single == fallback);
            }

            return true;
        }

        if (TryParseUInt32Flexible(trimmed, out uint interval))
        {
            input = new ImodInput(interval, null, null, interval == fallback);
            return true;
        }

        return false;
    }

    private static bool TryParseImodRoleIntervals(string text, out Dictionary<string, uint> roleValues)
    {
        roleValues = new Dictionary<string, uint>(StringComparer.OrdinalIgnoreCase);
        string trimmed = UnquotePowerShellString(text.Trim());
        if (trimmed.Length == 0 || (!trimmed.Contains('=') && !trimmed.Contains(':')))
        {
            return false;
        }

        string[] parts = trimmed.Split([',', ';'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        foreach (string part in parts)
        {
            int eq = part.IndexOf('=');
            int colon = part.IndexOf(':');
            int split = eq >= 0 && colon >= 0 ? Math.Min(eq, colon) : Math.Max(eq, colon);
            if (split <= 0 || split >= part.Length - 1)
            {
                roleValues.Clear();
                return false;
            }

            string rawRole = part[..split].Trim();
            string rawValue = part[(split + 1)..].Trim();
            if (!TryNormalizeImodRoleName(rawRole, out string role)
                || !TryParseImodInterval(rawValue, ImodDefaultInterval, out uint value))
            {
                roleValues.Clear();
                return false;
            }

            roleValues[role] = value & 0xFFFF;
        }

        return roleValues.Count > 0;
    }

    private static bool TryNormalizeImodRoleName(string text, out string role)
    {
        role = string.Empty;
        string compact = new string(text.Where(char.IsLetterOrDigit).ToArray()).ToUpperInvariant();
        role = compact switch
        {
            "MOUSE" => "Mouse",
            "KEYBOARD" => "Keyboard",
            "AUDIO" or "SPEAKER" or "SPEAKERS" or "MIC" or "MICROPHONE" => "Audio",
            "GAMEPAD" or "PAD" or "JOYSTICK" => "Gamepad",
            "WEBCAM" or "CAMERA" => "Webcam",
            _ => string.Empty,
        };

        return role.Length > 0;
    }

    private static string UnquotePowerShellString(string text)
    {
        string value = text.Trim();
        if (value.Length >= 2
            && ((value[0] == '"' && value[^1] == '"') || (value[0] == '\'' && value[^1] == '\'')))
        {
            value = value[1..^1];
        }

        return value.Replace("`\"", "\"", StringComparison.Ordinal)
            .Replace("`'", "'", StringComparison.Ordinal)
            .Trim();
    }

    private static string FormatImodRoleIntervals(IReadOnlyDictionary<string, uint> roleIntervals)
    {
        string[] preferredOrder = ["Mouse", "Keyboard", "Audio", "Gamepad", "Webcam"];
        List<string> parts = [];
        HashSet<string> emitted = new(StringComparer.OrdinalIgnoreCase);

        foreach (string role in preferredOrder)
        {
            if (roleIntervals.TryGetValue(role, out uint value))
            {
                parts.Add($"{role}={FormatImodValue(value)}");
                emitted.Add(role);
            }
        }

        foreach (KeyValuePair<string, uint> pair in roleIntervals.OrderBy(kvp => kvp.Key, StringComparer.OrdinalIgnoreCase))
        {
            if (emitted.Contains(pair.Key))
            {
                continue;
            }

            parts.Add($"{pair.Key}={FormatImodValue(pair.Value)}");
        }

        return string.Join(", ", parts);
    }

    private static bool TryParseBoolFlexible(string text, out bool value)
    {
        value = false;
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        string trimmed = text.Trim();
        if (trimmed.StartsWith("$", StringComparison.Ordinal))
        {
            trimmed = trimmed[1..].Trim();
        }

        if (trimmed.Equals("true", StringComparison.OrdinalIgnoreCase))
        {
            value = true;
            return true;
        }

        if (trimmed.Equals("false", StringComparison.OrdinalIgnoreCase))
        {
            value = false;
            return true;
        }

        if (TryParseUInt32Flexible(trimmed, out uint parsed))
        {
            value = parsed != 0;
            return true;
        }

        return false;
    }

    private static string FormatImodValue(uint value)
    {
        return $"0x{value:X}";
    }

    private static string FormatImodVector(IReadOnlyList<uint> values)
    {
        return string.Join(", ", values.Select(FormatImodValue));
    }

    private static string FormatNicItrConfigValue(ulong value)
    {
        return $"0x{value:X}";
    }

    private static string FormatNicItrConfigVector(IReadOnlyList<ulong> values)
    {
        return string.Join(", ", values.Select(FormatNicItrConfigValue));
    }

    private static string FormatPowerShellString(string value)
    {
        string escaped = value?.Replace("'", "''") ?? string.Empty;
        return $"'{escaped}'";
    }

    private static string FormatPowerShellBool(bool value)
    {
        return value ? "$true" : "$false";
    }

    private static string GetImodDriverSystemPathForScript()
    {
        return GetImodDriverSystemPath();
    }

    private static string GetImodOverrideKey(string instanceId)
    {
        if (string.IsNullOrWhiteSpace(instanceId))
        {
            return string.Empty;
        }

        int index = instanceId.IndexOf("DEV_", StringComparison.OrdinalIgnoreCase);
        if (index >= 0 && index + 8 <= instanceId.Length)
        {
            string candidate = instanceId.Substring(index, 8).ToUpperInvariant();
            if (candidate.Length == 8 && candidate[0..4] == "DEV_")
            {
                bool isHex = true;
                foreach (char ch in candidate.AsSpan(4, 4))
                {
                    if (!IsHexDigit(ch))
                    {
                        isHex = false;
                        break;
                    }
                }

                if (isHex)
                {
                    return candidate;
                }
            }
        }

        return instanceId;
    }

    private static bool IsHexDigit(char ch)
    {
        return (ch >= '0' && ch <= '9') || (ch >= 'A' && ch <= 'F') || (ch >= 'a' && ch <= 'f');
    }

    private uint GetEffectiveImodInterval(string instanceId, ImodConfig config)
    {
        uint interval = config.GlobalInterval;
        foreach (ImodConfigEntry entry in config.Overrides)
        {
            if (!string.IsNullOrWhiteSpace(entry.Hwid)
                && instanceId.Contains(entry.Hwid, StringComparison.OrdinalIgnoreCase)
                && (!entry.Enabled.HasValue || entry.Enabled.Value)
                && entry.Interval.HasValue)
            {
                interval = entry.Interval.Value;
            }
        }

        return interval;
    }

    private static ImodConfigEntry? FindImodOverride(string instanceId, ImodConfig config)
    {
        ImodConfigEntry? match = null;
        foreach (ImodConfigEntry entry in config.Overrides)
        {
            if (!string.IsNullOrWhiteSpace(entry.Hwid)
                && instanceId.Contains(entry.Hwid, StringComparison.OrdinalIgnoreCase))
            {
                match = entry;
            }
        }

        return match;
    }

    private void WriteImodScript(ImodConfig config, string path)
    {
        string configBlock = BuildImodConfigBlock(config);
        string scriptBody = ImodScriptTemplate.Replace(ImodScriptConfigToken, configBlock, StringComparison.Ordinal);

        if (File.Exists(path))
        {
            string existing = File.ReadAllText(path, Encoding.UTF8);
            if (existing.Contains(ImodScriptVersionMarker, StringComparison.Ordinal)
                && TryReplaceImodConfigBlock(existing, configBlock, out string updated))
            {
                scriptBody = updated;
            }
        }

        string? dir = Path.GetDirectoryName(path);
        if (!string.IsNullOrWhiteSpace(dir))
        {
            Directory.CreateDirectory(dir);
        }

        File.WriteAllText(path, scriptBody, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
    }

    private string BuildImodConfigBlock(ImodConfig config)
    {
        string kduPath = string.Empty;
        string kduDbPath = string.Empty;
        if (!EnsureImodStartupKduPayload(out kduPath, out kduDbPath, out string? kduError))
        {
            WriteLog($"IMOD.CONFIG.KDU: payload unavailable: {kduError}");
        }

        StringBuilder sb = new();
        sb.AppendLine(ImodScriptMarkerStart);
        sb.AppendLine($"$ImodDriverPath = {FormatPowerShellString(GetImodDriverSystemPathForScript())}");
        sb.AppendLine($"$ImodKduPath = {FormatPowerShellString(kduPath)}");
        sb.AppendLine($"$ImodKduDbPath = {FormatPowerShellString(kduDbPath)}");
        sb.AppendLine($"$ImodStartupLogEnabled = {FormatPowerShellBool(ImodStartupScriptLoggingEnabled)}");
        sb.AppendLine($"$ImodStartupVerboseLogEnabled = {FormatPowerShellBool(ImodStartupScriptVerboseLoggingEnabled)}");
        sb.AppendLine($"$globalInterval = {FormatImodValue(config.GlobalInterval)}");
        sb.AppendLine($"$globalHCSPARAMSOffset = {FormatImodValue(config.GlobalHcsparamsOffset)}");
        sb.AppendLine($"$globalRTSOFF = {FormatImodValue(config.GlobalRtsoff)}");
        sb.AppendLine($"$applyUsbImod = {FormatPowerShellBool(HasCustomUsbImod(config))}");
        sb.AppendLine("$userDefinedData = @{");

        foreach (ImodConfigEntry entry in config.Overrides.OrderBy(e => e.Hwid, StringComparer.OrdinalIgnoreCase))
        {
            if (string.IsNullOrWhiteSpace(entry.Hwid))
            {
                continue;
            }

            sb.AppendLine($"    \"{entry.Hwid}\" = @{{");
            if (entry.Enabled.HasValue)
            {
                string enabledText = entry.Enabled.Value ? "$true" : "$false";
                sb.AppendLine($"        \"ENABLED\" = {enabledText}");
            }
            if (entry.Interval.HasValue)
            {
                sb.AppendLine($"        \"INTERVAL\" = {FormatImodValue(entry.Interval.Value)}");
            }
            if (entry.Intervals is { Count: > 0 })
            {
                sb.AppendLine($"        \"INTERVALS\" = @({FormatImodVector(entry.Intervals)})");
            }
            if (entry.AdaptiveRoleBinding.HasValue)
            {
                string adaptiveText = entry.AdaptiveRoleBinding.Value ? "$true" : "$false";
                sb.AppendLine($"        \"ADAPTIVE_ROLE_BINDING\" = {adaptiveText}");
            }
            if (entry.RoleIntervals is { Count: > 0 })
            {
                sb.AppendLine($"        \"ROLE_INTERVALS\" = {FormatPowerShellString(FormatImodRoleIntervals(entry.RoleIntervals))}");
            }
            if (entry.HcsparamsOffset.HasValue)
            {
                sb.AppendLine($"        \"HCSPARAMS_OFFSET\" = {FormatImodValue(entry.HcsparamsOffset.Value)}");
            }
            if (entry.Rtsoff.HasValue)
            {
                sb.AppendLine($"        \"RTSOFF\" = {FormatImodValue(entry.Rtsoff.Value)}");
            }
            sb.AppendLine("    }");
        }

        sb.AppendLine("}");
        sb.AppendLine("$nicItrData = @(");
        foreach (NicItrConfigEntry entry in config.NicItrEntries.OrderBy(e => e.Hwid, StringComparer.OrdinalIgnoreCase))
        {
            if (string.IsNullOrWhiteSpace(entry.Hwid) || entry.Values.Count == 0)
            {
                continue;
            }

            sb.AppendLine("    @{");
            sb.AppendLine($"        \"HWID\" = {FormatPowerShellString(entry.Hwid)}");
            if (!string.IsNullOrWhiteSpace(entry.FamilyName))
            {
                sb.AppendLine($"        \"FAMILY\" = {FormatPowerShellString(entry.FamilyName)}");
            }
            sb.AppendLine($"        \"BASE_OFFSET\" = {FormatImodValue(entry.BaseOffset)}");
            sb.AppendLine($"        \"STRIDE\" = {FormatImodValue(entry.Stride)}");
            sb.AppendLine($"        \"QUEUES\" = {entry.Queues.ToString(CultureInfo.InvariantCulture)}");
            sb.AppendLine($"        \"WIDTH\" = {entry.Width.ToString(CultureInfo.InvariantCulture)}");
            sb.AppendLine($"        \"MASK\" = {FormatNicItrConfigValue(entry.Mask)}");
            sb.AppendLine($"        \"OR_BITS\" = {FormatNicItrConfigValue(entry.OrBits)}");
            sb.AppendLine($"        \"VALUES\" = @({FormatNicItrConfigVector(entry.Values)})");
            sb.AppendLine("    }");
        }
        sb.AppendLine(")");
        sb.AppendLine(ImodScriptMarkerEnd);
        sb.AppendLine();
        return sb.ToString();
    }

    private static bool TryReplaceImodConfigBlock(string existing, string configBlock, out string updated)
    {
        updated = string.Empty;
        int startIndex = existing.IndexOf(ImodScriptMarkerStart, StringComparison.Ordinal);
        if (startIndex < 0)
        {
            return false;
        }

        int endIndex = existing.IndexOf(ImodScriptMarkerEnd, startIndex, StringComparison.Ordinal);
        if (endIndex < 0)
        {
            return false;
        }

        int newlineIndex = existing.IndexOf('\n', endIndex);
        if (newlineIndex < 0)
        {
            newlineIndex = existing.Length;
        }
        else
        {
            newlineIndex += 1;
        }

        updated = existing.Substring(0, startIndex) + configBlock + existing[newlineIndex..];
        return true;
    }

    private static bool HasActiveImod(ImodConfig config)
    {
        if (config.HasScript)
        {
            return true;
        }

        if (config.NicItrEntries.Count > 0)
        {
            return true;
        }

        if (config.GlobalInterval != ImodDefaultInterval
            || config.GlobalHcsparamsOffset != ImodDefaultHcsparamsOffset
            || config.GlobalRtsoff != ImodDefaultRtsoff)
        {
            return true;
        }

        foreach (ImodConfigEntry entry in config.Overrides)
        {
            if (string.IsNullOrWhiteSpace(entry.Hwid))
            {
                continue;
            }

            if (entry.Enabled.HasValue && !entry.Enabled.Value)
            {
                continue;
            }

            if (entry.Interval.HasValue
                || entry.Intervals is { Count: > 0 }
                || entry.RoleIntervals is { Count: > 0 }
                || entry.HcsparamsOffset.HasValue
                || entry.Rtsoff.HasValue)
            {
                return true;
            }

            if (entry.Enabled.HasValue && entry.Enabled.Value)
            {
                return true;
            }
        }

        return false;
    }

    private static bool HasCustomImod(ImodConfig config)
    {
        return HasCustomUsbImod(config) || config.NicItrEntries.Count > 0;
    }

    private static bool HasCustomUsbImod(ImodConfig config)
    {
        if (config.GlobalInterval != ImodDefaultInterval
            || config.GlobalHcsparamsOffset != ImodDefaultHcsparamsOffset
            || config.GlobalRtsoff != ImodDefaultRtsoff)
        {
            return true;
        }

        foreach (ImodConfigEntry entry in config.Overrides)
        {
            if (string.IsNullOrWhiteSpace(entry.Hwid))
            {
                continue;
            }

            if (entry.Enabled.HasValue && !entry.Enabled.Value)
            {
                return true;
            }

            if (entry.Interval.HasValue && entry.Interval.Value != ImodDefaultInterval)
            {
                return true;
            }

            if (entry.Intervals is { Count: > 0 } intervals
                && intervals.Any(value => value != ImodDefaultInterval))
            {
                return true;
            }

            if (entry.RoleIntervals is { Count: > 0 } roleIntervals
                && roleIntervals.Values.Any(value => value != ImodDefaultInterval))
            {
                return true;
            }

            if (entry.HcsparamsOffset.HasValue && entry.HcsparamsOffset.Value != ImodDefaultHcsparamsOffset)
            {
                return true;
            }

            if (entry.Rtsoff.HasValue && entry.Rtsoff.Value != ImodDefaultRtsoff)
            {
                return true;
            }
        }

        return false;
    }

    private ImodApplyOutcome ApplyImodSettings(out string? note)
    {
        note = null;
        List<DeviceBlock> xhciBlocks = _blocks
            .Where(b => b.Kind == DeviceKind.USB && b.Device.UsbIsXhci && b.Device.UsbHasDevices && !b.Device.IsTestDevice)
            .ToList();
        if (xhciBlocks.Count == 0)
        {
            return ImodApplyOutcome.SkippedNoUsb;
        }

        List<DeviceBlock> targetBlocks = xhciBlocks.Where(b => IsUsbImodTarget(b.Device)).ToList();

        ResolveImodPaths(out string? scriptPath);
        if (string.IsNullOrWhiteSpace(scriptPath))
        {
            scriptPath = GetImodStartupPath();
        }

        bool scriptExists = File.Exists(scriptPath);
        ImodConfig config = scriptExists ? ParseImodScriptFile(scriptPath) : new ImodConfig();
        config.HasScript = scriptExists;
        List<string> invalidInputs = [];

        foreach (DeviceBlock block in targetBlocks)
        {
            string instanceId = block.Device.InstanceId;
            config.Overrides.RemoveAll(e =>
                string.Equals(e.Hwid, instanceId, StringComparison.OrdinalIgnoreCase)
                && e.Enabled == false
                && !e.Interval.HasValue
                && e.Intervals is not { Count: > 0 }
                && e.RoleIntervals is not { Count: > 0 }
                && !e.AdaptiveRoleBinding.HasValue
                && !e.HcsparamsOffset.HasValue
                && !e.Rtsoff.HasValue);

            string text = block.ImodBox.Text ?? string.Empty;
            text = text.Trim();
            string hwid = GetImodOverrideKey(block.Device.InstanceId);
            ImodConfigEntry? existing = null;
            foreach (ImodConfigEntry entry in config.Overrides)
            {
                if (string.Equals(entry.Hwid, hwid, StringComparison.OrdinalIgnoreCase))
                {
                    existing = entry;
                }
            }

            if (text.Length == 0)
            {
                config.Overrides.RemoveAll(e =>
                    string.Equals(e.Hwid, hwid, StringComparison.OrdinalIgnoreCase));
                block.ImodBox.Text = FormatImodValue(config.GlobalInterval);
                continue;
            }

            if (!TryParseImodInput(text, config.GlobalInterval, out ImodInput parsedInput))
            {
                string shortPnp = GetShortPnpId(block.Device.InstanceId);
                string label = string.IsNullOrWhiteSpace(shortPnp) ? block.Device.Name : $"{block.Device.Name} ({shortPnp})";
                invalidInputs.Add(label);
                parsedInput = new ImodInput(config.GlobalInterval, null, null, true);
            }

            if (parsedInput.IsDefault)
            {
                config.Overrides.RemoveAll(e =>
                    string.Equals(e.Hwid, hwid, StringComparison.OrdinalIgnoreCase));
                block.ImodBox.Text = FormatImodValue(config.GlobalInterval);
                continue;
            }

            block.ImodBox.Text = parsedInput.RoleIntervals is { Count: > 0 } roleIntervals
                ? FormatImodRoleIntervals(roleIntervals)
                : parsedInput.Intervals is { Count: > 0 } intervals
                    ? FormatImodVector(intervals)
                    : FormatImodValue(parsedInput.Interval ?? config.GlobalInterval);

            if (existing is not null)
            {
                config.Overrides.RemoveAll(e =>
                    !ReferenceEquals(e, existing)
                    && string.Equals(e.Hwid, hwid, StringComparison.OrdinalIgnoreCase));

                existing.Enabled = true;
                existing.Interval = parsedInput.Interval;
                existing.Intervals = parsedInput.Intervals;
                existing.RoleIntervals = parsedInput.RoleIntervals;
                existing.AdaptiveRoleBinding = parsedInput.RoleIntervals is { Count: > 0 };
            }
            else
            {
                config.Overrides.Add(new ImodConfigEntry
                {
                    Hwid = hwid,
                    Enabled = true,
                    Interval = parsedInput.Interval,
                    Intervals = parsedInput.Intervals,
                    RoleIntervals = parsedInput.RoleIntervals,
                    AdaptiveRoleBinding = parsedInput.RoleIntervals is { Count: > 0 },
                });
            }
        }

        foreach (DeviceBlock block in xhciBlocks.Where(b => !IsUsbImodTarget(b.Device)))
        {
            string hwid = block.Device.InstanceId;
            ImodConfigEntry? existing = null;
            foreach (ImodConfigEntry entry in config.Overrides)
            {
                if (string.Equals(entry.Hwid, hwid, StringComparison.OrdinalIgnoreCase))
                {
                    existing = entry;
                }
            }

            if (existing is null)
            {
                existing = new ImodConfigEntry { Hwid = hwid };
                config.Overrides.Add(existing);
            }
            else
            {
                config.Overrides.RemoveAll(e =>
                    !ReferenceEquals(e, existing)
                    && string.Equals(e.Hwid, hwid, StringComparison.OrdinalIgnoreCase));
            }

            existing.Enabled = false;
            existing.Interval = null;
            existing.Intervals = null;
            existing.AdaptiveRoleBinding = null;
            existing.RoleIntervals = null;
            existing.HcsparamsOffset = null;
            existing.Rtsoff = null;
            WriteLog($"IMOD.CONFIG: skip non-hid {block.Device.InstanceId} roles=\"{block.Device.UsbRoles}\"");
        }

        if (invalidInputs.Count > 0)
        {
            string shown = string.Join(", ", invalidInputs.Take(3));
            string suffix = invalidInputs.Count > 3 ? " ..." : string.Empty;
            ShowThemedInfo(
                $"Invalid IMOD interval value detected for: {shown}{suffix}\nValues have been reset to default ({FormatImodValue(config.GlobalInterval)}).");
        }

        bool hasCustomUsb = HasCustomUsbImod(config);
        bool hasCustom = hasCustomUsb || config.NicItrEntries.Count > 0;
        bool shouldApplyUsbLive = hasCustomUsb || !hasCustom;
        string startupPath = GetImodStartupPath();
        ImodApplyStats stats = new();
        if (shouldApplyUsbLive && !TryApplyImod(config, hasCustom, out stats, out string? applyError))
        {
            note = applyError is null ? "IMOD failed." : $"IMOD failed: {applyError}";
            WriteLog($"IMOD: {note}");
            return ImodApplyOutcome.Failed;
        }
        if (!shouldApplyUsbLive)
        {
            WriteLog("IMOD: USB live apply skipped (no custom USB IMOD config; NIC ITR startup config preserved)");
        }

        if (hasCustom)
        {
            try
            {
                WriteImodScript(config, startupPath);
                config.HasScript = true;
                WriteLog($"IMOD.CONFIG: standalone startup script saved {startupPath}");
            }
            catch (Exception ex)
            {
                note = $"IMOD applied, but failed to write startup script: {ex.Message}";
                WriteLog($"IMOD.CONFIG: write failed {startupPath}: {ex.Message}");
                return ImodApplyOutcome.Failed;
            }
        }
        else
        {
            RemoveImodPersistenceFiles();
            config.HasScript = false;
        }

        _imodConfigCache = config;
        _imodScriptPath = hasCustom ? startupPath : null;
        _imodConfigLoaded = true;

        if (!shouldApplyUsbLive)
        {
            note = "IMOD startup script saved. USB live apply skipped (no custom USB IMOD config).";
            WriteLog($"IMOD: {note}");
            return ImodApplyOutcome.Applied;
        }

        if (stats.ControllersFound == 0)
        {
            note = "IMOD skipped (no XHCI controllers found).";
            WriteLog($"IMOD: {note}");
            return ImodApplyOutcome.SkippedNoController;
        }

        if (stats.ControllersApplied == 0)
        {
            note = "IMOD skipped (no eligible USB controllers).";
            WriteLog($"IMOD: {note}");
            return ImodApplyOutcome.SkippedNoController;
        }

        if (stats.ReadFailures > 0 || stats.WriteFailures > 0)
        {
            note = $"IMOD applied to {stats.ControllersApplied} USB controller(s) with {stats.ReadFailures} read failure(s) and {stats.WriteFailures} write failure(s).";
        }
        else
        {
            note = $"IMOD applied to {stats.ControllersApplied} USB controller(s).";
        }

        return ImodApplyOutcome.Applied;
    }
}



