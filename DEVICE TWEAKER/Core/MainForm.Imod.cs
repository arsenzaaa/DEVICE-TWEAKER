using System.Globalization;
using System.Text;

namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private const uint ImodDefaultInterval = 0xC8;
    private const uint ImodDefaultHcsparamsOffset = 0x4;
    private const uint ImodDefaultRtsoff = 0x18;
    private const string ImodScriptFileName = "ApplyIMOD.ps1";
    private const string WinIoDriverName = "winio.sys";
    private const string ImodScriptMarkerStart = "$imodSettingsBegin = $true";
    private const string ImodScriptMarkerEnd = "$imodSettingsEnd = $true";
    private const string ImodScriptVersionMarker = "$imodScriptVersion = 11";
    private const string ImodScriptConfigToken = "{{IMOD_CONFIG_BLOCK}}";
    private static readonly string ImodScriptTemplate = """
    param(
        [switch]$verbose
    )
    
    $imodScriptVersion = 11
    
    {{IMOD_CONFIG_BLOCK}}
    
    function Resolve-WinioPath {
        param(
            [string]$preferred,
            [string]$scriptRoot
        )
    
        $candidates = @()
        if ($preferred) {
            $candidates += $preferred
        }
        if ($env:windir) {
            $candidates += (Join-Path $env:windir 'winio.sys')
        }
        if ($scriptRoot) {
            $candidates += (Join-Path $scriptRoot 'winio.sys')
        $candidates += (Join-Path (Join-Path $scriptRoot 'IMOD') 'winio.sys')
            $parent = Split-Path -Parent $scriptRoot
            if ($parent) {
                $candidates += (Join-Path $parent 'winio.sys')
                $candidates += (Join-Path (Join-Path $parent 'IMOD') 'winio.sys')
                $grand = Split-Path -Parent $parent
                if ($grand) {
                    $candidates += (Join-Path $grand 'winio.sys')
                    $candidates += (Join-Path (Join-Path $grand 'IMOD') 'winio.sys')
                }
            }
        }
    
        foreach ($candidate in $candidates) {
            if ($candidate -and (Test-Path $candidate -PathType Leaf)) {
                return $candidate
            }
        }
    
        return $null
    }
    
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
    
    $scriptRoot = $PSCommandPath
    if ($scriptRoot) {
        $scriptRoot = Split-Path -Parent $scriptRoot
    }
    
    $resolvedWinio = Resolve-WinioPath -preferred $winioPath -scriptRoot $scriptRoot
    if (-not $resolvedWinio) {
        Write-Host "error: winio.sys not found"
        exit 1
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
    
    $imodSource = @'
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Globalization;
    using System.IO;
    using System.Runtime.InteropServices;
    using System.Security.Principal;
    using System.Text;
    using System.Threading;
    
    namespace ImodScript
    {
        public static class ImodEngine
        {
            private const uint ImodDefaultInterval = 0xC8;
    
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
    
            private const int CrSuccess = 0x00000000;
            private const int ErrorInsufficientBuffer = 122;
            private const int ErrorNoMoreItems = 259;
    
            private const int ErrorServiceDoesNotExist = 1060;
            private const int ErrorServiceAlreadyRunning = 1056;
    
            private const uint ScManagerAllAccess = 0x000F003F;
            private const uint ServiceAllAccess = 0x000F01FF;
            private const uint ServiceKernelDriver = 0x00000001;
            private const uint ServiceDemandStart = 0x00000003;
            private const uint ServiceErrorNormal = 0x00000001;
            private const uint ServiceControlStop = 0x00000001;
            private const uint ServiceRunning = 0x00000004;
            private const uint ServiceStopped = 0x00000001;
            private const int ScStatusProcessInfo = 0;
    
            private const uint GenericRead = 0x80000000;
            private const uint GenericWrite = 0x40000000;
            private const uint FileShareRead = 0x00000001;
            private const uint FileShareWrite = 0x00000002;
            private const uint OpenExisting = 3;
            private const uint FileAttributeNormal = 0x00000080;
    
            private const uint FileDeviceWinIo = 0x00008010;
            private const uint WinIoIoctlIndex = 0x810;
            private const uint MethodBuffered = 0;
            private const uint FileAnyAccess = 0;
    
            private const string WinIoDevicePath = "\\\\.\\WINIO";
            private const string WinIoServiceName = "WINIO";
    
            private static readonly uint IoctlWinioMapPhysToLin =
                CtlCode(FileDeviceWinIo, WinIoIoctlIndex, MethodBuffered, FileAnyAccess);
            private static readonly uint IoctlWinioUnmapPhysAddr =
                CtlCode(FileDeviceWinIo, WinIoIoctlIndex + 1, MethodBuffered, FileAnyAccess);
            private static readonly uint IoctlWinioEnableDirectIo =
                CtlCode(FileDeviceWinIo, WinIoIoctlIndex + 2, MethodBuffered, FileAnyAccess);
            private static readonly uint IoctlWinioDisableDirectIo =
                CtlCode(FileDeviceWinIo, WinIoIoctlIndex + 3, MethodBuffered, FileAnyAccess);
    
            private static readonly IntPtr InvalidHandleValue = new IntPtr(-1);
    
            private class ImodOverride
            {
                public string Hwid;
                public uint? Interval;
                public uint? HcsparamsOffset;
                public uint? Rtsoff;
                public bool? Enabled;
            }
    
            private class ImodConfig
            {
                public uint GlobalInterval;
                public uint GlobalHcsparamsOffset;
                public uint GlobalRtsoff;
                public List<ImodOverride> Overrides;
            }
    
            private class ImodControllerInfo
            {
                public string DeviceId;
                public string Caption;
                public uint ProblemCode;
                public ulong BaseAddress;
                public bool HasBase;
                public string BaseError;
            }
    
            private class ImodApplyStats
            {
                public int ControllersFound;
                public int ControllersApplied;
                public int WriteFailures;
                public int SkippedDisabled;
                public int MissingBase;
                public int ReadFailures;
            }
    
            public static int Apply(uint globalInterval, uint globalHcsparamsOffset, uint globalRtsoff, IDictionary overrides, string driverPath, bool verbose)
            {
                if (!IsAdministrator())
                {
                    Console.WriteLine("error: administrator privileges required");
                    return 1;
                }
    
                if (string.IsNullOrWhiteSpace(driverPath) || !File.Exists(driverPath))
                {
                    Console.WriteLine("error: winio.sys not found");
                    return 1;
                }
    
                ImodConfig config = new ImodConfig();
                config.GlobalInterval = globalInterval == 0 ? ImodDefaultInterval : globalInterval;
                config.GlobalHcsparamsOffset = globalHcsparamsOffset;
                config.GlobalRtsoff = globalRtsoff;
                config.Overrides = ParseOverrides(overrides);
    
                ImodApplyStats stats;
                string error;
                if (!TryApplyImod(config, driverPath, verbose, out stats, out error))
                {
                    Console.WriteLine("error: " + error);
                    return 1;
                }
    
                if (verbose)
                {
                    Console.WriteLine("imod: controllers=" + stats.ControllersFound);
                    Console.WriteLine("imod: applied=" + stats.ControllersApplied + " read_failures=" + stats.ReadFailures +
                                      " write_failures=" + stats.WriteFailures + " skipped_disabled=" + stats.SkippedDisabled +
                                      " missing_base=" + stats.MissingBase);
                }
    
                return 0;
            }
    
            private static List<ImodOverride> ParseOverrides(IDictionary overrides)
            {
                List<ImodOverride> list = new List<ImodOverride>();
                if (overrides == null)
                {
                    return list;
                }
    
                foreach (DictionaryEntry entry in overrides)
                {
                    string hwid = entry.Key == null ? string.Empty : entry.Key.ToString();
                    if (string.IsNullOrWhiteSpace(hwid))
                    {
                        continue;
                    }
    
                    ImodOverride item = new ImodOverride();
                    item.Hwid = hwid;
    
                    IDictionary values = entry.Value as IDictionary;
                    if (values != null)
                    {
                        foreach (DictionaryEntry kv in values)
                        {
                            string key = kv.Key == null ? string.Empty : kv.Key.ToString();
                            if (string.IsNullOrWhiteSpace(key))
                            {
                                continue;
                            }
    
                            string upper = key.Trim().ToUpperInvariant();
                            if (upper == "ENABLED")
                            {
                                bool enabled;
                                if (TryConvertBool(kv.Value, out enabled))
                                {
                                    item.Enabled = enabled;
                                }
                            }
                            else if (upper == "INTERVAL")
                            {
                                uint parsed;
                                if (TryConvertUInt32(kv.Value, out parsed))
                                {
                                    item.Interval = parsed;
                                }
                            }
                            else if (upper == "HCSPARAMS_OFFSET" || upper == "HCSPARAPS_OFFSET")
                            {
                                uint parsed;
                                if (TryConvertUInt32(kv.Value, out parsed))
                                {
                                    item.HcsparamsOffset = parsed;
                                }
                            }
                            else if (upper == "RTSOFF")
                            {
                                uint parsed;
                                if (TryConvertUInt32(kv.Value, out parsed))
                                {
                                    item.Rtsoff = parsed;
                                }
                            }
                        }
                    }
    
                    list.Add(item);
                }
    
                return list;
            }
    
            private static bool TryConvertUInt32(object value, out uint result)
            {
                result = 0;
                if (value == null)
                {
                    return false;
                }
    
                if (value is uint)
                {
                    result = (uint)value;
                    return true;
                }
                if (value is int)
                {
                    int i = (int)value;
                    if (i < 0)
                    {
                        return false;
                    }
                    result = (uint)i;
                    return true;
                }
                if (value is long)
                {
                    long i = (long)value;
                    if (i < 0 || i > uint.MaxValue)
                    {
                        return false;
                    }
                    result = (uint)i;
                    return true;
                }
                if (value is ulong)
                {
                    ulong i = (ulong)value;
                    if (i > uint.MaxValue)
                    {
                        return false;
                    }
                    result = (uint)i;
                    return true;
                }
                if (value is short)
                {
                    short i = (short)value;
                    if (i < 0)
                    {
                        return false;
                    }
                    result = (uint)i;
                    return true;
                }
                if (value is byte)
                {
                    result = (byte)value;
                    return true;
                }
                if (value is bool)
                {
                    result = ((bool)value) ? 1u : 0u;
                    return true;
                }
    
                string text = value as string;
                if (text == null)
                {
                    text = value.ToString();
                }
                if (text == null)
                {
                    return false;
                }
    
                text = text.Trim();
                if (text.Length == 0)
                {
                    return false;
                }
    
                if (text.StartsWith("0x", StringComparison.OrdinalIgnoreCase))
                {
                    return uint.TryParse(text.Substring(2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out result);
                }
    
                return uint.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out result);
            }
    
            private static bool TryConvertBool(object value, out bool result)
            {
                result = false;
                if (value == null)
                {
                    return false;
                }
    
                if (value is bool)
                {
                    result = (bool)value;
                    return true;
                }
                if (value is int)
                {
                    result = (int)value != 0;
                    return true;
                }
                if (value is uint)
                {
                    result = (uint)value != 0;
                    return true;
                }
                if (value is long)
                {
                    result = (long)value != 0;
                    return true;
                }
                if (value is ulong)
                {
                    result = (ulong)value != 0;
                    return true;
                }
                if (value is short)
                {
                    result = (short)value != 0;
                    return true;
                }
                if (value is byte)
                {
                    result = (byte)value != 0;
                    return true;
                }
    
                string text = value as string;
                if (text == null)
                {
                    text = value.ToString();
                }
                if (text == null)
                {
                    return false;
                }
    
                text = text.Trim();
                if (string.Equals(text, "true", StringComparison.OrdinalIgnoreCase))
                {
                    result = true;
                    return true;
                }
                if (string.Equals(text, "false", StringComparison.OrdinalIgnoreCase))
                {
                    result = false;
                    return true;
                }
    
                uint parsed;
                if (TryConvertUInt32(text, out parsed))
                {
                    result = parsed != 0;
                    return true;
                }
    
                return false;
            }
    
            private static bool TryApplyImod(ImodConfig config, string driverPath, bool verbose, out ImodApplyStats stats, out string error)
            {
                stats = new ImodApplyStats();
                error = null;
    
                List<ImodControllerInfo> controllers;
                if (!TryEnumerateXhciControllers(out controllers, out error))
                {
                    return false;
                }
    
                stats.ControllersFound = controllers.Count;
                if (controllers.Count == 0)
                {
                    return true;
                }
    
                WinIoContext ctx;
                if (!WinIoContext.TryInitialize(driverPath, out ctx, out error))
                {
                    return false;
                }
    
                using (ctx)
                {
                    if (verbose)
                    {
                        Console.WriteLine("imod: controllers=" + controllers.Count);
                    }
    
                    foreach (ImodControllerInfo controller in controllers)
                    {
                        if (controller.ProblemCode == CmProbDisabled)
                        {
                            stats.SkippedDisabled++;
                            if (verbose)
                            {
                                Console.WriteLine("imod: skipped disabled " + controller.DeviceId);
                            }
                            continue;
                        }
    
                        if (!controller.HasBase)
                        {
                            stats.MissingBase++;
                            if (verbose && !string.IsNullOrEmpty(controller.BaseError))
                            {
                                Console.WriteLine("imod: " + controller.DeviceId + " base error: " + controller.BaseError);
                            }
                            continue;
                        }
    
                        uint desiredInterval = config.GlobalInterval;
                        uint hcsparamsOffset = config.GlobalHcsparamsOffset;
                        uint rtsoff = config.GlobalRtsoff;
                        bool enabled = true;
                        string overrideMatch = null;
    
                        foreach (ImodOverride entry in config.Overrides)
                        {
                            if (entry == null || string.IsNullOrWhiteSpace(entry.Hwid))
                            {
                                continue;
                            }
    
                            if (controller.DeviceId.IndexOf(entry.Hwid, StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                if (entry.Enabled.HasValue)
                                {
                                    enabled = entry.Enabled.Value;
                                }
                                if (entry.Interval.HasValue)
                                {
                                    desiredInterval = entry.Interval.Value;
                                }
                                if (entry.HcsparamsOffset.HasValue)
                                {
                                    hcsparamsOffset = entry.HcsparamsOffset.Value;
                                }
                                if (entry.Rtsoff.HasValue)
                                {
                                    rtsoff = entry.Rtsoff.Value;
                                }
                                overrideMatch = entry.Hwid;
                            }
                        }
    
                        if (!enabled)
                        {
                            stats.SkippedDisabled++;
                            if (verbose)
                            {
                                if (!string.IsNullOrEmpty(overrideMatch))
                                {
                                    Console.WriteLine("imod: skipped config-disabled " + controller.DeviceId + " (" + overrideMatch + ")");
                                }
                                else
                                {
                                    Console.WriteLine("imod: skipped config-disabled " + controller.DeviceId);
                                }
                            }
                            continue;
                        }
    
                        ulong capabilityAddress = controller.BaseAddress;
    
                        uint hcsparamsValue;
                        string ioError;
                        if (!TryReadPhys32(ctx, capabilityAddress + hcsparamsOffset, out hcsparamsValue, out ioError))
                        {
                            stats.ReadFailures++;
                            if (verbose)
                            {
                                Console.WriteLine("imod: read HCSPARAMS failed " + controller.DeviceId + ": " + ioError);
                            }
                            continue;
                        }
    
                        uint rtsoffValue;
                        if (!TryReadPhys32(ctx, capabilityAddress + rtsoff, out rtsoffValue, out ioError))
                        {
                            stats.ReadFailures++;
                            if (verbose)
                            {
                                Console.WriteLine("imod: read RTSOFF failed " + controller.DeviceId + ": " + ioError);
                            }
                            continue;
                        }
    
                        uint maxIntrs = (hcsparamsValue >> 8) & 0xFF;
                        ulong runtimeAddress = capabilityAddress + rtsoffValue;
    
                        uint writeFailures = 0;
                        for (uint i = 0; i < maxIntrs; ++i)
                        {
                            ulong interrupterAddress = runtimeAddress + 0x24 + (0x20 * i);
                            if (!TryWritePhys32(ctx, interrupterAddress, desiredInterval, out ioError))
                            {
                                writeFailures++;
                                if (verbose)
                                {
                                    Console.WriteLine("imod: write failed " + controller.DeviceId + " @ " + ToHex(interrupterAddress) + ": " + ioError);
                                }
                            }
                        }
    
                        stats.ControllersApplied++;
                        stats.WriteFailures += (int)writeFailures;
    
                        if (verbose)
                        {
                            Console.WriteLine("imod: " + controller.DeviceId + " writes=" + maxIntrs + " failures=" + writeFailures);
                        }
                    }
                }
    
                return true;
            }
    
            private static bool IsAdministrator()
            {
                try
                {
                    WindowsIdentity identity = WindowsIdentity.GetCurrent();
                    WindowsPrincipal principal = new WindowsPrincipal(identity);
                    return principal.IsInRole(WindowsBuiltInRole.Administrator);
                }
                catch
                {
                    return false;
                }
            }
    
            private static bool TryEnumerateXhciControllers(out List<ImodControllerInfo> controllers, out string error)
            {
                controllers = new List<ImodControllerInfo>();
                error = null;
    
                IntPtr devInfoSet = SetupDiGetClassDevsW(IntPtr.Zero, "PCI", IntPtr.Zero, DigcfPresent | DigcfAllClasses);
                if (devInfoSet == InvalidHandleValue)
                {
                    error = "failed to enumerate PCI devices: " + GetWin32ErrorMessage(Marshal.GetLastWin32Error());
                    return false;
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
    
                            error = "failed to enumerate device info: " + GetWin32ErrorMessage(lastError);
                            return false;
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
    
                        string caption = GetDeviceCaption(devInfoSet, ref devInfo);
                        uint problemCode;
                        TryGetDeviceProblemCode(devInfo.DevInst, out problemCode);
    
                        ulong baseAddress = 0;
                        string baseError;
                        bool hasBase = TryGetDeviceMemoryBase(devInfo.DevInst, out baseAddress, out baseError);
    
                        ImodControllerInfo info = new ImodControllerInfo();
                        info.DeviceId = instanceId;
                        info.Caption = caption;
                        info.ProblemCode = problemCode;
                        info.BaseAddress = baseAddress;
                        info.HasBase = hasBase;
                        info.BaseError = baseError ?? string.Empty;
                        controllers.Add(info);
                    }
                }
                finally
                {
                    SetupDiDestroyDeviceInfoList(devInfoSet);
                }
    
                return true;
            }
    
            private static bool IsXhciDevice(IntPtr devInfoSet, ref SP_DEVINFO_DATA devInfo)
            {
                string service;
                if (TryGetDeviceStringProperty(devInfoSet, ref devInfo, SpdrpService, out service))
                {
                    if (string.Equals(service, "USBXHCI", StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
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
    
            private static bool HasXhciClassCode(IEnumerable<string> ids)
            {
                foreach (string id in ids)
                {
                    if (id.IndexOf("CC_0C0330", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        id.IndexOf("CLASS_0C0330", StringComparison.OrdinalIgnoreCase) >= 0)
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
                problemCode = 0;
                uint status;
                int cr = CM_Get_DevNode_Status(out status, out problemCode, devInst, 0);
                return cr == CrSuccess;
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
                    ulong minBase = 0;
                    bool found = false;
                    uint[] resTypes = new uint[] { ResTypeMem, ResTypeMemLarge };
    
                    for (int resIndex = 0; resIndex < resTypes.Length; resIndex++)
                    {
                        uint resType = resTypes[resIndex];
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
                    int memSize = Marshal.SizeOf(typeof(MemDes));
                    if (data.Length < memSize)
                    {
                        return false;
                    }
    
                    MemDes mem = ReadStruct<MemDes>(data, 0);
                    ulong candidate = mem.MD_Alloc_Base;
                    if (candidate == 0 && mem.MD_Count > 0)
                    {
                        int offset = memSize;
                        int rangeSize = Marshal.SizeOf(typeof(MemRange));
                        if (data.Length >= offset + rangeSize)
                        {
                            MemRange range = ReadStruct<MemRange>(data, offset);
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
                    int memSize = Marshal.SizeOf(typeof(MemLargeDes));
                    if (data.Length < memSize)
                    {
                        return false;
                    }
    
                    MemLargeDes mem = ReadStruct<MemLargeDes>(data, 0);
                    ulong candidate = mem.MLD_Alloc_Base;
                    if (candidate == 0 && mem.MLD_Count > 0)
                    {
                        int offset = memSize;
                        int rangeSize = Marshal.SizeOf(typeof(MemLargeRange));
                        if (data.Length >= offset + rangeSize)
                        {
                            MemLargeRange range = ReadStruct<MemLargeRange>(data, offset);
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
    
            private static T ReadStruct<T>(byte[] data, int offset) where T : struct
            {
                int size = Marshal.SizeOf(typeof(T));
                if (data == null || data.Length < offset + size)
                {
                    return default(T);
                }
    
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
    
            private static bool TryGetDeviceStringProperty(IntPtr devInfoSet, ref SP_DEVINFO_DATA devInfo, uint property, out string value)
            {
                value = string.Empty;
                List<string> values;
                if (!TryGetDeviceMultiSzProperty(devInfoSet, ref devInfo, property, out values))
                {
                    return false;
                }
    
                if (values.Count == 0)
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
                for (int i = 0; i < parts.Length; i++)
                {
                    string trimmed = parts[i].Trim();
                    if (trimmed.Length > 0)
                    {
                        values.Add(trimmed);
                    }
                }
    
                return values.Count > 0;
            }
    
            private static bool TryGetDevicePropertyData(IntPtr devInfoSet, ref SP_DEVINFO_DATA devInfo, uint property, out byte[] data, out uint regType)
            {
                data = new byte[0];
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
                if (!SetupDiGetDeviceRegistryPropertyW(devInfoSet, ref devInfo, property, out regType, data, requiredSize, out requiredSize))
                {
                    return false;
                }
    
                return true;
            }
    
            private static bool TryGetDeviceInstanceId(IntPtr devInfoSet, ref SP_DEVINFO_DATA devInfo, out string instanceId)
            {
                instanceId = string.Empty;
    
                int requiredSize = 0;
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
                return !string.IsNullOrWhiteSpace(instanceId);
            }
    
            private static bool TryReadPhys32(WinIoContext ctx, ulong address, out uint value, out string error)
            {
                value = 0;
                error = null;
    
                PhysStruct phys;
                if (!TryMapPhysicalMemory(ctx, address, 4, out phys, out error))
                {
                    return false;
                }
    
                bool success = false;
                try
                {
                    int raw = Marshal.ReadInt32(new IntPtr(unchecked((long)phys.physMemLin)));
                    value = unchecked((uint)raw);
                    success = true;
                }
                finally
                {
                    string unmapError;
                    if (!TryUnmapPhysicalMemory(ctx, phys, out unmapError))
                    {
                        error = unmapError;
                        success = false;
                    }
                }
    
                return success;
            }
    
            private static bool TryWritePhys32(WinIoContext ctx, ulong address, uint value, out string error)
            {
                error = null;
    
                PhysStruct phys;
                if (!TryMapPhysicalMemory(ctx, address, 4, out phys, out error))
                {
                    return false;
                }
    
                bool success = false;
                try
                {
                    Marshal.WriteInt32(new IntPtr(unchecked((long)phys.physMemLin)), unchecked((int)value));
                    success = true;
                }
                finally
                {
                    string unmapError;
                    if (!TryUnmapPhysicalMemory(ctx, phys, out unmapError))
                    {
                        error = unmapError;
                        success = false;
                    }
                }
    
                return success;
            }
    
            private static bool TryMapPhysicalMemory(WinIoContext ctx, ulong address, ulong size, out PhysStruct phys, out string error)
            {
                error = null;
                phys = new PhysStruct();
                phys.physMemSizeInBytes = size;
                phys.physAddress = address;
    
                int bytesReturned = 0;
                if (!DeviceIoControl(
                    ctx.DriverHandle,
                    IoctlWinioMapPhysToLin,
                    ref phys,
                    Marshal.SizeOf(typeof(PhysStruct)),
                    ref phys,
                    Marshal.SizeOf(typeof(PhysStruct)),
                    out bytesReturned,
                    IntPtr.Zero))
                {
                    error = "failed to map physical memory: " + GetWin32ErrorMessage(Marshal.GetLastWin32Error());
                    return false;
                }
    
                if (phys.physMemLin == 0)
                {
                    error = "failed to map physical memory: returned null linear address";
                    return false;
                }
    
                return true;
            }
    
            private static bool TryUnmapPhysicalMemory(WinIoContext ctx, PhysStruct phys, out string error)
            {
                error = null;
    
                int bytesReturned = 0;
                if (!DeviceIoControl(
                    ctx.DriverHandle,
                    IoctlWinioUnmapPhysAddr,
                    ref phys,
                    Marshal.SizeOf(typeof(PhysStruct)),
                    ref phys,
                    Marshal.SizeOf(typeof(PhysStruct)),
                    out bytesReturned,
                    IntPtr.Zero))
                {
                    error = "failed to unmap physical memory: " + GetWin32ErrorMessage(Marshal.GetLastWin32Error());
                    return false;
                }
    
                return true;
            }
    
            private sealed class WinIoContext : IDisposable
            {
                public IntPtr DriverHandle = InvalidHandleValue;
                public bool Is64BitOS;
                public bool ServiceCreated;
                public string DriverPath;
    
                private WinIoContext(string driverPath)
                {
                    DriverPath = driverPath;
                }
    
                public static bool TryInitialize(string driverPath, out WinIoContext ctx, out string error)
                {
                    ctx = new WinIoContext(driverPath);
                    if (!ctx.Initialize(out error))
                    {
                        ctx.Dispose();
                        ctx = null;
                        return false;
                    }
    
                    return true;
                }
    
                private bool Initialize(out string error)
                {
                    error = null;
                    Is64BitOS = Environment.Is64BitOperatingSystem;
    
                    if (!EnsureWinIoService(out error))
                    {
                        return false;
                    }
    
                    DriverHandle = CreateFile(
                        WinIoDevicePath,
                        GenericRead | GenericWrite,
                        FileShareRead | FileShareWrite,
                        IntPtr.Zero,
                        OpenExisting,
                        FileAttributeNormal,
                        IntPtr.Zero);
    
                    if (DriverHandle == InvalidHandleValue)
                    {
                        error = "failed to open " + WinIoDevicePath + ": " + GetWin32ErrorMessage(Marshal.GetLastWin32Error());
                        return false;
                    }
    
                    if (!EnableDirectIo(out error))
                    {
                        return false;
                    }
    
                    return true;
                }
    
                private bool EnsureWinIoService(out string error)
                {
                    error = null;
    
                    IntPtr scm = OpenSCManager(null, null, ScManagerAllAccess);
                    if (scm == IntPtr.Zero)
                    {
                        error = "failed to open service manager: " + GetWin32ErrorMessage(Marshal.GetLastWin32Error());
                        return false;
                    }
    
                    try
                    {
                        IntPtr service = OpenService(scm, WinIoServiceName, ServiceAllAccess);
                        if (service == IntPtr.Zero)
                        {
                            int lastError = Marshal.GetLastWin32Error();
                            if (lastError != ErrorServiceDoesNotExist)
                            {
                                error = "failed to open WINIO service: " + GetWin32ErrorMessage(lastError);
                                return false;
                            }
    
                            service = CreateService(
                                scm,
                                WinIoServiceName,
                                WinIoServiceName,
                                ServiceAllAccess,
                                ServiceKernelDriver,
                                ServiceDemandStart,
                                ServiceErrorNormal,
                                DriverPath,
                                null,
                                IntPtr.Zero,
                                null,
                                null,
                                null);
                            if (service == IntPtr.Zero)
                            {
                                error = "failed to create WINIO service: " + GetWin32ErrorMessage(Marshal.GetLastWin32Error());
                                return false;
                            }
                            ServiceCreated = true;
                        }
    
                        try
                        {
                            bool wasRunning = false;
                            SERVICE_STATUS_PROCESS status;
                            if (QueryServiceStatus(service, out status))
                            {
                                wasRunning = status.dwCurrentState == ServiceRunning;
                            }
    
                            if (!wasRunning)
                            {
                                if (!StartService(service, 0, IntPtr.Zero))
                                {
                                    int lastError = Marshal.GetLastWin32Error();
                                    if (lastError != ErrorServiceAlreadyRunning)
                                    {
                                        error = "failed to start WINIO service: " + GetWin32ErrorMessage(lastError);
                                        return false;
                                    }
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
    
                    return true;
                }
    
                private bool EnableDirectIo(out string error)
                {
                    error = null;
                    if (Is64BitOS)
                    {
                        return true;
                    }
    
                    int bytesReturned = 0;
                    if (!DeviceIoControl(
                        DriverHandle,
                        IoctlWinioEnableDirectIo,
                        IntPtr.Zero,
                        0,
                        IntPtr.Zero,
                        0,
                        out bytesReturned,
                        IntPtr.Zero))
                    {
                        error = "failed to enable direct I/O: " + GetWin32ErrorMessage(Marshal.GetLastWin32Error());
                        return false;
                    }
    
                    return true;
                }
    
                private void DisableDirectIo()
                {
                    if (Is64BitOS || DriverHandle == InvalidHandleValue)
                    {
                        return;
                    }
    
                    int bytesReturned = 0;
                    DeviceIoControl(
                        DriverHandle,
                        IoctlWinioDisableDirectIo,
                        IntPtr.Zero,
                        0,
                        IntPtr.Zero,
                        0,
                        out bytesReturned,
                        IntPtr.Zero);
                }
    
                private void StopServiceIfNeeded()
                {
                    IntPtr scm = OpenSCManager(null, null, ScManagerAllAccess);
                    if (scm == IntPtr.Zero)
                    {
                        return;
                    }
    
                    try
                    {
                        IntPtr service = OpenService(scm, WinIoServiceName, ServiceAllAccess);
                        if (service == IntPtr.Zero)
                        {
                            return;
                        }
    
                        try
                        {
                            SERVICE_STATUS_PROCESS status;
                            if (QueryServiceStatus(service, out status) && status.dwCurrentState != ServiceStopped)
                            {
                                SERVICE_STATUS serviceStatus = new SERVICE_STATUS();
                                ControlService(service, ServiceControlStop, ref serviceStatus);
    
                                for (int i = 0; i < 25; i++)
                                {
                                    SERVICE_STATUS_PROCESS check;
                                    if (!QueryServiceStatus(service, out check) || check.dwCurrentState == ServiceStopped)
                                    {
                                        break;
                                    }
    
                                    Thread.Sleep(200);
                                }
                            }
    
                            if (ServiceCreated)
                            {
                                DeleteService(service);
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
    
                public void Dispose()
                {
                    DisableDirectIo();
    
                    if (DriverHandle != InvalidHandleValue)
                    {
                        CloseHandle(DriverHandle);
                        DriverHandle = InvalidHandleValue;
                    }
    
                    StopServiceIfNeeded();
                }
            }
    
            private static bool QueryServiceStatus(IntPtr service, out SERVICE_STATUS_PROCESS status)
            {
                status = new SERVICE_STATUS_PROCESS();
                uint bytesNeeded = 0;
                return QueryServiceStatusEx(service, ScStatusProcessInfo, ref status, (uint)Marshal.SizeOf(typeof(SERVICE_STATUS_PROCESS)), out bytesNeeded);
            }
    
            private static string ToHex(ulong value)
            {
                return "0x" + value.ToString("X");
            }
    
            private static string GetWin32ErrorMessage(int error)
            {
                return new Win32Exception(error).Message;
            }
    
            private static uint CtlCode(uint deviceType, uint function, uint method, uint access)
            {
                return (deviceType << 16) | (access << 14) | (function << 2) | method;
            }
    
            [DllImport("setupapi.dll", CharSet = CharSet.Unicode, SetLastError = true)]
            private static extern IntPtr SetupDiGetClassDevsW(
                IntPtr classGuid,
                string enumerator,
                IntPtr hwndParent,
                uint flags);
    
            [DllImport("setupapi.dll", SetLastError = true)]
            private static extern bool SetupDiEnumDeviceInfo(
                IntPtr deviceInfoSet,
                uint memberIndex,
                ref SP_DEVINFO_DATA deviceInfoData);
    
            [DllImport("setupapi.dll", CharSet = CharSet.Unicode, SetLastError = true)]
            private static extern bool SetupDiGetDeviceRegistryPropertyW(
                IntPtr deviceInfoSet,
                ref SP_DEVINFO_DATA deviceInfoData,
                uint property,
                out uint propertyRegDataType,
                [Out] byte[] propertyBuffer,
                uint propertyBufferSize,
                out uint requiredSize);
    
            [DllImport("setupapi.dll", CharSet = CharSet.Unicode, SetLastError = true)]
            private static extern bool SetupDiGetDeviceInstanceIdW(
                IntPtr deviceInfoSet,
                ref SP_DEVINFO_DATA deviceInfoData,
                StringBuilder deviceInstanceId,
                int deviceInstanceIdSize,
                out int requiredSize);
    
            [DllImport("setupapi.dll", SetLastError = true)]
            private static extern bool SetupDiDestroyDeviceInfoList(IntPtr deviceInfoSet);
    
            [DllImport("cfgmgr32.dll", SetLastError = true)]
            private static extern int CM_Get_DevNode_Status(out uint status, out uint problem, uint devInst, uint flags);
    
            [DllImport("cfgmgr32.dll", SetLastError = true)]
            private static extern int CM_Get_First_Log_Conf(out IntPtr logConf, uint devInst, uint flags);
    
            [DllImport("cfgmgr32.dll", SetLastError = true)]
            private static extern int CM_Get_Next_Res_Des(
                out IntPtr resDes,
                IntPtr logConfOrResDes,
                uint forResource,
                IntPtr resourceId,
                uint flags);
    
            [DllImport("cfgmgr32.dll", SetLastError = true)]
            private static extern int CM_Get_Res_Des_Data_Size(out uint dataSize, IntPtr resDes, uint flags);
    
            [DllImport("cfgmgr32.dll", SetLastError = true)]
            private static extern int CM_Get_Res_Des_Data(IntPtr resDes, [Out] byte[] buffer, uint bufferLen, uint flags);
    
            [DllImport("cfgmgr32.dll", SetLastError = true)]
            private static extern int CM_Free_Res_Des_Handle(IntPtr resDes);
    
            [DllImport("cfgmgr32.dll", SetLastError = true)]
            private static extern int CM_Free_Log_Conf_Handle(IntPtr logConf);
    
            [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
            private static extern IntPtr OpenSCManager(
                string machineName,
                string databaseName,
                uint desiredAccess);
    
            [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
            private static extern IntPtr OpenService(
                IntPtr scm,
                string serviceName,
                uint desiredAccess);
    
            [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
            private static extern IntPtr CreateService(
                IntPtr scm,
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
    
            [DllImport("advapi32.dll", SetLastError = true)]
            private static extern bool StartService(IntPtr service, uint numServiceArgs, IntPtr serviceArgVectors);
    
            [DllImport("advapi32.dll", SetLastError = true)]
            private static extern bool QueryServiceStatusEx(
                IntPtr service,
                int infoLevel,
                ref SERVICE_STATUS_PROCESS buffer,
                uint bufferSize,
                out uint bytesNeeded);
    
            [DllImport("advapi32.dll", SetLastError = true)]
            private static extern bool ControlService(IntPtr service, uint control, ref SERVICE_STATUS serviceStatus);
    
            [DllImport("advapi32.dll", SetLastError = true)]
            private static extern bool DeleteService(IntPtr service);
    
            [DllImport("advapi32.dll", SetLastError = true)]
            private static extern bool CloseServiceHandle(IntPtr handle);
    
            [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
            private static extern IntPtr CreateFile(
                string fileName,
                uint desiredAccess,
                uint shareMode,
                IntPtr securityAttributes,
                uint creationDisposition,
                uint flagsAndAttributes,
                IntPtr templateFile);
    
            [DllImport("kernel32.dll", SetLastError = true)]
            private static extern bool DeviceIoControl(
                IntPtr deviceHandle,
                uint ioControlCode,
                ref PhysStruct inBuffer,
                int inBufferSize,
                ref PhysStruct outBuffer,
                int outBufferSize,
                out int bytesReturned,
                IntPtr overlapped);
    
            [DllImport("kernel32.dll", SetLastError = true)]
            private static extern bool DeviceIoControl(
                IntPtr deviceHandle,
                uint ioControlCode,
                IntPtr inBuffer,
                int inBufferSize,
                IntPtr outBuffer,
                int outBufferSize,
                out int bytesReturned,
                IntPtr overlapped);
    
            [DllImport("kernel32.dll", SetLastError = true)]
            private static extern bool CloseHandle(IntPtr handle);
    
            [StructLayout(LayoutKind.Sequential)]
            private struct SP_DEVINFO_DATA
            {
                public uint cbSize;
                public Guid ClassGuid;
                public uint DevInst;
                public IntPtr Reserved;
            }
    
            [StructLayout(LayoutKind.Sequential, Pack = 1)]
            private struct PhysStruct
            {
                public ulong physMemSizeInBytes;
                public ulong physAddress;
                public ulong physicalMemoryHandle;
                public ulong physMemLin;
                public ulong physSection;
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
    
            [StructLayout(LayoutKind.Sequential)]
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
        }
    }
    '@
    
    if (-not ('ImodScript.ImodEngine' -as [type])) {
        try {
            Add-Type -TypeDefinition $imodSource -Language CSharp -ErrorAction Stop
        } catch {
            Write-Host "error: failed to compile IMOD engine: $($_.Exception.Message)"
            exit 1
        }
    }
    
    $exitCode = [ImodScript.ImodEngine]::Apply([uint32]$globalInterval, [uint32]$globalHCSPARAMSOffset, [uint32]$globalRTSOFF, $userDefinedData, $resolvedWinio, [bool]$verbose)
    exit $exitCode
    
    """;

    private ImodConfig? _imodConfigCache;
    private string? _imodScriptPath;
    private bool _imodConfigLoaded;

    private sealed class ImodConfigEntry
    {
        public required string Hwid { get; set; }
        public uint? Interval { get; set; }
        public uint? HcsparamsOffset { get; set; }
        public uint? Rtsoff { get; set; }
        public bool? Enabled { get; set; }
    }

    private sealed class ImodConfig
    {
        public uint GlobalInterval { get; set; } = ImodDefaultInterval;
        public uint GlobalHcsparamsOffset { get; set; } = ImodDefaultHcsparamsOffset;
        public uint GlobalRtsoff { get; set; } = ImodDefaultRtsoff;
        public List<ImodConfigEntry> Overrides { get; } = [];
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

    private string GetWinIoSystemPath()
    {
        string windows = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
        if (string.IsNullOrWhiteSpace(windows))
        {
            windows = GetScriptRoot();
        }

        return Path.Combine(windows, WinIoDriverName);
    }

    private void ResolveImodPaths(out string? scriptPath)
    {
        string startupPath = GetImodStartupPath();
        scriptPath = File.Exists(startupPath) ? startupPath : null;
    }

    private void RemoveImodPersistenceFiles()
    {
        DeleteFileIfExists(GetImodStartupPath(), "IMOD.CONFIG");
        DeleteFileIfExists(GetWinIoSystemPath(), "IMOD.DRIVER");
        DeleteFileIfExists(Path.Combine(GetScriptRoot(), WinIoDriverName), "IMOD.DRIVER");
    }

    private void DeleteFileIfExists(string path, string label)
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

    private static ImodConfig ParseImodScriptFile(string path)
    {
        ImodConfig config = new();
        ImodConfigEntry? currentDevice = null;
        bool inOverrides = false;

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

    private static string FormatPowerShellString(string value)
    {
        string escaped = value?.Replace("'", "''") ?? string.Empty;
        return $"'{escaped}'";
    }

    private static string GetWinIoSystemPathForScript()
    {
        string windows = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
        if (string.IsNullOrWhiteSpace(windows))
        {
            string baseDir = AppContext.BaseDirectory.TrimEnd(Path.DirectorySeparatorChar);
            windows = string.IsNullOrWhiteSpace(baseDir) ? Environment.CurrentDirectory : baseDir;
        }

        return Path.Combine(windows, WinIoDriverName);
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

        File.WriteAllText(path, scriptBody, Encoding.ASCII);
    }

    private static string BuildImodConfigBlock(ImodConfig config)
    {
        StringBuilder sb = new();
        sb.AppendLine(ImodScriptMarkerStart);
        sb.AppendLine($"$winioPath = {FormatPowerShellString(GetWinIoSystemPathForScript())}");
        sb.AppendLine($"$globalInterval = {FormatImodValue(config.GlobalInterval)}");
        sb.AppendLine($"$globalHCSPARAMSOffset = {FormatImodValue(config.GlobalHcsparamsOffset)}");
        sb.AppendLine($"$globalRTSOFF = {FormatImodValue(config.GlobalRtsoff)}");
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

            if (entry.Interval.HasValue || entry.HcsparamsOffset.HasValue || entry.Rtsoff.HasValue)
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

    internal static int ApplyImodFromScript(string? scriptPath, out string? note)
    {
        MainForm form = new(true);
        return form.ApplyImodFromScriptInternal(scriptPath, out note);
    }

    private int ApplyImodFromScriptInternal(string? scriptPath, out string? note)
    {
        note = null;
        bool explicitScript = !string.IsNullOrWhiteSpace(scriptPath);
        string? resolvedPath = scriptPath;
        if (string.IsNullOrWhiteSpace(resolvedPath))
        {
            ResolveImodPaths(out resolvedPath);
        }

        if (explicitScript && (string.IsNullOrWhiteSpace(resolvedPath) || !File.Exists(resolvedPath)))
        {
            note = "IMOD failed: ApplyIMOD.ps1 not found.";
            return 1;
        }

        ImodConfig config = !string.IsNullOrWhiteSpace(resolvedPath) && File.Exists(resolvedPath)
            ? ParseImodScriptFile(resolvedPath)
            : new ImodConfig();
        config.HasScript = !string.IsNullOrWhiteSpace(resolvedPath) && File.Exists(resolvedPath);

        if (!HasActiveImod(config))
        {
            note = "IMOD skipped (no configured values).";
            return 0;
        }

        bool persistDriver = config.HasScript || HasCustomImod(config);
        if (!TryApplyImod(config, persistDriver, out ImodApplyStats stats, out string? error))
        {
            note = error is null ? "IMOD failed." : $"IMOD failed: {error}";
            return 1;
        }

        if (stats.ControllersFound == 0)
        {
            note = "IMOD skipped (no XHCI controllers found).";
            return 0;
        }

        if (stats.ControllersApplied == 0)
        {
            note = "IMOD skipped (no eligible USB controllers).";
            return 0;
        }

        if (stats.ReadFailures > 0 || stats.WriteFailures > 0)
        {
            note = $"IMOD applied to {stats.ControllersApplied} USB controller(s) with {stats.ReadFailures} read failure(s) and {stats.WriteFailures} write failure(s).";
        }
        else
        {
            note = $"IMOD applied to {stats.ControllersApplied} USB controller(s).";
        }

        return 0;
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

            if (!TryParseImodInterval(text, config.GlobalInterval, out uint interval))
            {
                string shortPnp = GetShortPnpId(block.Device.InstanceId);
                string label = string.IsNullOrWhiteSpace(shortPnp) ? block.Device.Name : $"{block.Device.Name} ({shortPnp})";
                invalidInputs.Add(label);
                interval = config.GlobalInterval;
            }

            if (interval == config.GlobalInterval)
            {
                config.Overrides.RemoveAll(e =>
                    string.Equals(e.Hwid, hwid, StringComparison.OrdinalIgnoreCase));
                block.ImodBox.Text = FormatImodValue(config.GlobalInterval);
                continue;
            }

            block.ImodBox.Text = FormatImodValue(interval);

            if (existing is not null)
            {
                config.Overrides.RemoveAll(e =>
                    !ReferenceEquals(e, existing)
                    && string.Equals(e.Hwid, hwid, StringComparison.OrdinalIgnoreCase));

                existing.Enabled = true;
                existing.Interval = interval;
            }
            else
            {
                config.Overrides.Add(new ImodConfigEntry
                {
                    Hwid = hwid,
                    Enabled = true,
                    Interval = interval,
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
            existing.HcsparamsOffset = null;
            existing.Rtsoff = null;
            WriteLog($"IMOD.CONFIG: skip non-hid {block.Device.InstanceId} roles=\"{block.Device.UsbRoles}\"");
        }

        if (invalidInputs.Count > 0)
        {
            string shown = string.Join(", ", invalidInputs.Take(3));
            string suffix = invalidInputs.Count > 3 ? " ..." : string.Empty;
            MessageBox.Show(
                $"Invalid IMOD interval value detected for: {shown}{suffix}\nValues have been reset to default ({FormatImodValue(config.GlobalInterval)}).",
                "DEVICE TWEAKER",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }

        bool hasCustom = HasCustomImod(config);
        string startupPath = GetImodStartupPath();
        if (hasCustom)
        {
            try
            {
                WriteImodScript(config, startupPath);
                config.HasScript = true;
                WriteLog($"IMOD.CONFIG: script saved {startupPath}");
            }
            catch (Exception ex)
            {
                note = $"IMOD failed to write script: {ex.Message}";
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

        if (!TryApplyImod(config, hasCustom, out ImodApplyStats stats, out string? applyError))
        {
            note = applyError is null ? "IMOD failed." : $"IMOD failed: {applyError}";
            WriteLog($"IMOD: {note}");
            return ImodApplyOutcome.Failed;
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

