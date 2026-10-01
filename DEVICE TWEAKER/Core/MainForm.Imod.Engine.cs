using System.ComponentModel;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Threading;

namespace DeviceTweakerCS;

public sealed partial class MainForm
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

    private sealed class ImodControllerInfo
    {
        public string DeviceId { get; init; } = string.Empty;
        public string Caption { get; init; } = string.Empty;
        public uint ProblemCode { get; init; }
        public ulong BaseAddress { get; init; }
        public bool HasBase { get; init; }
        public string BaseError { get; init; } = string.Empty;
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
        [Out] byte[]? propertyBuffer,
        uint propertyBufferSize,
        out uint requiredSize);

    [DllImport("setupapi.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool SetupDiGetDeviceInstanceIdW(
        IntPtr deviceInfoSet,
        ref SP_DEVINFO_DATA deviceInfoData,
        StringBuilder? deviceInstanceId,
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
        string? machineName,
        string? databaseName,
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
        string? loadOrderGroup,
        IntPtr tagId,
        string? dependencies,
        string? serviceStartName,
        string? password);

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

    private sealed class ImodApplyStats
    {
        public int ControllersFound { get; set; }
        public int ControllersApplied { get; set; }
        public int WriteFailures { get; set; }
        public int SkippedDisabled { get; set; }
        public int MissingBase { get; set; }
        public int ReadFailures { get; set; }
    }

    private bool TryApplyImod(ImodConfig config, bool persistDriver, out ImodApplyStats stats, out string? error)
    {
        stats = new ImodApplyStats();
        error = null;

        if (!IsAdministrator())
        {
            error = "administrator privileges required";
            return false;
        }

        if (!EnsureWinIoDriverOnDisk(persistDriver, out string driverPath, out error))
        {
            return false;
        }

        bool cleanupDriver = !persistDriver;
        try
        {
            if (!TryEnumerateXhciControllers(out List<ImodControllerInfo> controllers, out error))
            {
                return false;
            }

            stats.ControllersFound = controllers.Count;
            if (controllers.Count == 0)
            {
                return true;
            }

            if (!WinIoContext.TryInitialize(driverPath, out WinIoContext? winio, out error))
            {
                return false;
            }

            WinIoContext winioContext = winio!;
            using (winioContext)
            {
                WriteLog($"IMOD: controllers={controllers.Count}");
                foreach (ImodControllerInfo controller in controllers)
                {
                    if (controller.ProblemCode == CmProbDisabled)
                    {
                        stats.SkippedDisabled++;
                        WriteLog($"IMOD: skipped disabled {controller.DeviceId}");
                        continue;
                    }

                    if (!controller.HasBase)
                    {
                        stats.MissingBase++;
                        if (!string.IsNullOrWhiteSpace(controller.BaseError))
                        {
                            WriteLog($"IMOD: {controller.DeviceId} base error: {controller.BaseError}");
                        }
                        continue;
                    }

                    uint desiredInterval = config.GlobalInterval;
                    uint hcsparamsOffset = config.GlobalHcsparamsOffset;
                    uint rtsoff = config.GlobalRtsoff;
                    bool enabled = true;
                    string? overrideMatch = null;

                    foreach (ImodConfigEntry entry in config.Overrides)
                    {
                        if (string.IsNullOrWhiteSpace(entry.Hwid))
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
                        WriteLog($"IMOD: skipped config-disabled {controller.DeviceId} ({overrideMatch})");
                        continue;
                    }

                    ulong capabilityAddress = controller.BaseAddress;

                    if (!TryReadPhys32(winioContext, capabilityAddress + hcsparamsOffset, out uint hcsparamsValue, out string? ioError))
                    {
                        stats.ReadFailures++;
                        WriteLog($"IMOD: read HCSPARAMS failed {controller.DeviceId}: {ioError}");
                        continue;
                    }

                    if (!TryReadPhys32(winioContext, capabilityAddress + rtsoff, out uint rtsoffValue, out ioError))
                    {
                        stats.ReadFailures++;
                        WriteLog($"IMOD: read RTSOFF failed {controller.DeviceId}: {ioError}");
                        continue;
                    }

                    uint maxIntrs = (hcsparamsValue >> 8) & 0xFF;
                    ulong runtimeAddress = capabilityAddress + rtsoffValue;

                    uint writeFailures = 0;
                    for (uint i = 0; i < maxIntrs; ++i)
                    {
                        ulong interrupterAddress = runtimeAddress + 0x24 + (0x20 * i);
                        if (!TryWritePhys32(winioContext, interrupterAddress, desiredInterval, out ioError))
                        {
                            writeFailures++;
                            WriteLog($"IMOD: write failed {controller.DeviceId} @ {ToHex(interrupterAddress)}: {ioError}");
                        }
                    }

                    stats.ControllersApplied++;
                    stats.WriteFailures += (int)writeFailures;

                    WriteLog($"IMOD: {controller.DeviceId} writes={maxIntrs} failures={writeFailures}");
                }
            }
        }
        finally
        {
            if (cleanupDriver)
            {
                DeleteFileIfExists(driverPath, "IMOD.DRIVER");
            }
        }

        return true;
    }

    private static bool IsAdministrator()
    {
        try
        {
            using WindowsIdentity identity = WindowsIdentity.GetCurrent();
            WindowsPrincipal principal = new(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }
        catch
        {
            return false;
        }
    }

    private bool EnsureWinIoDriverOnDisk(bool persistDriver, out string driverPath, out string? error)
    {
        error = null;
        driverPath = GetWinIoSystemPath();

        try
        {
            if (File.Exists(driverPath))
            {
                return true;
            }

            using Stream? resource = OpenWinIoResourceStream();
            if (resource is null)
            {
                error = "embedded winio.sys not found";
                return false;
            }

            using FileStream output = new(driverPath, FileMode.Create, FileAccess.Write, FileShare.Read);
            resource.CopyTo(output);
            WriteLog($"IMOD.DRIVER: extracted {driverPath}");
            return true;
        }
        catch (Exception ex)
        {
            error = $"failed to extract winio.sys: {ex.Message}";
            return false;
        }
    }

    private static Stream? OpenWinIoResourceStream()
    {
        Assembly asm = typeof(MainForm).Assembly;
        Stream? stream = asm.GetManifestResourceStream("DeviceTweakerCS.IMOD.winio.sys");
        if (stream is not null)
        {
            return stream;
        }

        foreach (string name in asm.GetManifestResourceNames())
        {
            if (name.EndsWith(".winio.sys", StringComparison.OrdinalIgnoreCase))
            {
                return asm.GetManifestResourceStream(name);
            }
        }

        return null;
    }

    private static string ToHex(ulong value)
    {
        return $"0x{value:X}";
    }

    private static bool TryEnumerateXhciControllers(out List<ImodControllerInfo> controllers, out string? error)
    {
        controllers = [];
        error = null;

        IntPtr devInfoSet = SetupDiGetClassDevsW(IntPtr.Zero, "PCI", IntPtr.Zero, DigcfPresent | DigcfAllClasses);
        if (devInfoSet == InvalidHandleValue)
        {
            error = $"failed to enumerate PCI devices: {GetWin32ErrorMessage(Marshal.GetLastWin32Error())}";
            return false;
        }

        try
        {
            for (uint index = 0; ; index++)
            {
                SP_DEVINFO_DATA devInfo = new()
                {
                    cbSize = (uint)Marshal.SizeOf<SP_DEVINFO_DATA>(),
                };

                if (!SetupDiEnumDeviceInfo(devInfoSet, index, ref devInfo))
                {
                    int lastError = Marshal.GetLastWin32Error();
                    if (lastError == ErrorNoMoreItems)
                    {
                        break;
                    }

                    error = $"failed to enumerate device info: {GetWin32ErrorMessage(lastError)}";
                    return false;
                }

                if (!IsXhciDevice(devInfoSet, ref devInfo))
                {
                    continue;
                }

                if (!TryGetDeviceInstanceId(devInfoSet, ref devInfo, out string instanceId))
                {
                    continue;
                }

                string caption = GetDeviceCaption(devInfoSet, ref devInfo);
                _ = TryGetDeviceProblemCode(devInfo.DevInst, out uint problemCode);

                ulong baseAddress = 0;
                bool hasBase = TryGetDeviceMemoryBase(devInfo.DevInst, out baseAddress, out string? baseError);

                controllers.Add(new ImodControllerInfo
                {
                    DeviceId = instanceId,
                    Caption = caption,
                    ProblemCode = problemCode,
                    BaseAddress = baseAddress,
                    HasBase = hasBase,
                    BaseError = baseError ?? string.Empty,
                });
            }
        }
        finally
        {
            _ = SetupDiDestroyDeviceInfoList(devInfoSet);
        }

        return true;
    }

    private static bool IsXhciDevice(IntPtr devInfoSet, ref SP_DEVINFO_DATA devInfo)
    {
        if (TryGetDeviceStringProperty(devInfoSet, ref devInfo, SpdrpService, out string service))
        {
            if (string.Equals(service, "USBXHCI", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        if (TryGetDeviceMultiSzProperty(devInfoSet, ref devInfo, SpdrpHardwareId, out List<string> ids)
            && HasXhciClassCode(ids))
        {
            return true;
        }

        if (TryGetDeviceMultiSzProperty(devInfoSet, ref devInfo, SpdrpCompatibleIds, out ids)
            && HasXhciClassCode(ids))
        {
            return true;
        }

        return false;
    }

    private static bool HasXhciClassCode(IEnumerable<string> ids)
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
        if (TryGetDeviceStringProperty(devInfoSet, ref devInfo, SpdrpFriendlyName, out string caption))
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
        int cr = CM_Get_DevNode_Status(out _, out uint problem, devInst, 0);
        if (cr != CrSuccess)
        {
            return false;
        }

        problemCode = problem;
        return true;
    }

    private static bool TryGetDeviceMemoryBase(uint devInst, out ulong baseAddress, out string? error)
    {
        baseAddress = 0;
        error = null;

        int cr = CM_Get_First_Log_Conf(out IntPtr logConf, devInst, AllocLogConf);
        if (cr != CrSuccess)
        {
            cr = CM_Get_First_Log_Conf(out logConf, devInst, BootLogConf);
        }
        if (cr != CrSuccess)
        {
            error = $"failed to query logical config (CONFIGRET {cr})";
            return false;
        }

        try
        {
            bool found = false;
            ulong minBase = 0;
            foreach (uint resType in new[] { ResTypeMem, ResTypeMemLarge })
            {
                int resCr = CM_Get_Next_Res_Des(out IntPtr resDes, logConf, resType, IntPtr.Zero, 0);
                while (resCr == CrSuccess)
                {
                    int sizeCr = CM_Get_Res_Des_Data_Size(out uint dataSize, resDes, 0);
                    if (sizeCr == CrSuccess && dataSize > 0)
                    {
                        byte[] buffer = new byte[dataSize];
                        if (CM_Get_Res_Des_Data(resDes, buffer, dataSize, 0) == CrSuccess)
                        {
                            if (TryExtractBaseFromResource(resType, buffer, out ulong candidate))
                            {
                                if (!found || candidate < minBase)
                                {
                                    minBase = candidate;
                                    found = true;
                                }
                            }
                        }
                    }

                    int nextCr = CM_Get_Next_Res_Des(out IntPtr nextResDes, resDes, resType, IntPtr.Zero, 0);
                    _ = CM_Free_Res_Des_Handle(resDes);
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
            _ = CM_Free_Log_Conf_Handle(logConf);
        }
    }

    private static bool TryExtractBaseFromResource(uint resType, byte[] data, out ulong baseAddress)
    {
        baseAddress = 0;

        if (resType == ResTypeMem)
        {
            if (data.Length < Marshal.SizeOf<MemDes>())
            {
                return false;
            }

            MemDes mem = MemoryMarshal.Read<MemDes>(data);
            ulong candidate = mem.MD_Alloc_Base;
            if (candidate == 0 && mem.MD_Count > 0)
            {
                int offset = Marshal.SizeOf<MemDes>();
                if (data.Length >= offset + Marshal.SizeOf<MemRange>())
                {
                    MemRange range = MemoryMarshal.Read<MemRange>(data.AsSpan(offset));
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
            if (data.Length < Marshal.SizeOf<MemLargeDes>())
            {
                return false;
            }

            MemLargeDes mem = MemoryMarshal.Read<MemLargeDes>(data);
            ulong candidate = mem.MLD_Alloc_Base;
            if (candidate == 0 && mem.MLD_Count > 0)
            {
                int offset = Marshal.SizeOf<MemLargeDes>();
                if (data.Length >= offset + Marshal.SizeOf<MemLargeRange>())
                {
                    MemLargeRange range = MemoryMarshal.Read<MemLargeRange>(data.AsSpan(offset));
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

    private static bool TryGetDeviceStringProperty(
        IntPtr devInfoSet,
        ref SP_DEVINFO_DATA devInfo,
        uint property,
        out string value)
    {
        value = string.Empty;
        if (!TryGetDeviceMultiSzProperty(devInfoSet, ref devInfo, property, out List<string> values))
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

    private static bool TryGetDeviceMultiSzProperty(
        IntPtr devInfoSet,
        ref SP_DEVINFO_DATA devInfo,
        uint property,
        out List<string> values)
    {
        values = [];
        if (!TryGetDevicePropertyData(devInfoSet, ref devInfo, property, out byte[] data, out uint regType))
        {
            return false;
        }

        if (regType != RegMultiSz && regType != RegSz)
        {
            return false;
        }

        string text = Encoding.Unicode.GetString(data);
        string[] parts = text.Split('\0', StringSplitOptions.RemoveEmptyEntries);
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

    private static bool TryGetDevicePropertyData(
        IntPtr devInfoSet,
        ref SP_DEVINFO_DATA devInfo,
        uint property,
        out byte[] data,
        out uint regType)
    {
        data = [];
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
        if (!SetupDiGetDeviceRegistryPropertyW(devInfoSet, ref devInfo, property, out regType, data, requiredSize, out _))
        {
            return false;
        }

        return true;
    }

    private static bool TryGetDeviceInstanceId(
        IntPtr devInfoSet,
        ref SP_DEVINFO_DATA devInfo,
        out string instanceId)
    {
        instanceId = string.Empty;
        int requiredSize = 0;
        _ = SetupDiGetDeviceInstanceIdW(devInfoSet, ref devInfo, null, 0, out requiredSize);
        int err = Marshal.GetLastWin32Error();
        if (err != ErrorInsufficientBuffer || requiredSize <= 0)
        {
            return false;
        }

        StringBuilder buffer = new(requiredSize);
        if (!SetupDiGetDeviceInstanceIdW(devInfoSet, ref devInfo, buffer, buffer.Capacity, out _))
        {
            return false;
        }

        instanceId = buffer.ToString();
        return !string.IsNullOrWhiteSpace(instanceId);
    }

    private static bool TryReadPhys32(WinIoContext ctx, ulong address, out uint value, out string? error)
    {
        value = 0;
        error = null;

        if (!TryMapPhysicalMemory(ctx, address, 4, out PhysStruct phys, out error))
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
            if (!TryUnmapPhysicalMemory(ctx, phys, out string? unmapError))
            {
                error = unmapError;
                success = false;
            }
        }

        return success;
    }

    private static bool TryWritePhys32(WinIoContext ctx, ulong address, uint value, out string? error)
    {
        error = null;
        if (!TryMapPhysicalMemory(ctx, address, 4, out PhysStruct phys, out error))
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
            if (!TryUnmapPhysicalMemory(ctx, phys, out string? unmapError))
            {
                error = unmapError;
                success = false;
            }
        }

        return success;
    }

    private static bool TryMapPhysicalMemory(WinIoContext ctx, ulong address, ulong size, out PhysStruct phys, out string? error)
    {
        error = null;
        phys = new PhysStruct
        {
            physMemSizeInBytes = size,
            physAddress = address,
        };

        int bytesReturned = 0;
        if (!DeviceIoControl(
                ctx.DriverHandle,
                IoctlWinioMapPhysToLin,
                ref phys,
                Marshal.SizeOf<PhysStruct>(),
                ref phys,
                Marshal.SizeOf<PhysStruct>(),
                out bytesReturned,
                IntPtr.Zero))
        {
            error = $"failed to map physical memory: {GetWin32ErrorMessage(Marshal.GetLastWin32Error())}";
            return false;
        }

        if (phys.physMemLin == 0)
        {
            error = "failed to map physical memory: returned null linear address";
            return false;
        }

        return true;
    }

    private static bool TryUnmapPhysicalMemory(WinIoContext ctx, PhysStruct phys, out string? error)
    {
        error = null;
        int bytesReturned = 0;
        if (!DeviceIoControl(
                ctx.DriverHandle,
                IoctlWinioUnmapPhysAddr,
                ref phys,
                Marshal.SizeOf<PhysStruct>(),
                ref phys,
                Marshal.SizeOf<PhysStruct>(),
                out bytesReturned,
                IntPtr.Zero))
        {
            error = $"failed to unmap physical memory: {GetWin32ErrorMessage(Marshal.GetLastWin32Error())}";
            return false;
        }

        return true;
    }

    private sealed class WinIoContext : IDisposable
    {
        public IntPtr DriverHandle { get; private set; } = InvalidHandleValue;
        public bool Is64BitOS { get; private set; }
        public bool ServiceCreated { get; private set; }
        public string DriverPath { get; }

        private WinIoContext(string driverPath)
        {
            DriverPath = driverPath;
        }

        public static bool TryInitialize(string driverPath, out WinIoContext? ctx, out string? error)
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

        private bool Initialize(out string? error)
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
                error = $"failed to open {WinIoDevicePath}: {GetWin32ErrorMessage(Marshal.GetLastWin32Error())}";
                return false;
            }

            if (!EnableDirectIo(out error))
            {
                return false;
            }

            return true;
        }

        private bool EnsureWinIoService(out string? error)
        {
            error = null;
            IntPtr scm = OpenSCManager(null, null, ScManagerAllAccess);
            if (scm == IntPtr.Zero)
            {
                error = $"failed to open service manager: {GetWin32ErrorMessage(Marshal.GetLastWin32Error())}";
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
                        error = $"failed to open WINIO service: {GetWin32ErrorMessage(lastError)}";
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
                        error = $"failed to create WINIO service: {GetWin32ErrorMessage(Marshal.GetLastWin32Error())}";
                        return false;
                    }
                    ServiceCreated = true;
                }

                try
                {
                    bool wasRunning = false;
                    if (QueryServiceStatus(service, out SERVICE_STATUS_PROCESS status))
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
                                error = $"failed to start WINIO service: {GetWin32ErrorMessage(lastError)}";
                                return false;
                            }
                        }
                    }
                }
                finally
                {
                    _ = CloseServiceHandle(service);
                }
            }
            finally
            {
                _ = CloseServiceHandle(scm);
            }

            return true;
        }

        private bool EnableDirectIo(out string? error)
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
                error = $"failed to enable direct I/O: {GetWin32ErrorMessage(Marshal.GetLastWin32Error())}";
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

            _ = DeviceIoControl(
                DriverHandle,
                IoctlWinioDisableDirectIo,
                IntPtr.Zero,
                0,
                IntPtr.Zero,
                0,
                out _,
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
                    if (QueryServiceStatus(service, out SERVICE_STATUS_PROCESS status)
                        && status.dwCurrentState != ServiceStopped)
                    {
                        SERVICE_STATUS serviceStatus = new();
                        _ = ControlService(service, ServiceControlStop, ref serviceStatus);

                        for (int i = 0; i < 25; i++)
                        {
                            if (!QueryServiceStatus(service, out SERVICE_STATUS_PROCESS check)
                                || check.dwCurrentState == ServiceStopped)
                            {
                                break;
                            }

                            Thread.Sleep(200);
                        }
                    }

                    if (ServiceCreated)
                    {
                        _ = DeleteService(service);
                    }
                }
                finally
                {
                    _ = CloseServiceHandle(service);
                }
            }
            finally
            {
                _ = CloseServiceHandle(scm);
            }
        }

        public void Dispose()
        {
            DisableDirectIo();

            if (DriverHandle != InvalidHandleValue)
            {
                _ = CloseHandle(DriverHandle);
                DriverHandle = InvalidHandleValue;
            }

            StopServiceIfNeeded();
        }
    }

    private static bool QueryServiceStatus(IntPtr service, out SERVICE_STATUS_PROCESS status)
    {
        status = new SERVICE_STATUS_PROCESS();
        return QueryServiceStatusEx(service, ScStatusProcessInfo, ref status, (uint)Marshal.SizeOf<SERVICE_STATUS_PROCESS>(), out _);
    }

    private static string GetWin32ErrorMessage(int error)
    {
        return new Win32Exception(error).Message;
    }

    private static uint CtlCode(uint deviceType, uint function, uint method, uint access)
    {
        return (deviceType << 16) | (access << 14) | (function << 2) | method;
    }

    private static readonly IntPtr InvalidHandleValue = new(-1);

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
