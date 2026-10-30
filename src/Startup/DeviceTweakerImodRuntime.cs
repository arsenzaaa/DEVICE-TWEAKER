// Hardware core shared by DEVICE TWEAKER and Startup\ApplyIMOD.ps1, which compiles this file
// with Add-Type in Windows PowerShell 5.1. Keep C# 5 syntax.
// The app compiles it with nullable annotations on; C# 5 has no #nullable, so the
// related warnings are silenced instead.
#if NET8_0_OR_GREATER
#pragma warning disable CS8600, CS8601, CS8602, CS8603, CS8604, CS8605, CS8618, CS8625
#endif
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;

// Register access for the IMOD / NIC ITR sequences: DTIMOD in the app and the logon
// script, an in-memory bus in tests. Sizes are in bytes (1, 2 or 4).
public interface IDeviceRegisters
{
    bool TryRead(ulong address, uint size, out ulong value, out string error);
    bool TryWrite(ulong address, uint size, ulong value, out string error);
}

public sealed class DeviceTweakerImodController
{
    public string DeviceId;
    public string Caption;
    public uint ProblemCode;
    public ulong BaseAddress;
    public ulong MemoryLength;
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
    private const uint FileWriteAccess = 2;
    private const int OpenRetryCount = 10;
    private const int OpenRetryDelayMs = 100;
    private const string ImodDriverDevicePath = "\\\\.\\DeviceTweakerImod2";
    private const string ImodDriverServiceName = "DeviceTweakerImod2";
    private static readonly IntPtr InvalidHandleValue = new IntPtr(-1);
    private static readonly uint IoctlImodReadPhysicalMemory = CtlCode(FileDeviceImod, ImodIoctlIndex + 2, MethodBuffered, FileAnyAccess);
    // DTIMOD writes only inside memory ranges Windows assigned to the named PCI devnode.
    internal static readonly uint IoctlImodWriteDevice = CtlCode(FileDeviceImod, ImodIoctlIndex + 5, MethodBuffered, FileWriteAccess);
    private const uint ImodDeviceWriteVersion = 1;
    private const int ImodDeviceWriteRequestSize = 1048;
    private const int ImodDeviceWriteReplySize = 8 + 65536;

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

    [DllImport("kernel32.dll", SetLastError = true, EntryPoint = "DeviceIoControl")]
    private static extern bool DeviceIoControlBytes(IntPtr deviceHandle, uint ioControlCode, byte[] inBuffer, int inBufferSize, byte[] outBuffer, int outBufferSize, out int bytesReturned, IntPtr overlapped);

    [DllImport("cfgmgr32.dll", CharSet = CharSet.Unicode)]
    private static extern int CM_Locate_DevNodeW(out uint devInst, string deviceId, uint flags);

    [DllImport("cfgmgr32.dll", CharSet = CharSet.Unicode)]
    private static extern int CM_Get_DevNode_PropertyW(uint devInst, ref DevPropKey key, out uint type, byte[] buffer, ref uint size, uint flags);

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
                ulong memoryLength;
                string baseError;
                bool hasBase = TryGetDeviceMemoryRange(devInfo.DevInst, out baseAddress, out memoryLength, out baseError);

                DeviceTweakerImodController controller = new DeviceTweakerImodController();
                controller.DeviceId = instanceId;
                controller.Caption = GetDeviceCaption(devInfoSet, ref devInfo);
                controller.ProblemCode = problemCode;
                controller.BaseAddress = baseAddress;
                controller.MemoryLength = memoryLength;
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
                ulong memoryLength;
                string baseError;
                bool hasBase = TryGetDeviceMemoryRange(devInfo.DevInst, out baseAddress, out memoryLength, out baseError);

                DeviceTweakerImodController device = new DeviceTweakerImodController();
                device.DeviceId = instanceId;
                device.Caption = GetDeviceCaption(devInfoSet, ref devInfo);
                device.ProblemCode = problemCode;
                device.BaseAddress = baseAddress;
                device.MemoryLength = memoryLength;
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

    // Memory range of one exact PCI devnode (no enumeration of the whole bus).
    public static bool TryGetPciMemoryRange(string instanceId, out ulong baseAddress, out ulong memoryLength, out string error)
    {
        baseAddress = 0;
        memoryLength = 0;
        uint devInst;
        int cr = CM_Locate_DevNodeW(out devInst, instanceId, 0);
        if (cr != CrSuccess)
        {
            error = "PCI device not found: " + instanceId + " (CM 0x" + cr.ToString("X", CultureInfo.InvariantCulture) + ")";
            return false;
        }

        return TryGetDeviceMemoryRange(devInst, out baseAddress, out memoryLength, out error);
    }

    private static bool RangeContains(ulong memoryLength, uint offset, ulong size)
    {
        ulong value = offset;
        return value <= memoryLength && size <= memoryLength - value;
    }

    internal static bool TryResolveXhciRuntimeLayout(
        IDeviceRegisters registers,
        ulong capabilityAddress,
        ulong memoryLength,
        uint hcsparamsOffset,
        uint rtsoffOffset,
        out uint maxIntrs,
        out ulong runtimeAddress,
        out uint hcsparamsValue,
        out uint rtsoffValue,
        out string error)
    {
        maxIntrs = 0;
        runtimeAddress = 0;
        hcsparamsValue = 0;
        rtsoffValue = 0;
        error = null;

        if (capabilityAddress == 0 || memoryLength < 0x20UL)
        {
            error = "missing or undersized PCI memory resource (length=0x" + memoryLength.ToString("X", CultureInfo.InvariantCulture) + ")";
            return false;
        }
        if (!RangeContains(memoryLength, 0, 4) || !RangeContains(memoryLength, hcsparamsOffset, 4) || !RangeContains(memoryLength, rtsoffOffset, 4))
        {
            error = "capability offsets exceed PCI memory resource (HCS=0x" + hcsparamsOffset.ToString("X", CultureInfo.InvariantCulture)
                + ", RTS=0x" + rtsoffOffset.ToString("X", CultureInfo.InvariantCulture) + ", length=0x" + memoryLength.ToString("X", CultureInfo.InvariantCulture) + ")";
            return false;
        }

        uint capabilityHeader;
        if (!TryRead32(registers, capabilityAddress, out capabilityHeader, out error))
        {
            error = "failed to read xHCI capability header: " + error;
            return false;
        }
        uint capabilityLength = capabilityHeader & 0xFFU;
        uint hciVersion = capabilityHeader >> 16;
        if (capabilityLength < 0x20U || capabilityLength > memoryLength)
        {
            error = "invalid xHCI CAPLENGTH 0x" + capabilityLength.ToString("X", CultureInfo.InvariantCulture) + " for BAR length 0x" + memoryLength.ToString("X", CultureInfo.InvariantCulture);
            return false;
        }
        if (hciVersion < 0x0090U || hciVersion > 0x0120U)
        {
            error = "unexpected xHCI version 0x" + hciVersion.ToString("X4", CultureInfo.InvariantCulture);
            return false;
        }

        if (!TryRead32(registers, capabilityAddress + hcsparamsOffset, out hcsparamsValue, out error))
        {
            error = "failed to read HCSPARAMS: " + error;
            return false;
        }
        maxIntrs = (hcsparamsValue >> 8) & 0x7FFU;
        if (maxIntrs == 0)
        {
            error = "invalid MaxIntrs value 0";
            return false;
        }

        if (!TryRead32(registers, capabilityAddress + rtsoffOffset, out rtsoffValue, out error))
        {
            error = "failed to read RTSOFF: " + error;
            return false;
        }
        // RTSOFF is 32-bit and MaxIntrs at most 2047, so the end cannot overflow.
        ulong runtimeOffset = ((ulong)rtsoffValue) & 0xFFFFFFE0UL;
        ulong finalRegisterEnd = runtimeOffset + 0x24UL + (0x20UL * ((ulong)maxIntrs - 1UL)) + 4UL;
        if (runtimeOffset == 0 || finalRegisterEnd > memoryLength)
        {
            error = "xHCI runtime registers exceed PCI memory resource (RTSOFF=0x" + runtimeOffset.ToString("X", CultureInfo.InvariantCulture)
                + ", end=0x" + finalRegisterEnd.ToString("X", CultureInfo.InvariantCulture) + ", length=0x" + memoryLength.ToString("X", CultureInfo.InvariantCulture) + ")";
            return false;
        }

        runtimeAddress = capabilityAddress + runtimeOffset;
        return true;
    }

    public static string ApplyController(string deviceId, ulong capabilityAddress, ulong memoryLength, uint hcsparamsOffset, uint rtsoff, uint interval, uint[] intervals)
    {
        DeviceWriteTarget writeTarget;
        string targetError;
        if (!DeviceWriteTarget.TryResolve(deviceId, out writeTarget, out targetError))
        {
            return "error: " + targetError;
        }

        IntPtr handle;
        string openError;
        if (!TryOpenDriverDevice(OpenRetryCount, out handle, out openError))
        {
            return "error: " + openError;
        }

        try
        {
            if (IsOlderDriverLoaded(handle))
            {
                return "error: " + OlderDriverMessage;
            }

            DtimodRegisters registers = new DtimodRegisters(handle, writeTarget);
            string ioError;
            uint maxIntrs;
            ulong runtimeAddress;
            uint hcsparamsValue;
            uint rtsoffValue;
            if (!TryResolveXhciRuntimeLayout(registers, capabilityAddress, memoryLength, hcsparamsOffset, rtsoff, out maxIntrs, out runtimeAddress, out hcsparamsValue, out rtsoffValue, out ioError))
            {
                return "error: unsafe or invalid xHCI MMIO layout: " + ioError;
            }
            uint writeCount = intervals != null && intervals.Length > 0
                ? Math.Min(maxIntrs, (uint)intervals.Length)
                : maxIntrs;
            uint failures = 0;

            for (uint i = 0; i < writeCount; i++)
            {
                ulong interrupterAddress = runtimeAddress + 0x24 + (0x20 * i);
                uint target = intervals != null && intervals.Length > 0 ? intervals[(int)i] : interval;
                if (!TryWriteImodInterval(registers, interrupterAddress, target, out ioError))
                {
                    failures++;
                }
            }

            string mode = intervals != null && intervals.Length > 0 ? "vector=" + intervals.Length : "interval=0x" + interval.ToString("X", CultureInfo.InvariantCulture);
            return "writes=" + writeCount + "/" + maxIntrs
                + " hcsparams=0x" + hcsparamsValue.ToString("X", CultureInfo.InvariantCulture)
                + " rtsoff=0x" + rtsoffValue.ToString("X", CultureInfo.InvariantCulture)
                + " runtime=0x" + runtimeAddress.ToString("X", CultureInfo.InvariantCulture)
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
        ulong memoryLength,
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
            if (!TryReadXhciInterrupterTopology(new DtimodRegisters(handle, null), capabilityAddress, memoryLength, hcsparamsOffset, out maxIntrs, out topology, out topologyDetail))
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
                        shown.Add("port" + pair.Key + "->I" + intr + "=" + selectedRole + ":0x" + selectedValue.ToString("X", CultureInfo.InvariantCulture));
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

    public static string ApplyNicItr(string deviceId, ulong baseAddress, ulong memoryLength, uint baseOffset, uint stride, uint queues, uint width, ulong mask, ulong orBits, ulong[] values)
    {
        if (queues == 0)
        {
            return "error: queues=0";
        }
        if (width != 16 && width != 32)
        {
            return "error: unsupported width=" + width;
        }
        ulong registerWidth = width / 8U;
        ulong lastOffset = (ulong)baseOffset + ((ulong)(queues - 1U) * stride);
        if (baseAddress == 0 || memoryLength == 0 || lastOffset > memoryLength || registerWidth > memoryLength - lastOffset)
        {
            return "error: NIC ITR register range exceeds PCI memory resource";
        }
        if (values == null || values.Length == 0)
        {
            return "error: no values";
        }

        DeviceWriteTarget target;
        string targetError;
        if (!DeviceWriteTarget.TryResolve(deviceId, out target, out targetError))
        {
            return "error: " + targetError;
        }

        IntPtr handle;
        string openError;
        if (!TryOpenDriverDevice(OpenRetryCount, out handle, out openError))
        {
            return "error: " + openError;
        }

        try
        {
            if (IsOlderDriverLoaded(handle))
            {
                return "error: " + OlderDriverMessage;
            }

            uint size = width == 16 ? 2U : 4U;
            StringBuilder applied = new StringBuilder();
            for (uint q = 0; q < queues; q++)
            {
                ulong selected = q < values.Length ? values[q] : values[0];
                ulong address = baseAddress + baseOffset + (stride * q);
                if (applied.Length > 0)
                {
                    applied.Append(", ");
                }
                applied.Append("Q").Append(q).Append("@0x").Append(address.ToString("X", CultureInfo.InvariantCulture)).Append("=0x").Append(((selected & mask) | orBits).ToString("X", CultureInfo.InvariantCulture));
            }

            string ioError;
            if (!TryProgramNicItr(new DtimodRegisters(handle, target), baseAddress, baseOffset, stride, queues, size, mask, orBits, values, null, out ioError))
            {
                return "error: " + ioError + " values=[" + applied.ToString() + "]";
            }

            return "writes=" + queues + " width=" + width + " values=[" + applied.ToString() + "] verified";
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

    internal static bool TryRead32(IDeviceRegisters registers, ulong address, out uint value, out string error)
    {
        value = 0;
        ulong raw;
        if (!registers.TryRead(address, 4, out raw, out error))
        {
            return false;
        }
        value = unchecked((uint)raw);
        return true;
    }

    internal static bool TryRead64(IDeviceRegisters registers, ulong address, out ulong value, out string error)
    {
        value = 0;
        uint low;
        uint high;
        if (!TryRead32(registers, address, out low, out error) || !TryRead32(registers, address + 4, out high, out error))
        {
            return false;
        }
        value = (ulong)low | ((ulong)high << 32);
        return true;
    }

    // Writes the IMOD interval (low 16 bits) and keeps the counter (high 16 bits);
    // the interval is read back. Used by the app, the logon script and tests.
    public static bool TryWriteImodInterval(IDeviceRegisters registers, ulong address, uint interval, out string error)
    {
        ulong currentValue;
        if (!registers.TryRead(address, 4, out currentValue, out error))
        {
            return false;
        }
        uint mergedValue = ((uint)currentValue & 0xFFFF0000U) | (interval & 0xFFFFU);
        if (!registers.TryWrite(address, 4, mergedValue, out error))
        {
            return false;
        }
        ulong readBack;
        string readError;
        if (!registers.TryRead(address, 4, out readBack, out readError))
        {
            error = "IMOD write completed, but readback failed: " + readError;
            return false;
        }
        uint expectedInterval = interval & 0xFFFFU;
        uint actualInterval = (uint)readBack & 0xFFFFU;
        if (actualInterval != expectedInterval)
        {
            error = string.Format(CultureInfo.InvariantCulture, "IMOD readback mismatch at 0x{0:X}: expected interval 0x{1:X}, read 0x{2:X}", address, expectedInterval, actualInterval);
            return false;
        }
        error = null;
        return true;
    }

    // Programs one ITR register per queue and reads each back; the first failure stops
    // the remaining queues. labels name the queues in errors (Q0, IRQ0, ...).
    public static bool TryProgramNicItr(IDeviceRegisters registers, ulong baseAddress, uint baseOffset, uint stride,
        uint queues, uint size, ulong mask, ulong orBits, ulong[] values, string[] labels, out string error)
    {
        error = null;
        if (values == null || values.Length == 0)
        {
            error = "no values";
            return false;
        }
        for (uint q = 0; q < queues; q++)
        {
            string label = labels != null && q < labels.Length ? labels[q] : "Q" + q;
            ulong selected = q < values.Length ? values[q] : values[0];
            ulong finalValue = (selected & mask) | orBits;
            ulong address = baseAddress + baseOffset + (stride * q);
            if (!registers.TryWrite(address, size, finalValue, out error))
            {
                return false;
            }
            ulong readback;
            if (!registers.TryRead(address, size, out readback, out error))
            {
                error = "write completed but readback failed at " + label + ": " + error;
                return false;
            }
            ulong expected = selected & mask;
            ulong actual = readback & mask;
            if (actual != expected)
            {
                error = string.Format(CultureInfo.InvariantCulture, "write verification failed at {0}: expected 0x{1:X}, read 0x{2:X}", label, expected, actual);
                return false;
            }
        }
        return true;
    }

    internal static bool TryReadPhysicalMemory(IntPtr handle, ulong address, uint size, out ulong value, out string error)
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
        if (bytesReturned < Marshal.SizeOf(typeof(PhysAccessStruct)))
        {
            error = "incomplete DTIMOD read reply";
            return false;
        }
        value = access.value;
        return true;
    }

    // Exact PCI devnode DTIMOD accepts writes for: its PDO name and instance ID.
    public sealed class DeviceWriteTarget
    {
        public readonly string InstanceId;
        public readonly string PdoName;

        public DeviceWriteTarget(string instanceId, string pdoName)
        {
            InstanceId = instanceId;
            PdoName = pdoName;
        }

        public static bool TryResolve(string instanceId, out DeviceWriteTarget target, out string error)
        {
            target = null;
            error = null;
            uint devInst;
            int cr = CM_Locate_DevNodeW(out devInst, instanceId, 0);
            if (cr != CrSuccess)
            {
                error = "device not found: " + instanceId + " (CM 0x" + cr.ToString("X", CultureInfo.InvariantCulture) + ")";
                return false;
            }

            // DEVPKEY_Device_PDOName {a45c254e-df1c-4efd-8020-67d146a850e0}, 16.
            DevPropKey key = new DevPropKey();
            key.fmtid = new Guid("a45c254e-df1c-4efd-8020-67d146a850e0");
            key.pid = 16;
            byte[] buffer = new byte[512];
            uint size = (uint)buffer.Length;
            uint type;
            cr = CM_Get_DevNode_PropertyW(devInst, ref key, out type, buffer, ref size, 0);
            if (cr != CrSuccess || size < 4)
            {
                error = "PDO name not available for " + instanceId + " (CM 0x" + cr.ToString("X", CultureInfo.InvariantCulture) + ")";
                return false;
            }

            target = new DeviceWriteTarget(instanceId, Encoding.Unicode.GetString(buffer, 0, (int)size - 2));
            return true;
        }
    }

    // DTIMOD access for one device: reads anywhere, writes only inside its memory.
    public sealed class DtimodRegisters : IDeviceRegisters
    {
        private readonly IntPtr handle;
        private readonly DeviceWriteTarget target;

        public DtimodRegisters(IntPtr handle, DeviceWriteTarget target)
        {
            this.handle = handle;
            this.target = target;
        }

        public bool TryRead(ulong address, uint size, out ulong value, out string error)
        {
            return TryReadPhysicalMemory(handle, address, size, out value, out error);
        }

        public bool TryWrite(ulong address, uint size, ulong value, out string error)
        {
            if (target == null)
            {
                error = "no target device was selected for the DTIMOD write";
                return false;
            }
            return TryWritePhysicalMemory(handle, target, address, size, value, out error);
        }
    }

    // struct ImodDeviceWrite in IMOD\Driver\imod_driver.h; null when the request is invalid.
    internal static byte[] BuildDeviceWriteRequest(string pdoName, string instanceId, ulong address, uint size, ulong value)
    {
        if (pdoName == null || instanceId == null
            || !pdoName.StartsWith("\\Device\\", StringComparison.OrdinalIgnoreCase)
            || !instanceId.StartsWith("PCI\\", StringComparison.OrdinalIgnoreCase)
            || pdoName.Length >= 256 || instanceId.Length >= 256
            || pdoName.IndexOf('\0') >= 0 || instanceId.IndexOf('\0') >= 0
            || (size != 1 && size != 2 && size != 4))
        {
            return null;
        }

        byte[] request = new byte[ImodDeviceWriteRequestSize];
        BitConverter.GetBytes(ImodDeviceWriteVersion).CopyTo(request, 0);
        BitConverter.GetBytes(size).CopyTo(request, 4);
        BitConverter.GetBytes(address).CopyTo(request, 8);
        BitConverter.GetBytes(value).CopyTo(request, 16);
        Encoding.Unicode.GetBytes(pdoName).CopyTo(request, 24);
        Encoding.Unicode.GetBytes(instanceId).CopyTo(request, 536);
        return request;
    }

    internal const string OlderDriverMessage = "An older DTIMOD driver is loaded. Restart Windows, then press CHECK again.";

    // The DTIMOD build without device-checked writes is still in memory until Windows
    // restarts. It does not know IOCTL_IMOD_WRITE_DEVICE (ERROR_INVALID_FUNCTION); the
    // current driver refuses an empty request with ERROR_INVALID_PARAMETER. Nothing is written.
    public static bool IsOlderDriverLoaded(IntPtr handle)
    {
        int bytesReturned;
        if (DeviceIoControlBytes(handle, IoctlImodWriteDevice, null, 0, null, 0, out bytesReturned, IntPtr.Zero))
        {
            return false;
        }
        return Marshal.GetLastWin32Error() == 1;
    }

    internal static string DescribeWriteFailure(int win32Error, ulong address, string instanceId)
    {
        if (win32Error == 5)
        {
            return "DTIMOD refused the write: 0x" + address.ToString("X", CultureInfo.InvariantCulture) + " is outside the memory of " + instanceId;
        }
        if (win32Error == 1)
        {
            return OlderDriverMessage;
        }
        return "failed to write physical memory via driver: " + GetWin32ErrorMessage(win32Error);
    }

    private static bool TryWritePhysicalMemory(IntPtr handle, DeviceWriteTarget target, ulong address, uint size, ulong value, out string error)
    {
        error = null;
        byte[] request = BuildDeviceWriteRequest(target.PdoName, target.InstanceId, address, size, value);
        if (request == null)
        {
            error = "invalid DTIMOD write request for " + target.InstanceId;
            return false;
        }
        byte[] reply = new byte[ImodDeviceWriteReplySize];
        int bytesReturned;
        if (!DeviceIoControlBytes(handle, IoctlImodWriteDevice, request, request.Length, reply, reply.Length, out bytesReturned, IntPtr.Zero))
        {
            error = DescribeWriteFailure(Marshal.GetLastWin32Error(), address, target.InstanceId);
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

    // Highest-priority role first (Mouse, Keyboard, ...), then any other role by name.
    internal static bool TrySelectAdaptiveRoleInterval(
        HashSet<string> roles,
        IReadOnlyDictionary<string, uint> roleIntervals,
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

        List<string> others = new List<string>(roles);
        others.Sort(StringComparer.OrdinalIgnoreCase);
        foreach (string role in others)
        {
            if (roleIntervals.TryGetValue(role, out selectedValue))
            {
                selectedRole = role;
                return true;
            }
        }

        return false;
    }

    // Interrupter of every active non-hub device, keyed by root port and by device address.
    // Transfer events go to the Interrupter Target of the endpoint's TRBs, so that one wins;
    // the slot context target is used when no endpoint ring can be read.
    internal static bool TryReadXhciInterrupterTopology(
        IDeviceRegisters registers,
        ulong capabilityAddress,
        ulong memoryLength,
        uint hcsparamsOffset,
        out uint maxIntrs,
        out XhciInterrupterTopology topology,
        out string detail)
    {
        maxIntrs = 0;
        topology = new XhciInterrupterTopology();
        detail = null;

        if (capabilityAddress == 0 || memoryLength < 0x40UL
            || !RangeContains(memoryLength, hcsparamsOffset, 4)
            || !RangeContains(memoryLength, 0x10, 4))
        {
            detail = "xHCI capability registers exceed PCI memory resource";
            return false;
        }

        uint capReg;
        string ioError;
        if (!TryRead32(registers, capabilityAddress, out capReg, out ioError))
        {
            detail = "failed to read xHCI CAPLENGTH: " + ioError;
            return false;
        }

        uint capLength = capReg & 0xFFU;
        uint hciVersion = capReg >> 16;
        if (capLength < 0x20U || capLength > memoryLength)
        {
            detail = "invalid xHCI CAPLENGTH 0x" + capLength.ToString("X", CultureInfo.InvariantCulture);
            return false;
        }
        if (hciVersion < 0x0090U || hciVersion > 0x0120U)
        {
            detail = "unexpected xHCI version 0x" + hciVersion.ToString("X4", CultureInfo.InvariantCulture);
            return false;
        }

        uint hcsparamsValue;
        if (!TryRead32(registers, capabilityAddress + hcsparamsOffset, out hcsparamsValue, out ioError))
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
        if (!TryRead32(registers, capabilityAddress + 0x10, out hccparamsValue, out ioError))
        {
            detail = "failed to read xHCI HCCPARAMS1: " + ioError;
            return false;
        }

        uint contextSize = ((hccparamsValue >> 2) & 0x1U) != 0 ? 64U : 32U;
        if ((ulong)capLength + 0x38UL > memoryLength)
        {
            detail = "xHCI operational registers exceed PCI memory resource";
            return false;
        }
        ulong operationalAddress = capabilityAddress + capLength;
        ulong dcbaap;
        if (!TryRead64(registers, operationalAddress + 0x30, out dcbaap, out ioError))
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

        topology.MaxSlots = maxSlots;
        topology.ContextSize = contextSize;
        for (uint slot = 1; slot <= maxSlots; slot++)
        {
            ulong deviceContext;
            if (!TryRead64(registers, dcbaap + ((ulong)slot * 8UL), out deviceContext, out ioError))
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
            if (!TryRead32(registers, deviceContext, out slotDword0, out ioError)
                || !TryRead32(registers, deviceContext + 0x04, out slotDword1, out ioError)
                || !TryRead32(registers, deviceContext + 0x08, out slotDword2, out ioError)
                || !TryRead32(registers, deviceContext + 0x0C, out slotDword3, out ioError))
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

            uint endpointInterrupter;
            if (TryReadEndpointInterrupterTarget(registers, deviceContext, contextSize, contextEntries, maxIntrs, out endpointInterrupter))
            {
                interrupter = endpointInterrupter;
                topology.EndpointTargetCount++;
            }
            else
            {
                topology.SlotTargetCount++;
            }

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
            + ", intrs=" + maxIntrs
            + ", ctx=" + contextSize
            + ", byAddress=" + topology.ByDeviceAddress.Count
            + ", byRootPort=" + topology.ByRootPort.Count
            + ", endpointTargets=" + topology.EndpointTargetCount
            + ", slotTargets=" + topology.SlotTargetCount
            + ", " + FormatInterrupterMap("addr", topology.ByDeviceAddress)
            + ", " + FormatInterrupterMap("rootPort", topology.ByRootPort);
        return true;
    }

    private static bool TryReadEndpointInterrupterTarget(
        IDeviceRegisters registers,
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
            if (!TryRead32(registers, endpointContext, out epDword0, out ioError)
                || !TryRead32(registers, endpointContext + 0x04, out epDword1, out ioError))
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
            if (!TryRead64(registers, endpointContext + 0x08, out transferRing, out ioError))
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
            if (!TryRead32(registers, transferRing + 0x08, out trbDword2, out ioError)
                || !TryRead32(registers, transferRing + 0x0C, out trbDword3, out ioError))
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

    internal static string FormatInterrupterMap(string name, Dictionary<uint, List<uint>> map)
    {
        if (map.Count == 0)
        {
            return name + "Map=-";
        }

        const int maxShown = 16;
        List<uint> keys = new List<uint>(map.Keys);
        keys.Sort();
        List<string> parts = new List<string>();
        foreach (uint key in keys)
        {
            if (parts.Count >= maxShown)
            {
                break;
            }

            List<uint> interrupters = new List<uint>(map[key]);
            interrupters.Sort();
            List<string> shown = new List<string>();
            foreach (uint interrupter in interrupters)
            {
                shown.Add("I" + interrupter);
            }

            parts.Add(key + ":" + (shown.Count == 0 ? "-" : string.Join("/", shown.ToArray())));
        }

        string suffix = map.Count > maxShown ? ", +" + (map.Count - maxShown) + " more" : string.Empty;
        return name + "Map=[" + string.Join(", ", parts.ToArray()) + suffix + "]";
    }

    internal sealed class XhciInterrupterTopology
    {
        public readonly Dictionary<uint, List<uint>> ByRootPort = new Dictionary<uint, List<uint>>();
        public readonly Dictionary<uint, List<uint>> ByDeviceAddress = new Dictionary<uint, List<uint>>();
        public int EndpointTargetCount;
        public int SlotTargetCount;
        public uint MaxSlots;
        public uint ContextSize;
    }

    internal static readonly string[] AdaptiveRolePriority = new[] { "Mouse", "Keyboard", "Audio", "Gamepad", "Webcam" };

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

    private static bool TryGetDeviceMemoryRange(uint devInst, out ulong baseAddress, out ulong length, out string error)
    {
        baseAddress = 0;
        length = 0;
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
                            ulong candidateLength;
                            if (TryExtractMemoryRangeFromResource(resType, buffer, out candidate, out candidateLength))
                            {
                                if (candidateLength > length)
                                {
                                    baseAddress = candidate;
                                    length = candidateLength;
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

            if (baseAddress == 0 || length == 0)
            {
                error = "no memory resource found";
                return false;
            }

            return true;
        }
        finally
        {
            CM_Free_Log_Conf_Handle(logConf);
        }
    }

    private static bool TryExtractMemoryRangeFromResource(uint resType, byte[] data, out ulong baseAddress, out ulong length)
    {
        baseAddress = 0;
        length = 0;
        if (resType == ResTypeMem)
        {
            if (data.Length < Marshal.SizeOf(typeof(MemDes)))
            {
                return false;
            }
            MemDes mem = BytesToStruct<MemDes>(data, 0);
            ulong candidate = mem.MD_Alloc_Base;
            ulong candidateEnd = mem.MD_Alloc_End;
            if (candidate == 0 && mem.MD_Count > 0)
            {
                int offset = Marshal.SizeOf(typeof(MemDes));
                if (data.Length >= offset + Marshal.SizeOf(typeof(MemRange)))
                {
                    MemRange range = BytesToStruct<MemRange>(data, offset);
                    candidate = range.MR_Min;
                    candidateEnd = range.MR_nBytes > 0 ? candidate + range.MR_nBytes - 1UL : range.MR_Max;
                }
            }
            if (candidate == 0)
            {
                return false;
            }
            baseAddress = candidate;
            length = candidateEnd >= candidate ? candidateEnd - candidate + 1UL : 0;
            if (length == 0)
            {
                return false;
            }
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
            ulong candidateEnd = mem.MLD_Alloc_End;
            if (candidate == 0 && mem.MLD_Count > 0)
            {
                int offset = Marshal.SizeOf(typeof(MemLargeDes));
                if (data.Length >= offset + Marshal.SizeOf(typeof(MemLargeRange)))
                {
                    MemLargeRange range = BytesToStruct<MemLargeRange>(data, offset);
                    candidate = range.MLR_Min;
                    candidateEnd = range.MLR_nBytes > 0 ? candidate + range.MLR_nBytes - 1UL : range.MLR_Max;
                }
            }
            if (candidate == 0)
            {
                return false;
            }
            baseAddress = candidate;
            length = candidateEnd >= candidate ? candidateEnd - candidate + 1UL : 0;
            if (length == 0)
            {
                return false;
            }
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

    [StructLayout(LayoutKind.Sequential)]
    private struct DevPropKey
    {
        public Guid fmtid;
        public uint pid;
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
