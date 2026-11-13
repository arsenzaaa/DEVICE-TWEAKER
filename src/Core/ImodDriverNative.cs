using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Threading;

namespace DeviceTweakerCS;

// Native device/service access shared by driver sessions and the loader.
internal static class ImodDriverNative
{
    [DllImport("advapi32.dll", SetLastError = true)]
    internal static extern bool ControlService(IntPtr service, uint control, ref SERVICE_STATUS serviceStatus);
    [DllImport("advapi32.dll", SetLastError = true)]
    internal static extern bool DeleteService(IntPtr service);

    internal const uint CmProbDisabled = 0x00000016;

    internal const uint DigcfPresent = 0x00000002;

    internal const int ErrorInsufficientBuffer = 122;
    internal const int ErrorNoMoreItems = 259;

    internal const int ErrorServiceDoesNotExist = 1060;
    internal const int ErrorServiceNotActive = 1062;
    internal const int ErrorServiceAlreadyRunning = 1056;

    internal const uint ScManagerAllAccess = 0x000F003F;
    internal const uint ServiceAllAccess = 0x000F01FF;
    internal const uint ServiceKernelDriver = 0x00000001;
    internal const uint ServiceDemandStart = 0x00000003;
    internal const uint ServiceErrorNormal = 0x00000001;
    internal const uint ServiceControlStop = 0x00000001;
    internal const uint ServiceRunning = 0x00000004;
    internal const uint ServiceStopped = 0x00000001;
    internal const uint ServiceStopPending = 0x00000003;
    internal const int ScStatusProcessInfo = 0;

    internal const uint GenericRead = 0x80000000;
    internal const uint GenericWrite = 0x40000000;
    internal const uint FileShareRead = 0x00000001;
    internal const uint FileShareWrite = 0x00000002;
    internal const uint OpenExisting = 3;
    internal const uint FileAttributeNormal = 0x00000080;

    internal const string ImodDriverDevicePath = DriverIdentity.DevicePath;
    internal const string ImodDriverServiceName = DriverIdentity.ServiceName;
    internal const int ImodDriverOpenRetryCount = 10;
    internal const int ImodDriverOpenRetryDelayMs = 100;
    internal const int ImodKduTimeoutMs = 60000;
    internal const string ImodKduFileName = "kdu.exe";
    internal const string ImodKduDatabaseFileName = "drv64.dll";
    internal const string ImodKduDisableEnv = "DEVICE_TWEAKER_DISABLE_KDU_FALLBACK";

    internal sealed class ImodControllerInfo
    {
        public string DeviceId { get; init; } = string.Empty;
        public string Caption { get; init; } = string.Empty;
        public uint ProblemCode { get; init; }
        public ulong BaseAddress { get; init; }
        public ulong MemoryLength { get; init; }
        public bool HasBase { get; init; }
        public string BaseError { get; init; } = string.Empty;
    }

    [DllImport("setupapi.dll", SetLastError = true)]
    internal static extern bool SetupDiDestroyDeviceInfoList(IntPtr deviceInfoSet);

    [DllImport("cfgmgr32.dll", SetLastError = true)]
    internal static extern int CM_Get_DevNode_Status(out uint status, out uint problem, uint devInst, uint flags);

    [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    internal static extern IntPtr OpenSCManager(
        string? machineName,
        string? databaseName,
        uint desiredAccess);

    [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    internal static extern IntPtr OpenService(
        IntPtr scm,
        string serviceName,
        uint desiredAccess);

    [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    internal static extern IntPtr CreateService(
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

    [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    internal static extern bool ChangeServiceConfig(
        IntPtr service,
        uint serviceType,
        uint startType,
        uint errorControl,
        string? binaryPathName,
        string? loadOrderGroup,
        IntPtr tagId,
        string? dependencies,
        string? serviceStartName,
        string? password,
        string? displayName);

    [DllImport("advapi32.dll", SetLastError = true)]
    internal static extern bool StartService(IntPtr service, uint numServiceArgs, IntPtr serviceArgVectors);

    [DllImport("advapi32.dll", SetLastError = true)]
    internal static extern bool QueryServiceStatusEx(
        IntPtr service,
        int infoLevel,
        ref SERVICE_STATUS_PROCESS buffer,
        uint bufferSize,
        out uint bytesNeeded);

    [DllImport("advapi32.dll", SetLastError = true)]
    internal static extern bool CloseServiceHandle(IntPtr handle);

    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    internal static extern IntPtr CreateFile(
        string fileName,
        uint desiredAccess,
        uint shareMode,
        IntPtr securityAttributes,
        uint creationDisposition,
        uint flagsAndAttributes,
        IntPtr templateFile);

    [DllImport("kernel32.dll", SetLastError = true)]
    internal static extern bool CloseHandle(IntPtr handle);

    internal static bool TryOpenImodDriverDevice(out IntPtr handle, out string? error)
    {
        handle = InvalidHandleValue;
        error = null;
        if (QaRun.TryBlock("IMOD DRIVER DEVICE OPEN")) { error = "QA read-only session: driver access skipped"; return false; }
        int lastError = 0;

        for (int attempt = 0; attempt < ImodDriverOpenRetryCount; attempt++)
        {
            handle = CreateFile(
                ImodDriverDevicePath,
                GenericRead | GenericWrite,
                FileShareRead | FileShareWrite,
                IntPtr.Zero,
                OpenExisting,
                FileAttributeNormal,
                IntPtr.Zero);

            if (handle != InvalidHandleValue)
            {
                return true;
            }

            lastError = Marshal.GetLastWin32Error();
            Thread.Sleep(ImodDriverOpenRetryDelayMs);
        }

        error = $"failed to open {ImodDriverDevicePath}: {GetWin32ErrorMessage(lastError)}";
        return false;
    }

    internal static bool IsImodDriverAlreadyAvailable()
    {
        if (!TryOpenImodDriverDeviceOnce(out IntPtr handle, out _))
        {
            return false;
        }

        _ = CloseHandle(handle);
        return true;
    }

    internal static bool TryOpenImodDriverDeviceOnce(out IntPtr handle, out string? error)
    {
        if (QaRun.TryBlock("IMOD DRIVER DEVICE OPEN")) { handle = InvalidHandleValue; error = "QA read-only session: driver access skipped"; return false; }
        handle = CreateFile(
            ImodDriverDevicePath,
            GenericRead | GenericWrite,
            FileShareRead | FileShareWrite,
            IntPtr.Zero,
            OpenExisting,
            FileAttributeNormal,
            IntPtr.Zero);

        if (handle != InvalidHandleValue)
        {
            error = null;
            return true;
        }

        error = FormatImodDeviceOpenError(Marshal.GetLastWin32Error());
        return false;
    }

    internal static string FormatImodDeviceOpenError(int errorCode)
    {
        string detail = $"failed to open {ImodDriverDevicePath} (Win32 {errorCode}): {GetWin32ErrorMessage(errorCode)}";
        return errorCode is 2 or 3 ? $"driver not loaded; press CHECK. {detail}" : detail;
    }

    internal static bool QueryServiceStatus(IntPtr service, out SERVICE_STATUS_PROCESS status)
    {
        status = new SERVICE_STATUS_PROCESS();
        return QueryServiceStatusEx(service, ScStatusProcessInfo, ref status, (uint)Marshal.SizeOf<SERVICE_STATUS_PROCESS>(), out _);
    }

    internal static string GetWin32ErrorMessage(int error)
    {
        return new Win32Exception(error).Message;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct SERVICE_STATUS_PROCESS
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
    internal struct SERVICE_STATUS
    {
        public uint dwServiceType;
        public uint dwCurrentState;
        public uint dwControlsAccepted;
        public uint dwWin32ExitCode;
        public uint dwServiceSpecificExitCode;
        public uint dwCheckPoint;
        public uint dwWaitHint;
    }
    internal static readonly IntPtr InvalidHandleValue = new(-1);
}
