using System.Runtime.InteropServices;
using System.Threading;
using static DeviceTweakerCS.ImodDriverNative;

namespace DeviceTweakerCS;

internal delegate bool ImodDriverLoader(string path, Action<string>? log, out string? error);

internal sealed class ImodDriverContext : IDisposable
{
    internal const int StartupScriptVersion = 31;
    public IntPtr DriverHandle { get; private set; } = InvalidHandleValue;
    public bool ServiceCreated { get; private set; }
    public bool ServiceStartedByContext { get; private set; }
    public bool InitializedSuccessfully { get; private set; }
    public string DriverPath { get; }
    // Device whose memory ranges DTIMOD accepts for writes.
    public DeviceTweakerImodRuntime.DeviceWriteTarget? WriteTarget { get; private set; }
    // Every DTIMOD read and write goes through here, so a read-only QA process stops
    // before the driver. Without a target the shared runtime refuses writes; reads need none.
    public IDeviceRegisters Registers
    {
        get
        {
            QaRun.DemandWrite("PUBLISHED PHYSICAL ACCESS");
            return new DeviceTweakerImodRuntime.DtimodRegisters(DriverHandle, WriteTarget!);
        }
    }
    private readonly Action<string>? _log;
    private ImodDriverLoader? _loadDriver;
    private int _disposed;

    public bool OlderDriverLoaded => _olderDriverLoaded ??= DeviceTweakerImodRuntime.IsOlderDriverLoaded(DriverHandle);
    private bool? _olderDriverLoaded;

    public bool TrySetWriteTarget(string instanceId, out string? error)
    {
        WriteTarget = null;
        if (OlderDriverLoaded)
        {
            error = DeviceTweakerImodRuntime.OlderDriverMessage;
            return false;
        }

        if (!DeviceTweakerImodRuntime.DeviceWriteTarget.TryResolve(instanceId, out DeviceTweakerImodRuntime.DeviceWriteTarget target, out error))
        {
            return false;
        }

        WriteTarget = target;
        return true;
    }

    private ImodDriverContext(string driverPath, Action<string>? log)
    {
        DriverPath = driverPath;
        _log = log;
    }

    public static bool TryInitialize(string driverPath, Action<string>? log, ImodDriverLoader loadDriver, out ImodDriverContext? ctx, out string? error)
    {
        if (QaRun.TryBlock("PUBLISHED DRIVER INITIALIZE")) { ctx = null; error = "QA read-only session"; return false; }
        DriverRename.Prepare(DriverRename.Native(driverPath), false);
        log?.Invoke($"IMOD.RUNTIME.PROFILE: startup={StartupScriptVersion} loader=1.4.5 irq-source=pci-existing-device-or-windows");
        ctx = new ImodDriverContext(driverPath, log) { _loadDriver = loadDriver };
        if (!ctx.Initialize(out error))
        {
            ctx.Dispose();
            ctx = null;
            return false;
        }

        return true;
    }

    public static bool TryOpenExisting(Action<string>? log, out ImodDriverContext? ctx, out string? error)
    {
        ctx = null;
        if (QaRun.TryBlock("IMOD DRIVER READ")) { error = "QA read-only session: driver access skipped"; return false; }
        IntPtr handle = InvalidHandleValue;
        string? openError = null;
        DriverReadyResult state = DriverReadyPolicy.Ensure(false,
            () => TryOpenImodDriverDeviceOnce(out handle, out openError),
            () => throw new InvalidOperationException("Readback requested a loader"),
            () => throw new InvalidOperationException("Readback requested KDU"),
            () => throw new InvalidOperationException("Readback requested SCM"));
        if (state == DriverReadyResult.Unavailable) { error = openError; return false; }
        if (!DriverRuntime.QueryInfo(handle, out string info)) { CloseHandle(handle); error = info; return false; }
        ctx = new ImodDriverContext(ImodDriverPaths.SystemPath(), log) { DriverHandle = handle };
        error = null;
        return true;
    }

    private bool Initialize(out string? error)
    {
        error = null;
        IntPtr deviceHandle;

        if (!EnsureImodDriverService(out error))
        {
            string loadError = error ?? "failed to load IMOD driver";
            _log?.Invoke($"IMOD.DRIVER: load failed path={DriverPath} error={loadError}");
            return false;
        }

        if (!TryOpenImodDriverDevice(out deviceHandle, out error))
        {
            return false;
        }

        DriverHandle = deviceHandle;
        InitializedSuccessfully = true;
        return true;
    }

    private bool EnsureImodDriverService(out string? error)
    {
        error = null;
        _log?.Invoke("IMOD.DRIVER: primary loader=kdu");

        if (TryOpenImodDriverDeviceOnce(out IntPtr existingHandle, out string? openError))
        {
            _ = CloseHandle(existingHandle);
            _log?.Invoke("IMOD.DRIVER: device already available");
            _log?.Invoke("IMOD.DRIVER.SUMMARY: final=existing loader=already-loaded");
            return true;
        }

        _log?.Invoke($"IMOD.DRIVER.PRELOAD: deviceAvailable=false loaderNext=kdu detail=\"{SanitizeLogValue(openError)}\"");
        if (_loadDriver!(DriverPath, _log, out string? kduError))
        {
            _log?.Invoke("IMOD.DRIVER.SUMMARY: final=kdu loader=kdu");
            return true;
        }

        DriverLoadGate.AssertLoaderIdle();
        _log?.Invoke($"IMOD.DRIVER.KDU.WARN: load failed; falling back to service: {kduError}");
        if (TryEnsureImodDriverService(out string? serviceError))
        {
            _log?.Invoke($"IMOD.DRIVER.SUMMARY: final=service loader=service kdu_error={kduError ?? "none"}");
            return true;
        }

        error = $"kdu={kduError ?? "unknown"}; service={serviceError ?? "unknown"}";
        _log?.Invoke($"IMOD.DRIVER: service load failed: {serviceError}");
        _log?.Invoke($"IMOD.DRIVER.SUMMARY: final=failed loader=kdu,service error={error}");
        return false;
    }

    private bool TryEnsureImodDriverService(out string? error)
    {
        QaRun.DemandWrite("SCM FALLBACK");
        error = null;
        IntPtr scm = OpenSCManager(null, null, ScManagerAllAccess);
        if (scm == IntPtr.Zero)
        {
            error = $"failed to open service manager: {GetWin32ErrorMessage(Marshal.GetLastWin32Error())}";
            return false;
        }

        IntPtr service = IntPtr.Zero;
        try
        {
            service = OpenService(scm, ImodDriverServiceName, ServiceAllAccess);
            if (service == IntPtr.Zero)
            {
                int lastError = Marshal.GetLastWin32Error();
                if (lastError != ErrorServiceDoesNotExist)
                {
                    error = $"failed to open IMOD driver service: {GetWin32ErrorMessage(lastError)}";
                    return false;
                }

                service = CreateService(
                    scm,
                    ImodDriverServiceName,
                    ImodDriverServiceName,
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
                    error = $"failed to create IMOD driver service: {GetWin32ErrorMessage(Marshal.GetLastWin32Error())}";
                    return false;
                }
                ServiceCreated = true;
            }
            else if (!ChangeServiceConfig(
                         service,
                         ServiceKernelDriver,
                         ServiceDemandStart,
                         ServiceErrorNormal,
                         DriverPath,
                         null,
                         IntPtr.Zero,
                         null,
                         null,
                         null,
                         ImodDriverServiceName))
            {
                error = $"failed to configure IMOD driver service path: {GetWin32ErrorMessage(Marshal.GetLastWin32Error())}";
                return false;
            }

            if (QueryServiceStatus(service, out SERVICE_STATUS_PROCESS status)
                && status.dwCurrentState == ServiceRunning)
            {
                if (TryOpenImodDriverDeviceOnce(out IntPtr runningHandle, out string? runningOpenError))
                {
                    _ = CloseHandle(runningHandle);
                    _log?.Invoke($"IMOD.DRIVER: service already running {ImodDriverServiceName}");
                    return true;
                }

                _log?.Invoke($"IMOD.DRIVER: service running but device unavailable, restarting service: {runningOpenError}");
                if (!StopService(service, out string? stopError))
                {
                    error = stopError ?? "failed to restart stale IMOD driver service";
                    return false;
                }
            }

            if (!StartService(service, 0, IntPtr.Zero))
            {
                int lastError = Marshal.GetLastWin32Error();
                if (lastError == ErrorServiceAlreadyRunning)
                {
                    _log?.Invoke($"IMOD.DRIVER: service already running {ImodDriverServiceName}");
                    return true;
                }

                error = $"failed to start IMOD driver service: {GetWin32ErrorMessage(lastError)} (code {lastError})";
                _log?.Invoke(
                    "IMOD.DRIVER: service start failed "
                    + $"path={DriverPath} error={GetWin32ErrorMessage(lastError)} code={lastError}");
                return false;
            }

            ServiceStartedByContext = true;
            _log?.Invoke($"IMOD.DRIVER: service started {ImodDriverServiceName}");

        }
        finally
        {
            if (service != IntPtr.Zero) _ = CloseServiceHandle(service);
            _ = CloseServiceHandle(scm);
        }

        return true;
    }

    private static bool StopService(IntPtr service, out string? error)
    {
        error = null;
        if (!QueryServiceStatus(service, out SERVICE_STATUS_PROCESS status)
            || status.dwCurrentState == ServiceStopped)
        {
            return true;
        }

        if (status.dwCurrentState != ServiceStopPending)
        {
            SERVICE_STATUS serviceStatus = new();
            if (!ControlService(service, ServiceControlStop, ref serviceStatus))
            {
                int lastError = Marshal.GetLastWin32Error();
                if (lastError != ErrorServiceNotActive)
                {
                    error = $"failed to stop IMOD driver service: {GetWin32ErrorMessage(lastError)}";
                    return false;
                }
            }
        }

        for (int i = 0; i < 25; i++)
        {
            if (!QueryServiceStatus(service, out SERVICE_STATUS_PROCESS check)
                || check.dwCurrentState == ServiceStopped)
            {
                return true;
            }

            Thread.Sleep(200);
        }

        error = "timed out while stopping IMOD driver service";
        return false;
    }

    private void StopServiceIfNeeded()
    {
        QaRun.DemandWrite("OWNED SERVICE CLEANUP");
        IntPtr scm = OpenSCManager(null, null, ScManagerAllAccess);
        if (scm == IntPtr.Zero)
        {
            return;
        }

        try
        {
            IntPtr service = OpenService(scm, ImodDriverServiceName, ServiceAllAccess);
            if (service == IntPtr.Zero)
            {
                return;
            }

            try
            {
                if (ServiceStartedByContext || ServiceCreated)
                {
                    _ = StopService(service, out _);
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

    private void CleanupDriverArtifacts(string driverPath)
    {
        QaRun.DemandWrite("OWNED ARTIFACT CLEANUP");
        if (string.IsNullOrWhiteSpace(driverPath))
        {
            return;
        }

        if (ImodDriverPaths.IsSystemPath(driverPath))
        {
            return;
        }

        try
        {
            if (File.Exists(driverPath))
            {
                File.Delete(driverPath);
            }
        }
        catch (Exception ex)
        {
            _log?.Invoke($"IMOD.DRIVER.CLEANUP.ERROR: file=\"{driverPath}\" error=\"{ex.Message}\"");
        }

        string? driverDirectory = Path.GetDirectoryName(driverPath);
        if (string.IsNullOrWhiteSpace(driverDirectory))
        {
            return;
        }

        try
        {
            if (Directory.Exists(driverDirectory))
            {
                Directory.Delete(driverDirectory, recursive: true);
            }
        }
        catch (Exception ex)
        {
            _log?.Invoke($"IMOD.DRIVER.CLEANUP.ERROR: directory=\"{driverDirectory}\" error=\"{ex.Message}\"");
        }

        string? stagingRoot = Path.GetDirectoryName(driverDirectory);
        if (string.IsNullOrWhiteSpace(stagingRoot))
        {
            return;
        }

        try
        {
            if (Directory.Exists(stagingRoot))
            {
                Directory.Delete(stagingRoot, recursive: true);
            }
        }
        catch (Exception ex)
        {
            _log?.Invoke($"IMOD.DRIVER.CLEANUP.ERROR: stagingRoot=\"{stagingRoot}\" error=\"{ex.Message}\"");
        }
    }

    public void Dispose()
    {
        if (Interlocked.Exchange(ref _disposed, 1) != 0) return;
        try
        {
            _log?.Invoke($"IMOD.DRIVER.CLEANUP: created={ServiceCreated} started={ServiceStartedByContext} initialized={InitializedSuccessfully}");
        }
        finally
        {
            LegacyServiceCleanup.Run(ServiceCreated, ServiceStartedByContext, InitializedSuccessfully,
                () => {
                    if (DriverHandle == InvalidHandleValue) return;
                    _ = CloseHandle(DriverHandle);
                    DriverHandle = InvalidHandleValue;
                }, StopServiceIfNeeded, () => CleanupDriverArtifacts(DriverPath),
                () => { ServiceCreated = false; ServiceStartedByContext = false; });
        }
    }
    private static string SanitizeLogValue(string? value)
        => string.IsNullOrWhiteSpace(value) ? string.Empty : value.Replace("\r\n", " | ").Replace("\n", " | ").Replace("\r", " | ").Trim().Replace("\"", "'");
}
