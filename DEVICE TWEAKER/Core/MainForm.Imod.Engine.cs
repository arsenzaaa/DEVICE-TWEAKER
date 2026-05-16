using System.ComponentModel;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Security.Cryptography;
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
    private const int ErrorServiceNotActive = 1062;
    private const int ErrorServiceAlreadyRunning = 1056;

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

    private const string ImodDriverDevicePath = "\\\\.\\DeviceTweakerImod2";
    private const string ImodDriverServiceName = "DeviceTweakerImod2";
    private const int ImodDriverOpenRetryCount = 10;
    private const int ImodDriverOpenRetryDelayMs = 100;
    private const int ImodKduTimeoutMs = 60000;
    private const string ImodKduFileName = "kdu.exe";
    private const string ImodKduDatabaseFileName = "drv64.dll";
    private const string ImodKduDisableEnv = "DEVICE_TWEAKER_DISABLE_KDU_FALLBACK";
    private static readonly uint IoctlImodMapPhysicalMemory =
        CtlCode(FileDeviceImod, ImodIoctlIndex, MethodBuffered, FileAnyAccess);
    private static readonly uint IoctlImodUnmapPhysicalMemory =
        CtlCode(FileDeviceImod, ImodIoctlIndex + 1, MethodBuffered, FileAnyAccess);
    private static readonly uint IoctlImodReadPhysicalMemory =
        CtlCode(FileDeviceImod, ImodIoctlIndex + 2, MethodBuffered, FileAnyAccess);
    private static readonly uint IoctlImodWritePhysicalMemory =
        CtlCode(FileDeviceImod, ImodIoctlIndex + 3, MethodBuffered, FileAnyAccess);

    private sealed class ImodControllerInfo
    {
        public string DeviceId { get; init; } = string.Empty;
        public string Caption { get; init; } = string.Empty;
        public uint ProblemCode { get; init; }
        public ulong BaseAddress { get; init; }
        public bool HasBase { get; init; }
        public string BaseError { get; init; } = string.Empty;
    }

    private sealed class XhciInterrupterTopology
    {
        public Dictionary<uint, List<uint>> ByRootPort { get; } = [];
        public Dictionary<uint, List<uint>> ByDeviceAddress { get; } = [];
        public int EndpointTargetCount { get; set; }
        public int SlotTargetCount { get; set; }
        public uint MaxSlots { get; set; }
        public uint ContextSize { get; set; }
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

    [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool ChangeServiceConfig(
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
        ref PhysAccessStruct inBuffer,
        int inBufferSize,
        ref PhysAccessStruct outBuffer,
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

        if (!EnsureImodDriverOnDisk(persistDriver, out string driverPath, out error))
        {
            return false;
        }

        bool cleanupDriver = false;
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

        if (!ImodDriverContext.TryInitialize(driverPath, WriteLog, out ImodDriverContext? driverContext, out error))
        {
            if (IsImodKernelCiBlockedLoadError(error))
            {
                string ciState = GetImodKernelCiBlockState();
                error = $"{error} (kernel CI blocked; {ciState})";
            }
            LogImodDriverLoadDiagnostics(driverPath, error);
            return false;
        }

            ClearImodKernelCiBlockStatus();
            ImodDriverContext imodDriver = driverContext!;
            using (imodDriver)
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
                    List<uint>? desiredIntervals = null;
                    ImodConfigEntry? adaptiveEntry = null;
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
                                desiredIntervals = null;
                            }
                            if (entry.Intervals is { Count: > 0 })
                            {
                                desiredIntervals = entry.Intervals;
                            }
                            if (entry.AdaptiveRoleBinding == true && entry.RoleIntervals is { Count: > 0 })
                            {
                                adaptiveEntry = entry;
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

                    if (!TryReadPhys32(imodDriver, capabilityAddress + hcsparamsOffset, out uint hcsparamsValue, out string? ioError))
                    {
                        stats.ReadFailures++;
                        WriteLog($"IMOD: read HCSPARAMS failed {controller.DeviceId}: {ioError}");
                        continue;
                    }

                    if (!TryReadPhys32(imodDriver, capabilityAddress + rtsoff, out uint rtsoffValue, out ioError))
                    {
                        stats.ReadFailures++;
                        WriteLog($"IMOD: read RTSOFF failed {controller.DeviceId}: {ioError}");
                        continue;
                    }

                    uint maxIntrs = (hcsparamsValue >> 8) & 0x7FF;
                    ulong runtimeAddress = capabilityAddress + rtsoffValue;

                    if (adaptiveEntry?.RoleIntervals is { Count: > 0 } roleIntervals)
                    {
                        if (TryBuildAdaptiveImodIntervals(
                                controller,
                                imodDriver,
                                maxIntrs,
                                roleIntervals,
                                desiredInterval,
                                out List<uint> adaptiveIntervals,
                                out string adaptiveDetail))
                        {
                            desiredIntervals = adaptiveIntervals;
                            adaptiveEntry.Intervals = adaptiveIntervals;
                            WriteLog($"IMOD.ADAPTIVE.APPLY: {controller.DeviceId} {adaptiveDetail}");
                        }
                        else
                        {
                            WriteLog($"IMOD.ADAPTIVE.FALLBACK: {controller.DeviceId} {adaptiveDetail}");
                        }
                    }

                    uint writeFailures = 0;
                    uint writeCount = desiredIntervals is { Count: > 0 }
                        ? Math.Min(maxIntrs, (uint)desiredIntervals.Count)
                        : maxIntrs;
                    for (uint i = 0; i < writeCount; ++i)
                    {
                        ulong interrupterAddress = runtimeAddress + 0x24 + (0x20 * i);
                        uint targetInterval = desiredIntervals is { Count: > 0 }
                            ? desiredIntervals[(int)i]
                            : desiredInterval;
                        if (!TryWriteImodInterval(imodDriver, interrupterAddress, targetInterval, out ioError))
                        {
                            writeFailures++;
                            WriteLog($"IMOD: write failed {controller.DeviceId} @ {ToHex(interrupterAddress)}: {ioError}");
                        }
                    }

                    stats.ControllersApplied++;
                    stats.WriteFailures += (int)writeFailures;

                    string modeText = desiredIntervals is { Count: > 0 }
                        ? $"vector={desiredIntervals.Count}"
                        : $"interval={FormatImodValue(desiredInterval)}";
                    WriteLog($"IMOD: {controller.DeviceId} writes={writeCount}/{maxIntrs} {modeText} failures={writeFailures}");
                }
            }
        }
        finally
        {
            if (cleanupDriver && !IsImodDriverSystemPath(driverPath))
            {
                DeleteFileIfExists(driverPath, "IMOD.DRIVER");
            }
        }

        return true;
    }

    private void RefreshImodCurrentValues(bool showReadingStatus = true, string reason = "refresh")
    {
        _ = RefreshImodCurrentValuesAsync(showReadingStatus, reason);
    }

    private async Task RefreshImodCurrentValuesAsync(bool showReadingStatus = true, string reason = "refresh")
    {
        List<DeviceBlock> targetBlocks = _blocks
            .Where(b => IsUsbImodTarget(b.Device) && !b.Device.IsTestDevice)
            .ToList();
        int generation = Interlocked.Increment(ref _imodReadbackGeneration);

        if (targetBlocks.Count == 0)
        {
            WriteLog($"IMOD.READBACK.START: reason={reason} targets=0");
            return;
        }

        WriteLog($"IMOD.READBACK.START: reason={reason} targets={targetBlocks.Count} showReading={showReadingStatus}");

        if (TryGetCachedImodKernelCiBlockStatus(out string blockedStatus, out string blockedDetail))
        {
            foreach (DeviceBlock block in targetBlocks)
            {
                block.ImodCurrentLabel.Text = blockedStatus;
                block.ImodCurrentLabel.Tag = $"{blockedStatus}\r\n{blockedDetail}";
                block.ImodCurrentLabel.ForeColor = _statusDanger;
                SetImodStatusTooltip(block.ImodCurrentLabel);
                if (block.ImodMapLabel is not null)
                {
                    block.ImodMapLabel.Text = blockedDetail;
                    block.ImodMapLabel.Tag = block.ImodMapLabel.Text;
                    block.ImodMapLabel.ForeColor = _statusDanger;
                    SetImodStatusTooltip(block.ImodMapLabel);
                }
                SetImodDetailsVisibility(block, forceVisible: true);
                LogImodUiSnapshot(block, reason);
            }

            return;
        }

        if (showReadingStatus)
        {
            foreach (DeviceBlock block in targetBlocks)
            {
                block.ImodCurrentLabel.Text = "current: reading...";
                block.ImodCurrentLabel.Tag = block.ImodCurrentLabel.Text;
                block.ImodCurrentLabel.ForeColor = _statusInactive;
                if (block.ImodMapLabel is not null)
                {
                    block.ImodMapLabel.Text = "devices: reading...";
                    block.ImodMapLabel.Tag = block.ImodMapLabel.Text;
                    block.ImodMapLabel.ForeColor = _mutedText;
                    SetImodStatusTooltip(block.ImodMapLabel);
                }
            }
        }

        EnsureImodConfigLoaded();
        ImodConfig config = _imodConfigCache ?? new ImodConfig();
        (bool ok, Dictionary<string, List<uint>> valuesByDeviceId, Dictionary<string, string> mapByDeviceId, Dictionary<string, string> mapDetailByDeviceId, string? error) readback =
            await Task.Run(() =>
            {
                bool ok = TryReadCurrentImodValues(
                    config,
                    out Dictionary<string, List<uint>> values,
                    out Dictionary<string, string> map,
                    out Dictionary<string, string> mapDetail,
                    out string? error);
                return (ok, values, map, mapDetail, error);
            });

        if (IsDisposed || generation != Volatile.Read(ref _imodReadbackGeneration))
        {
            WriteLog($"IMOD.READBACK.STALE: reason={reason} generation={generation}");
            return;
        }

        if (!readback.ok)
        {
            CacheImodKernelCiBlockStatus(readback.error);
            string statusText = FormatImodReadbackStatus(readback.error);
            string detailText = IsImodAttentionStatus(statusText)
                ? FormatImodReadbackDetail(readback.error)
                : "devices: unavailable";
            bool isBlockedStatus =
                statusText.Contains("admin required", StringComparison.OrdinalIgnoreCase)
                || statusText.Contains("driver blocked", StringComparison.OrdinalIgnoreCase)
                || statusText.Contains("kernel CI blocked", StringComparison.OrdinalIgnoreCase)
                || statusText.Contains("signature blocked", StringComparison.OrdinalIgnoreCase)
                || statusText.Contains("driver not loaded", StringComparison.OrdinalIgnoreCase);
            foreach (DeviceBlock block in targetBlocks)
            {
                block.ImodCurrentLabel.Text = statusText;
                block.ImodCurrentLabel.Tag = isBlockedStatus
                    ? $"{statusText}\r\n{detailText}"
                    : block.ImodCurrentLabel.Text;
                block.ImodCurrentLabel.ForeColor = isBlockedStatus ? _statusDanger : _statusInactive;
                SetImodStatusTooltip(block.ImodCurrentLabel);
                if (block.ImodMapLabel is not null)
                {
                    block.ImodMapLabel.Text = detailText;
                    block.ImodMapLabel.Tag = block.ImodMapLabel.Text;
                    block.ImodMapLabel.ForeColor = isBlockedStatus ? _statusDanger : _mutedText;
                    SetImodStatusTooltip(block.ImodMapLabel);
                }
                SetImodDetailsVisibility(block, forceVisible: isBlockedStatus);
                LogImodUiSnapshot(block, reason);
            }

            WriteLog($"IMOD.READBACK: failed: {readback.error}");
            return;
        }

        foreach (DeviceBlock block in targetBlocks)
        {
            string key = NormalizeInstanceId(block.Device.InstanceId);
            if (!readback.valuesByDeviceId.TryGetValue(key, out List<uint>? values) || values.Count == 0)
            {
                block.ImodCurrentLabel.Text = "current: unavailable";
                block.ImodCurrentLabel.Tag = block.ImodCurrentLabel.Text;
                block.ImodCurrentLabel.ForeColor = _statusInactive;
                SetImodStatusTooltip(block.ImodCurrentLabel);
                if (block.ImodMapLabel is not null)
                {
                    block.ImodMapLabel.Text = "devices: unavailable";
                    block.ImodMapLabel.Tag = block.ImodMapLabel.Text;
                    block.ImodMapLabel.ForeColor = _mutedText;
                    SetImodStatusTooltip(block.ImodMapLabel);
                }
                SetImodDetailsVisibility(block, forceVisible: false);
                LogImodUiSnapshot(block, reason);
                WriteLog($"IMOD.READBACK: no values for {block.Device.InstanceId}");
                continue;
            }

            block.ImodCurrentLabel.Text = $"current: {FormatImodValueList(values)}";
            block.ImodCurrentLabel.Tag = $"current raw: {FormatImodValueListForLog(values)}";
            block.ImodCurrentLabel.ForeColor = _statusActive;
            SetImodStatusTooltip(block.ImodCurrentLabel);
            if (block.ImodMapLabel is not null)
            {
                if (readback.mapByDeviceId.TryGetValue(key, out string? mapText))
                {
                    block.ImodMapLabel.Text = mapText;
                    block.ImodMapLabel.Tag = readback.mapDetailByDeviceId.TryGetValue(key, out string? mapDetail)
                        ? mapDetail
                        : block.ImodMapLabel.Text;
                    block.ImodMapLabel.ForeColor = _mutedText;
                }
                else
                {
                    block.ImodMapLabel.Text = "devices: unavailable";
                    block.ImodMapLabel.Tag = block.ImodMapLabel.Text;
                    block.ImodMapLabel.ForeColor = _mutedText;
                }

                SetImodStatusTooltip(block.ImodMapLabel);
            }
            SetImodDetailsVisibility(block, forceVisible: false);
            LogImodUiSnapshot(block, reason, values, block.ImodMapLabel?.Tag as string);
            WriteLog($"IMOD.READBACK: {block.Device.InstanceId} values={FormatImodValueListForLog(values)}");
        }
    }

    private void LogImodUiSnapshot(DeviceBlock block, string reason, IReadOnlyList<uint>? values = null, string? mapDetail = null)
    {
        string value = block.ImodBox.Text?.Trim() ?? string.Empty;
        string defaultText = FlattenLogText(block.ImodDefaultLabel.Text);
        string currentText = FlattenLogText(block.ImodCurrentLabel.Text);
        string devicesAndInterrupts = FlattenLogText(block.ImodMapLabel?.Text);
        string detail = FlattenLogText(mapDetail ?? (block.ImodMapLabel?.Tag as string));
        string baseRoles = block.Device.UsbRoles ?? string.Empty;
        string effectiveRoles = GetEffectiveUsbRolesForController(block.Device, block.Device.InstanceId);
        string rawValues = values is { Count: > 0 }
            ? FormatImodValueListForLog(values)
            : string.Empty;

        WriteLog(
            "IMOD.UI: "
            + $"reason={reason} "
            + $"controller={block.Device.InstanceId} "
            + $"input=\"{SanitizeLogValue(value)}\" "
            + $"default=\"{SanitizeLogValue(defaultText)}\" "
            + $"current=\"{SanitizeLogValue(currentText)}\" "
            + $"baseRoles=\"{SanitizeLogValue(baseRoles)}\" "
            + $"effectiveRoles=\"{SanitizeLogValue(effectiveRoles)}\" "
            + $"map=\"{SanitizeLogValue(devicesAndInterrupts)}\" "
            + $"raw=\"{SanitizeLogValue(rawValues)}\" "
            + $"detail=\"{SanitizeLogValue(detail)}\"");
    }

    private void SetImodStatusTooltip(Control label)
    {
        try
        {
            if (label is TextBox { Multiline: true } or RichTextBox { Multiline: true })
            {
                _copyToolTip.SetToolTip(label, string.Empty);
                return;
            }

            string text;
            text = label.Tag as string ?? label.Text;
            _copyToolTip.SetToolTip(label, text);
        }
        catch
        {
        }
    }

    private static bool IsImodAttentionStatus(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        return text.Contains("driver blocked", StringComparison.OrdinalIgnoreCase)
            || text.Contains("kernel CI blocked", StringComparison.OrdinalIgnoreCase)
            || text.Contains("signature blocked", StringComparison.OrdinalIgnoreCase)
            || text.Contains("driver not loaded", StringComparison.OrdinalIgnoreCase)
            || text.Contains("admin required", StringComparison.OrdinalIgnoreCase);
    }

    private void SetImodDetailsVisibility(DeviceBlock block, bool forceVisible)
    {
        bool showDetails = block.ImodAutoCheck.Visible
            && (block.ImodAutoCheck.Checked || forceVisible || IsImodAttentionStatus(block.ImodCurrentLabel.Text));

        block.ImodCurrentLabel.Visible = showDetails;
        block.ImodDefaultLabel.Visible = showDetails;
        if (block.ImodMapLabel is not null)
        {
            block.ImodMapLabel.Visible = showDetails;
        }

        RelayoutDeviceBlockChrome(block);
        LayoutBlocks();
    }

    private void RelayoutDeviceBlockChrome(DeviceBlock block)
    {
        if (block.Group.IsDisposed)
        {
            return;
        }

        if (block.RelayoutAction is not null)
        {
            block.RelayoutAction();
            return;
        }

        Panel? settingsPanel = block.ImodCurrentLabel.Parent as Panel;
        Control? cpuPanel = block.CpuBoxes.Count > 0 ? block.CpuBoxes[0].Parent : null;
        if (settingsPanel is null || cpuPanel is null)
        {
            return;
        }

        int maxCpuPanelWidth = Math.Max(UiScale(220), block.Group.Width - cpuPanel.Left - UiScale(24));
        if (cpuPanel.Width > maxCpuPanelWidth)
        {
            cpuPanel.Width = maxCpuPanelWidth;
        }

        int visibleSettingsRight = 0;
        int visibleSettingsBottom = 0;
        foreach (Control child in settingsPanel.Controls)
        {
            if (!child.Visible)
            {
                continue;
            }

            visibleSettingsRight = Math.Max(visibleSettingsRight, child.Right);
            visibleSettingsBottom = Math.Max(visibleSettingsBottom, child.Bottom);
        }

        bool stackedSettings = settingsPanel.Left <= UiScale(24);
        if (!stackedSettings && block.Group.Width - settingsPanel.Left - UiScale(24) < UiScale(320))
        {
            stackedSettings = true;
            settingsPanel.Left = UiScale(18);
        }

        int availableSettingsWidth = Math.Max(UiScale(120), block.Group.Width - settingsPanel.Left - UiScale(24));
        int settingsMinWidth = Math.Min(UiScale(420), availableSettingsWidth);
        Size currentSettingsSize = new(
            Math.Min(Math.Max(visibleSettingsRight + UiScale(8), settingsMinWidth), availableSettingsWidth),
            Math.Max(visibleSettingsBottom + UiScale(8), UiScale(24)));
        settingsPanel.Size = currentSettingsSize;

        int settingsTop = stackedSettings
            ? block.IrqLabel.Bottom + UiScale(18)
            : cpuPanel.Top + Math.Max(0, (cpuPanel.Height - currentSettingsSize.Height) / 2);
        settingsPanel.Location = new Point(settingsPanel.Left, settingsTop);

        int infoY = Math.Max(block.IrqLabel.Bottom + UiScale(14), cpuPanel.Bottom + UiScale(18));
        infoY = Math.Max(infoY, settingsPanel.Bottom + UiScale(10));
        block.InfoLabel.Location = new Point(block.InfoLabel.Left, infoY);
        block.InfoLabel.Size = new Size(block.Group.Width - UiScale(40), block.InfoLabel.Height);

        block.Group.Height = Math.Max(
            Math.Max(cpuPanel.Bottom + UiScale(110), settingsPanel.Bottom + UiScale(20)),
            block.InfoLabel.Bottom + UiScale(20));
    }

    private bool TryGetCachedImodKernelCiBlockStatus(out string statusText, out string detailText)
    {
        statusText = _imodKernelCiBlockStatus ?? string.Empty;
        detailText = _imodKernelCiBlockDetail ?? string.Empty;
        if (string.IsNullOrWhiteSpace(statusText))
        {
            return false;
        }

        if ((DateTime.UtcNow - _imodKernelCiBlockStatusUtc) > TimeSpan.FromMinutes(5))
        {
            ClearImodKernelCiBlockStatus();
            statusText = string.Empty;
            detailText = string.Empty;
            return false;
        }

        if (string.IsNullOrWhiteSpace(detailText))
        {
            detailText = FormatImodReadbackDetail(statusText);
        }

        return true;
    }

    private void CacheImodKernelCiBlockStatus(string? error)
    {
        if (!IsImodSignatureRejectedError(error))
        {
            return;
        }

        _imodKernelCiBlockStatus = FormatImodKernelBlockStatus();
        _imodKernelCiBlockDetail = FormatImodReadbackDetail(error);
        _imodKernelCiBlockStatusUtc = DateTime.UtcNow;
    }

    private void ClearImodKernelCiBlockStatus()
    {
        _imodKernelCiBlockStatus = null;
        _imodKernelCiBlockDetail = null;
        _imodKernelCiBlockStatusUtc = default;
    }

    private bool TryReadCurrentImodValues(
        ImodConfig config,
        out Dictionary<string, List<uint>> valuesByDeviceId,
        out Dictionary<string, string> mapByDeviceId,
        out Dictionary<string, string> mapDetailByDeviceId,
        out string? error)
    {
        const uint readbackLimit = 64;

        valuesByDeviceId = new Dictionary<string, List<uint>>(StringComparer.OrdinalIgnoreCase);
        mapByDeviceId = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        mapDetailByDeviceId = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        error = null;

        if (!IsAdministrator())
        {
            error = "administrator privileges required";
            return false;
        }

        bool persistDriver = config.HasScript || HasCustomImod(config);
        if (!EnsureImodDriverOnDisk(persistDriver, out string driverPath, out error))
        {
            return false;
        }

        bool cleanupDriver = false;
        try
        {
            if (!TryEnumerateXhciControllers(out List<ImodControllerInfo> controllers, out error))
            {
                return false;
            }

            if (controllers.Count == 0)
            {
                return true;
            }

        if (!ImodDriverContext.TryInitialize(driverPath, WriteLog, out ImodDriverContext? driverContext, out error))
            {
                LogImodDriverLoadDiagnostics(driverPath, error);
                return false;
            }

            ClearImodKernelCiBlockStatus();
            using ImodDriverContext imodDriver = driverContext!;
            foreach (ImodControllerInfo controller in controllers)
            {
                if (controller.ProblemCode == CmProbDisabled || !controller.HasBase)
                {
                    continue;
                }

                uint hcsparamsOffset = config.GlobalHcsparamsOffset;
                uint rtsoff = config.GlobalRtsoff;

                foreach (ImodConfigEntry entry in config.Overrides)
                {
                    if (string.IsNullOrWhiteSpace(entry.Hwid)
                        || controller.DeviceId.IndexOf(entry.Hwid, StringComparison.OrdinalIgnoreCase) < 0)
                    {
                        continue;
                    }

                    if (entry.HcsparamsOffset.HasValue)
                    {
                        hcsparamsOffset = entry.HcsparamsOffset.Value;
                    }
                    if (entry.Rtsoff.HasValue)
                    {
                        rtsoff = entry.Rtsoff.Value;
                    }
                }

                ulong capabilityAddress = controller.BaseAddress;
                if (!TryReadPhys32(imodDriver, capabilityAddress + hcsparamsOffset, out uint hcsparamsValue, out string? ioError))
                {
                    WriteLog($"IMOD.READBACK: read HCSPARAMS failed {controller.DeviceId}: {ioError}");
                    continue;
                }

                if (!TryReadPhys32(imodDriver, capabilityAddress + rtsoff, out uint rtsoffValue, out ioError))
                {
                    WriteLog($"IMOD.READBACK: read RTSOFF failed {controller.DeviceId}: {ioError}");
                    continue;
                }

                uint maxIntrs = (hcsparamsValue >> 8) & 0x7FF;
                uint readCount = Math.Min(maxIntrs, readbackLimit);
                ulong runtimeAddress = capabilityAddress + rtsoffValue;
                List<uint> values = [];

                for (uint i = 0; i < readCount; i++)
                {
                    ulong interrupterAddress = runtimeAddress + 0x24 + (0x20 * i);
                    if (!TryReadPhys32(imodDriver, interrupterAddress, out uint registerValue, out ioError))
                    {
                        WriteLog($"IMOD.READBACK: read failed {controller.DeviceId} @ {ToHex(interrupterAddress)}: {ioError}");
                        continue;
                    }

                    values.Add(registerValue & 0xFFFF);
                }

                if (values.Count > 0)
                {
                    string normalizedControllerId = NormalizeInstanceId(controller.DeviceId);
                    valuesByDeviceId[normalizedControllerId] = values;
                    if (TryFormatImodInterrupterRoleMap(controller, imodDriver, maxIntrs, values, out string mapText, out string mapDetail))
                    {
                        mapByDeviceId[normalizedControllerId] = mapText;
                        mapDetailByDeviceId[normalizedControllerId] = mapDetail;
                        WriteLog($"IMOD.MAP: {controller.DeviceId} {mapDetail}");
                    }
                    else
                    {
                        WriteLog($"IMOD.MAP: unavailable {controller.DeviceId}: {mapDetail}");
                    }

                    if (maxIntrs > readbackLimit)
                    {
                        WriteLog($"IMOD.READBACK: {controller.DeviceId} read first {readbackLimit} of {maxIntrs} interrupters");
                    }
                }
            }
        }
        finally
        {
            if (cleanupDriver && !IsImodDriverSystemPath(driverPath))
            {
                DeleteFileIfExists(driverPath, "IMOD.DRIVER");
            }
        }

        return true;
    }

    private static string FormatImodValueList(IReadOnlyList<uint> values)
    {
        if (values.Count == 0)
        {
            return "-";
        }

        uint first = values[0];
        if (values.All(v => v == first))
        {
            return values.Count > 1
                ? $"{FormatImodValue(first)} x{values.Count}"
                : FormatImodValue(first);
        }

        return "varies by interrupter";
    }

    private static string FormatImodValueListForLog(IReadOnlyList<uint> values)
    {
        if (values.Count == 0)
        {
            return "-";
        }

        return string.Join(",", values.Select(value => $"0x{value & 0xFFFF:X4}"));
    }

    private static string FormatImodReadbackStatus(string? error)
    {
        if (string.IsNullOrWhiteSpace(error))
        {
            return "current: unavailable";
        }

        if (IsImodSignatureRejectedError(error))
        {
            return FormatImodKernelBlockStatus();
        }

        if (error.Contains("signature", StringComparison.OrdinalIgnoreCase)
            || error.Contains("digital", StringComparison.OrdinalIgnoreCase)
            || error.Contains("подпис", StringComparison.OrdinalIgnoreCase)
            || error.Contains("цифров", StringComparison.OrdinalIgnoreCase))
        {
            return FormatImodKernelBlockStatus();
        }

        if (error.Contains("administrator", StringComparison.OrdinalIgnoreCase)
            || error.Contains("админист", StringComparison.OrdinalIgnoreCase))
        {
            return "current: admin required";
        }

        return "current: unavailable";
    }

    private static string FormatImodKernelBlockStatus()
    {
        return "current: driver blocked";
    }

    private static string FormatImodReadbackDetail(string? error)
    {
        if (string.IsNullOrWhiteSpace(error))
        {
            return "devices: unavailable";
        }

        if (!IsImodSignatureRejectedError(error))
        {
            return "devices: unavailable";
        }

        if (error.Contains("virus", StringComparison.OrdinalIgnoreCase)
            || error.Contains("potentially unwanted", StringComparison.OrdinalIgnoreCase)
            || error.Contains("PUA", StringComparison.OrdinalIgnoreCase))
        {
            return "DTIMOD.sys blocked by Windows Defender";
        }

        if (error.Contains("KDU", StringComparison.OrdinalIgnoreCase))
        {
            return "DTIMOD.sys loader blocked";
        }

        return "DTIMOD.sys blocked by Windows";
    }

    private bool TryBuildAdaptiveImodIntervals(
        ImodControllerInfo controller,
        ImodDriverContext imodDriver,
        uint maxIntrs,
        IReadOnlyDictionary<string, uint> roleIntervals,
        uint fallbackInterval,
        out List<uint> intervals,
        out string detail)
    {
        intervals = [];
        detail = "no adaptive role map";

        if (maxIntrs == 0)
        {
            detail = "controller reports zero interrupters";
            return false;
        }

        if (!TryResolveAdaptiveRolesByInterrupter(
                controller,
                imodDriver,
                maxIntrs,
                out Dictionary<uint, HashSet<string>> rolesByInterrupter,
                out detail,
                allowFallback: false))
        {
            return false;
        }

        int count = (int)Math.Min(maxIntrs, 2048);
        uint fallback = fallbackInterval & 0xFFFF;
        for (int i = 0; i < count; i++)
        {
            uint target = fallback;
            if (rolesByInterrupter.TryGetValue((uint)i, out HashSet<string>? roles)
                && TrySelectAdaptiveRoleInterval(roles, roleIntervals, out _, out uint roleValue))
            {
                target = roleValue;
            }

            intervals.Add(target & 0xFFFF);
        }

        List<string> shown = [];
        List<string> shared = [];
        foreach (KeyValuePair<uint, HashSet<string>> pair in rolesByInterrupter.OrderBy(kvp => kvp.Key))
        {
            bool hasSelectedRole = TrySelectAdaptiveRoleInterval(pair.Value, roleIntervals, out string selectedRole, out uint selectedValue);
            if (!hasSelectedRole)
            {
                selectedRole = string.Join("+", pair.Value.OrderBy(x => x, StringComparer.OrdinalIgnoreCase));
                selectedValue = fallback;
            }
            else if (pair.Value.Count > 1)
            {
                List<string> configuredSharedRoles = pair.Value
                    .Where(role => roleIntervals.ContainsKey(role))
                    .OrderBy(role => role, StringComparer.OrdinalIgnoreCase)
                    .ToList();
                List<string> losingRoles = configuredSharedRoles
                    .Where(role => !string.Equals(role, selectedRole, StringComparison.OrdinalIgnoreCase))
                    .ToList();
                if (losingRoles.Count > 0)
                {
                    shared.Add($"intr{pair.Key}: {selectedRole} wins over {string.Join("+", losingRoles)}");
                }
            }

            shown.Add($"intr{pair.Key}={selectedRole}:{FormatImodValue(selectedValue)}");
            if (shown.Count >= 8)
            {
                break;
            }
        }

        string suffix = rolesByInterrupter.Count > shown.Count ? $" +{rolesByInterrupter.Count - shown.Count} more" : string.Empty;
        detail = $"roles=[{string.Join(", ", shown)}{suffix}]";
        if (shared.Count > 0)
        {
            string sharedSuffix = shared.Count > 4 ? $" +{shared.Count - 4} more" : string.Empty;
            detail += $"; shared=[{string.Join(", ", shared.Take(4))}{sharedSuffix}]";
        }

        return true;
    }

    private bool TryFormatImodInterrupterRoleMap(
        ImodControllerInfo controller,
        ImodDriverContext imodDriver,
        uint maxIntrs,
        IReadOnlyList<uint> currentIntervals,
        out string mapText,
        out string detail)
    {
        mapText = string.Empty;
        detail = string.Empty;
        if (currentIntervals.Count == 0)
        {
            detail = "no IMOD values were read";
            return false;
        }

        Dictionary<uint, HashSet<string>> rolesByInterrupter = [];
        string resolveDetail = string.Empty;
        bool hasRoleMap = TryResolveAdaptiveRolesByInterrupter(
            controller,
            imodDriver,
            maxIntrs,
            out rolesByInterrupter,
            out resolveDetail,
            allowFallback: false);
        if (!hasRoleMap)
        {
            WriteLog($"IMOD.MAP.ROLES: unavailable {controller.DeviceId}: {resolveDetail}");
        }

        Dictionary<string, List<string>> displayLabelsByRole = BuildAdaptiveRoleDisplayLabels(controller.DeviceId);
        Dictionary<uint, List<string>> displayLabelsByInterrupter = [];
        HashSet<string> mappedLabelsSeen = new(StringComparer.OrdinalIgnoreCase);

        if (hasRoleMap)
        {
            foreach (KeyValuePair<uint, HashSet<string>> pair in rolesByInterrupter.OrderBy(static kvp => kvp.Key))
            {
                if (pair.Key >= currentIntervals.Count || pair.Value.Count == 0)
                {
                    continue;
                }

                List<string> mappedLabels = [];
                HashSet<string> localLabelsSeen = new(StringComparer.OrdinalIgnoreCase);
                foreach (string label in FormatAdaptiveDisplayRoles(pair.Value, displayLabelsByRole))
                {
                    string cleanLabel = FormatImodDeviceMapLabel(label);
                    if (localLabelsSeen.Add(cleanLabel))
                    {
                        mappedLabels.Add(cleanLabel);
                        mappedLabelsSeen.Add(cleanLabel);
                    }
                }

                if (mappedLabels.Count > 0)
                {
                    displayLabelsByInterrupter[pair.Key] = mappedLabels;
                }
            }
        }

        string unknownDeviceValue = FormatUnknownImodDeviceValue(currentIntervals);
        List<string> unknownLabels = [];
        HashSet<string> unknownLabelsSeen = new(StringComparer.OrdinalIgnoreCase);
        foreach (string label in BuildControllerImodDeviceDisplayLabels(controller.DeviceId))
        {
            string cleanLabel = FormatImodDeviceMapLabel(label);
            if (!mappedLabelsSeen.Contains(cleanLabel) && unknownLabelsSeen.Add(cleanLabel))
            {
                unknownLabels.Add(cleanLabel);
            }
        }

        string mappedDeviceLine = FormatMappedImodDeviceLine(currentIntervals, displayLabelsByInterrupter, unknownLabels, unknownDeviceValue);
        string visibleInterruptLines = FormatVisibleImodInterrupterLines(currentIntervals);
        string detailedInterruptLine = FormatDetailedImodInterrupterMap(currentIntervals, displayLabelsByInterrupter);
        string rawLine = FormatRawImodInterrupterMap(currentIntervals);
        string timeLine = FormatTimeImodInterrupterMap(currentIntervals);

        mapText = $"{mappedDeviceLine}\r\n\r\n{visibleInterruptLines}";
        detail = string.IsNullOrWhiteSpace(resolveDetail)
            ? $"{mappedDeviceLine}\r\n\r\n{visibleInterruptLines}\r\n{detailedInterruptLine}\r\n{rawLine}\r\n{timeLine}"
            : $"{mappedDeviceLine}\r\n\r\n{visibleInterruptLines}\r\n{detailedInterruptLine}\r\n{rawLine}\r\n{timeLine}\r\nmap: {resolveDetail}";
        return true;
    }

    private List<string> BuildControllerImodDeviceDisplayLabels(string controllerDeviceId)
    {
        Dictionary<string, List<string>> labelsByRole = BuildAdaptiveRoleDisplayLabels(controllerDeviceId);
        List<string> labels = [];
        HashSet<string> emitted = new(StringComparer.OrdinalIgnoreCase);

        foreach (string role in AdaptiveRolePriority)
        {
            if (!labelsByRole.TryGetValue(role, out List<string>? roleLabels))
            {
                continue;
            }

            foreach (string label in roleLabels)
            {
                if (emitted.Add(label))
                {
                    labels.Add(label);
                }
            }
        }

        foreach (KeyValuePair<string, List<string>> pair in labelsByRole.OrderBy(static kvp => kvp.Key, StringComparer.OrdinalIgnoreCase))
        {
            foreach (string label in pair.Value)
            {
                if (emitted.Add(label))
                {
                    labels.Add(label);
                }
            }
        }

        return labels;
    }

    private static string FormatRawImodInterrupterMap(IReadOnlyList<uint> values)
    {
        if (values.Count == 0)
        {
            return "raw: -";
        }

        return "raw: " + string.Join(" | ", values.Select((value, index) => $"I{index}:{FormatImodValue(value)}"));
    }

    private static string FormatMappedImodDeviceLine(
        IReadOnlyList<uint> values,
        IReadOnlyDictionary<uint, List<string>> labelsByInterrupter,
        IReadOnlyList<string> unknownLabels,
        string unknownDeviceValue)
    {
        List<string> parts = [];
        foreach (KeyValuePair<uint, List<string>> pair in labelsByInterrupter.OrderBy(static kvp => kvp.Key))
        {
            if (pair.Key >= values.Count || pair.Value.Count == 0)
            {
                continue;
            }

            uint value = values[(int)pair.Key];
            foreach (string label in pair.Value)
            {
                parts.Add($"{label} -> intr{pair.Key}={FormatImodValue(value)}/{FormatUsbImodTime(value)}");
            }
        }

        if (unknownLabels.Count > 0)
        {
            foreach (string label in unknownLabels)
            {
                parts.Add($"{label} -> intr?={unknownDeviceValue}");
            }
        }

        return parts.Count > 0
            ? FormatDeviceFirstImodLines(parts)
            : "devices: no exact USB role map";
    }

    private static string FormatDeviceFirstImodLines(IReadOnlyList<string> parts)
    {
        const int itemsPerLine = 2;
        List<string> lines = [];
        for (int i = 0; i < parts.Count; i += itemsPerLine)
        {
            string prefix = i == 0 ? "devices: " : "         ";
            lines.Add(prefix + string.Join(" | ", parts.Skip(i).Take(itemsPerLine)));
        }

        return string.Join("\r\n", lines);
    }

    private static string FormatVisibleImodInterrupterLines(IReadOnlyList<uint> values)
    {
        const int chunkSize = 4;
        const int maxVisible = 8;

        if (values.Count == 0)
        {
            return "interrupters: -";
        }

        int visibleCount = Math.Min(values.Count, maxVisible);
        List<string> lines = [];

        for (int start = 0; start < visibleCount; start += chunkSize)
        {
            int end = Math.Min(start + chunkSize, visibleCount);
            List<string> parts = [];
            for (int i = start; i < end; i++)
            {
                parts.Add($"intr{i}={FormatImodValue(values[i])}/{FormatUsbImodTime(values[i])}");
            }

            string suffix = string.Empty;
            if (end == visibleCount && values.Count > visibleCount)
            {
                suffix = $" | +{values.Count - visibleCount} more";
            }

            lines.Add($"interrupters {start}-{end - 1}: {string.Join(" | ", parts)}{suffix}");
        }

        return string.Join("\r\n", lines);
    }

    private static string FormatDetailedImodInterrupterMap(
        IReadOnlyList<uint> values,
        IReadOnlyDictionary<uint, List<string>> labelsByInterrupter)
    {
        if (values.Count == 0)
        {
            return "detail: -";
        }

        List<string> parts = [];
        for (int i = 0; i < values.Count; i++)
        {
            string labelText = labelsByInterrupter.TryGetValue((uint)i, out List<string>? labels) && labels.Count > 0
                ? $" {string.Join("/", labels)}"
                : string.Empty;
            parts.Add($"I{i}{labelText}={FormatImodValue(values[i])} ({FormatUsbImodTime(values[i])})");
        }

        return "detail: " + string.Join(" | ", parts);
    }

    private static string FormatTimeImodInterrupterMap(IReadOnlyList<uint> values)
    {
        if (values.Count == 0)
        {
            return "time: -";
        }

        return "time: " + string.Join(" | ", values.Select((value, index) => $"I{index}:{FormatUsbImodTime(value)}"));
    }

    private static string FormatUnknownImodDeviceValue(IReadOnlyList<uint> values)
    {
        if (values.Count == 0)
        {
            return "-";
        }

        uint first = values[0];
        return values.All(value => value == first)
            ? FormatImodValue(first)
            : "mixed";
    }

    private static string FormatImodDeviceMapLabel(string label)
    {
        string clean = string.IsNullOrWhiteSpace(label) ? "Device" : label.Trim();
        clean = clean.Replace(" | ", "/", StringComparison.Ordinal);
        clean = clean.Replace("=", "-", StringComparison.Ordinal);
        return clean;
    }

    private Dictionary<string, List<string>> BuildAdaptiveRoleDisplayLabels(string controllerDeviceId)
    {
        Dictionary<string, List<string>> labelsByRole = new(StringComparer.OrdinalIgnoreCase);
        foreach (DeviceBlock block in _blocks)
        {
            DeviceInfo device = block.Device;
            if (device.Kind != DeviceKind.USB || device.IsTestDevice)
            {
                continue;
            }

            if (!HostControllerPathMatchesDeviceId(device.InstanceId, controllerDeviceId) &&
                !HostControllerPathMatchesDeviceId(controllerDeviceId, device.InstanceId))
            {
                continue;
            }

            AddAdaptiveRoleDisplayLabels(GetEffectiveUsbRolesForController(device, controllerDeviceId), labelsByRole);
            AddAdaptiveRoleDisplayLabels(device.AudioEndpoints, labelsByRole);
        }

        return labelsByRole;
    }

    private string GetEffectiveUsbRolesForController(DeviceInfo device, string controllerDeviceId)
    {
        string controllerKey = NormalizeInstanceId(controllerDeviceId);
        if (!string.IsNullOrWhiteSpace(controllerKey)
            && _usbRoleOverrideByController.TryGetValue(controllerKey, out string? controllerRoles)
            && !string.IsNullOrWhiteSpace(controllerRoles))
        {
            return MergeUsbRolePollingOverrideWithCurrentRoles(device.UsbRoles, controllerRoles);
        }

        string deviceKey = NormalizeInstanceId(device.InstanceId);
        if (!string.IsNullOrWhiteSpace(deviceKey)
            && _usbRoleOverrideByController.TryGetValue(deviceKey, out string? deviceRoles)
            && !string.IsNullOrWhiteSpace(deviceRoles))
        {
            return MergeUsbRolePollingOverrideWithCurrentRoles(device.UsbRoles, deviceRoles);
        }

        return device.UsbRoles;
    }

    private static string MergeUsbRolePollingOverrideWithCurrentRoles(string currentRoles, string overrideRoles)
    {
        if (string.IsNullOrWhiteSpace(currentRoles) || string.IsNullOrWhiteSpace(overrideRoles))
        {
            return currentRoles;
        }

        Dictionary<string, string> pollingOverrides = new(StringComparer.OrdinalIgnoreCase);
        foreach (string rawEntry in overrideRoles.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            string entry = rawEntry.Trim();
            foreach (string role in new[] { "Mouse", "Keyboard" })
            {
                if (!IsUsbRoleText(entry, role))
                {
                    continue;
                }

                string tag = entry[role.Length..].Trim();
                if (tag.Length > 0)
                {
                    pollingOverrides[role] = tag;
                }
            }
        }

        return pollingOverrides.Count > 0
            ? ApplyUsbRolePollingOverrides(currentRoles, pollingOverrides)
            : currentRoles;
    }

    private static IEnumerable<string> FormatAdaptiveDisplayRoles(
        IEnumerable<string> roles,
        IReadOnlyDictionary<string, List<string>> labelsByRole)
    {
        HashSet<string> emitted = new(StringComparer.OrdinalIgnoreCase);
        foreach (string role in OrderAdaptiveRoles(roles))
        {
            if (labelsByRole.TryGetValue(role, out List<string>? labels) && labels.Count > 0)
            {
                foreach (string label in labels)
                {
                    if (emitted.Add(label))
                    {
                        yield return label;
                    }
                }
            }
            else if (emitted.Add(role))
            {
                yield return role;
            }
        }
    }

    private static void AddAdaptiveRoleDisplayLabels(string? text, Dictionary<string, List<string>> labelsByRole)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return;
        }

        foreach (string rawPart in text.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            string label = rawPart.Trim();
            if (label.Length == 0 || !TryGetAdaptiveRoleBase(label, out string role))
            {
                continue;
            }

            if (!labelsByRole.TryGetValue(role, out List<string>? labels))
            {
                labels = [];
                labelsByRole[role] = labels;
            }

            if (!labels.Contains(label, StringComparer.OrdinalIgnoreCase))
            {
                labels.Add(label);
            }
        }
    }

    private static bool TryGetAdaptiveRoleBase(string text, out string role)
    {
        role = string.Empty;
        if (text.Contains("Mouse", StringComparison.OrdinalIgnoreCase))
        {
            role = "Mouse";
            return true;
        }

        if (text.Contains("Keyboard", StringComparison.OrdinalIgnoreCase))
        {
            role = "Keyboard";
            return true;
        }

        if (text.Contains("Audio", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("Microphone", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("Speaker", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("Headphone", StringComparison.OrdinalIgnoreCase))
        {
            role = "Audio";
            return true;
        }

        if (text.Contains("Gamepad", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("Joystick", StringComparison.OrdinalIgnoreCase))
        {
            role = "Gamepad";
            return true;
        }

        if (text.Contains("Camera", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("Webcam", StringComparison.OrdinalIgnoreCase))
        {
            role = "Webcam";
            return true;
        }

        return false;
    }

    private static string FormatUsbImodTime(uint rawValue)
    {
        ulong ns = (rawValue & 0xFFFF) * 250UL;
        if (ns >= 1000)
        {
            double us = ns / 1000.0;
            return us.ToString(us % 1 == 0 ? "0" : "0.###", System.Globalization.CultureInfo.InvariantCulture) + "us";
        }

        return ns.ToString(System.Globalization.CultureInfo.InvariantCulture) + " ns";
    }

    private static IEnumerable<string> OrderAdaptiveRoles(IEnumerable<string> roles)
    {
        return roles
            .OrderBy(static role => AdaptiveRolePriorityIndex(role))
            .ThenBy(static role => role, StringComparer.OrdinalIgnoreCase);
    }

    private static int AdaptiveRolePriorityIndex(string role)
    {
        for (int i = 0; i < AdaptiveRolePriority.Length; i++)
        {
            if (AdaptiveRolePriority[i].Equals(role, StringComparison.OrdinalIgnoreCase))
            {
                return i;
            }
        }

        return AdaptiveRolePriority.Length;
    }

    private bool TryResolveAdaptiveRolesByInterrupter(
        ImodControllerInfo controller,
        ImodDriverContext imodDriver,
        uint maxIntrs,
        out Dictionary<uint, HashSet<string>> rolesByInterrupter,
        out string detail,
        bool allowFallback = true)
    {
        rolesByInterrupter = [];

        if (!TryBuildExactAdaptiveRolesByInterrupter(
                controller,
                imodDriver,
                maxIntrs,
                out rolesByInterrupter,
                out Dictionary<uint, List<uint>> interruptersByRootPort,
                out detail))
        {
            if (!allowFallback)
            {
                return false;
            }

            if (!TryReadXhciRootPortInterrupters(controller, imodDriver, maxIntrs, out interruptersByRootPort, out string fallbackReadDetail))
            {
                detail = $"{detail}; {fallbackReadDetail}";
                return false;
            }
        }
        else
        {
            return true;
        }

        HashSet<string> fallbackRoles = BuildAdaptiveControllerRoleFallback(controller.DeviceId);
        if (fallbackRoles.Count == 0)
        {
            detail = "no exact USB role/interrupter map and no controller role fallback";
            return false;
        }

        if (!TryAssignFallbackRolesToActiveInterrupters(fallbackRoles, interruptersByRootPort, maxIntrs, rolesByInterrupter, out detail))
        {
            return false;
        }

        detail = "fallback controller roles assigned to active interrupters: " + detail;
        return true;
    }

    private bool TryBuildExactAdaptiveRolesByInterrupter(
        ImodControllerInfo controller,
        ImodDriverContext imodDriver,
        uint maxIntrs,
        out Dictionary<uint, HashSet<string>> rolesByInterrupter,
        out Dictionary<uint, List<uint>> interruptersByRootPort,
        out string detail)
    {
        rolesByInterrupter = [];
        interruptersByRootPort = [];
        detail = "no exact USB role/interrupter map";

        if (!TryReadXhciInterrupterTopology(controller, imodDriver, maxIntrs, out XhciInterrupterTopology topology, out string topologyDetail))
        {
            detail = topologyDetail;
            return false;
        }

        interruptersByRootPort = topology.ByRootPort;

        List<UsbEndpointInfo> endpoints;
        try
        {
            endpoints = UsbTopologyInterop.EnumerateEndpoints();
        }
        catch (Exception ex)
        {
            detail = $"USB endpoint enumeration failed: {ex.Message}; {topologyDetail}";
            return false;
        }

        int addressMatches = 0;
        int rootPortMatches = 0;
        int classifiedEndpoints = 0;
        int filteredGamepadEndpoints = 0;
        HashSet<string> controllerRoles = BuildAdaptiveControllerRoleFallback(controller.DeviceId);
        foreach (UsbEndpointInfo endpoint in endpoints)
        {
            if (!HostControllerPathMatchesDeviceId(endpoint.HostControllerPath, controller.DeviceId)
                || !TryClassifyAdaptiveUsbRole(endpoint, out string role))
            {
                continue;
            }

            if (role.Equals("Gamepad", StringComparison.OrdinalIgnoreCase)
                && !controllerRoles.Contains("Gamepad"))
            {
                filteredGamepadEndpoints++;
                continue;
            }

            classifiedEndpoints++;
            bool mapped = false;
            if (endpoint.DeviceAddress > 0
                && topology.ByDeviceAddress.TryGetValue((uint)endpoint.DeviceAddress, out List<uint>? addressIntrs))
            {
                AddAdaptiveRoleToInterrupters(rolesByInterrupter, addressIntrs, role);
                addressMatches++;
                mapped = true;
            }

            if (!mapped
                && TryGetUsbRootPort(endpoint.TopologyPath, out uint rootPort)
                && topology.ByRootPort.TryGetValue(rootPort, out List<uint>? rootIntrs))
            {
                AddAdaptiveRoleToInterrupters(rolesByInterrupter, rootIntrs, role);
                rootPortMatches++;
            }
        }

        if (rolesByInterrupter.Count == 0)
        {
            detail = $"no USB endpoints matched active xHCI interrupters; classified={classifiedEndpoints}; {topologyDetail}";
            return false;
        }

        detail =
            $"exact xHCI map: addressMatches={addressMatches}, rootPortMatches={rootPortMatches}, endpointTargets={topology.EndpointTargetCount}, slotTargets={topology.SlotTargetCount}, classified={classifiedEndpoints}, filteredGenericHid={filteredGamepadEndpoints}; {topologyDetail}";
        return true;
    }

    private static void AddAdaptiveRoleToInterrupters(
        Dictionary<uint, HashSet<string>> rolesByInterrupter,
        IEnumerable<uint> interrupters,
        string role)
    {
        foreach (uint intr in interrupters)
        {
            if (!rolesByInterrupter.TryGetValue(intr, out HashSet<string>? roles))
            {
                roles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                rolesByInterrupter[intr] = roles;
            }

            roles.Add(role);
        }
    }

    private static Dictionary<uint, HashSet<string>> BuildAdaptiveRootPortRoles(string controllerDeviceId)
    {
        Dictionary<uint, HashSet<string>> rolesByRootPort = [];
        List<UsbEndpointInfo> endpoints;
        try
        {
            endpoints = UsbTopologyInterop.EnumerateEndpoints();
        }
        catch
        {
            return rolesByRootPort;
        }

        foreach (UsbEndpointInfo endpoint in endpoints)
        {
            if (!HostControllerPathMatchesDeviceId(endpoint.HostControllerPath, controllerDeviceId)
                || !TryGetUsbRootPort(endpoint.TopologyPath, out uint rootPort)
                || !TryClassifyAdaptiveUsbRole(endpoint, out string role))
            {
                continue;
            }

            if (!rolesByRootPort.TryGetValue(rootPort, out HashSet<string>? roles))
            {
                roles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                rolesByRootPort[rootPort] = roles;
            }

            roles.Add(role);
        }

        return rolesByRootPort;
    }

    private HashSet<string> BuildAdaptiveControllerRoleFallback(string controllerDeviceId)
    {
        HashSet<string> roles = new(StringComparer.OrdinalIgnoreCase);
        foreach (DeviceBlock block in _blocks)
        {
            DeviceInfo device = block.Device;
            if (device.Kind != DeviceKind.USB || device.IsTestDevice)
            {
                continue;
            }

            if (!HostControllerPathMatchesDeviceId(device.InstanceId, controllerDeviceId) &&
                !HostControllerPathMatchesDeviceId(controllerDeviceId, device.InstanceId))
            {
                continue;
            }

            AddAdaptiveRolesFromText(GetEffectiveUsbRolesForController(device, controllerDeviceId), roles);
            AddAdaptiveRolesFromText(device.AudioEndpoints, roles);
        }

        return roles;
    }

    private static bool TryAssignFallbackRolesToActiveInterrupters(
        HashSet<string> fallbackRoles,
        Dictionary<uint, List<uint>> interruptersByRootPort,
        uint maxIntrs,
        Dictionary<uint, HashSet<string>> rolesByInterrupter,
        out string detail)
    {
        List<uint> activeInterrupters = interruptersByRootPort
            .Values
            .SelectMany(static intrs => intrs)
            .Where(intr => intr < maxIntrs)
            .Distinct()
            .OrderBy(static intr => intr)
            .ToList();

        if (activeInterrupters.Count == 0)
        {
            detail = "no active xHCI interrupters were discovered";
            return false;
        }

        List<string> orderedRoles = AdaptiveRolePriority
            .Where(role => fallbackRoles.Contains(role))
            .ToList();

        foreach (string role in fallbackRoles.OrderBy(static role => role, StringComparer.OrdinalIgnoreCase))
        {
            if (!orderedRoles.Contains(role, StringComparer.OrdinalIgnoreCase))
            {
                orderedRoles.Add(role);
            }
        }

        for (int i = 0; i < orderedRoles.Count; i++)
        {
            uint intr = activeInterrupters[Math.Min(i, activeInterrupters.Count - 1)];
            if (!rolesByInterrupter.TryGetValue(intr, out HashSet<string>? roles))
            {
                roles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                rolesByInterrupter[intr] = roles;
            }

            roles.Add(orderedRoles[i]);
        }

        detail = string.Join(", ", rolesByInterrupter
            .OrderBy(static kvp => kvp.Key)
            .Select(kvp => $"intr{kvp.Key}={string.Join("+", kvp.Value.OrderBy(static role => role, StringComparer.OrdinalIgnoreCase))}"));
        return rolesByInterrupter.Count > 0;
    }

    private static void AddAdaptiveRolesFromText(string? text, HashSet<string> roles)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return;
        }

        if (text.Contains("Mouse", StringComparison.OrdinalIgnoreCase))
        {
            roles.Add("Mouse");
        }

        if (text.Contains("Keyboard", StringComparison.OrdinalIgnoreCase))
        {
            roles.Add("Keyboard");
        }

        if (text.Contains("Audio", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("Microphone", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("Speaker", StringComparison.OrdinalIgnoreCase))
        {
            roles.Add("Audio");
        }

        if (text.Contains("Gamepad", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("Joystick", StringComparison.OrdinalIgnoreCase))
        {
            roles.Add("Gamepad");
        }

        if (text.Contains("Camera", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("Webcam", StringComparison.OrdinalIgnoreCase))
        {
            roles.Add("Webcam");
        }
    }

    private static bool TryReadXhciRootPortInterrupters(
        ImodControllerInfo controller,
        ImodDriverContext imodDriver,
        uint maxIntrs,
        out Dictionary<uint, List<uint>> interruptersByRootPort,
        out string detail)
    {
        interruptersByRootPort = [];
        if (!TryReadXhciInterrupterTopology(controller, imodDriver, maxIntrs, out XhciInterrupterTopology topology, out detail))
        {
            return false;
        }

        interruptersByRootPort = topology.ByRootPort;
        if (interruptersByRootPort.Count == 0)
        {
            detail = "no active xHCI slot root-port interrupter map was found";
            return false;
        }

        detail = $"ok; endpointTargets={topology.EndpointTargetCount}, slotTargets={topology.SlotTargetCount}";
        return true;
    }

    private static bool TryReadXhciInterrupterTopology(
        ImodControllerInfo controller,
        ImodDriverContext imodDriver,
        uint maxIntrs,
        out XhciInterrupterTopology topology,
        out string detail)
    {
        topology = new XhciInterrupterTopology();
        ulong capabilityAddress = controller.BaseAddress;

        if (!TryReadPhys32(imodDriver, capabilityAddress, out uint capReg, out string? ioError))
        {
            detail = $"failed to read xHCI CAPLENGTH: {ioError}";
            return false;
        }

        uint capLength = capReg & 0xFF;
        if (capLength == 0)
        {
            detail = "xHCI CAPLENGTH is zero";
            return false;
        }

        if (!TryReadPhys32(imodDriver, capabilityAddress + ImodDefaultHcsparamsOffset, out uint hcsparamsValue, out ioError))
        {
            detail = $"failed to read xHCI HCSPARAMS1: {ioError}";
            return false;
        }

        uint maxSlots = hcsparamsValue & 0xFF;
        if (maxSlots == 0)
        {
            detail = "xHCI MaxSlots is zero";
            return false;
        }

        if (!TryReadPhys32(imodDriver, capabilityAddress + 0x10, out uint hccparamsValue, out ioError))
        {
            detail = $"failed to read xHCI HCCPARAMS1: {ioError}";
            return false;
        }

        uint contextSize = ((hccparamsValue >> 2) & 0x1) != 0 ? 64u : 32u;
        ulong operationalAddress = capabilityAddress + capLength;
        if (!TryReadPhys64ForAdaptive(imodDriver, operationalAddress + 0x30, out ulong dcbaap, out ioError))
        {
            detail = $"failed to read xHCI DCBAAP: {ioError}";
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
            if (!TryReadPhys64ForAdaptive(imodDriver, dcbaap + ((ulong)slot * 8), out ulong deviceContext, out _))
            {
                continue;
            }

            deviceContext &= 0xFFFFFFFFFFFFFFC0UL;
            if (deviceContext == 0)
            {
                continue;
            }

            if (!TryReadPhys32(imodDriver, deviceContext, out uint slotDword0, out _)
                || !TryReadPhys32(imodDriver, deviceContext + 0x04, out uint slotDword1, out _)
                || !TryReadPhys32(imodDriver, deviceContext + 0x08, out uint slotDword2, out _)
                || !TryReadPhys32(imodDriver, deviceContext + 0x0C, out uint slotDword3, out _))
            {
                continue;
            }

            bool isHub = ((slotDword0 >> 26) & 0x1) != 0;
            uint contextEntries = (slotDword0 >> 27) & 0x1F;
            uint slotState = (slotDword3 >> 27) & 0x1F;
            uint rootPort = (slotDword1 >> 16) & 0xFF;
            uint deviceAddress = slotDword3 & 0xFF;
            uint interrupter = (slotDword2 >> 22) & 0x3FF;

            if (slotState < 2 || isHub || rootPort == 0)
            {
                continue;
            }

            if (TryReadEndpointInterrupterTarget(
                    imodDriver,
                    deviceContext,
                    contextSize,
                    contextEntries,
                    maxIntrs,
                    out uint endpointInterrupter))
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

        detail =
            $"slots={maxSlots}, ctx={contextSize}, byAddress={topology.ByDeviceAddress.Count}, byRootPort={topology.ByRootPort.Count}, endpointTargets={topology.EndpointTargetCount}, slotTargets={topology.SlotTargetCount}, "
            + FormatXhciTopologyMap("addr", topology.ByDeviceAddress)
            + ", "
            + FormatXhciTopologyMap("rootPort", topology.ByRootPort);
        return true;
    }

    private static string FormatXhciTopologyMap(string name, IReadOnlyDictionary<uint, List<uint>> map)
    {
        if (map.Count == 0)
        {
            return $"{name}Map=-";
        }

        const int maxShown = 16;
        List<string> parts = [];
        foreach (KeyValuePair<uint, List<uint>> pair in map.OrderBy(static kvp => kvp.Key))
        {
            string intrs = pair.Value.Count == 0
                ? "-"
                : string.Join("/", pair.Value.OrderBy(static intr => intr).Select(static intr => $"I{intr}"));
            parts.Add($"{pair.Key}:{intrs}");
            if (parts.Count >= maxShown)
            {
                break;
            }
        }

        string suffix = map.Count > maxShown ? $", +{map.Count - maxShown} more" : string.Empty;
        return $"{name}Map=[{string.Join(", ", parts)}{suffix}]";
    }

    private static bool TryReadEndpointInterrupterTarget(
        ImodDriverContext imodDriver,
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
            if (!TryReadPhys32(imodDriver, endpointContext, out uint epDword0, out _)
                || !TryReadPhys32(imodDriver, endpointContext + 0x04, out uint epDword1, out _))
            {
                continue;
            }

            uint endpointState = epDword0 & 0x7;
            uint endpointType = (epDword1 >> 3) & 0x7;
            if (endpointState == 0 || endpointType is not (3 or 5 or 7))
            {
                continue;
            }

            if (!TryReadPhys64ForAdaptive(imodDriver, endpointContext + 0x08, out ulong transferRing, out _))
            {
                continue;
            }

            transferRing &= 0xFFFFFFFFFFFFFFF0UL;
            if (transferRing == 0)
            {
                continue;
            }

            if (!TryReadPhys32(imodDriver, transferRing + 0x08, out uint trbDword2, out _)
                || !TryReadPhys32(imodDriver, transferRing + 0x0C, out uint trbDword3, out _))
            {
                continue;
            }

            uint trbType = (trbDword3 >> 10) & 0x3F;
            if (trbType is not (1 or 3 or 5))
            {
                continue;
            }

            uint target = (trbDword2 >> 22) & 0x3FF;
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
        if (!map.TryGetValue(key, out List<uint>? intrs))
        {
            intrs = [];
            map[key] = intrs;
        }

        if (!intrs.Contains(interrupter))
        {
            intrs.Add(interrupter);
        }
    }

    private static bool TryReadPhys64ForAdaptive(ImodDriverContext imodDriver, ulong address, out ulong value, out string? error)
    {
        value = 0;
        if (!TryReadPhys32(imodDriver, address, out uint low, out error))
        {
            return false;
        }

        if (!TryReadPhys32(imodDriver, address + 4, out uint high, out error))
        {
            return false;
        }

        value = low | ((ulong)high << 32);
        return true;
    }

    private static bool TrySelectAdaptiveRoleInterval(
        HashSet<string> roles,
        IReadOnlyDictionary<string, uint> roleIntervals,
        out string selectedRole,
        out uint selectedValue)
    {
        selectedRole = string.Empty;
        selectedValue = 0;
        foreach (string role in AdaptiveRolePriority)
        {
            if (roles.Contains(role) && roleIntervals.TryGetValue(role, out selectedValue))
            {
                selectedRole = role;
                return true;
            }
        }

        foreach (string role in roles.OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
        {
            if (roleIntervals.TryGetValue(role, out selectedValue))
            {
                selectedRole = role;
                return true;
            }
        }

        return false;
    }

    private static readonly string[] AdaptiveRolePriority = ["Mouse", "Keyboard", "Audio", "Gamepad", "Webcam"];

    private static bool HostControllerPathMatchesDeviceId(string hostControllerPath, string controllerDeviceId)
    {
        string pathKey = CompactDeviceMatchText(hostControllerPath.Replace('#', '\\'));
        string idKey = CompactDeviceMatchText(controllerDeviceId);
        return pathKey.Length > 0 && idKey.Length > 0 && pathKey.Contains(idKey, StringComparison.OrdinalIgnoreCase);
    }

    private static string CompactDeviceMatchText(string text)
    {
        StringBuilder sb = new(text.Length);
        foreach (char ch in text)
        {
            if (char.IsLetterOrDigit(ch))
            {
                sb.Append(char.ToUpperInvariant(ch));
            }
        }

        return sb.ToString();
    }

    private static bool TryGetUsbRootPort(string topologyPath, out uint rootPort)
    {
        rootPort = 0;
        if (string.IsNullOrWhiteSpace(topologyPath))
        {
            return false;
        }

        foreach (string segment in topologyPath.Split(['/', '\\'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            if (!segment.StartsWith("Port", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            string suffix = segment[4..];
            int end = 0;
            while (end < suffix.Length && char.IsDigit(suffix[end]))
            {
                end++;
            }

            if (end > 0 && uint.TryParse(suffix[..end], out rootPort) && rootPort > 0)
            {
                return true;
            }
        }

        int index = 0;
        while (index >= 0 && index < topologyPath.Length)
        {
            index = topologyPath.IndexOf(".P", index, StringComparison.OrdinalIgnoreCase);
            if (index < 0)
            {
                break;
            }

            int start = index + 2;
            int end = start;
            while (end < topologyPath.Length && char.IsDigit(topologyPath[end]))
            {
                end++;
            }

            if (end > start && uint.TryParse(topologyPath[start..end], out rootPort) && rootPort > 0)
            {
                return true;
            }

            index = start;
        }

        return false;
    }

    private static bool TryClassifyAdaptiveUsbRole(UsbEndpointInfo endpoint, out string role)
    {
        role = string.Empty;
        string cls = endpoint.InterfaceClass.Trim();
        string sub = endpoint.InterfaceSubClass.Trim();
        string proto = endpoint.InterfaceProtocol.Trim();

        if (IsUsbDescriptorValue(cls, "03") || cls.Contains("HID", StringComparison.OrdinalIgnoreCase))
        {
            if (IsUsbDescriptorValue(proto, "02") || proto.Contains("Mouse", StringComparison.OrdinalIgnoreCase))
            {
                role = "Mouse";
                return true;
            }

            if (IsUsbDescriptorValue(proto, "01") || proto.Contains("Keyboard", StringComparison.OrdinalIgnoreCase))
            {
                role = "Keyboard";
                return true;
            }

            if (!IsUsbDescriptorValue(sub, "01"))
            {
                role = "Gamepad";
                return true;
            }
        }

        if (IsUsbDescriptorValue(cls, "01") || cls.Contains("Audio", StringComparison.OrdinalIgnoreCase))
        {
            role = "Audio";
            return true;
        }

        if (IsUsbDescriptorValue(cls, "0E") || cls.Contains("Video", StringComparison.OrdinalIgnoreCase))
        {
            role = "Webcam";
            return true;
        }

        return false;
    }

    private static bool IsUsbDescriptorValue(string text, string expectedHex)
    {
        string compact = new string(text.Where(Uri.IsHexDigit).ToArray()).TrimStart('0');
        string expected = expectedHex.TrimStart('0');
        if (compact.Length == 0)
        {
            compact = "0";
        }
        if (expected.Length == 0)
        {
            expected = "0";
        }

        return string.Equals(compact, expected, StringComparison.OrdinalIgnoreCase);
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

    private bool EnsureImodDriverOnDisk(bool persistDriver, out string driverPath, out string? error)
    {
        _ = persistDriver;
        error = null;
        driverPath = GetImodDriverSystemPath();

        try
        {
            using Stream? resource = OpenImodDriverResourceStream();
            if (resource is null)
            {
                error = "embedded DTIMOD.sys not found";
                return false;
            }

            using MemoryStream buffer = new();
            resource.CopyTo(buffer);
            byte[] driverBytes = buffer.ToArray();
            string embeddedHash = ComputeSha256(driverBytes);
            string? expectedHash = ReadEmbeddedImodDriverHash();

            if (!string.IsNullOrWhiteSpace(expectedHash)
                && !HashEquals(embeddedHash, expectedHash))
            {
                error = $"embedded DTIMOD.sys hash mismatch: actual={embeddedHash}, expected={expectedHash}";
                return false;
            }

            string? driverDir = Path.GetDirectoryName(driverPath);
            if (string.IsNullOrWhiteSpace(driverDir))
            {
                error = "failed to resolve IMOD driver directory";
                return false;
            }

            if (!IsImodDriverSystemPath(driverPath))
            {
                string? driverRoot = Path.GetDirectoryName(driverDir);
                if (string.IsNullOrWhiteSpace(driverRoot))
                {
                    error = "failed to resolve IMOD driver root directory";
                    return false;
                }

                if (!TrySecureImodDriverDirectory(driverRoot, out error))
                {
                    return false;
                }

                if (!TrySecureImodDriverDirectory(driverDir, out error))
                {
                    return false;
                }
            }
            else if (!Directory.Exists(driverDir))
            {
                error = "Windows directory for IMOD driver not found";
                return false;
            }

            if (File.Exists(driverPath))
            {
                string existingHash = ComputeFileSha256(driverPath);
                if (HashEquals(existingHash, expectedHash ?? embeddedHash))
                {
                    WriteLog($"IMOD.DRIVER: using existing staged {driverPath} sha256={existingHash}");
                    return true;
                }

                WriteLog($"IMOD.DRIVER: replacing staged {driverPath} sha256={existingHash} -> {embeddedHash}");
            }

            File.WriteAllBytes(driverPath, driverBytes);

            string writtenHash = ComputeFileSha256(driverPath);
            if (!HashEquals(writtenHash, expectedHash ?? embeddedHash))
            {
                error = $"written DTIMOD.sys hash mismatch: actual={writtenHash}, expected={expectedHash ?? embeddedHash}";
                return false;
            }

            WriteLog($"IMOD.DRIVER: staged {driverPath} sha256={writtenHash}");
            return true;
        }
        catch (Exception ex)
        {
            error = $"failed to stage DTIMOD.sys: {ex.Message}";
            return false;
        }
    }

    private static bool TrySecureImodDriverDirectory(string directoryPath, out string? error)
    {
        error = null;

        try
        {
            Directory.CreateDirectory(directoryPath);

            DirectorySecurity security = new();
            security.SetAccessRuleProtection(isProtected: true, preserveInheritance: false);

            void AddRule(WellKnownSidType sidType)
            {
                security.AddAccessRule(new FileSystemAccessRule(
                    new SecurityIdentifier(sidType, null),
                    FileSystemRights.FullControl,
                    InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit,
                    PropagationFlags.None,
                    AccessControlType.Allow));
            }

            AddRule(WellKnownSidType.LocalSystemSid);
            AddRule(WellKnownSidType.BuiltinAdministratorsSid);

            SecurityIdentifier? userSid = WindowsIdentity.GetCurrent().User;
            if (userSid is not null)
            {
                security.AddAccessRule(new FileSystemAccessRule(
                    userSid,
                    FileSystemRights.FullControl,
                    InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit,
                    PropagationFlags.None,
                    AccessControlType.Allow));
            }

            new DirectoryInfo(directoryPath).SetAccessControl(security);
            return true;
        }
        catch (Exception ex)
        {
            error = $"failed to secure IMOD temp directory {directoryPath}: {ex.Message}";
            return false;
        }
    }

    private static Stream? OpenImodDriverResourceStream()
    {
        Assembly asm = typeof(MainForm).Assembly;
        Stream? stream = asm.GetManifestResourceStream("DeviceTweakerCS.IMOD.DTIMOD.sys");
        if (stream is not null)
        {
            return stream;
        }

        foreach (string name in asm.GetManifestResourceNames())
        {
            if (name.EndsWith(".DTIMOD.sys", StringComparison.OrdinalIgnoreCase))
            {
                return asm.GetManifestResourceStream(name);
            }
        }

        return null;
    }

    private static string? ReadEmbeddedImodDriverHash()
    {
        using Stream? stream = OpenImodDriverHashResourceStream();
        if (stream is null)
        {
            return null;
        }

        using StreamReader reader = new(stream, Encoding.ASCII, detectEncodingFromByteOrderMarks: true);
        return NormalizeHash(reader.ReadToEnd());
    }

    private static Stream? OpenImodDriverHashResourceStream()
    {
        Assembly asm = typeof(MainForm).Assembly;
        Stream? stream = asm.GetManifestResourceStream("DeviceTweakerCS.IMOD.DTIMOD.sys.sha256");
        if (stream is not null)
        {
            return stream;
        }

        foreach (string name in asm.GetManifestResourceNames())
        {
            if (name.EndsWith(".DTIMOD.sys.sha256", StringComparison.OrdinalIgnoreCase))
            {
                return asm.GetManifestResourceStream(name);
            }
        }

        return null;
    }

    private static bool TryOpenImodDriverDevice(out IntPtr handle, out string? error)
    {
        handle = InvalidHandleValue;
        error = null;
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

    private static bool TryOpenImodDriverDeviceOnce(out IntPtr handle, out string? error)
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
            error = null;
            return true;
        }

        error = $"failed to open {ImodDriverDevicePath}: {GetWin32ErrorMessage(Marshal.GetLastWin32Error())}";
        return false;
    }

    private static string ComputeFileSha256(string path)
    {
        using FileStream stream = File.OpenRead(path);
        return NormalizeHash(Convert.ToHexString(SHA256.HashData(stream))) ?? string.Empty;
    }

    private static string ComputeSha256(byte[] data)
    {
        return NormalizeHash(Convert.ToHexString(SHA256.HashData(data))) ?? string.Empty;
    }

    private static string? NormalizeHash(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        StringBuilder sb = new(value.Length);
        foreach (char ch in value)
        {
            if (Uri.IsHexDigit(ch))
            {
                sb.Append(char.ToUpperInvariant(ch));
            }
        }

        return sb.Length == 0 ? null : sb.ToString();
    }

    private static bool HashEquals(string? left, string? right)
    {
        string? normalizedLeft = NormalizeHash(left);
        string? normalizedRight = NormalizeHash(right);
        return normalizedLeft is not null
            && normalizedRight is not null
            && string.Equals(normalizedLeft, normalizedRight, StringComparison.Ordinal);
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

    private static bool TryReadPhys32(ImodDriverContext ctx, ulong address, out uint value, out string? error)
    {
        value = 0;

        if (!TryReadPhysicalMemory(ctx, address, 4, out ulong raw, out error))
        {
            return false;
        }

        value = unchecked((uint)raw);
        return true;
    }

    private static bool TryWritePhys32(ImodDriverContext ctx, ulong address, uint value, out string? error)
    {
        return TryWritePhysicalMemory(ctx, address, 4, value, out error);
    }

    private static bool TryWriteImodInterval(ImodDriverContext ctx, ulong address, uint interval, out string? error)
    {
        if (!TryReadPhys32(ctx, address, out uint currentValue, out error))
        {
            return false;
        }

        uint mergedValue = (currentValue & 0xFFFF0000) | (interval & 0xFFFF);
        return TryWritePhys32(ctx, address, mergedValue, out error);
    }

    private static bool TryMapPhysicalMemory(ImodDriverContext ctx, ulong address, ulong size, out PhysStruct phys, out string? error)
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
                IoctlImodMapPhysicalMemory,
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

        if (bytesReturned < Marshal.SizeOf<PhysStruct>())
        {
            error = "failed to map physical memory: incomplete ioctl response";
            return false;
        }

        if (phys.physMemLin == 0)
        {
            error = "failed to map physical memory: returned null linear address";
            return false;
        }

        return true;
    }

    private static bool TryReadPhysicalMemory(
        ImodDriverContext ctx,
        ulong address,
        uint size,
        out ulong value,
        out string? error)
    {
        value = 0;
        error = null;
        PhysAccessStruct access = new()
        {
            physAddress = address,
            accessSizeInBytes = size,
        };

        int bytesReturned = 0;
        if (!DeviceIoControl(
                ctx.DriverHandle,
                IoctlImodReadPhysicalMemory,
                ref access,
                Marshal.SizeOf<PhysAccessStruct>(),
                ref access,
                Marshal.SizeOf<PhysAccessStruct>(),
                out bytesReturned,
                IntPtr.Zero))
        {
            error = $"failed to read physical memory via driver: {GetWin32ErrorMessage(Marshal.GetLastWin32Error())}";
            return false;
        }

        if (bytesReturned < Marshal.SizeOf<PhysAccessStruct>())
        {
            error = "failed to read physical memory via driver: incomplete ioctl response";
            return false;
        }

        value = access.value;
        return true;
    }

    private static bool TryWritePhysicalMemory(
        ImodDriverContext ctx,
        ulong address,
        uint size,
        ulong value,
        out string? error)
    {
        error = null;
        PhysAccessStruct access = new()
        {
            physAddress = address,
            accessSizeInBytes = size,
            value = value,
        };

        int bytesReturned = 0;
        if (!DeviceIoControl(
                ctx.DriverHandle,
                IoctlImodWritePhysicalMemory,
                ref access,
                Marshal.SizeOf<PhysAccessStruct>(),
                ref access,
                Marshal.SizeOf<PhysAccessStruct>(),
                out bytesReturned,
                IntPtr.Zero))
        {
            error = $"failed to write physical memory via driver: {GetWin32ErrorMessage(Marshal.GetLastWin32Error())}";
            return false;
        }

        return true;
    }

    private static bool TryUnmapPhysicalMemory(ImodDriverContext ctx, PhysStruct phys, out string? error)
    {
        error = null;
        int bytesReturned = 0;
        if (!DeviceIoControl(
                ctx.DriverHandle,
                IoctlImodUnmapPhysicalMemory,
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

    private static bool TryLoadImodDriverWithKduFallback(string driverPath, Action<string>? log, out string? error)
    {
        error = null;

        if (IsImodKduFallbackDisabled())
        {
            error = $"{ImodKduDisableEnv}=1";
            log?.Invoke($"IMOD.DRIVER.KDU: skipped reason={error}");
            return false;
        }

        if (!File.Exists(driverPath))
        {
            error = $"driver missing: {driverPath}";
            log?.Invoke($"IMOD.DRIVER.KDU: skipped reason=driver_missing path={driverPath}");
            return false;
        }

        if (!EnsureImodKduPayloadOnDisk(log, out string kduPath, out error))
        {
            log?.Invoke($"IMOD.DRIVER.KDU: payload unavailable error={CompactImodLogValue(error)}");
            return false;
        }

        string kduDirectory = Path.GetDirectoryName(kduPath) ?? string.Empty;
        string kduDatabasePath = Path.Combine(kduDirectory, ImodKduDatabaseFileName);
        if (!File.Exists(kduDatabasePath))
        {
            error = $"drv64.dll missing next to kdu.exe: {kduDatabasePath}";
            log?.Invoke($"IMOD.DRIVER.KDU: skipped reason=db_missing path={kduDatabasePath}");
            return false;
        }

        log?.Invoke($"IMOD.DRIVER.KDU: map start kdu={kduPath} driver={driverPath}");
        if (!RunImodKduMap(kduPath, driverPath, log, out error))
        {
            log?.Invoke($"IMOD.DRIVER.KDU: map failed error={CompactImodLogValue(error)}");
            return false;
        }

        if (TryOpenImodDriverDevice(out IntPtr handle, out string? openError))
        {
            _ = CloseHandle(handle);
            log?.Invoke("IMOD.DRIVER.KDU: device available");
            return true;
        }

        error = $"mapped but device unavailable: {openError}";
        log?.Invoke($"IMOD.DRIVER.KDU: map result unusable error={CompactImodLogValue(error)}");
        return false;
    }

    private static bool IsImodKduFallbackDisabled()
    {
        string? value = Environment.GetEnvironmentVariable(ImodKduDisableEnv);
        return string.Equals(value, "1", StringComparison.OrdinalIgnoreCase)
            || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase)
            || string.Equals(value, "yes", StringComparison.OrdinalIgnoreCase);
    }

    private static bool EnsureImodKduPayloadOnDisk(Action<string>? log, out string kduPath, out string? error)
    {
        kduPath = string.Empty;
        error = null;

        string payloadRoot = Path.Combine(Path.GetTempPath(), "DeviceTweaker", "IMOD", "Loader");
        string kduTarget = Path.Combine(payloadRoot, ImodKduFileName);
        string dbTarget = Path.Combine(payloadRoot, ImodKduDatabaseFileName);

        if (TrySecureImodDriverDirectory(payloadRoot, out error)
            && TryWriteEmbeddedImodResource("DeviceTweakerCS.IMOD.Loader.kdu.exe", ".kdu.exe", kduTarget, out error)
            && TryWriteEmbeddedImodResource("DeviceTweakerCS.IMOD.Loader.drv64.dll", ".drv64.dll", dbTarget, out error))
        {
            kduPath = kduTarget;
            return true;
        }

        log?.Invoke($"IMOD.DRIVER.KDU: embedded payload unavailable error={CompactImodLogValue(error)}");
        if (TryFindExternalKdu(out string externalKduPath, out error))
        {
            kduPath = externalKduPath;
            return true;
        }

        return false;
    }

    private static bool TryWriteEmbeddedImodResource(string exactName, string suffix, string targetPath, out string? error)
    {
        error = null;
        using Stream? resource = OpenManifestResourceStreamExactOrSuffix(exactName, suffix);
        if (resource is null)
        {
            error = $"embedded resource missing: {exactName}";
            return false;
        }

        try
        {
            string? directory = Path.GetDirectoryName(targetPath);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            using MemoryStream buffer = new();
            resource.CopyTo(buffer);
            byte[] bytes = buffer.ToArray();
            string payloadHash = ComputeSha256(bytes);

            if (File.Exists(targetPath))
            {
                string existingHash = ComputeFileSha256(targetPath);
                if (HashEquals(existingHash, payloadHash))
                {
                    return true;
                }
            }

            File.WriteAllBytes(targetPath, bytes);
            return true;
        }
        catch (Exception ex)
        {
            error = $"failed to write {targetPath}: {ex.Message}";
            return false;
        }
    }

    private static Stream? OpenManifestResourceStreamExactOrSuffix(string exactName, string suffix)
    {
        Assembly asm = typeof(MainForm).Assembly;
        Stream? stream = asm.GetManifestResourceStream(exactName);
        if (stream is not null)
        {
            return stream;
        }

        foreach (string name in asm.GetManifestResourceNames())
        {
            if (name.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
            {
                return asm.GetManifestResourceStream(name);
            }
        }

        return null;
    }

    private static bool TryFindExternalKdu(out string kduPath, out string? error)
    {
        kduPath = string.Empty;
        error = null;

        IEnumerable<string?> rawCandidates =
        [
            Environment.GetEnvironmentVariable("DEVICE_TWEAKER_KDU_EXE"),
            Environment.GetEnvironmentVariable("KDU_EXE"),
            Environment.GetEnvironmentVariable("KDU_PATH"),
            Path.Combine(AppContext.BaseDirectory, ImodKduFileName),
            Path.Combine(AppContext.BaseDirectory, "Tools", "kdu", ImodKduFileName),
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
                "HYPE-MODE",
                "HYPE-MODE",
                "Tools",
                "kdu",
                ImodKduFileName)
        ];

        foreach (string? rawCandidate in rawCandidates)
        {
            if (string.IsNullOrWhiteSpace(rawCandidate))
            {
                continue;
            }

            string candidate = rawCandidate;
            if (Directory.Exists(candidate))
            {
                candidate = Path.Combine(candidate, ImodKduFileName);
            }

            if (!File.Exists(candidate))
            {
                continue;
            }

            string dbPath = Path.Combine(Path.GetDirectoryName(candidate) ?? string.Empty, ImodKduDatabaseFileName);
            if (!File.Exists(dbPath))
            {
                error = $"external KDU db missing: {dbPath}";
                continue;
            }

            kduPath = candidate;
            return true;
        }

        error ??= "kdu.exe not found";
        return false;
    }

    private static bool RunImodKduMap(string kduPath, string driverPath, Action<string>? log, out string? error)
    {
        error = null;

        try
        {
            using Process process = new();
            process.StartInfo = new ProcessStartInfo
            {
                FileName = kduPath,
                Arguments = $"-scv 3 -drvn {ImodDriverServiceName} -drvr {ImodDriverServiceName} -map {QuoteProcessArgument(driverPath)}",
                WorkingDirectory = Path.GetDirectoryName(kduPath) ?? AppContext.BaseDirectory,
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
            };

            if (!process.Start())
            {
                error = "failed to start kdu.exe";
                return false;
            }

            Task<string> stdout = process.StandardOutput.ReadToEndAsync();
            Task<string> stderr = process.StandardError.ReadToEndAsync();
            if (!process.WaitForExit(ImodKduTimeoutMs))
            {
                TryKillProcess(process);
                error = "kdu.exe timed out";
                return false;
            }

            _ = Task.WaitAll(new Task[] { stdout, stderr }, 2000);
            string output = stdout.IsCompletedSuccessfully ? stdout.Result : string.Empty;
            string errOutput = stderr.IsCompletedSuccessfully ? stderr.Result : string.Empty;
            string combinedOutput = CompactImodLogValue((output + " " + errOutput).Trim());
            log?.Invoke($"IMOD.DRIVER.KDU: exit code={process.ExitCode} output={combinedOutput}");

            if (TryOpenImodDriverDevice(out IntPtr mappedHandle, out string? mappedOpenError))
            {
                _ = CloseHandle(mappedHandle);
                if (process.ExitCode != 0)
                {
                    log?.Invoke(
                        "IMOD.DRIVER.KDU: non-zero exit ignored because device is available "
                        + $"code={process.ExitCode}");
                }

                return true;
            }

            if (process.ExitCode != 0)
            {
                error = $"kdu.exe failed code={process.ExitCode} device={mappedOpenError} output={combinedOutput}";
                return false;
            }

            error = $"kdu.exe exited successfully but device is unavailable: {mappedOpenError}";
            return false;
        }
        catch (Exception ex)
        {
            error = $"{ex.GetType().Name}: {ex.Message}";
            return false;
        }
    }

    private static void TryKillProcess(Process process)
    {
        try
        {
            if (!process.HasExited)
            {
                process.Kill(entireProcessTree: true);
            }
        }
        catch
        {
        }
    }

    private static string QuoteProcessArgument(string value)
    {
        return "\"" + value.Replace("\"", "\\\"", StringComparison.Ordinal) + "\"";
    }

    private sealed class ImodDriverContext : IDisposable
    {
        public IntPtr DriverHandle { get; private set; } = InvalidHandleValue;
        public bool ServiceCreated { get; private set; }
        public bool ServiceStartedByContext { get; private set; }
        public bool InitializedSuccessfully { get; private set; }
        public string DriverPath { get; }
        private readonly Action<string>? _log;

        private ImodDriverContext(string driverPath, Action<string>? log)
        {
            DriverPath = driverPath;
            _log = log;
        }

        public static bool TryInitialize(string driverPath, Action<string>? log, out ImodDriverContext? ctx, out string? error)
        {
            ctx = new ImodDriverContext(driverPath, log);
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

            _log?.Invoke($"IMOD.DRIVER: device not open before KDU start: {openError}");
            if (TryLoadImodDriverWithKduFallback(DriverPath, _log, out string? kduError))
            {
                _log?.Invoke("IMOD.DRIVER.SUMMARY: final=kdu loader=kdu");
                return true;
            }

            _log?.Invoke($"IMOD.DRIVER: KDU load failed, falling back to service: {kduError}");
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
            error = null;
            IntPtr scm = OpenSCManager(null, null, ScManagerAllAccess);
            if (scm == IntPtr.Zero)
            {
                error = $"failed to open service manager: {GetWin32ErrorMessage(Marshal.GetLastWin32Error())}";
                return false;
            }

            try
            {
                IntPtr service = OpenService(scm, ImodDriverServiceName, ServiceAllAccess);
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

                try
                {
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
                    _ = CloseServiceHandle(service);
                }
            }
            finally
            {
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

        private static void CleanupDriverArtifacts(string driverPath)
        {
            if (string.IsNullOrWhiteSpace(driverPath))
            {
                return;
            }

            if (IsImodDriverSystemPath(driverPath))
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
            catch
            {
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
            catch
            {
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
            catch
            {
            }
        }

        public void Dispose()
        {
            if (DriverHandle != InvalidHandleValue)
            {
                _ = CloseHandle(DriverHandle);
                DriverHandle = InvalidHandleValue;
            }

            if (ServiceStartedByContext || ServiceCreated)
            {
                StopServiceIfNeeded();
            }

            if (InitializedSuccessfully && (ServiceStartedByContext || ServiceCreated))
            {
                CleanupDriverArtifacts(DriverPath);
            }
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

    [StructLayout(LayoutKind.Sequential, Pack = 1)]
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
