using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Security.Cryptography;
using System.Security.Principal;
using System.Text;
using System.Threading;

using static DeviceTweakerCS.ImodDriverNative;

namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private sealed class ImodApplyStats
    {
        public int ControllersFound { get; set; }
        public int ControllersApplied { get; set; }
        public int ControllersPartiallyApplied { get; set; }
        public int ControllersFailed { get; set; }
        public int WritesAttempted { get; set; }
        public int WritesSucceeded { get; set; }
        public int WriteFailures { get; set; }
        public int SkippedDisabled { get; set; }
        public int MissingBase { get; set; }
        public int ReadFailures { get; set; }
        public string? DriverError { get; set; }
    }

    private bool TryApplyImod(ImodConfig config, ImodRoleSnapshot roleSnapshot, bool persistDriver, out ImodApplyStats stats, out string? error)
    {
        DemandRealWrite("USB IMOD");
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

        if (!TryEnumerateXhciControllers(out List<ImodControllerInfo> controllers, out error))
        {
            return false;
        }

        stats.ControllersFound = controllers.Count;
        if (controllers.Count == 0)
        {
            return true;
        }

        if (!ImodDriverContext.TryInitialize(driverPath, WriteLog, TryLoadImodDriverWithKduFallback, out ImodDriverContext? driverContext, out error))
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

                if (!imodDriver.TrySetDeviceTarget(controller.DeviceId, out string? targetError))
                {
                    stats.ControllersFailed++;
                    WriteLog($"IMOD: {controller.DeviceId} skipped: {targetError}");
                    if (imodDriver.OlderDriverLoaded)
                    {
                        // Every controller would fail the same way until Windows restarts.
                        stats.DriverError = targetError;
                        break;
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

                if (!DeviceTweakerImodRuntime.TryResolveXhciRuntimeLayout(
                        imodDriver.Registers,
                        controller.BaseAddress,
                        controller.MemoryLength,
                        hcsparamsOffset,
                        rtsoff,
                        out uint maxIntrs,
                        out ulong runtimeAddress,
                        out _,
                        out _,
                        out string? ioError))
                {
                    stats.ReadFailures++;
                    WriteLog($"IMOD: unsafe or invalid xHCI MMIO layout {controller.DeviceId}: {ioError}");
                    continue;
                }

                if (adaptiveEntry?.RoleIntervals is { Count: > 0 } roleIntervals)
                {
                    if (TryBuildAdaptiveImodIntervals(
                            roleSnapshot, controller,
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
                        WriteLog($"IMOD.ADAPTIVE.FALLBACK.WARN: {controller.DeviceId} {adaptiveDetail}");
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

                stats.WritesAttempted += (int)writeCount;
                stats.WritesSucceeded += (int)(writeCount - writeFailures);
                stats.WriteFailures += (int)writeFailures;
                if (writeCount > 0 && writeFailures == 0)
                {
                    stats.ControllersApplied++;
                }
                else if (writeCount > 0 && writeFailures < writeCount)
                {
                    stats.ControllersPartiallyApplied++;
                }
                else
                {
                    stats.ControllersFailed++;
                }

                string modeText = desiredIntervals is { Count: > 0 }
                    ? $"vector={desiredIntervals.Count}"
                    : $"interval={FormatImodValue(desiredInterval)}";
                WriteLog($"IMOD: {controller.DeviceId} writes={writeCount}/{maxIntrs} {modeText} failures={writeFailures}");
            }
        }

        return true;
    }

    private async Task CheckImodDriverFromBlock(DeviceBlock block)
    {
        WriteLog($"UI: CHECK IMOD button clicked device={block.Device.InstanceId}");
        if (block.Device.IsTestDevice)
        {
            block.ImodReadbackRequested = true;
            RefreshTestImodPreview(block, "test-imod-check");
            RevealImodCheckResult(block);
            return;
        }

        if (TryBlockBusyHardwareWrite("IMOD CHECK") || TryBlockSandboxHardwareWrite("IMOD CHECK"))
        {
            return;
        }
        using IDisposable? operation = TryBeginDeviceOperation("IMOD CHECK");
        if (operation is null) return;
        if (!ConfirmDriverUse()) return;

        block.ImodReadbackRequested = true;
        (bool ok, string? error) = await CheckImodDriverAsync();
        if (_uiDisposing || IsDisposed || Disposing || block.Group.IsDisposed) return;
        if (!ok)
        {
            string status = "current: driver load failed";
            block.ImodCurrentLabel.Text = status;
            block.ImodCurrentLabel.Tag = "DTIMOD driver unavailable. Open the session log for diagnostics.";
            block.ImodCurrentLabel.ForeColor = _statusDanger;
            SetImodStatusTooltip(block.ImodCurrentLabel);
            if (block.ImodMapLabel is not null)
            {
                block.ImodMapLabel.Text = "devices: driver unavailable";
                block.ImodMapLabel.Tag = "DTIMOD driver unavailable. Open the session log for diagnostics.";
                block.ImodMapLabel.ForeColor = _statusDanger;
                SetImodStatusTooltip(block.ImodMapLabel);
            }

            SetImodDetailsVisibility(block, forceVisible: true);
            WriteLog($"IMOD.CHECK: failed device={block.Device.InstanceId} error={error}");
            OperationReport report = new();
            report.MarkNoChangesMade();
            report.AddError(
                "USB IMOD",
                FormatImodUnavailableUserMessage(error, includeNotChanged: false),
                error ?? "Unknown DTIMOD driver load error.");
            ShowOperationResult(
                report,
                string.Empty,
                "The driver check did not complete.",
                operationName: "USB IMOD CHECK");
            return;
        }

        WriteLog($"IMOD.CHECK: ok device={block.Device.InstanceId}");
        await AwaitBackgroundUiTasksAsync(null,
            RefreshImodCurrentValuesAsync(showReadingStatus: false, reason: "imod-check", checkedBlock: block));
    }

    private void RevealImodCheckResult(DeviceBlock block)
    {
        if (!block.ImodDetailsExpanded && block.ImodDetailsButton is { IsDisposed: false } details)
        {
            details.PerformClick();
        }
    }

    private void RefreshImodCurrentValues(bool showReadingStatus = true, string reason = "refresh")
    {
        _ = RefreshImodCurrentValuesAsync(showReadingStatus, reason);
    }

    private async Task RefreshImodCurrentValuesAsync(bool showReadingStatus = true, string reason = "refresh", DeviceBlock? checkedBlock = null)
    {
        List<DeviceBlock> testBlocks = _blocks
            .Where(b => IsUsbImodTarget(b.Device) && b.Device.IsTestDevice)
            .ToList();
        foreach (DeviceBlock block in testBlocks)
        {
            RefreshTestImodPreview(block, reason);
        }

        List<DeviceBlock> targetBlocks = _blocks
            .Where(b => IsUsbImodTarget(b.Device) && !b.Device.IsTestDevice)
            .ToList();
        int generation = Interlocked.Increment(ref _imodReadbackGeneration);

        if (targetBlocks.Count == 0)
        {
            WriteLog($"IMOD.READBACK.START: reason={reason} targets=0 testTargets={testBlocks.Count}");
            return;
        }

        WriteLog($"IMOD.READBACK.START: reason={reason} targets={targetBlocks.Count} testTargets={testBlocks.Count} showReading={showReadingStatus}");

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
        ImodRoleSnapshot roleSnapshot = CaptureImodRoleSnapshot();
        (bool ok, Dictionary<string, List<uint>> valuesByDeviceId, Dictionary<string, string> mapByDeviceId, Dictionary<string, string> mapDetailByDeviceId, string? error) readback =
            await Task.Run(() =>
            {
                bool ok = TryReadCurrentImodValues(
                    config, roleSnapshot,
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
            string detailText = IsImodAttentionStatus(statusText) || IsImodNeedsCheckStatus(readback.error)
                ? FormatImodReadbackDetail(readback.error)
                : "devices: unavailable";
            bool isBlockedStatus = IsImodBlockedStatus(statusText) || IsImodBlockedStatus(detailText);
            bool needsCheck = IsImodNeedsCheckStatus(statusText) || IsImodNeedsCheckStatus(detailText) || IsImodNeedsCheckStatus(readback.error);
            bool isAdminRequired = statusText.Contains("admin required", StringComparison.OrdinalIgnoreCase)
                || detailText.Contains("admin required", StringComparison.OrdinalIgnoreCase);
            bool showAttention = isBlockedStatus || needsCheck || isAdminRequired;
            foreach (DeviceBlock block in targetBlocks)
            {
                block.ImodCurrentLabel.Text = statusText;
                block.ImodCurrentLabel.Tag = showAttention
                    ? $"{statusText}\r\n{detailText}"
                    : block.ImodCurrentLabel.Text;
                block.ImodCurrentLabel.ForeColor = isBlockedStatus
                    ? _statusDanger
                    : _statusInactive;
                SetImodStatusTooltip(block.ImodCurrentLabel);
                if (block.ImodMapLabel is not null)
                {
                    block.ImodMapLabel.Text = detailText;
                    block.ImodMapLabel.Tag = block.ImodMapLabel.Text;
                    block.ImodMapLabel.ForeColor = isBlockedStatus ? _statusDanger : _mutedText;
                    SetImodStatusTooltip(block.ImodMapLabel);
                }
                SetImodDetailsVisibility(block, forceVisible: showAttention);
                LogImodUiSnapshot(block, reason);
            }

            WriteLog(QaRun.IsReadOnly && readback.error?.Contains("QA read-only session:", StringComparison.Ordinal) == true
                ? "IMOD.READBACK.SKIPPED: reason=qa-read-only hardwareValidation=SKIPPED"
                : IsImodNeedsCheckStatus(readback.error)
                ? "IMOD.READBACK.SKIPPED: reason=driver-not-loaded action=press-CHECK"
                : $"IMOD.READBACK.FAILED: error=\"{SanitizeLogValue(readback.error)}\"");
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

            block.ImodMapReadAt = DateTime.Now;
            block.ImodMapStale = false;
            block.ImodCurrentLabel.Text = $"current: {ImodMapText.ValueList(values)}";
            block.ImodCurrentLabel.Tag = FormatImodReadbackTooltip($"current raw: {ImodMapText.ValueListForLog(values)}", block.ImodMapReadAt.Value);
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
            if (ReferenceEquals(block, checkedBlock))
            {
                RevealImodCheckResult(block);
            }
            LogImodUiSnapshot(block, reason, values, block.ImodMapLabel?.Tag as string);
            WriteLog($"IMOD.READBACK: {block.Device.InstanceId} values={ImodMapText.ValueListForLog(values)}");
        }
    }

    private void LogImodUiSnapshot(DeviceBlock block, string reason, IReadOnlyList<uint>? values = null, string? mapDetail = null)
    {
        string value = block.ImodBox.Text?.Trim() ?? string.Empty;
        string defaultText = FlattenLogText(block.ImodDefaultLabel.Text);
        string currentText = FlattenLogText(block.ImodCurrentLabel.Text);
        string baseRoles = block.Device.UsbRoles ?? string.Empty;
        string effectiveRoles = ImodRoleMap.EffectiveUsbRoles(baseRoles, block.Device.InstanceId, block.Device.InstanceId, _usbRoleOverrideByController);
        string rawValues = values is { Count: > 0 }
            ? ImodMapText.ValueListForLog(values)
            : string.Empty;
        string detail = (mapDetail ?? (block.ImodMapLabel?.Tag as string) ?? block.ImodMapLabel?.Text ?? string.Empty).Trim();

        WriteLog(
            "IMOD.UI: "
            + $"reason={reason} "
            + $"controller={block.Device.InstanceId} "
            + $"input=\"{SanitizeLogValue(value)}\" "
            + $"default=\"{SanitizeLogValue(defaultText)}\" "
            + $"current=\"{SanitizeLogValue(currentText)}\" "
            + $"baseRoles=\"{SanitizeLogValue(baseRoles)}\" "
            + $"effectiveRoles=\"{SanitizeLogValue(effectiveRoles)}\" "
            + $"raw=\"{SanitizeLogValue(rawValues)}\" "
            + $"checked={block.ImodAutoCheck.Checked} currentVisible={block.ImodCurrentLabel.Visible} "
            + $"detailsExpanded={block.ImodDetailsExpanded} mapVisible={block.ImodMapLabel?.Visible == true} "
            + $"detailLines={CountNonEmptyLines(detail)}");

        if (!string.IsNullOrWhiteSpace(detail))
        {
            WriteLog(
                $"IMOD.UI.DETAIL: reason={reason} controller={block.Device.InstanceId}\r\n"
                + detail);
        }
    }

    private static int CountNonEmptyLines(string value)
        => value.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).Length;

    private const string ImodMapStaleSuffix = "devices changed";
    private const string ImodMapStaleHint = "Devices were added or removed after this read. A USB device may now use another interrupter; press CHECK to read again.";

    private static string FormatImodReadTime(DateTime readAt)
        => readAt.ToString("HH:mm", CultureInfo.InvariantCulture);

    private static string FormatImodReadbackTooltip(string rawValues, DateTime readAt)
        => $"Last read: {FormatImodReadTime(readAt)}\r\n{rawValues}";

    // Called after Windows reports added or removed devices: a re-plugged USB
    // device may now use another interrupter, so the shown map is outdated.
    private void MarkImodMapsStale()
    {
        foreach (DeviceBlock block in _blocks.Where(block => block.ImodMapReadAt is not null && !block.ImodMapStale))
        {
            block.ImodMapStale = true;
            string text = block.ImodCurrentLabel.Text ?? string.Empty;
            block.ImodCurrentLabel.Text = $"{text} · {ImodMapStaleSuffix}";
            block.ImodCurrentLabel.Tag = $"{ImodMapStaleHint}\r\nLast read: {FormatImodReadTime(block.ImodMapReadAt!.Value)}";
            block.ImodCurrentLabel.ForeColor = _statusWarn;
            SetImodStatusTooltip(block.ImodCurrentLabel);
            block.RelayoutAction?.Invoke();
            WriteLog($"IMOD.MAP.STALE: {block.Device.InstanceId} readAt={FormatImodReadTime(block.ImodMapReadAt!.Value)}");
        }
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
            // Tooltip only; the control can already be disposed.
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
            || text.Contains("press CHECK", StringComparison.OrdinalIgnoreCase)
            || text.Contains("admin required", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsImodBlockedStatus(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        return text.Contains("driver blocked", StringComparison.OrdinalIgnoreCase)
            || text.Contains("kernel CI blocked", StringComparison.OrdinalIgnoreCase)
            || text.Contains("signature blocked", StringComparison.OrdinalIgnoreCase)
            || text.Contains("Defender", StringComparison.OrdinalIgnoreCase)
            || text.Contains("blocked by Windows", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsImodNeedsCheckStatus(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        return text.Contains("press CHECK", StringComparison.OrdinalIgnoreCase)
            || text.Contains("driver not loaded", StringComparison.OrdinalIgnoreCase);
    }

    private void SetImodDetailsVisibility(DeviceBlock block, bool forceVisible)
    {
        bool showDetails = block.ImodAutoCheck.Visible
            && (block.ImodAutoCheck.Checked || block.ImodReadbackRequested || forceVisible || IsImodAttentionStatus(GetSourceControlText(block.ImodCurrentLabel)));

        if (block.ImodReadbackPanel is not null)
        {
            block.ImodReadbackPanel.Visible = showDetails;
        }
        block.ImodCurrentLabel.Visible = showDetails;
        block.ImodDefaultLabel.Visible = showDetails;
        if (block.ImodDetailsButton is not null)
        {
            block.ImodDetailsButton.Visible = showDetails;
        }
        if (block.ImodMapLabel is not null)
        {
            block.ImodMapLabel.Visible = showDetails && block.ImodDetailsExpanded;
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

        bool stackedSettings = settingsPanel.Left <= UiScale(24);
        if (!stackedSettings && block.Group.Width - settingsPanel.Left - UiScale(24) < UiScale(320))
        {
            stackedSettings = true;
            settingsPanel.Left = UiScale(18);
        }

        int availableSettingsWidth = Math.Max(UiScale(120), block.Group.Width - settingsPanel.Left - UiScale(24));
        if (block.ImodMapLabel is not null && block.ImodMapLabel.Visible)
        {
            int mapWidth = Math.Max(UiScale(120), availableSettingsWidth - block.ImodMapLabel.Left);
            FitImodMapLabel(block.ImodMapLabel, mapWidth);
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

        int settingsMinWidth = Math.Min(UiScale(420), availableSettingsWidth);
        Size currentSettingsSize = new(
            Math.Min(Math.Max(visibleSettingsRight + UiScale(8), settingsMinWidth), availableSettingsWidth),
            Math.Max(visibleSettingsBottom + UiScale(8), UiScale(24)));
        settingsPanel.Size = currentSettingsSize;

        int settingsTop = stackedSettings
            ? Math.Max(cpuPanel.Bottom, block.IrqLabel.Bottom) + UiScale(18)
            : cpuPanel.Top;
        settingsPanel.Location = new Point(settingsPanel.Left, settingsTop);

        if (!stackedSettings)
        {
            int maskY = cpuPanel.Bottom + UiScale(12);
            block.AffinityLabel.Location = new Point(block.AffinityLabel.Left, maskY);
            block.IrqLabel.Location = new Point(block.IrqLabel.Left, maskY + UiScale(20));
        }

        int infoY = Math.Max(block.IrqLabel.Bottom, settingsPanel.Bottom) + UiScale(12);

        block.InfoLabel.Location = new Point(block.InfoLabel.Left, infoY);
        block.InfoLabel.Size = new Size(block.Group.Width - UiScale(40), block.InfoLabel.Height);

        block.Group.Height = block.InfoLabel.Bottom + UiScale(20);
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
        ImodConfig config, ImodRoleSnapshot roleSnapshot,
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

        // Read-only hardware observation: never stage, migrate, load or delete files.
        {
            if (!TryEnumerateXhciControllers(out List<ImodControllerInfo> controllers, out error))
            {
                return false;
            }

            if (controllers.Count == 0)
            {
                return true;
            }

            if (!ImodDriverContext.TryOpenExisting(WriteLog, out ImodDriverContext? driverContext, out error))
            {
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

                if (!imodDriver.TrySetDeviceTarget(controller.DeviceId, out string? targetError))
                {
                    if (imodDriver.OlderDriverLoaded) { error = targetError; return false; }
                    WriteLog($"IMOD.READBACK: target unavailable {controller.DeviceId}: {targetError}");
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

                if (!DeviceTweakerImodRuntime.TryResolveXhciRuntimeLayout(
                        imodDriver.Registers,
                        controller.BaseAddress,
                        controller.MemoryLength,
                        hcsparamsOffset,
                        rtsoff,
                        out uint maxIntrs,
                        out ulong runtimeAddress,
                        out _,
                        out _,
                        out string? ioError))
                {
                    WriteLog($"IMOD.READBACK: unsafe or invalid xHCI MMIO layout {controller.DeviceId}: {ioError}");
                    continue;
                }

                uint readCount = Math.Min(maxIntrs, readbackLimit);
                if (!ImodReadbackService.TryRead(imodDriver.Registers, runtimeAddress, readCount,
                        out List<uint> values, out string? readError))
                {
                    WriteLog($"IMOD.READBACK: incomplete controller read {controller.DeviceId}: {readError}");
                    continue;
                }

                if (values.Count > 0)
                {
                    string normalizedControllerId = NormalizeInstanceId(controller.DeviceId);
                    valuesByDeviceId[normalizedControllerId] = values;
                    if (TryFormatImodInterrupterRoleMap(roleSnapshot, controller, imodDriver, maxIntrs, values, out string mapText, out string mapDetail))
                    {
                        mapByDeviceId[normalizedControllerId] = mapText;
                        mapDetailByDeviceId[normalizedControllerId] = mapDetail;
                        WriteLog($"IMOD.MAP: {controller.DeviceId} {mapDetail}");
                    }
                    else
                    {
                        WriteLog($"IMOD.MAP.WARN: unavailable {controller.DeviceId}: {mapDetail}");
                    }

                    if (maxIntrs > readbackLimit)
                    {
                        WriteLog($"IMOD.READBACK: {controller.DeviceId} read first {readbackLimit} of {maxIntrs} interrupters");
                    }
                }
            }
        }
        return true;
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

        if (error.Contains("press CHECK", StringComparison.OrdinalIgnoreCase)
            || error.Contains("driver not loaded", StringComparison.OrdinalIgnoreCase))
        {
            return "current: press CHECK";
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

        if (error.Contains("press CHECK", StringComparison.OrdinalIgnoreCase)
            || error.Contains("driver not loaded", StringComparison.OrdinalIgnoreCase))
        {
            return "devices: press CHECK";
        }

        if (error.Contains("administrator", StringComparison.OrdinalIgnoreCase)
            || error.Contains("админист", StringComparison.OrdinalIgnoreCase))
        {
            return "devices: admin required";
        }

        if (!IsImodSignatureRejectedError(error))
        {
            return "devices: unavailable";
        }

        if (error.Contains("virus", StringComparison.OrdinalIgnoreCase)
            || error.Contains("potentially unwanted", StringComparison.OrdinalIgnoreCase)
            || error.Contains("PUA", StringComparison.OrdinalIgnoreCase)
            || error.Contains("Defender", StringComparison.OrdinalIgnoreCase)
            || error.Contains("quarantine", StringComparison.OrdinalIgnoreCase))
        {
            return "DTIMOD.sys rejected by Windows or antivirus";
        }

        if (error.Contains("KDU", StringComparison.OrdinalIgnoreCase))
        {
            return "DTIMOD.sys loader rejected by Windows";
        }

        return "DTIMOD.sys rejected by Windows";
    }

    private bool TryBuildAdaptiveImodIntervals(
        ImodRoleSnapshot roleSnapshot, ImodControllerInfo controller,
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
                roleSnapshot, controller,
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
                && DeviceTweakerImodRuntime.TrySelectAdaptiveRoleInterval(roles, roleIntervals, out _, out uint roleValue))
            {
                target = roleValue;
            }

            intervals.Add(target & 0xFFFF);
        }

        List<string> shown = [];
        List<string> shared = [];
        foreach (KeyValuePair<uint, HashSet<string>> pair in rolesByInterrupter.OrderBy(kvp => kvp.Key))
        {
            bool hasSelectedRole = DeviceTweakerImodRuntime.TrySelectAdaptiveRoleInterval(pair.Value, roleIntervals, out string selectedRole, out uint selectedValue);
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
        ImodRoleSnapshot roleSnapshot, ImodControllerInfo controller,
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
            roleSnapshot, controller,
            imodDriver,
            maxIntrs,
            out rolesByInterrupter,
            out resolveDetail,
            allowFallback: false);
        if (!hasRoleMap)
        {
            WriteLog($"IMOD.MAP.ROLES.WARN: unavailable {controller.DeviceId}: {resolveDetail}");
        }

        (mapText, detail) = ImodMapText.Compose(currentIntervals, hasRoleMap ? rolesByInterrupter : null,
            roleSnapshot, controller.DeviceId, resolveDetail);
        return true;
    }

    private ImodRoleSnapshot CaptureImodRoleSnapshot() => new(
        _blocks.Where(block => block.Kind == DeviceKind.USB && !block.Device.IsTestDevice)
            .Select(block => new ImodRoleDevice(block.Device.InstanceId, block.Device.UsbRoles, block.Device.AudioEndpoints)),
        _usbRoleOverrideByController);

    private bool TryResolveAdaptiveRolesByInterrupter(
        ImodRoleSnapshot roleSnapshot, ImodControllerInfo controller,
        ImodDriverContext imodDriver,
        uint maxIntrs,
        out Dictionary<uint, HashSet<string>> rolesByInterrupter,
        out string detail,
        bool allowFallback = true)
    {
        rolesByInterrupter = [];

        if (!TryBuildExactAdaptiveRolesByInterrupter(
                roleSnapshot, controller,
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

        HashSet<string> fallbackRoles = ImodRoleMap.ControllerRoleFallback(roleSnapshot, controller.DeviceId);
        if (fallbackRoles.Count == 0)
        {
            detail = "no exact USB role/interrupter map and no controller role fallback";
            return false;
        }

        if (!ImodRoleMap.TryAssignFallbackRoles(fallbackRoles, interruptersByRootPort, maxIntrs, rolesByInterrupter, out detail))
        {
            return false;
        }

        detail = "fallback controller roles assigned to active interrupters: " + detail;
        return true;
    }

    private bool TryBuildExactAdaptiveRolesByInterrupter(
        ImodRoleSnapshot roleSnapshot, ImodControllerInfo controller,
        ImodDriverContext imodDriver,
        uint maxIntrs,
        out Dictionary<uint, HashSet<string>> rolesByInterrupter,
        out Dictionary<uint, List<uint>> interruptersByRootPort,
        out string detail)
    {
        rolesByInterrupter = [];
        interruptersByRootPort = [];
        detail = "no exact USB role/interrupter map";

        if (!TryReadXhciTopology(controller, imodDriver, out DeviceTweakerImodRuntime.XhciInterrupterTopology topology, out string topologyDetail))
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

        HashSet<string> controllerRoles = ImodRoleMap.ControllerRoleFallback(roleSnapshot, controller.DeviceId);
        rolesByInterrupter = ImodRoleMap.MapEndpoints(endpoints, topology, controller.DeviceId, controllerRoles,
            out ImodRoleMap.EndpointMapCounts counts);

        if (rolesByInterrupter.Count == 0)
        {
            detail = $"no USB endpoints matched active xHCI interrupters; classified={counts.Classified}; {topologyDetail}";
            return false;
        }

        detail =
            $"exact xHCI map: addressMatches={counts.AddressMatches}, rootPortMatches={counts.RootPortMatches}, endpointTargets={topology.EndpointTargetCount}, slotTargets={topology.SlotTargetCount}, classified={counts.Classified}, filteredGenericHid={counts.FilteredGamepads}; {topologyDetail}";
        return true;
    }

    private static bool TryReadXhciTopology(
        ImodControllerInfo controller,
        ImodDriverContext imodDriver,
        out DeviceTweakerImodRuntime.XhciInterrupterTopology topology,
        out string detail)
    {
        return DeviceTweakerImodRuntime.TryReadXhciInterrupterTopology(
            imodDriver.Registers, controller.BaseAddress, controller.MemoryLength, ImodDefaultHcsparamsOffset,
            out _, out topology, out detail);
    }

    private static bool TryReadXhciRootPortInterrupters(
        ImodControllerInfo controller,
        ImodDriverContext imodDriver,
        uint maxIntrs,
        out Dictionary<uint, List<uint>> interruptersByRootPort,
        out string detail)
    {
        interruptersByRootPort = [];
        if (!TryReadXhciTopology(controller, imodDriver, out DeviceTweakerImodRuntime.XhciInterrupterTopology topology, out detail))
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

    private async Task<(bool Ok, string? Error)> CheckImodDriverAsync()
    {
        if (!ConfirmDriverUse()) return (false, "Driver access cancelled.");
        DeviceBlock[] blocks = _blocks.Where(block => !block.Device.IsTestDevice).ToArray();
        string[] instances = blocks.Select(block => block.Device.InstanceId).ToArray();
        Dictionary<string, InterruptModeResult> results = new(StringComparer.OrdinalIgnoreCase);
        (bool ok, string? error) result = await Task.Run(() =>
        {
            try
            {
                bool ok = TryCheckImodDriver(instances, results, out string? error);
                return (ok, error);
            }
            catch (Exception error)
            {
                WriteLog($"IMOD.DRIVER.CHECK.FAILURE: {FlattenLogText(error.ToString())}");
                return (false, error.Message);
            }
        });
        if (!_uiDisposing && !IsDisposed && !Disposing)
            foreach (DeviceBlock block in blocks)
                if (!block.Group.IsDisposed && _blocks.Contains(block)
                    && results.TryGetValue(block.Device.InstanceId, out InterruptModeResult? value))
                    SetInterruptModeResult(block, value);
        return result;
    }

    private bool TryCheckImodDriver(string[] instances, Dictionary<string, InterruptModeResult> results, out string? error)
    {
        error = null;
        if (!IsAdministrator())
        {
            error = "administrator privileges required";
            return false;
        }

        if (!EnsureImodDriverOnDisk(persistDriver: true, out string driverPath, out error))
        {
            return false;
        }

        if (!ImodDriverContext.TryInitialize(driverPath, WriteLog, TryLoadImodDriverWithKduFallback, out ImodDriverContext? driverContext, out error))
        {
            if (IsImodKernelCiBlockedLoadError(error))
            {
                string ciState = GetImodKernelCiBlockState();
                error = $"{error} (kernel CI blocked; {ciState})";
            }

            LogImodDriverLoadDiagnostics(driverPath, error);
            return false;
        }

        using (driverContext)
        {
            ClearImodKernelCiBlockStatus();
            foreach (string instance in instances)
                results[instance] = PciInterruptReader.Query(driverContext!.DriverHandle, instance);
            WriteLog($"IMOD.DRIVER.CHECK: ok path={driverPath}");
            return true;
        }
    }

    private bool EnsureImodDriverOnDisk(bool persistDriver, out string driverPath, out string? error)
    {
        if (QaRun.TryBlock("PUBLISHED DRIVER STAGING"))
        { driverPath = string.Empty; error = "QA read-only session: driver access skipped"; return false; }
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

            if (File.Exists(driverPath))
            {
                string existingHash = ComputeFileSha256(driverPath);
                if (HashEquals(existingHash, expectedHash ?? embeddedHash))
                {
                    WriteLog($"IMOD.DRIVER: using existing staged {driverPath} sha256={existingHash}");
                    return true;
                }

                if (!persistDriver)
                {
                    WriteLog(
                        $"IMOD.DRIVER: hash mismatch persistDriver=false -> keep existing {driverPath} " +
                        $"sha256={existingHash} (no replace on refresh/read)");
                    return true;
                }

                WriteLog($"IMOD.DRIVER: replacing staged {driverPath} sha256={existingHash} -> {embeddedHash}");
            }
            else if (!persistDriver)
            {
                error = "DTIMOD.sys is not staged; press CHECK to load the driver";
                WriteLog($"IMOD.DRIVER: missing {driverPath} persistDriver=false -> skip stage");
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

            WriteAllBytesAtomic(driverPath, driverBytes);

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
        QaRun.DemandWrite("TrySecureImodDriverDirectory");
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

    // Startup\DeviceTweakerImodRuntime.cs enumerates controllers for the app and the logon script.
    private static bool TryEnumerateXhciControllers(out List<ImodControllerInfo> controllers, out string? error)
    {
        controllers = [];
        try
        {
            foreach (DeviceTweakerImodController controller in DeviceTweakerImodRuntime.EnumerateXhciControllers())
            {
                controllers.Add(new ImodControllerInfo
                {
                    DeviceId = controller.DeviceId,
                    Caption = controller.Caption ?? string.Empty,
                    ProblemCode = controller.ProblemCode,
                    BaseAddress = controller.BaseAddress,
                    MemoryLength = controller.MemoryLength,
                    HasBase = controller.HasBase,
                    BaseError = controller.BaseError ?? string.Empty,
                });
            }
        }
        catch (InvalidOperationException ex)
        {
            error = ex.Message;
            return false;
        }

        error = null;
        return true;
    }

    private static bool TryWriteImodInterval(ImodDriverContext ctx, ulong address, uint interval, out string? error)
    {
        QaRun.DemandWrite("PUBLISHED PHYSICAL ACCESS");
        return DeviceTweakerImodRuntime.TryWriteImodInterval(ctx.Registers, address, interval, out error);
    }

    private static bool TryLoadImodDriverWithKduFallback(string driverPath, Action<string>? log, out string? error)
    {
        QaRun.DemandWrite("PUBLISHED RUNTIME: TryLoadImodDriverWithKduFallback");
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

        if (!TryVerifyEmbeddedImodResourceFile("DeviceTweakerCS.IMOD.Loader.kdu.exe", ".kdu.exe", kduPath, out error)
            || !TryVerifyEmbeddedImodResourceFile("DeviceTweakerCS.IMOD.Loader.drv64.dll", ".drv64.dll", kduDatabasePath, out error))
        {
            log?.Invoke($"IMOD.DRIVER.KDU: payload verification failed error={CompactImodLogValue(error)}");
            return false;
        }

        string? expectedDriverHash = ReadEmbeddedImodDriverHash();
        if (string.IsNullOrWhiteSpace(expectedDriverHash)
            || !HashEquals(ComputeFileSha256(driverPath), expectedDriverHash))
        {
            error = "DTIMOD.sys changed after staging or its embedded hash is unavailable";
            log?.Invoke($"IMOD.DRIVER.KDU: driver verification failed path={driverPath}");
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
        QaRun.DemandWrite("PUBLISHED RUNTIME: EnsureImodKduPayloadOnDisk");
        kduPath = string.Empty;
        error = null;

        string commonData = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);
        if (string.IsNullOrWhiteSpace(commonData))
        {
            commonData = Path.GetTempPath();
        }

        string payloadRoot = Path.Combine(commonData, "DEVICE TWEAKER", "IMOD", "Loader");
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
        return false;
    }

    private static bool TryWriteEmbeddedImodResource(string exactName, string suffix, string targetPath, out string? error)
    {
        QaRun.DemandWrite("PUBLISHED RUNTIME: TryWriteEmbeddedImodResource");
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

            WriteAllBytesAtomic(targetPath, bytes);
            return true;
        }
        catch (Exception ex)
        {
            error = $"failed to write {targetPath}: {ex.Message}";
            return false;
        }
    }

    private static bool TryVerifyEmbeddedImodResourceFile(string exactName, string suffix, string targetPath, out string? error)
    {
        QaRun.DemandWrite("PUBLISHED RUNTIME: TryVerifyEmbeddedImodResourceFile");
        error = null;
        using Stream? resource = OpenManifestResourceStreamExactOrSuffix(exactName, suffix);
        if (resource is null)
        {
            error = $"embedded resource missing: {exactName}";
            return false;
        }

        if (!File.Exists(targetPath))
        {
            error = $"staged payload missing: {targetPath}";
            return false;
        }

        try
        {
            string expectedHash = NormalizeHash(Convert.ToHexString(SHA256.HashData(resource))) ?? string.Empty;
            string actualHash = ComputeFileSha256(targetPath);
            if (!HashEquals(actualHash, expectedHash))
            {
                error = $"staged payload hash mismatch: {Path.GetFileName(targetPath)}";
                return false;
            }

            return true;
        }
        catch (Exception ex)
        {
            error = $"failed to verify {targetPath}: {ex.Message}";
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

    private static bool RunImodKduMap(string kduPath, string driverPath, Action<string>? log, out string? error)
    {
        QaRun.DemandWrite("RunImodKduMap");
        error = null;
        Stopwatch timer = Stopwatch.StartNew();

        try
        {
            using Process process = new();
            process.StartInfo = new ProcessStartInfo
            {
                FileName = kduPath,
                Arguments = $"-scv 3 -drvn DeviceTweakerImod2 -drvr DeviceTweakerImod2 -map {QuoteProcessArgument(driverPath)}",
                WorkingDirectory = Path.GetDirectoryName(kduPath) ?? AppContext.BaseDirectory,
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
            };

            DriverLoadGate.BeginLoader();
            bool started;
            try { started = process.Start(); }
            catch { DriverLoadGate.CompleteLoader(); throw; }
            if (!started)
            {
                DriverLoadGate.CompleteLoader();
                error = "failed to start kdu.exe";
                return false;
            }

            Task<string> stdout = process.StandardOutput.ReadToEndAsync();
            Task<string> stderr = process.StandardError.ReadToEndAsync();
            if (!process.WaitForExit(ImodKduTimeoutMs))
            {
                _ = Task.WaitAll(new Task[] { stdout, stderr }, 2000);
                string timeoutOutput = stdout.IsCompletedSuccessfully ? stdout.Result : string.Empty;
                string timeoutError = stderr.IsCompletedSuccessfully ? stderr.Result : string.Empty;
                log?.Invoke($"IMOD.DRIVER.KDU: timeout elapsedMs={timer.ElapsedMilliseconds}");
                LogImodKduOutput(log, "STDOUT", timeoutOutput);
                LogImodKduOutput(log, "STDERR", timeoutError);
                error = "KDU has not finished. Restart Windows before loading DTIMOD again. See the session log.";
                return false;
            }

            DriverLoadGate.CompleteLoader();

            _ = Task.WaitAll(new Task[] { stdout, stderr }, 2000);
            string output = stdout.IsCompletedSuccessfully ? stdout.Result : string.Empty;
            string errOutput = stderr.IsCompletedSuccessfully ? stderr.Result : string.Empty;
            log?.Invoke($"IMOD.DRIVER.KDU: exit code={process.ExitCode} elapsedMs={timer.ElapsedMilliseconds}");
            LogImodKduOutput(log, "STDOUT", output);
            LogImodKduOutput(log, "STDERR", errOutput);

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
                error = $"kdu.exe failed code={process.ExitCode}; DTIMOD device unavailable: {mappedOpenError}. See full KDU output in the session log.";
                return false;
            }

            error = $"kdu.exe exited successfully but DTIMOD device was not created: {mappedOpenError}. See full KDU output in the session log.";
            return false;
        }
        catch (Exception ex)
        {
            log?.Invoke($"IMOD.DRIVER.KDU.EXCEPTION: {FlattenLogText(ex.ToString())}");
            error = $"{ex.GetType().Name}: {ex.Message}";
            return false;
        }
    }

    private static void LogImodKduOutput(Action<string>? log, string streamName, string? content)
    {
        if (log is null)
        {
            return;
        }

        string normalized = (content ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n');
        string[] lines = normalized.Split('\n');
        log($"IMOD.DRIVER.KDU.{streamName}: BEGIN chars={normalized.Length} lines={lines.Length}");
        if (normalized.Length == 0)
        {
            log($"IMOD.DRIVER.KDU.{streamName}: <empty>");
        }
        else
        {
            for (int i = 0; i < lines.Length; i++)
            {
                log($"IMOD.DRIVER.KDU.{streamName}: {i + 1:D4}: {lines[i]}");
            }
        }
        log($"IMOD.DRIVER.KDU.{streamName}: END");
    }

    private static string QuoteProcessArgument(string value)
    {
        return "\"" + value.Replace("\"", "\\\"", StringComparison.Ordinal) + "\"";
    }

}
