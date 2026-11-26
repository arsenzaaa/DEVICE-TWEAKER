using System.Diagnostics;
using System.Globalization;

namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private const int WmNcActivate = 0x0086;
    private const int WmDeviceChange = 0x0219;
    private const int DbtDevNodesChanged = 0x0007;
    private const int RawPollingUpdateIntervalMs = 250;
    private const int RawPollingMinSamples = 32;
    private const int RawPollingLiveMinSamples = 12;
    private const int RawPollingLiveLogIntervalMs = 15000;

    private readonly object _rawPollingLock = new();
    private readonly Dictionary<IntPtr, RawPollingState> _rawPollingStates = [];
    private readonly HashSet<IntPtr> _ignoredRawDeviceHandles = [];
    private readonly List<RawPollingState> _rawPollingSnapshot = [];
    private readonly Dictionary<string, Dictionary<string, string>> _rawPollingByController = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, Dictionary<string, string>> _rawLivePollingByController = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, string> _usbRoleOverrideByController = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, string> _rawPollingUiSnapshotByController = new(StringComparer.OrdinalIgnoreCase);
    private RawInputInterop.RawInputBackgroundMonitor? _rawInputMonitor;
    private System.Windows.Forms.Timer? _rawPollingTimer;
    private System.Windows.Forms.Timer? _deviceChangeTimer;
    private bool _rawPollingInitialized;
    private bool _cachedRawMouseThrottleValid;
    private int _cachedRawMouseThrottleDuration;
    private long _cachedRawMouseThrottleReadTick;

    private System.Windows.Forms.Timer CreateDeviceChangeTimer()
    {
        System.Windows.Forms.Timer timer = new() { Interval = 1000 };
        timer.Tick += (_, _) =>
        {
            timer.Stop();
            if (!IsDisposed)
            {
                MarkImodMapsStale();
            }
        };
        return timer;
    }

    protected override void WndProc(ref Message m)
    {
#if DEVELOPER_QA
        if (TryHandleVisualQa(ref m)) return;
#endif
        // Windows 10 ignores DWMWA_CAPTION_COLOR for parts of the inactive
        // caption path and repaints the owner grey while a modal dialog is
        // active. Keep only the non-client visual state active; WinForms still
        // disables the owner and all focus/modal semantics remain unchanged.
        if (m.Msg == WmNcActivate && m.WParam == IntPtr.Zero && WindowState != FormWindowState.Minimized)
        {
            m.WParam = new IntPtr(1);
        }

        // Top-level windows receive this without registration whenever a device
        // is added or removed; a burst of them collapses into one update.
        if (!_uiDisposing && !IsDisposed && !Disposing && m.Msg == WmDeviceChange && m.WParam == DbtDevNodesChanged)
        {
            InvalidateDeviceTopology();
            _deviceChangeTimer ??= CreateDeviceChangeTimer();
            _deviceChangeTimer.Stop();
            _deviceChangeTimer.Start();
        }

        base.WndProc(ref m);
    }

    private void InitializeRawPolling()
    {
        if (_rawPollingInitialized || !IsHandleCreated)
        {
            return;
        }

        _rawPollingInitialized = true;
        // Receive WM_INPUT on a dedicated background message-only HWND thread so
        // 1K-8K mouse/keyboard packets never flood the WinForms UI message loop.
        RawInputInterop.RawInputBackgroundMonitor monitor = new(OnBackgroundRawInputSample);
        if (!monitor.Registered)
        {
            monitor.Dispose();
            _rawPollingInitialized = false;
            WriteLog("USBPOLL.RAW.WARN: registration failed");
            return;
        }

        _rawInputMonitor = monitor;
        _rawPollingTimer = new System.Windows.Forms.Timer
        {
            Interval = RawPollingUpdateIntervalMs,
        };
        _rawPollingTimer.Tick += (_, _) => UpdateRawPollingMeasurements();
        _rawPollingTimer.Start();
        WriteLog("USBPOLL.RAW: raw input measurement enabled");
        WriteLog($"USBPOLL.THROTTLE: {GetRawMouseThrottleStatus()}");
    }

    private void DisposeRawPolling()
    {
        _rawPollingInitialized = false;
        _rawInputMonitor?.Dispose();
        _rawInputMonitor = null;

        _rawPollingTimer?.Stop();
        _rawPollingTimer?.Dispose();
        _rawPollingTimer = null;
        _deviceChangeTimer?.Stop();
        _deviceChangeTimer?.Dispose();
        _deviceChangeTimer = null;
    }

    private void OnBackgroundRawInputSample(RawInputDeviceKind kind, IntPtr deviceHandle)
    {
        if (_uiDisposing || IsDisposed || Disposing)
        {
            return;
        }

        RawPollingState? state;
        bool resolveController = false;
        lock (_rawPollingLock)
        {
            if (_ignoredRawDeviceHandles.Contains(deviceHandle))
            {
                return;
            }

            _ = _rawPollingStates.TryGetValue(deviceHandle, out state);
        }

        if (state is null)
        {
            if (!RawInputInterop.TryGetDeviceIdentity(deviceHandle, out string deviceName, out string rawInstanceId))
            {
                lock (_rawPollingLock)
                {
                    _ignoredRawDeviceHandles.Add(deviceHandle);
                }
                return;
            }

            string instanceId = NormalizeInstanceId(rawInstanceId);
            if (string.IsNullOrWhiteSpace(instanceId))
            {
                lock (_rawPollingLock)
                {
                    _ignoredRawDeviceHandles.Add(deviceHandle);
                }
                return;
            }

            string role = kind switch
            {
                RawInputDeviceKind.Mouse => "Mouse",
                RawInputDeviceKind.Keyboard => "Keyboard",
                _ => string.Empty,
            };
            if (string.IsNullOrWhiteSpace(role))
            {
                return;
            }

            RawPollingState created = new()
            {
                Role = role,
                InstanceId = instanceId,
                ControllerResolveStarted = true,
            };
            lock (_rawPollingLock)
            {
                if (_rawPollingStates.TryGetValue(deviceHandle, out RawPollingState? existing))
                {
                    state = existing;
                }
                else
                {
                    _rawPollingStates[deviceHandle] = created;
                    state = created;
                    resolveController = true;
                }
            }

            if (resolveController)
            {
                WriteLog($"USBPOLL.RAW: device role={role} inst={instanceId} name=\"{deviceName}\"");
                QueueRawPollingControllerResolve(state);
            }
        }

        state.RecordSampleTick(Stopwatch.GetTimestamp(), Stopwatch.Frequency);
    }

    private void QueueRawPollingControllerResolve(RawPollingState state)
    {
        string instanceId = state.InstanceId;
        string role = state.Role;
        _ = Task.Run(() =>
        {
            try
            {
                string? controller = ResolveRawPollingController(instanceId);
                if (string.IsNullOrWhiteSpace(controller) || _uiDisposing || IsDisposed || Disposing)
                {
                    return;
                }

                state.ControllerId = controller;
                WriteLog($"USBPOLL.RAW: device controller role={role} inst={instanceId} controller={controller}");
            }
            catch
            {
                // Ignore background PnP lookup failures during shutdown or device removal.
            }
        });
    }

    private string? ResolveRawPollingController(string instanceId)
    {
        string? controller = FindUsbControllerFor(instanceId, new Dictionary<string, WmiPnPDevice>(StringComparer.OrdinalIgnoreCase));
        return string.IsNullOrWhiteSpace(controller) ? null : NormalizeInstanceId(controller);
    }

    private void UpdateRawPollingMeasurements()
    {
        bool changed = false;
        bool liveChanged = false;

        _rawPollingSnapshot.Clear();
        lock (_rawPollingLock)
        {
            foreach (RawPollingState state in _rawPollingStates.Values)
            {
                _rawPollingSnapshot.Add(state);
            }
        }

        foreach (RawPollingState state in _rawPollingSnapshot)
        {
            if (string.IsNullOrWhiteSpace(state.ControllerId))
            {
                continue;
            }

            if (state.TryEstimateLive(out double liveHertz))
            {
                liveHertz = NormalizeLivePollingHz(liveHertz);
                string liveTag = FormatPollingRateTag(liveHertz);
                if (RawPollingService.TryConfirmLive(state, liveHertz, liveTag))
                {
                    if (!_rawLivePollingByController.TryGetValue(state.ControllerId, out Dictionary<string, string>? liveRoles))
                    {
                        liveRoles = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        _rawLivePollingByController[state.ControllerId] = liveRoles;
                    }

                    liveRoles[state.Role] = liveTag;
                    liveChanged = true;
                    long now = Stopwatch.GetTimestamp();
                    if (state.LastLiveLogTimestamp == 0
                        || Stopwatch.GetElapsedTime(state.LastLiveLogTimestamp, now).TotalMilliseconds >= RawPollingLiveLogIntervalMs)
                    {
                        state.LastLiveLogTimestamp = now;
                        WriteLog($"USBPOLL.RAW.LIVE: {state.Role} {liveTag} hz={liveHertz.ToString("0.##", CultureInfo.InvariantCulture)} controller={state.ControllerId} inst={state.InstanceId} samples={state.LiveSampleCount}");
                    }
                }
            }

            if (state.StableSampleCount < RawPollingMinSamples || !state.TryEstimateStable(out double hertz))
            {
                continue;
            }

            if (!TrySnapStandardPollingRate(hertz, out double snapped))
            {
                continue;
            }

            string tag = FormatPollingRateTag(snapped);
            if (!RawPollingService.TryConfirmStable(state, snapped, tag)) continue;
            if (!_rawPollingByController.TryGetValue(state.ControllerId, out Dictionary<string, string>? roles))
            {
                roles = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                _rawPollingByController[state.ControllerId] = roles;
            }

            roles[state.Role] = tag;
            changed = true;
            WriteLog($"USBPOLL.RAW: measured {state.Role} {tag} hz={hertz.ToString("0.##", CultureInfo.InvariantCulture)} controller={state.ControllerId} inst={state.InstanceId} samples={state.StableSampleCount}");
        }

        if (changed || liveChanged)
        {
            ApplyRawPollingOverridesToBlocks();
        }
    }

    internal static bool TryEstimateRawPollingHz(IEnumerable<double> intervalsMs, out double hertz)
        => RawPollingService.TryEstimateRawPollingHz(intervalsMs, out hertz);

    internal static bool TryEstimateLivePollingHz(IEnumerable<double> intervalsMs, out double hertz)
        => RawPollingService.TryEstimateLivePollingHz(intervalsMs, out hertz);

    private double NormalizeLivePollingHz(double hertz)
    {
        if (TryGetCachedRawMouseThrottleDuration(out int duration))
        {
            double throttleCap = 1000d / duration;
            if (IsClosePollingRate(hertz, throttleCap, 0.35d))
            {
                return throttleCap;
            }
        }

        double[] liveRates = [50d, 62.5d, 100d, 125d, 200d, 250d, 500d, 1000d, 2000d, 4000d, 8000d];
        double best = hertz;
        double bestError = double.MaxValue;
        foreach (double rate in liveRates)
        {
            double error = Math.Abs(hertz - rate) / rate;
            if (error < bestError)
            {
                bestError = error;
                best = rate;
            }
        }

        return bestError <= 0.15d ? best : hertz;
    }

    private bool TryGetCachedRawMouseThrottleDuration(out int duration)
    {
        long now = Stopwatch.GetTimestamp();
        if (_cachedRawMouseThrottleReadTick == 0
            || Stopwatch.GetElapsedTime(_cachedRawMouseThrottleReadTick, now).TotalSeconds >= 5)
        {
            _cachedRawMouseThrottleReadTick = now;
            _cachedRawMouseThrottleValid = TryReadRawMouseThrottleDuration(out _cachedRawMouseThrottleDuration);
        }

        duration = _cachedRawMouseThrottleDuration;
        return _cachedRawMouseThrottleValid;
    }

    private static bool IsClosePollingRate(double value, double target, double maxRelativeError)
        => RawPollingService.IsClosePollingRate(value, target, maxRelativeError);

    private static bool IsPollingDowngrade(double candidateHertz, double currentHertz)
        => RawPollingService.IsPollingDowngrade(candidateHertz, currentHertz);

    private static bool IsLiveRateConsistentWithStableRate(double liveHertz, double stableHertz)
        => RawPollingService.IsLiveRateConsistentWithStableRate(liveHertz, stableHertz);

    private static bool TrySnapStandardPollingRate(double hertz, out double snapped)
        => RawPollingService.TrySnapStandardPollingRate(hertz, out snapped);

    private void ApplyRawPollingOverridesToBlocks()
    {
        bool imodRoleLabelsChanged = false;

        foreach (DeviceBlock block in _blocks)
        {
            if (block.Device.Kind != DeviceKind.USB || string.IsNullOrWhiteSpace(block.Device.UsbRoles))
            {
                continue;
            }

            string controllerKey = NormalizeInstanceId(block.Device.InstanceId);
            bool hasPolling = _rawPollingByController.TryGetValue(controllerKey, out Dictionary<string, string>? overrides)
                && overrides.Count > 0;
            bool hasLivePolling = _rawLivePollingByController.TryGetValue(controllerKey, out Dictionary<string, string>? liveOverrides)
                && liveOverrides.Count > 0;

            if (!hasPolling && !hasLivePolling)
            {
                continue;
            }

            string roles = block.Device.UsbRoles;
            string polling = block.Device.UsbPollingRates;
            if (hasPolling && overrides is not null)
            {
                roles = ApplyUsbRolePollingOverrides(roles, overrides);
                polling = FormatUsbPollingRoleSummary(roles.Split(',', StringSplitOptions.RemoveEmptyEntries));
            }

            string livePolling = hasLivePolling && liveOverrides is not null
                ? FormatLivePollingSummary(liveOverrides)
                : string.Empty;

            string title = BuildDeviceBlockTitle(block.Device, roles);
            if (!string.Equals(block.TitleLabel.Text, title, StringComparison.Ordinal))
            {
                block.TitleLabel.Text = title;
            }

            UpdateBlockInfoText(block, roles, polling, livePolling);

            if (hasPolling && !string.Equals(
                    _usbRoleOverrideByController.TryGetValue(controllerKey, out string? previousRoles) ? previousRoles : string.Empty,
                    roles,
                    StringComparison.OrdinalIgnoreCase))
            {
                _usbRoleOverrideByController[controllerKey] = roles;
                if (IsUsbImodTarget(block.Device) && !block.Device.IsTestDevice)
                {
                    imodRoleLabelsChanged = true;
                }
            }

            string uiSnapshot = $"roles=\"{roles}\" polling=\"{polling}\"";
            if (!_rawPollingUiSnapshotByController.TryGetValue(controllerKey, out string? previousSnapshot)
                || !string.Equals(previousSnapshot, uiSnapshot, StringComparison.Ordinal))
            {
                _rawPollingUiSnapshotByController[controllerKey] = uiSnapshot;
                WriteLog($"USBPOLL.RAW: UI updated controller={controllerKey} {uiSnapshot} live=\"{livePolling}\"");
            }
        }

        if (imodRoleLabelsChanged)
        {
            RefreshImodCurrentValues(showReadingStatus: false, reason: "raw-role-update");
        }
    }

    private static string FormatLivePollingSummary(IReadOnlyDictionary<string, string> liveOverrides)
    {
        return string.Join(
            ", ",
            liveOverrides
                .OrderBy(k => UsbRoleDisplay.Rank(k.Key))
                .ThenBy(k => k.Key, StringComparer.OrdinalIgnoreCase)
                .Select(k => $"{k.Key}: {k.Value}"));
    }

    internal static string ApplyUsbRolePollingOverrides(string rolesText, IReadOnlyDictionary<string, string> overrides)
        => RawPollingService.ApplyUsbRolePollingOverrides(rolesText, overrides);

    private static bool ShouldKeepExistingUsbPollingRole(string roleText, string candidateTag)
        => RawPollingService.ShouldKeepExistingUsbPollingRole(roleText, candidateTag);

    private static bool TryExtractPollingRateFromRole(string roleText, out string role, out double hertz)
        => RawPollingService.TryExtractPollingRateFromRole(roleText, out role, out hertz);

    private static bool TryParsePollingRateTag(string tag, out double hertz)
        => RawPollingService.TryParsePollingRateTag(tag, out hertz);
}
