using System.Diagnostics;
using System.Globalization;

namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private const int WmNcActivate = 0x0086;
    private const int RawPollingUpdateIntervalMs = 250;
    private const int RawPollingMinSamples = 32;
    private const int RawPollingLiveMinSamples = 12;
    private const int RawPollingLiveLogIntervalMs = 15000;

    private readonly Dictionary<IntPtr, RawPollingState> _rawPollingStates = [];
    private readonly Dictionary<string, Dictionary<string, string>> _rawPollingByController = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, Dictionary<string, string>> _rawLivePollingByController = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, string> _usbRoleOverrideByController = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, string> _rawPollingUiSnapshotByController = new(StringComparer.OrdinalIgnoreCase);
    private System.Windows.Forms.Timer? _rawPollingTimer;
    private bool _rawPollingInitialized;

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

        if (m.Msg == RawInputInterop.WmInput)
        {
            try
            {
                HandleRawInputMessage(m.LParam);
            }
            catch (Exception ex)
            {
                WriteLog($"USBPOLL.RAW.WARN: failed to process raw input: {ex.Message}");
            }
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
        if (!RawInputInterop.RegisterMouseAndKeyboard(Handle))
        {
            _rawPollingInitialized = false;
            WriteLog("USBPOLL.RAW.WARN: registration failed");
            return;
        }

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
        _rawPollingTimer?.Stop();
        _rawPollingTimer?.Dispose();
        _rawPollingTimer = null;
    }

    private void HandleRawInputMessage(IntPtr lParam)
    {
        if (!RawInputInterop.TryGetMessage(lParam, out RawInputMessage message))
        {
            return;
        }

        string role = message.Kind switch
        {
            RawInputDeviceKind.Mouse => "Mouse",
            RawInputDeviceKind.Keyboard => "Keyboard",
            _ => string.Empty,
        };

        if (string.IsNullOrWhiteSpace(role))
        {
            return;
        }

        string instanceId = NormalizeInstanceId(message.InstanceId);
        if (string.IsNullOrWhiteSpace(instanceId))
        {
            return;
        }

        if (!_rawPollingStates.TryGetValue(message.DeviceHandle, out RawPollingState? state))
        {
            state = new RawPollingState
            {
                Role = role,
                InstanceId = instanceId,
            };
            _rawPollingStates[message.DeviceHandle] = state;
            WriteLog($"USBPOLL.RAW: device role={role} inst={instanceId} name=\"{message.DeviceName}\"");
        }

        if (string.IsNullOrWhiteSpace(state.ControllerId))
        {
            state.ControllerId = ResolveRawPollingController(instanceId);
            if (!string.IsNullOrWhiteSpace(state.ControllerId))
            {
                WriteLog($"USBPOLL.RAW: device controller role={role} inst={instanceId} controller={state.ControllerId}");
            }
        }

        long now = Stopwatch.GetTimestamp();
        if (state.LastTick != 0)
        {
            double intervalMs = (now - state.LastTick) * 1000d / Stopwatch.Frequency;
            if (intervalMs > 500d)
            {
                state.IntervalsMs.Clear();
                state.LiveIntervalsMs.Clear();
            }
            else if (intervalMs >= 0.05d && intervalMs <= 200d)
            {
                state.LiveIntervalsMs.Enqueue(intervalMs);
                while (state.LiveIntervalsMs.Count > 128)
                {
                    _ = state.LiveIntervalsMs.Dequeue();
                }

                if (intervalMs <= 12.5d)
                {
                    state.IntervalsMs.Enqueue(intervalMs);
                    while (state.IntervalsMs.Count > 256)
                    {
                        _ = state.IntervalsMs.Dequeue();
                    }
                }
            }
        }

        state.LastTick = now;
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

        foreach (RawPollingState state in _rawPollingStates.Values)
        {
            if (string.IsNullOrWhiteSpace(state.ControllerId))
            {
                continue;
            }

            if (TryEstimateLivePollingHz(state.LiveIntervalsMs, out double liveHertz))
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
                        WriteLog($"USBPOLL.RAW.LIVE: {state.Role} {liveTag} hz={liveHertz.ToString("0.##", CultureInfo.InvariantCulture)} controller={state.ControllerId} inst={state.InstanceId} samples={state.LiveIntervalsMs.Count}");
                    }
                }
            }

            if (state.IntervalsMs.Count < RawPollingMinSamples || !TryEstimateRawPollingHz(state.IntervalsMs, out double hertz))
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
            WriteLog($"USBPOLL.RAW: measured {state.Role} {tag} hz={hertz.ToString("0.##", CultureInfo.InvariantCulture)} controller={state.ControllerId} inst={state.InstanceId} samples={state.IntervalsMs.Count}");
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
        if (TryReadRawMouseThrottleDuration(out int duration))
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
