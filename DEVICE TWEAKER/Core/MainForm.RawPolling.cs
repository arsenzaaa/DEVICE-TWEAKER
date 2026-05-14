using System.Diagnostics;
using System.Globalization;

namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private const int RawPollingUpdateIntervalMs = 250;
    private const int RawPollingMinSamples = 32;
    private const int RawPollingLiveMinSamples = 12;
    private const int RawPollingConfirmTicks = 2;
    private const int RawPollingDowngradeConfirmTicks = 8;

    private sealed class RawPollingState
    {
        public required string Role { get; init; }
        public required string InstanceId { get; init; }
        public string? ControllerId { get; set; }
        public long LastTick { get; set; }
        public Queue<double> IntervalsMs { get; } = [];
        public Queue<double> LiveIntervalsMs { get; } = [];
        public string CandidateTag { get; set; } = string.Empty;
        public int CandidateCount { get; set; }
        public string LastDisplayedTag { get; set; } = string.Empty;
        public double LastDisplayedHertz { get; set; }
        public string LastLiveTag { get; set; } = string.Empty;
        public double LastLiveHertz { get; set; }
    }

    private readonly Dictionary<IntPtr, RawPollingState> _rawPollingStates = [];
    private readonly Dictionary<string, Dictionary<string, string>> _rawPollingByController = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, Dictionary<string, string>> _rawLivePollingByController = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, string> _usbRoleOverrideByController = new(StringComparer.OrdinalIgnoreCase);
    private System.Windows.Forms.Timer? _rawPollingTimer;
    private bool _rawPollingInitialized;

    protected override void WndProc(ref Message m)
    {
        if (m.Msg == RawInputInterop.WmInput)
        {
            try
            {
                HandleRawInputMessage(m.LParam);
            }
            catch (Exception ex)
            {
                WriteLog($"USBPOLL.RAW: failed to process raw input: {ex.Message}");
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
            WriteLog("USBPOLL.RAW: registration failed");
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
                state.LastLiveHertz = liveHertz;
                string liveTag = FormatPollingRateTag(liveHertz);
                if (!string.Equals(state.LastLiveTag, liveTag, StringComparison.OrdinalIgnoreCase))
                {
                    state.LastLiveTag = liveTag;
                    if (!_rawLivePollingByController.TryGetValue(state.ControllerId, out Dictionary<string, string>? liveRoles))
                    {
                        liveRoles = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        _rawLivePollingByController[state.ControllerId] = liveRoles;
                    }

                    liveRoles[state.Role] = liveTag;
                    liveChanged = true;
                    WriteLog($"USBPOLL.RAW.LIVE: {state.Role} {liveTag} hz={liveHertz.ToString("0.##", CultureInfo.InvariantCulture)} controller={state.ControllerId} inst={state.InstanceId} samples={state.LiveIntervalsMs.Count}");
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
            if (string.Equals(state.CandidateTag, tag, StringComparison.OrdinalIgnoreCase))
            {
                state.CandidateCount++;
            }
            else
            {
                state.CandidateTag = tag;
                state.CandidateCount = 1;
            }

            if (state.CandidateCount < RawPollingConfirmTicks)
            {
                continue;
            }

            if (IsPollingDowngrade(snapped, state.LastDisplayedHertz)
                && (state.CandidateCount < RawPollingDowngradeConfirmTicks
                    || !IsLiveRateConsistentWithStableRate(state.LastLiveHertz, snapped)))
            {
                continue;
            }

            if (string.Equals(state.LastDisplayedTag, tag, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            state.LastDisplayedTag = tag;
            state.LastDisplayedHertz = snapped;
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

    private static bool TryEstimateRawPollingHz(IEnumerable<double> intervalsMs, out double hertz)
    {
        hertz = 0;
        double[] samples = intervalsMs
            .Where(v => v >= 0.05d && v <= 12.5d)
            .OrderBy(v => v)
            .ToArray();

        if (samples.Length < RawPollingMinSamples)
        {
            return false;
        }

        double activeIntervalMs = samples[samples.Length / 2];
        if (activeIntervalMs <= 0)
        {
            return false;
        }

        hertz = 1000d / activeIntervalMs;
        return hertz >= 100d;
    }

    private static bool TryEstimateLivePollingHz(IEnumerable<double> intervalsMs, out double hertz)
    {
        hertz = 0;
        double[] samples = intervalsMs
            .Where(v => v >= 0.05d && v <= 200d)
            .OrderBy(v => v)
            .ToArray();

        if (samples.Length < RawPollingLiveMinSamples)
        {
            return false;
        }

        double medianMs = samples[samples.Length / 2];
        if (medianMs <= 0)
        {
            return false;
        }

        hertz = 1000d / medianMs;
        return hertz > 0;
    }

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
    {
        return target > 0
            && Math.Abs(value - target) / target <= maxRelativeError;
    }

    private static bool IsPollingDowngrade(double candidateHertz, double currentHertz)
    {
        return currentHertz > 0
            && candidateHertz < currentHertz * 0.80d;
    }

    private static bool IsLiveRateConsistentWithStableRate(double liveHertz, double stableHertz)
    {
        return liveHertz > 0
            && stableHertz > 0
            && liveHertz >= stableHertz * 0.80d;
    }

    private static bool TrySnapStandardPollingRate(double hertz, out double snapped)
    {
        snapped = 0;
        double[] standardRates = [125d, 250d, 500d, 1000d, 2000d, 4000d, 8000d];
        double best = standardRates[0];
        double bestError = double.MaxValue;

        foreach (double rate in standardRates)
        {
            double error = Math.Abs(hertz - rate) / rate;
            if (error < bestError)
            {
                bestError = error;
                best = rate;
            }
        }

        if (bestError > 0.20d)
        {
            return false;
        }

        snapped = best;
        return true;
    }

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

            WriteLog($"USBPOLL.RAW: UI updated controller={controllerKey} roles=\"{roles}\" polling=\"{polling}\" live=\"{livePolling}\"");
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
                .OrderBy(k => k.Key, StringComparer.OrdinalIgnoreCase)
                .Select(k => $"{k.Key}: {k.Value}"));
    }

    private static string ApplyUsbRolePollingOverrides(string rolesText, IReadOnlyDictionary<string, string> overrides)
    {
        List<string> roles = [];
        foreach (string rawEntry in rolesText.Split(',', StringSplitOptions.RemoveEmptyEntries))
        {
            string entry = rawEntry.Trim();
            foreach (KeyValuePair<string, string> overrideEntry in overrides)
            {
                if (IsUsbRoleText(entry, overrideEntry.Key))
                {
                    if (ShouldKeepExistingUsbPollingRole(entry, overrideEntry.Value))
                    {
                        break;
                    }

                    entry = $"{overrideEntry.Key} {overrideEntry.Value}";
                    break;
                }
            }

            roles.Add(entry);
        }

        return string.Join(", ", roles);
    }

    private static bool ShouldKeepExistingUsbPollingRole(string roleText, string candidateTag)
    {
        if (!TryExtractPollingRateFromRole(roleText, out string role, out double existingHertz)
            || !TryParsePollingRateTag(candidateTag, out double candidateHertz))
        {
            return false;
        }

        return string.Equals(role, "Keyboard", StringComparison.OrdinalIgnoreCase)
            && candidateHertz < existingHertz;
    }

    private static bool TryExtractPollingRateFromRole(string roleText, out string role, out double hertz)
    {
        role = string.Empty;
        hertz = 0;

        string trimmed = roleText.Trim();
        foreach (string knownRole in new[] { "Mouse", "Keyboard" })
        {
            if (!trimmed.StartsWith(knownRole, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            string tag = trimmed[knownRole.Length..].Trim();
            if (!TryParsePollingRateTag(tag, out hertz))
            {
                return false;
            }

            role = knownRole;
            return true;
        }

        return false;
    }

    private static bool TryParsePollingRateTag(string tag, out double hertz)
    {
        hertz = 0;
        string trimmed = tag.Trim();
        if (trimmed.Length == 0 || trimmed.Equals("scanning", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        double multiplier = 1d;
        if (trimmed.EndsWith("Hz", StringComparison.OrdinalIgnoreCase))
        {
            trimmed = trimmed[..^2];
        }
        else if (trimmed.EndsWith("K", StringComparison.OrdinalIgnoreCase))
        {
            trimmed = trimmed[..^1];
            multiplier = 1000d;
        }

        return double.TryParse(trimmed, NumberStyles.Float, CultureInfo.InvariantCulture, out double value)
            && value > 0
            && (hertz = value * multiplier) > 0;
    }
}
