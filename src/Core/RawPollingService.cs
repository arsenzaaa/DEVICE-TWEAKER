using System.Globalization;

namespace DeviceTweakerCS;

internal sealed class RawPollingState
{
    private readonly object _sync = new();
    public required string Role { get; init; }
    public required string InstanceId { get; init; }
    private string? _controllerId;
    public string? ControllerId
    {
        get => Volatile.Read(ref _controllerId);
        set => Volatile.Write(ref _controllerId, value);
    }
    public bool ControllerResolveStarted { get; set; }
    public long LastTick { get; set; }
    private readonly Queue<double> _intervalsMs = [];
    private readonly Queue<double> _liveIntervalsMs = [];
    private readonly double[] _stableScratch = new double[256];
    private readonly double[] _liveScratch = new double[128];
    public int StableSampleCount { get { lock (_sync) return _intervalsMs.Count; } }
    public int LiveSampleCount { get { lock (_sync) return _liveIntervalsMs.Count; } }
    private bool _stableDirty = true, _liveDirty = true;
    private bool _hasStableEstimate, _hasLiveEstimate;
    private double _stableHertz, _liveHertz;
    public string CandidateTag { get; set; } = string.Empty;
    public int CandidateCount { get; set; }
    public string LastDisplayedTag { get; set; } = string.Empty;
    public double LastDisplayedHertz { get; set; }
    public string LastLiveTag { get; set; } = string.Empty;
    public double LastLiveHertz { get; set; }
    public string LiveCandidateTag { get; set; } = string.Empty;
    public int LiveCandidateCount { get; set; }
    public long LastLiveLogTimestamp { get; set; }

    internal void RecordSampleTick(long now, long frequency)
    {
        lock (_sync)
        {
            if (LastTick != 0)
            {
                double intervalMs = (now - LastTick) * 1000d / frequency;
                RecordIntervalUnsafe(intervalMs);
            }

            LastTick = now;
        }
    }

    internal void RecordInterval(double intervalMs)
    {
        lock (_sync)
        {
            RecordIntervalUnsafe(intervalMs);
        }
    }

    private void RecordIntervalUnsafe(double intervalMs)
    {
        if (intervalMs > 500d)
        {
            _intervalsMs.Clear();
            _liveIntervalsMs.Clear();
            _stableDirty = _liveDirty = true;
        }
        else if (intervalMs >= 0.05d && intervalMs <= 200d)
        {
            _liveIntervalsMs.Enqueue(intervalMs);
            while (_liveIntervalsMs.Count > 128) _ = _liveIntervalsMs.Dequeue();
            _liveDirty = true;
            if (intervalMs <= 12.5d)
            {
                _intervalsMs.Enqueue(intervalMs);
                while (_intervalsMs.Count > 256) _ = _intervalsMs.Dequeue();
                _stableDirty = true;
            }
        }
    }

    // Only sample changes invalidate estimates. Confirmation still runs on
    // every UI timer tick, including when no new input has arrived.
    internal bool TryEstimateLive(out double hertz)
    {
        lock (_sync)
        {
            if (_liveDirty)
            {
                _hasLiveEstimate = RawPollingService.TryEstimateLivePollingHz(_liveIntervalsMs, _liveScratch, out _liveHertz);
                _liveDirty = false;
            }
            hertz = _liveHertz;
            return _hasLiveEstimate;
        }
    }

    internal bool TryEstimateStable(out double hertz)
    {
        lock (_sync)
        {
            if (_stableDirty)
            {
                _hasStableEstimate = RawPollingService.TryEstimateRawPollingHz(_intervalsMs, _stableScratch, out _stableHertz);
                _stableDirty = false;
            }
            hertz = _stableHertz;
            return _hasStableEstimate;
        }
    }
}

internal static class RawPollingService
{
    private const int RawPollingMinSamples = 32;
    private const int RawPollingLiveMinSamples = 12;
    internal static bool TryConfirmLive(RawPollingState state, double hertz, string tag)
    {
        if (string.Equals(state.LiveCandidateTag, tag, StringComparison.OrdinalIgnoreCase)) state.LiveCandidateCount++;
        else { state.LiveCandidateTag = tag; state.LiveCandidateCount = 1; }
        if (state.LiveCandidateCount < 4 || string.Equals(state.LastLiveTag, tag, StringComparison.OrdinalIgnoreCase)) return false;
        state.LastLiveTag = tag; state.LastLiveHertz = hertz; return true;
    }
    internal static bool TryConfirmStable(RawPollingState state, double hertz, string tag)
    {
        if (string.Equals(state.CandidateTag, tag, StringComparison.OrdinalIgnoreCase)) state.CandidateCount++;
        else { state.CandidateTag = tag; state.CandidateCount = 1; }
        if (state.CandidateCount < 2) return false;
        if (IsPollingDowngrade(hertz, state.LastDisplayedHertz) &&
            (state.CandidateCount < 8 || !IsLiveRateConsistentWithStableRate(state.LastLiveHertz, hertz))) return false;
        if (string.Equals(state.LastDisplayedTag, tag, StringComparison.OrdinalIgnoreCase)) return false;
        state.LastDisplayedTag = tag; state.LastDisplayedHertz = hertz; return true;
    }
    internal static bool TryEstimateRawPollingHz(Queue<double> intervalsMs, double[] scratch, out double hertz)
        => TryEstimateMedianHz(intervalsMs, scratch, 12.5d, RawPollingMinSamples, 100d, out hertz);

    internal static bool TryEstimateLivePollingHz(Queue<double> intervalsMs, double[] scratch, out double hertz)
        => TryEstimateMedianHz(intervalsMs, scratch, 200d, RawPollingLiveMinSamples, double.Epsilon, out hertz);

    internal static bool TryEstimateRawPollingHz(IEnumerable<double> intervalsMs, out double hertz)
        => TryEstimateMedianHz(intervalsMs, null, 12.5d, RawPollingMinSamples, 100d, out hertz);

    internal static bool TryEstimateLivePollingHz(IEnumerable<double> intervalsMs, out double hertz)
        => TryEstimateMedianHz(intervalsMs, null, 200d, RawPollingLiveMinSamples, double.Epsilon, out hertz);

    private static bool TryEstimateMedianHz(
        IEnumerable<double> intervalsMs,
        double[]? scratch,
        double maxIntervalMs,
        int minSamples,
        double minHertz,
        out double hertz)
    {
        hertz = 0;
        double[] buffer = scratch ?? new double[256];
        int count = 0;
        foreach (double v in intervalsMs)
        {
            if (v >= 0.05d && v <= maxIntervalMs)
            {
                if (count == buffer.Length)
                {
                    Array.Resize(ref buffer, buffer.Length * 2);
                }

                buffer[count++] = v;
            }
        }

        if (count < minSamples)
        {
            return false;
        }

        Array.Sort(buffer, 0, count);
        double medianMs = buffer[count / 2];
        if (medianMs <= 0)
        {
            return false;
        }

        hertz = 1000d / medianMs;
        return hertz >= minHertz;
    }
    internal static bool IsClosePollingRate(double value, double target, double maxRelativeError)
    {
        return target > 0
            && Math.Abs(value - target) / target <= maxRelativeError;
    }
    internal static bool IsPollingDowngrade(double candidateHertz, double currentHertz)
    {
        return currentHertz > 0
            && candidateHertz < currentHertz * 0.80d;
    }
    internal static bool IsLiveRateConsistentWithStableRate(double liveHertz, double stableHertz)
    {
        return liveHertz > 0
            && stableHertz > 0
            && liveHertz >= stableHertz * 0.80d;
    }
    internal static bool TrySnapStandardPollingRate(double hertz, out double snapped)
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
    internal static string ApplyUsbRolePollingOverrides(string rolesText, IReadOnlyDictionary<string, string> overrides)
    {
        List<string> roles = [];
        HashSet<string> seen = new(StringComparer.OrdinalIgnoreCase);
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

            if (seen.Add(entry)) roles.Add(entry);
        }

        return string.Join(", ", UsbRoleDisplay.Order(roles));
    }
    internal static bool ShouldKeepExistingUsbPollingRole(string roleText, string candidateTag)
    {
        if (!TryExtractPollingRateFromRole(roleText, out string role, out double existingHertz)
            || !TryParsePollingRateTag(candidateTag, out double candidateHertz))
        {
            return false;
        }

        // The endpoint rate is only a maximum. Once the mouse cadence has been
        // confirmed, show that measurement even if the configured rate is lower.
        // Keyboard Raw Input is activity-dependent, so keep its endpoint rate.
        return role.Equals("Keyboard", StringComparison.OrdinalIgnoreCase)
            && candidateHertz < existingHertz;
    }
    internal static bool TryExtractPollingRateFromRole(string roleText, out string role, out double hertz)
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
    internal static bool TryParsePollingRateTag(string tag, out double hertz)
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
    internal static bool IsUsbRoleText(string text, string role)
    {
        string trimmed = text.Trim();
        return string.Equals(trimmed, role, StringComparison.OrdinalIgnoreCase)
            || trimmed.StartsWith(role + " ", StringComparison.OrdinalIgnoreCase);
    }
}
