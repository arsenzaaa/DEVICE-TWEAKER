using System.Globalization;

namespace DeviceTweakerCS;

internal sealed class RawPollingState
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
    public string LiveCandidateTag { get; set; } = string.Empty;
    public int LiveCandidateCount { get; set; }
    public long LastLiveLogTimestamp { get; set; }
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
    internal static bool TryEstimateRawPollingHz(IEnumerable<double> intervalsMs, out double hertz)
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
    internal static bool TryEstimateLivePollingHz(IEnumerable<double> intervalsMs, out double hertz)
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
