using System.Globalization;
using System.Text.RegularExpressions;

namespace DeviceTweakerCS;

internal enum NicItrTimingKind
{
    None,
    IntelEitr,
    IntelItr,
    RealtekIntrMit,
    RealtekIntrMitV2,
}

internal sealed record NicItrProfile(
    string FamilyName,
    string VendorId,
    string[] DeviceIds,
    uint BaseOffset,
    uint Stride,
    int MaxQueues,
    int ReadWidth,
    ulong ReadMask,
    ulong WriteOrBits,
    NicItrTimingKind TimingKind,
    string[]? VectorLabels = null,
    uint IntelEitrUnitUs = 2);

internal interface INicItrRegisters
{
    bool TryWrite(ulong address, int width, ulong value, out string? error);
    bool TryRead(ulong address, int width, out ulong value, out string? error);
}

internal static class NicItrService
{
    internal static readonly NicItrProfile[] NicItrProfiles =
    [
        new("Intel I225/I226 (EITR)", "8086", ["15F2", "15F3", "15F7", "15F8", "0D9F", "5502", "125B", "125C", "125D", "5503"], 0x1680, 0x4, 5, 32, 0x00007FFC, 0x80000000, NicItrTimingKind.IntelEitr, ["Other", "IRQ0", "IRQ1", "IRQ2", "IRQ3"]),
        new("Intel I210/I211 (EITR)", "8086", ["1533", "1536", "1537", "1538", "1539", "157B", "157C", "1F40", "1F41", "1F45"], 0x1680, 0x4, 5, 32, 0x00007FFC, 0x80000000, NicItrTimingKind.IntelEitr, IntelEitrUnitUs: 1),
        new("Intel I350 (EITR)", "8086", ["1521", "1522", "1523", "1524"], 0x1680, 0x4, 25, 32, 0x00007FFC, 0x80000000, NicItrTimingKind.IntelEitr, IntelEitrUnitUs: 1),
        new("Intel 82580 (EITR)", "8086", ["150E", "150F", "1510", "1511"], 0x1680, 0x4, 10, 32, 0x00007FFC, 0x80000000, NicItrTimingKind.IntelEitr),
        new("Intel 82576 (EITR)", "8086", ["1516", "1518", "1526"], 0x1680, 0x4, 25, 32, 0x00007FFC, 0x80000000, NicItrTimingKind.IntelEitr),
        new("Killer E3100 (EITR)", "8086", ["3100", "3101", "3102"], 0x1680, 0x4, 5, 32, 0x00007FFC, 0x80000000, NicItrTimingKind.IntelEitr, ["Other", "IRQ0", "IRQ1", "IRQ2", "IRQ3"]),
        new("Intel I219 (ITR)", "8086", ["15B7", "15B8", "15B9", "15D7", "15D8", "15E3", "15BB", "15BC", "15BD", "15BE", "0D4C", "0D4D", "0D4E", "0D4F", "0D53", "0D55", "0D5C", "0D5D", "0D5E", "0D5F", "15FB", "15FC", "1A1E", "1A1F", "550A", "550B", "550C", "550D", "550E", "550F", "0DC5", "0DC6", "0DC7", "0DC8", "1A1C", "1A1D", "15F9", "15FA", "3166", "3167", "3197", "3198", "4DF4", "4B33", "4B34", "4DC2", "4DC3", "54B4", "54B5", "54B6", "54B7", "0126", "153A", "153B", "1559", "155A", "156F", "1570", "15D6"], 0x00C4, 0x0, 1, 32, 0x0000FFFF, 0, NicItrTimingKind.IntelItr),
        new("Realtek RTL8111/8168", "10EC", ["8168", "8161", "8136", "8167", "8169"], 0x00E2, 0x0, 1, 16, 0xFFFF, 0, NicItrTimingKind.RealtekIntrMit),
        new("Realtek RTL8125/8126", "10EC", ["8125", "8162", "8126"], 0x0A00, 0x8, 4, 32, 0x7F7F7F7F, 0, NicItrTimingKind.RealtekIntrMitV2),
        new("Realtek RTL8168KB", "10EC", ["3000"], 0x00E2, 0x0, 1, 16, 0xFFFF, 0, NicItrTimingKind.RealtekIntrMit),
        new("Killer E2500/E2600", "10EC", ["2600", "2502", "2500"], 0x00E2, 0x0, 1, 16, 0xFFFF, 0, NicItrTimingKind.RealtekIntrMit),
    ];

    internal static NicItrProfile? TryGetNicItrProfile(string instanceId)
    {
        if (!TryGetPciVenDev(instanceId, out string ven, out string dev))
        {
            return null;
        }

        foreach (NicItrProfile profile in NicItrProfiles)
        {
            if (string.Equals(profile.VendorId, ven, StringComparison.OrdinalIgnoreCase)
                && profile.DeviceIds.Any(id => string.Equals(id, dev, StringComparison.OrdinalIgnoreCase)))
            {
                return profile;
            }
        }

        return null;
    }

    internal static bool TryProgramNicItrRegisters(
        NicItrProfile profile, ulong baseAddress, IReadOnlyList<ulong> values,
        INicItrRegisters registers, out string? error)
    {
        error = null;
        if (values.Count == 0)
        {
            error = "no values";
            return false;
        }
        for (int q = 0; q < profile.MaxQueues; q++)
        {
            ulong selected = q < values.Count ? values[q] : values[0];
            ulong finalValue = (selected & profile.ReadMask) | profile.WriteOrBits;
            ulong address = baseAddress + profile.BaseOffset + (profile.Stride * (uint)q);
            if (!registers.TryWrite(address, profile.ReadWidth, finalValue, out error))
            {
                return false;
            }

            if (!registers.TryRead(address, profile.ReadWidth, out ulong readback, out error))
            {
                error = $"write completed but readback failed at {GetNicItrVectorLabel(profile, q)}: {error}";
                return false;
            }

            ulong expected = selected & profile.ReadMask;
            ulong actual = readback & profile.ReadMask;
            if (actual != expected)
            {
                error = $"write verification failed at {GetNicItrVectorLabel(profile, q)}: expected {FormatNicItrValue(expected, profile)}, read {FormatNicItrValue(actual, profile)}";
                return false;
            }
        }

        return true;
    }
    internal static bool TryValidateNicItrMemoryRange(NicItrProfile profile, ulong memoryLength, out string? error)
    {
        ulong registerWidth = (ulong)(profile.ReadWidth / 8);
        ulong lastOffset = profile.BaseOffset + ((ulong)Math.Max(0, profile.MaxQueues - 1) * profile.Stride);
        if (registerWidth == 0 || memoryLength == 0 || lastOffset > memoryLength || registerWidth > memoryLength - lastOffset)
        {
            error = $"NIC ITR register range exceeds the allocated BAR (required end=0x{lastOffset + registerWidth:X}, BAR length=0x{memoryLength:X})";
            return false;
        }

        error = null;
        return true;
    }
    internal static bool TryGetPciVenDev(string instanceId, out string vendorId, out string deviceId)
    {
        vendorId = string.Empty;
        deviceId = string.Empty;
        Match match = Regex.Match(
            instanceId ?? string.Empty,
            "VEN_([0-9A-Fa-f]{4})&DEV_([0-9A-Fa-f]{4})",
            RegexOptions.CultureInvariant);
        if (!match.Success)
        {
            return false;
        }

        vendorId = match.Groups[1].Value.ToUpperInvariant();
        deviceId = match.Groups[2].Value.ToUpperInvariant();
        return true;
    }
    internal static bool TryParseNicItrInput(string text, NicItrProfile profile, out List<ulong> values)
    {
        values = [];
        string[] parts = (text ?? string.Empty)
            .Split([',', ';', ' ', '\t'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length == 0 || (parts.Length != 1 && parts.Length != profile.MaxQueues))
        {
            return false;
        }

        foreach (string part in parts)
        {
            if (!TryParseUInt64Flexible(part, out ulong parsed))
            {
                values.Clear();
                return false;
            }

            values.Add(parsed & profile.ReadMask);
        }

        return values.Count > 0;
    }
    internal static bool TryParseUInt64Flexible(string text, out ulong value)
    {
        value = 0;
        string trimmed = text.Trim();
        if (trimmed.StartsWith("0x", StringComparison.OrdinalIgnoreCase))
        {
            return ulong.TryParse(trimmed[2..], NumberStyles.HexNumber, CultureInfo.InvariantCulture, out value);
        }

        return ulong.TryParse(trimmed, NumberStyles.Integer, CultureInfo.InvariantCulture, out value);
    }
    internal static string FormatNicItrValueList(IReadOnlyList<ulong> values, NicItrProfile profile)
    {
        if (values.Count == 0)
        {
            return "-";
        }

        string[] formatted = values.Select(v => FormatNicItrValue(v, profile)).ToArray();
        return string.Join(", ", formatted);
    }
    internal static string FormatNicItrValue(ulong value, NicItrProfile profile)
    {
        value &= profile.ReadMask;
        if (value == 0)
        {
            return "0x0";
        }

        int digits = profile.ReadWidth == 16 ? 4 : 8;
        return "0x" + value.ToString($"X{digits}", CultureInfo.InvariantCulture);
    }
    internal static string FormatNicItrLabeledValueList(IReadOnlyList<ulong> values, NicItrProfile profile)
    {
        if (values.Count == 0)
        {
            return "-";
        }

        string[] formatted = values
            .Select((value, index) => $"{GetNicItrVectorLabel(profile, index)}={FormatNicItrValue(value, profile)}")
            .ToArray();
        return string.Join(", ", formatted);
    }
    internal static string GetNicItrVectorLabel(NicItrProfile profile, int index)
    {
        if (profile.VectorLabels is not null && index >= 0 && index < profile.VectorLabels.Length)
        {
            return profile.VectorLabels[index];
        }

        return profile.MaxQueues == 1
            ? "ITR"
            : profile.TimingKind == NicItrTimingKind.IntelEitr ? $"EITR{index}" : $"Q{index}";
    }
    internal static string FormatNicItrQueueSummary(IReadOnlyList<ulong> values, NicItrProfile profile)
    {
        if (values.Count == 0)
        {
            return "current: -";
        }

        if (profile.MaxQueues == 1 || values.Count == 1)
        {
            return $"current: {FormatNicItrValue(values[0], profile)} | {FormatNicItrTimingDetail(values[0], profile)}";
        }

        string[] states = values
            .Select((value, index) =>
            {
                string state = (value & profile.ReadMask) == 0 ? "off" : "active";
                return $"{GetNicItrVectorLabel(profile, index)} {state}";
            })
            .ToArray();

        int stateWidth = states.Select(state => state.Length).DefaultIfEmpty(1).Max();
        for (int index = 0; index < states.Length - 1; index++)
        {
            states[index] = states[index].PadRight(stateWidth);
        }

        return "current: " + string.Join(" | ", states);
    }
    internal static string FormatNicItrQueueDetailText(IReadOnlyList<ulong> values, NicItrProfile profile, bool preview = false)
    {
        if (values.Count == 0)
        {
            return preview ? "preview: -" : "queues: -";
        }

        if (values.Count > 4)
        {
            string prefix = preview ? "preview: " : "queues: ";
            return prefix + FormatNicItrTimingList(values, profile);
        }

        string[] labels = values.Select((_, index) => GetNicItrVectorLabel(profile, index)).ToArray();
        string[] rawValues = values.Select(value => FormatNicItrValue(value, profile)).ToArray();
        string[] details = values.Select(value => FormatNicItrTimingDetail(value, profile)).ToArray();
        int labelWidth = labels.Select(label => label.Length).DefaultIfEmpty(1).Max();
        int rawWidth = rawValues.Select(raw => raw.Length).DefaultIfEmpty(1).Max();
        int rxWidth = details
            .Select(detail => detail.Split([", "], 2, StringSplitOptions.None)[0])
            .Where(detail => detail.StartsWith("RX ", StringComparison.Ordinal))
            .Select(detail => detail.Length)
            .DefaultIfEmpty(0)
            .Max();

        string[] rows = values
            .Select((_, index) =>
            {
                string detail = details[index];
                string[] timingColumns = detail.Split([", "], 2, StringSplitOptions.None);
                if (timingColumns.Length == 2 && timingColumns[0].StartsWith("RX ", StringComparison.Ordinal))
                {
                    detail = $"{timingColumns[0].PadRight(rxWidth)} | {timingColumns[1]}";
                }

                return $"{labels[index].PadRight(labelWidth)}: {rawValues[index].PadRight(rawWidth)} | {detail}";
            })
            .ToArray();

        return string.Join(Environment.NewLine, rows);
    }
    internal static string FormatNicItrTimingList(IReadOnlyList<ulong> values, NicItrProfile profile)
    {
        if (values.Count == 0 || profile.TimingKind == NicItrTimingKind.None)
        {
            return "n/a";
        }

        string[] formatted = values
            .Select((value, index) => $"{GetNicItrVectorLabel(profile, index)}={FormatNicItrTiming(value, profile)}")
            .ToArray();
        return string.Join(", ", formatted);
    }
    internal static string FormatNicItrTiming(ulong raw, NicItrProfile profile)
    {
        raw &= profile.ReadMask;
        return profile.TimingKind switch
        {
            NicItrTimingKind.IntelEitr => FormatDurationUs(((raw >> 2) & 0x1FFF) * profile.IntelEitrUnitUs),
            NicItrTimingKind.IntelItr => FormatDurationNs(raw * 256),
            NicItrTimingKind.RealtekIntrMit => FormatRealtekIntrMit(raw, timerUnitUs: 125, extended: false),
            NicItrTimingKind.RealtekIntrMitV2 => FormatRealtekIntrMit(raw, timerUnitUs: 1, extended: true),
            _ => "n/a",
        };
    }
    internal static string FormatNicItrTimingDetail(ulong raw, NicItrProfile profile)
    {
        raw &= profile.ReadMask;
        if (raw == 0)
        {
            return "Off";
        }

        return profile.TimingKind switch
        {
            NicItrTimingKind.RealtekIntrMit => FormatRealtekIntrMitDetail(raw, timerUnitUs: 125, extended: false),
            NicItrTimingKind.RealtekIntrMitV2 => FormatRealtekIntrMitDetail(raw, timerUnitUs: 1, extended: true),
            NicItrTimingKind.IntelEitr => "Delay " + FormatNicItrTiming(raw, profile),
            NicItrTimingKind.IntelItr => "Delay " + FormatNicItrTiming(raw, profile),
            _ => FormatNicItrTiming(raw, profile),
        };
    }
    internal static string FormatDurationUs(ulong us)
    {
        if (us == 0)
        {
            return "Off";
        }

        return us >= 1000
            ? (us / 1000.0).ToString("0.###", CultureInfo.InvariantCulture) + " ms"
            : us.ToString(CultureInfo.InvariantCulture) + " us";
    }
    internal static string FormatDurationNs(ulong ns)
    {
        if (ns == 0)
        {
            return "Off";
        }

        if (ns >= 1_000_000)
        {
            return (ns / 1_000_000.0).ToString("0.###", CultureInfo.InvariantCulture) + " ms";
        }

        if (ns >= 1000)
        {
            return (ns / 1000.0).ToString("0.###", CultureInfo.InvariantCulture) + " us";
        }

        return ns.ToString(CultureInfo.InvariantCulture) + " ns";
    }
    internal static string FormatRealtekIntrMit(ulong raw, uint timerUnitUs, bool extended)
    {
        if (raw == 0)
        {
            return "Off";
        }

        ulong timerMask = extended ? 0x7FUL : 0xFUL;
        ulong frameMask = extended ? 0x7FUL : 0xFUL;
        int rxFrameShift = extended ? 8 : 4;
        int txTimerShift = extended ? 16 : 8;
        int txFrameShift = extended ? 24 : 12;

        ulong rxTimer = raw & timerMask;
        ulong rxFrames = (raw >> rxFrameShift) & frameMask;
        ulong txTimer = (raw >> txTimerShift) & timerMask;
        ulong txFrames = (raw >> txFrameShift) & frameMask;

        string rx = $"{rxTimer * timerUnitUs}us/{rxFrames}f";
        string tx = $"{txTimer * timerUnitUs}us/{txFrames}f";
        return $"Rx:{rx} Tx:{tx}";
    }
    internal static string FormatRealtekIntrMitDetail(ulong raw, uint timerUnitUs, bool extended)
    {
        if (raw == 0)
        {
            return "Off";
        }

        ulong timerMask = extended ? 0x7FUL : 0xFUL;
        ulong frameMask = extended ? 0x7FUL : 0xFUL;
        int rxFrameShift = extended ? 8 : 4;
        int txTimerShift = extended ? 16 : 8;
        int txFrameShift = extended ? 24 : 12;

        ulong rxTimer = raw & timerMask;
        ulong rxFrames = (raw >> rxFrameShift) & frameMask;
        ulong txTimer = (raw >> txTimerShift) & timerMask;
        ulong txFrames = (raw >> txFrameShift) & frameMask;

        string rx = $"RX {rxTimer * timerUnitUs} us / {rxFrames}f";
        string tx = $"TX {txTimer * timerUnitUs} us / {txFrames}f";
        return $"{rx}, {tx}";
    }
}
