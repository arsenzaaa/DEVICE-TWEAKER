using System.Globalization;

namespace DeviceTweakerCS;

// Text of the USB IMOD readback: the compact map in the device card and the
// detailed lines in its tooltip and the log. Values are raw IMODI fields (250 ns units).
internal static class ImodMapText
{
    internal static string Value(uint value) => $"0x{value:X}";

    internal static string UsbTime(uint rawValue)
    {
        ulong ns = (rawValue & 0xFFFF) * 250UL;
        if (ns >= 1000)
        {
            double us = ns / 1000.0;
            return us.ToString(us % 1 == 0 ? "0" : "0.###", CultureInfo.InvariantCulture) + " us";
        }

        return ns.ToString(CultureInfo.InvariantCulture) + " ns";
    }

    // Map text for one controller: the card shows device and interrupter lines, the
    // detail adds every interrupter, raw values, times and how the roles were found.
    internal static (string MapText, string Detail) Compose(
        IReadOnlyList<uint> currentIntervals,
        IReadOnlyDictionary<uint, HashSet<string>>? rolesByInterrupter,
        ImodRoleSnapshot roleSnapshot,
        string controllerDeviceId,
        string resolveDetail)
    {
        Dictionary<string, List<string>> displayLabelsByRole = ImodRoleMap.DisplayLabelsByRole(roleSnapshot, controllerDeviceId);
        Dictionary<uint, List<string>> displayLabelsByInterrupter = [];
        HashSet<string> mappedLabelsSeen = new(StringComparer.OrdinalIgnoreCase);

        if (rolesByInterrupter is not null)
        {
            foreach (KeyValuePair<uint, HashSet<string>> pair in rolesByInterrupter.OrderBy(static kvp => kvp.Key))
            {
                if (pair.Key >= currentIntervals.Count || pair.Value.Count == 0)
                {
                    continue;
                }

                List<string> mappedLabels = [];
                HashSet<string> localLabelsSeen = new(StringComparer.OrdinalIgnoreCase);
                foreach (string label in ImodRoleMap.DisplayRoles(pair.Value, displayLabelsByRole))
                {
                    string cleanLabel = DeviceMapLabel(label);
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

        string unknownDeviceValue = UnknownDeviceValue(currentIntervals);
        List<string> unknownLabels = [];
        HashSet<string> unknownLabelsSeen = new(StringComparer.OrdinalIgnoreCase);
        foreach (string label in ImodRoleMap.ControllerDeviceDisplayLabels(roleSnapshot, controllerDeviceId))
        {
            string cleanLabel = DeviceMapLabel(label);
            if (!mappedLabelsSeen.Contains(cleanLabel) && unknownLabelsSeen.Add(cleanLabel))
            {
                unknownLabels.Add(cleanLabel);
            }
        }

        string mappedDeviceLine = MappedDeviceLine(currentIntervals, displayLabelsByInterrupter, unknownLabels, unknownDeviceValue);
        string visibleInterruptLines = VisibleInterrupterLines(currentIntervals);
        string detailedInterruptLine = Detailed(currentIntervals, displayLabelsByInterrupter);
        string rawLine = Raw(currentIntervals);
        string timeLine = Time(currentIntervals);

        string mapText = $"{mappedDeviceLine}\r\n\r\n{visibleInterruptLines}";
        string detail = string.IsNullOrWhiteSpace(resolveDetail)
            ? $"{mappedDeviceLine}\r\n\r\n{visibleInterruptLines}\r\n{detailedInterruptLine}\r\n{rawLine}\r\n{timeLine}"
            : $"{mappedDeviceLine}\r\n\r\n{visibleInterruptLines}\r\n{detailedInterruptLine}\r\n{rawLine}\r\n{timeLine}\r\nmap: {resolveDetail}";
        return (mapText, detail);
    }

    internal static string ValueList(IReadOnlyList<uint> values)
    {
        if (values.Count == 0)
        {
            return "-";
        }

        uint first = values[0];
        if (values.All(v => v == first))
        {
            return values.Count > 1
                ? $"{Value(first)} x{values.Count}"
                : Value(first);
        }

        return "varies by interrupter";
    }

    internal static string ValueListForLog(IReadOnlyList<uint> values)
    {
        if (values.Count == 0)
        {
            return "-";
        }

        return string.Join(",", values.Select(value => $"0x{value & 0xFFFF:X4}"));
    }

    internal static string Raw(IReadOnlyList<uint> values)
    {
        if (values.Count == 0)
        {
            return "raw: -";
        }

        return "raw: " + string.Join(" | ", values.Select((value, index) => $"I{index}:{Value(value)}"));
    }

    internal static string MappedDeviceLine(
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
                parts.Add($"{label} -> intr{pair.Key}={Value(value)}/{UsbTime(value)}");
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
            ? DeviceFirstLines(parts)
            : "devices: no exact USB role map";
    }

    // Aligns "label -> intrN=value/time" parts into columns under one "devices:" header.
    internal static string DeviceFirstLines(IReadOnlyList<string> parts)
    {
        List<(string Role, string Intr, string Value, string Time)> rows = [];
        foreach (string part in parts)
        {
            string[] roleAndValue = part.Split([" -> "], 2, StringSplitOptions.None);
            string role = roleAndValue[0].Trim();
            string assignment = roleAndValue.Length == 2 ? roleAndValue[1].Trim() : string.Empty;
            string[] intrAndValue = assignment.Split(['='], 2);
            string intr = intrAndValue[0].Trim();
            string valueAndTime = intrAndValue.Length == 2 ? intrAndValue[1].Trim() : string.Empty;
            string[] valueParts = valueAndTime.Split(['/'], 2);
            rows.Add((
                role,
                intr,
                valueParts[0].Trim(),
                valueParts.Length == 2 ? valueParts[1].Trim() : string.Empty));
        }

        int roleWidth = rows.Select(row => row.Role.Length).DefaultIfEmpty(0).Max();
        int intrWidth = rows.Select(row => row.Intr.Length).DefaultIfEmpty(0).Max();
        int valueWidth = rows.Select(row => row.Value.Length).DefaultIfEmpty(0).Max();
        return string.Join("\r\n", rows.Select((row, index) =>
        {
            string prefix = index == 0 ? "devices: " : "         ";
            string time = string.IsNullOrEmpty(row.Time) ? string.Empty : $"  {row.Time}";
            return $"{prefix}{row.Role.PadRight(roleWidth)} -> {row.Intr.PadRight(intrWidth)}  {row.Value.PadRight(valueWidth)}{time}";
        }));
    }

    internal static string VisibleInterrupterLines(IReadOnlyList<uint> values)
    {
        // Keep lines short enough for the map box width so WordWrap never splits
        // mid-token (e.g. "intr3=" on one line and "0xC8/50us" on the next).
        const int chunkSize = 2;
        const int maxVisible = 8;

        if (values.Count == 0)
        {
            return "interrupters: -";
        }

        int visibleCount = Math.Min(values.Count, maxVisible);
        List<string> lines = [];
        List<(string Intr, string Value, string Time)> formatted = Enumerable.Range(0, visibleCount)
            .Select(index => ($"intr{index}", Value(values[index]), UsbTime(values[index])))
            .ToList();
        int intrWidth = formatted.Select(part => part.Intr.Length).DefaultIfEmpty(0).Max();
        int valueWidth = formatted.Select(part => part.Value.Length).DefaultIfEmpty(0).Max();
        int timeWidth = formatted.Select(part => part.Time.Length).DefaultIfEmpty(0).Max();

        for (int start = 0; start < visibleCount; start += chunkSize)
        {
            int end = Math.Min(start + chunkSize, visibleCount);
            IReadOnlyList<(string Intr, string Value, string Time)> parts = formatted.Skip(start).Take(end - start).ToList();

            string suffix = string.Empty;
            if (end == visibleCount && values.Count > visibleCount)
            {
                suffix = $" | +{values.Count - visibleCount} more";
            }

            string[] columns = parts
                .Select(part => $"{part.Intr.PadRight(intrWidth)}  {part.Value.PadRight(valueWidth)}  {part.Time.PadRight(timeWidth)}")
                .ToArray();
            columns[^1] = columns[^1].TrimEnd();
            string rowText = string.Join(" | ", columns);
            lines.Add($"interrupters {start}-{end - 1}: {rowText}{suffix}");
        }

        return string.Join("\r\n", lines);
    }

    internal static string Detailed(
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
            parts.Add($"I{i}{labelText}={Value(values[i])} ({UsbTime(values[i])})");
        }

        return "detail: " + string.Join(" | ", parts);
    }

    internal static string Time(IReadOnlyList<uint> values)
    {
        if (values.Count == 0)
        {
            return "time: -";
        }

        return "time: " + string.Join(" | ", values.Select((value, index) => $"I{index}:{UsbTime(value)}"));
    }

    internal static string UnknownDeviceValue(IReadOnlyList<uint> values)
    {
        if (values.Count == 0)
        {
            return "-";
        }

        uint first = values[0];
        return values.All(value => value == first)
            ? Value(first)
            : "mixed";
    }

    private static string DeviceMapLabel(string label)
    {
        string clean = string.IsNullOrWhiteSpace(label) ? "Device" : label.Trim();
        clean = clean.Replace(" | ", "/", StringComparison.Ordinal);
        clean = clean.Replace("=", "-", StringComparison.Ordinal);
        return clean;
    }
}
