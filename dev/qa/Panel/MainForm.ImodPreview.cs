using System.Text.RegularExpressions;

namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private void RefreshTestImodPreview(DeviceBlock block, string reason = "test-preview")
    {
        if (!block.Device.IsTestDevice || !IsUsbImodTarget(block.Device))
        {
            return;
        }

        EnsureImodConfigLoaded();
        uint fallback = (_imodConfigCache?.GlobalInterval ?? ImodDefaultInterval) & 0xFFFF;
        TestDeviceState testState = EnsureTestDeviceState(block.Device);
        if (reason.Contains("set", StringComparison.OrdinalIgnoreCase))
        {
            testState.ImodValue = block.ImodBox.Text?.Trim() ?? FormatImodValue(fallback);
        }
        else if (reason.Contains("delete", StringComparison.OrdinalIgnoreCase)
            || reason.Contains("reset", StringComparison.OrdinalIgnoreCase))
        {
            testState.ImodValue = FormatImodValue(fallback);
        }
        string text = string.IsNullOrWhiteSpace(testState.ImodValue)
            ? FormatImodValue(fallback)
            : testState.ImodValue.Trim();
        if (!TryParseImodInput(text, fallback, out ImodInput parsedInput))
        {
            block.ImodCurrentLabel.Text = "current: invalid input";
            block.ImodCurrentLabel.Tag = block.ImodCurrentLabel.Text;
            block.ImodCurrentLabel.ForeColor = _statusDanger;
            if (block.ImodMapLabel is not null)
            {
                block.ImodMapLabel.Text = "devices: invalid IMOD input";
                block.ImodMapLabel.Tag = block.ImodMapLabel.Text;
                block.ImodMapLabel.ForeColor = _statusDanger;
                SetImodStatusTooltip(block.ImodMapLabel);
            }

            SetImodStatusTooltip(block.ImodCurrentLabel);
            SetImodDetailsVisibility(block, forceVisible: true);
            LogImodUiSnapshot(block, reason);
            WriteLog($"IMOD.TEST.PREVIEW: {block.Device.InstanceId} invalid input=\"{SanitizeLogValue(text)}\"");
            return;
        }

        List<uint> values = BuildTestImodPreviewValues(block.Device, parsedInput, fallback);
        Dictionary<uint, List<string>> labelsByInterrupter = BuildTestImodPreviewLabels(block.Device);
        string mappedDeviceLine = FormatMappedImodDeviceLine(values, labelsByInterrupter, [], FormatUnknownImodDeviceValue(values));
        string visibleInterruptLines = FormatVisibleImodInterrupterLines(values);
        string detailText =
            $"{mappedDeviceLine}\r\n\r\n{visibleInterruptLines}\r\n{FormatDetailedImodInterrupterMap(values, labelsByInterrupter)}\r\n{FormatRawImodInterrupterMap(values)}\r\n{FormatTimeImodInterrupterMap(values)}\r\nsource: test preview only; no driver read/write";

        block.ImodDefaultLabel.Text = $"default: {FormatImodValue(fallback)}";
        block.ImodDefaultLabel.Tag = FormatImodValue(fallback);
        block.ImodCurrentLabel.Text = $"current: {FormatImodValueList(values)}";
        block.ImodCurrentLabel.Tag = $"test preview raw: {FormatImodValueListForLog(values)}";
        block.ImodCurrentLabel.ForeColor = _statusActive;
        SetImodStatusTooltip(block.ImodCurrentLabel);

        if (block.ImodMapLabel is not null)
        {
            block.ImodMapLabel.Text = $"{mappedDeviceLine}\r\n\r\n{visibleInterruptLines}";
            block.ImodMapLabel.Tag = detailText;
            block.ImodMapLabel.ForeColor = _mutedText;
            SetImodStatusTooltip(block.ImodMapLabel);
        }

        if (!block.ImodAutoCheck.Checked && !reason.Equals("test-imod-check", StringComparison.Ordinal))
        {
            block.SuppressImodEvents++;
            try
            {
                block.ImodAutoCheck.Checked = true;
            }
            finally
            {
                block.SuppressImodEvents--;
            }
        }

        SetImodDetailsVisibility(block, forceVisible: false);
        LogImodUiSnapshot(block, reason, values, detailText);
        WriteLog($"IMOD.TEST.PREVIEW: {block.Device.InstanceId} values={FormatImodValueListForLog(values)} reason={reason}");
    }

    private static List<uint> BuildTestImodPreviewValues(DeviceInfo device, ImodInput input, uint fallback)
    {
        const int previewInterrupters = 8;
        uint normalizedFallback = fallback & 0xFFFF;

        if (input.Intervals is { Count: > 0 } intervals)
        {
            List<uint> values = intervals
                .Take(previewInterrupters)
                .Select(value => value & 0xFFFF)
                .ToList();
            while (values.Count < previewInterrupters)
            {
                values.Add(normalizedFallback);
            }

            return values;
        }

        if (input.RoleIntervals is { Count: > 0 } roleIntervals)
        {
            List<uint> values = Enumerable.Repeat(normalizedFallback, previewInterrupters).ToList();
            Dictionary<string, List<string>> labelsByRole = new(StringComparer.OrdinalIgnoreCase);
            AddAdaptiveRoleDisplayLabels(device.UsbRoles, labelsByRole);
            AddAdaptiveRoleDisplayLabels(device.AudioEndpoints, labelsByRole);

            uint currentIntr = 0;
            foreach (string role in AdaptiveRolePriority)
            {
                if (!labelsByRole.ContainsKey(role))
                {
                    continue;
                }

                if (TrySelectAdaptiveRoleInterval(new HashSet<string>(StringComparer.OrdinalIgnoreCase) { role }, roleIntervals, out _, out uint selectedValue))
                {
                    if (currentIntr < previewInterrupters)
                    {
                        values[(int)currentIntr] = selectedValue & 0xFFFF;
                    }
                }
                currentIntr++;
            }

            return values;
        }

        uint interval = (input.Interval ?? normalizedFallback) & 0xFFFF;
        return Enumerable.Repeat(interval, previewInterrupters).ToList();
    }

    private static Dictionary<uint, List<string>> BuildTestImodPreviewLabels(DeviceInfo device)
    {
        Dictionary<string, List<string>> labelsByRole = new(StringComparer.OrdinalIgnoreCase);
        AddAdaptiveRoleDisplayLabels(device.UsbRoles, labelsByRole);
        AddAdaptiveRoleDisplayLabels(device.AudioEndpoints, labelsByRole);

        Dictionary<uint, List<string>> result = [];
        HashSet<string> emitted = new(StringComparer.OrdinalIgnoreCase);
        uint currentIntr = 0;

        foreach (string role in AdaptiveRolePriority)
        {
            if (!labelsByRole.TryGetValue(role, out List<string>? roleLabels))
            {
                continue;
            }

            bool addedAny = false;
            foreach (string label in roleLabels)
            {
                if (emitted.Add(label))
                {
                    if (!result.ContainsKey(currentIntr))
                    {
                        result[currentIntr] = [];
                    }
                    result[currentIntr].Add(label);
                    addedAny = true;
                }
            }

            if (addedAny)
            {
                currentIntr++;
            }
        }

        foreach (KeyValuePair<string, List<string>> pair in labelsByRole.OrderBy(static kvp => kvp.Key, StringComparer.OrdinalIgnoreCase))
        {
            bool addedAny = false;
            foreach (string label in pair.Value)
            {
                if (emitted.Add(label))
                {
                    if (!result.ContainsKey(currentIntr))
                    {
                        result[currentIntr] = [];
                    }
                    result[currentIntr].Add(label);
                    addedAny = true;
                }
            }

            if (addedAny)
            {
                currentIntr++;
            }
        }

        return result;
    }

    private static HashSet<string> BuildTestImodPreviewRoles(DeviceInfo device)
    {
        Dictionary<string, List<string>> labelsByRole = new(StringComparer.OrdinalIgnoreCase);
        AddAdaptiveRoleDisplayLabels(device.UsbRoles, labelsByRole);
        AddAdaptiveRoleDisplayLabels(device.AudioEndpoints, labelsByRole);
        return new HashSet<string>(labelsByRole.Keys, StringComparer.OrdinalIgnoreCase);
    }

}
