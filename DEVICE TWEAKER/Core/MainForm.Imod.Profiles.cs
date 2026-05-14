namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private const string ImodModeSingle = "XHCI";
    private const string ImodModeVector = "Interrupters";
    private const string ImodModeRoles = "Devices";

    private static string GetDefaultRoleImodText(DeviceInfo? device = null)
    {
        Dictionary<string, uint> roleIntervals = new(StringComparer.OrdinalIgnoreCase)
        {
            ["Mouse"] = 0x0,
            ["Keyboard"] = ImodDefaultInterval,
            ["Audio"] = 0xFA0,
            ["Gamepad"] = ImodDefaultInterval,
        };

        if (device is not null)
        {
            HashSet<string> availableRoles = new(StringComparer.OrdinalIgnoreCase);
            foreach (string role in GetImodDeviceEditorRoles(device))
            {
                availableRoles.Add(role);
            }

            if (availableRoles.Count > 0)
            {
                Dictionary<string, uint> filtered = new(StringComparer.OrdinalIgnoreCase);
                foreach (string role in ImodDeviceEditorRoleOrder)
                {
                    if (availableRoles.Contains(role) && roleIntervals.TryGetValue(role, out uint value))
                    {
                        filtered[role] = value;
                    }
                }

                if (filtered.Count > 0)
                {
                    roleIntervals = filtered;
                }
            }
        }

        return FormatImodRoleIntervals(roleIntervals);
    }

    private void InitializeImodSelectors(DeviceBlock block)
    {
        if (block.ImodModeCombo is null)
        {
            return;
        }

        block.SuppressImodEvents++;
        try
        {
            if (block.ImodModeCombo.Items.Count == 0)
            {
                block.ImodModeCombo.Items.AddRange([ImodModeSingle, ImodModeVector, ImodModeRoles]);
            }

            block.ImodModeCombo.SelectedItem = ImodModeSingle;
        }
        finally
        {
            block.SuppressImodEvents--;
        }
    }

    private void HandleImodModeChanged(DeviceBlock block)
    {
        if (block.SuppressImodEvents > 0 || block.ImodModeCombo is null)
        {
            return;
        }

        string mode = block.ImodModeCombo.SelectedItem?.ToString() ?? ImodModeSingle;
        string currentText = block.ImodBox.Text?.Trim() ?? string.Empty;
        string nextText = currentText;

        if (mode == ImodModeRoles)
        {
            if (!TryParseImodRoleIntervals(currentText, out _))
            {
                nextText = GetDefaultRoleImodText(block.Device);
            }
        }
        else if (mode == ImodModeVector)
        {
            if (!TryParseImodIntervalList(currentText, out List<uint> vector) || vector.Count <= 1)
            {
                uint seed = TryParseUInt32Flexible(currentText, out uint parsed) ? parsed : ImodDefaultInterval;
                nextText = FormatImodVector([seed, seed]);
            }
        }
        else
        {
            if (TryParseImodRoleIntervals(currentText, out Dictionary<string, uint> roleValues) && roleValues.Count > 0)
            {
                nextText = FormatImodValue(ImodDefaultInterval);
            }
            else if (TryParseImodIntervalList(currentText, out List<uint> vector) && vector.Count > 0)
            {
                nextText = FormatImodValue(vector[0]);
            }
        }

        if (!string.Equals(nextText, currentText, StringComparison.Ordinal))
        {
            block.SuppressImodEvents++;
            try
            {
                block.ImodBox.Text = nextText;
                block.ImodAutoCheck.Checked = true;
            }
            finally
            {
                block.SuppressImodEvents--;
            }
        }

        UpdateImodSelectorsFromText(block);
    }

    private void UpdateImodSelectorsFromText(DeviceBlock block)
    {
        if (block.ImodModeCombo is null)
        {
            return;
        }

        string text = block.ImodBox.Text?.Trim() ?? string.Empty;
        string mode = GetImodTextMode(text);

        block.SuppressImodEvents++;
        try
        {
            block.ImodModeCombo.SelectedItem = mode;
        }
        finally
        {
            block.SuppressImodEvents--;
        }
    }

    private static string GetImodTextMode(string text)
    {
        if (TryParseImodRoleIntervals(text, out Dictionary<string, uint> roleValues) && roleValues.Count > 0)
        {
            return ImodModeRoles;
        }

        if (TryParseImodIntervalList(text, out List<uint> values) && values.Count > 1)
        {
            return ImodModeVector;
        }

        return ImodModeSingle;
    }
}
