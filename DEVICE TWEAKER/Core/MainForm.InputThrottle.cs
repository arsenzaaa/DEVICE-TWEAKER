using Microsoft.Win32;
using System.Globalization;

namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private const string RawMouseThrottleValueName = "RawMouseThrottleDuration";
    private const int DefaultRawMouseThrottleDuration = 20;

    private readonly RawMouseThrottlePreset[] _rawMouseThrottlePresets =
    [
        new("50Hz", 20),
        new("100Hz", 10),
        new("125Hz", 8),
        new("200Hz", 5),
        new("250Hz", 4),
        new("500Hz", 2),
        new("1K", 1),
    ];

    private bool _rawMouseThrottleUiRefreshing;

    private sealed record RawMouseThrottlePreset(string Label, int Duration)
    {
        public override string ToString() => Label;
    }

    private sealed record RawMouseThrottleState(bool Exists, bool IsValid, int Duration, string RawValue);

    private bool HasMouseThrottleContext(DeviceInfo device)
    {
        return device.Kind == DeviceKind.USB
            && device.UsbRoles.Split(',', StringSplitOptions.RemoveEmptyEntries)
                .Any(role => IsUsbRoleText(role, "Mouse"));
    }

    private bool TryReadRawMouseThrottleDuration(out int duration)
    {
        RawMouseThrottleState state = ReadRawMouseThrottleState();
        duration = state.Duration;
        return state.IsValid;
    }

    private string GetRawMouseThrottleStatus()
    {
        RawMouseThrottleState state = ReadRawMouseThrottleState();
        if (!state.Exists)
        {
            return $"Raw input throttle: disabled ({RawMouseThrottleValueName} missing)";
        }

        if (!state.IsValid)
        {
            return $"Raw input throttle: invalid ({RawMouseThrottleValueName}={state.RawValue}, expected DWORD 1-20)";
        }

        return $"Raw input throttle: enabled (current {GetRawMouseThrottleCapTag(state.Duration)}, DWORD={state.Duration})";
    }

    private RawMouseThrottleState ReadRawMouseThrottleState()
    {
        try
        {
            using RegistryKey? mouseKey = Registry.CurrentUser.OpenSubKey(@"Control Panel\Mouse");
            object? value = mouseKey?.GetValue(RawMouseThrottleValueName);
            if (value is null)
            {
                return new RawMouseThrottleState(false, false, 0, "missing");
            }

            string rawValue = FormatRawMouseThrottleRegistryValue(value);
            if (!TryParseRegistryInt(value, out int duration) || duration < 1 || duration > 20)
            {
                return new RawMouseThrottleState(true, false, 0, rawValue);
            }

            return new RawMouseThrottleState(true, true, duration, rawValue);
        }
        catch (Exception ex)
        {
            return new RawMouseThrottleState(true, false, 0, $"read failed: {ex.Message}");
        }
    }

    private void LoadRawMouseThrottleControls(DeviceBlock block)
    {
        if (block.RawMouseThrottleCheck is null || block.RawMouseThrottleCombo is null)
        {
            return;
        }

        _rawMouseThrottleUiRefreshing = true;
        try
        {
            if (block.RawMouseThrottleCombo.Items.Count == 0)
            {
                block.RawMouseThrottleCombo.Items.AddRange(_rawMouseThrottlePresets);
            }

            RawMouseThrottleState state = ReadRawMouseThrottleState();
            RawMouseThrottlePreset selected = GetRawMouseThrottlePreset(
                state.IsValid ? state.Duration : DefaultRawMouseThrottleDuration);
            if (!block.RawMouseThrottleCombo.Items.Contains(selected))
            {
                block.RawMouseThrottleCombo.Items.Add(selected);
            }

            block.RawMouseThrottleCheck.Checked = state.IsValid;
            block.RawMouseThrottleCombo.SelectedItem = selected;
            block.RawMouseThrottleCombo.Enabled = true;
            UpdateRawMouseThrottleStatusLabel(block, state);
        }
        finally
        {
            _rawMouseThrottleUiRefreshing = false;
        }
    }

    private void HandleRawMouseThrottleChanged(DeviceBlock block)
    {
        if (_rawMouseThrottleUiRefreshing
            || block.RawMouseThrottleCheck is null
            || block.RawMouseThrottleCombo is null)
        {
            return;
        }

        int duration = GetSelectedRawMouseThrottleDuration(block);
        try
        {
            SetRawMouseThrottle(block.RawMouseThrottleCheck.Checked, duration);
        }
        catch (Exception ex)
        {
            WriteLog($"USBPOLL.THROTTLE.SET: failed: {ex.Message}");
            ShowThemedInfo($"Failed to update RawMouseThrottleDuration.\n{ex.Message}");
        }

        RefreshRawMouseThrottleUi();
    }

    private int GetSelectedRawMouseThrottleDuration(DeviceBlock block)
    {
        if (block.RawMouseThrottleCombo?.SelectedItem is RawMouseThrottlePreset preset)
        {
            return preset.Duration;
        }

        return DefaultRawMouseThrottleDuration;
    }

    private RawMouseThrottlePreset GetRawMouseThrottlePreset(int duration)
    {
        foreach (RawMouseThrottlePreset preset in _rawMouseThrottlePresets)
        {
            if (preset.Duration == duration)
            {
                return preset;
            }
        }

        return new RawMouseThrottlePreset(GetRawMouseThrottleCapTag(duration), duration);
    }

    private void SetRawMouseThrottle(bool enabled, int duration)
    {
        using RegistryKey? mouseKey = Registry.CurrentUser.CreateSubKey(@"Control Panel\Mouse", writable: true);
        if (mouseKey is null)
        {
            throw new InvalidOperationException(@"HKCU\Control Panel\Mouse is unavailable.");
        }

        if (!enabled)
        {
            mouseKey.DeleteValue(RawMouseThrottleValueName, throwOnMissingValue: false);
            WriteLog($"USBPOLL.THROTTLE.SET: disabled, deleted {RawMouseThrottleValueName}");
            return;
        }

        int clampedDuration = Math.Clamp(duration, 1, 20);
        mouseKey.SetValue(RawMouseThrottleValueName, clampedDuration, RegistryValueKind.DWord);
        WriteLog($"USBPOLL.THROTTLE.SET: enabled {RawMouseThrottleValueName}={clampedDuration} cap={GetRawMouseThrottleCapTag(clampedDuration)}");
    }

    private void ApplyAutoRawMouseThrottle()
    {
        try
        {
            SetRawMouseThrottle(enabled: true, DefaultRawMouseThrottleDuration);
            RefreshRawMouseThrottleUi();
            WriteLog($"AUTO.THROTTLE: {RawMouseThrottleValueName}={DefaultRawMouseThrottleDuration} cap=50Hz");
        }
        catch (Exception ex)
        {
            WriteLog($"AUTO.THROTTLE: failed: {ex.Message}");
        }
    }

    private void RefreshRawMouseThrottleUi()
    {
        foreach (DeviceBlock block in _blocks)
        {
            if (block.RawMouseThrottleCheck is null || block.RawMouseThrottleCombo is null)
            {
                continue;
            }

            LoadRawMouseThrottleControls(block);
            UpdateRawMouseThrottleInfoLine(block);
        }
    }

    private void UpdateRawMouseThrottleStatusLabel(DeviceBlock block, RawMouseThrottleState state)
    {
        if (block.RawMouseThrottleStatusLabel is null)
        {
            return;
        }

        if (!state.Exists)
        {
            block.RawMouseThrottleStatusLabel.Text = "current: off";
            block.RawMouseThrottleStatusLabel.ForeColor = _statusInactive;
            return;
        }

        if (!state.IsValid)
        {
            block.RawMouseThrottleStatusLabel.Text = $"current: invalid ({CompactRawMouseThrottleValue(state.RawValue)})";
            block.RawMouseThrottleStatusLabel.ForeColor = _statusDanger;
            return;
        }

        block.RawMouseThrottleStatusLabel.Text = $"current: {GetRawMouseThrottleCapTag(state.Duration)} (DWORD={state.Duration})";
        block.RawMouseThrottleStatusLabel.ForeColor = _statusActive;
    }

    private void UpdateRawMouseThrottleInfoLine(DeviceBlock block)
    {
        string status = GetRawMouseThrottleStatus();
        string text = block.InfoLabel.Text ?? string.Empty;
        if (string.IsNullOrWhiteSpace(text))
        {
            block.InfoLabel.Text = status;
            return;
        }

        string normalized = text.Replace("\r\n", "\n");
        string[] lines = normalized.Split('\n');
        bool replaced = false;
        for (int i = 0; i < lines.Length; i++)
        {
            if (lines[i].StartsWith("Raw input throttle:", StringComparison.OrdinalIgnoreCase))
            {
                lines[i] = status;
                replaced = true;
            }
        }

        block.InfoLabel.Text = replaced
            ? string.Join(Environment.NewLine, lines)
            : text + Environment.NewLine + status;
    }

    private static string GetRawMouseThrottleCapTag(int duration)
    {
        if (duration <= 0)
        {
            duration = DefaultRawMouseThrottleDuration;
        }

        return FormatPollingRateTag(1000d / duration);
    }

    private static string FormatRawMouseThrottleRegistryValue(object value)
    {
        return value switch
        {
            byte[] bytes => "0x" + Convert.ToHexString(bytes),
            IFormattable formattable => formattable.ToString(null, CultureInfo.InvariantCulture),
            _ => value.ToString() ?? "unknown",
        };
    }

    private static string CompactRawMouseThrottleValue(string value)
    {
        string compact = value.Replace('\r', ' ').Replace('\n', ' ').Trim();
        return compact.Length <= 48 ? compact : compact[..48] + "...";
    }
}
