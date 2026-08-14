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

    /// <summary>
    /// RawMouseThrottle* is a Windows 11 input-stack feature from the July 2023
    /// update (KB5028185 / Moment 3; preview KB5027303). Win10 has no throttle
    /// path in win32k. Registry writes are inert no-ops for the OS.
    /// </summary>
    private static bool SupportsRawMouseThrottleOs()
    {
        if (!OperatingSystem.IsWindowsVersionAtLeast(10, 0, 22621))
        {
            return false;
        }

        // The feature first shipped in 22621.1928 (KB5027303 preview) and was
        // then serviced broadly by KB5028185. A build-only check incorrectly
        // enables the control on an unpatched 22H2 installation.
        if (Environment.OSVersion.Version.Build > 22621)
        {
            return true;
        }

        try
        {
            using RegistryKey? currentVersion = Registry.LocalMachine.OpenSubKey(
                @"SOFTWARE\Microsoft\Windows NT\CurrentVersion");
            object? ubrValue = currentVersion?.GetValue("UBR");
            int ubr = ubrValue is int value
                ? value
                : int.TryParse(ubrValue?.ToString(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed)
                    ? parsed
                    : 0;
            return ubr >= 1928;
        }
        catch
        {
            return false;
        }
    }

    private static string GetRawMouseThrottleToolTip(bool osSupported)
    {
        const string whatItDoes =
            "Windows limits how often background apps receive raw mouse input.\n"
            + "This reduces the message load created by throttling and coalescing input.\n"
            + "Mice with a polling rate from 1 to 8 kHz can create a high message load.\n"
            + "Foreground apps keep the full input rate.\n"
            + "Background raw input listeners receive a lower message rate.";

        const string how =
            "The setting writes RawMouseThrottleDuration to HKCU\\Control Panel\\Mouse.\n"
            + "The accepted DWORD range is from 1 to 20 milliseconds.\n"
            + "A larger value produces a lower rate for background listeners.\n"
            + "20 milliseconds is approximately 50 Hz. 8 milliseconds is approximately 125 Hz. 1 millisecond is approximately 1 kHz.";

        const string when =
            "The setting first appeared in the Windows 11 22H2 preview build 22621.1928 with KB5027303.\n"
            + "It was included in the July 2023 update KB5028185.";

        if (osSupported)
        {
            return whatItDoes + "\n\n" + how + "\n\n" + when;
        }

        return whatItDoes + "\n\n" + when
            + "\n\nThis setting is not available on Windows 10 or an older unpatched Windows 11 22H2 build.";
    }

    private bool HasMouseThrottleContext(DeviceInfo device)
    {
        // Show for USB mice on all OS versions; enable only on Win11+ (see SupportsRawMouseThrottleOs).
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

            bool osSupported = SupportsRawMouseThrottleOs();
            RawMouseThrottlePreset selected = GetRawMouseThrottlePreset(DefaultRawMouseThrottleDuration);
            if (!osSupported)
            {
                if (!block.RawMouseThrottleCombo.Items.Contains(selected))
                {
                    block.RawMouseThrottleCombo.Items.Add(selected);
                }

                block.RawMouseThrottleCheck.Checked = false;
                block.RawMouseThrottleCombo.SelectedItem = selected;
                block.RawMouseThrottleCheck.Enabled = false;
                block.RawMouseThrottleCombo.Enabled = false;
                if (block.RawMouseThrottleStatusLabel is not null)
                {
                    block.RawMouseThrottleStatusLabel.Text = "current: unavailable (Windows 11 22H2+)";
                    block.RawMouseThrottleStatusLabel.ForeColor = _statusInactive;
                }

                return;
            }

            RawMouseThrottleState state = ReadRawMouseThrottleState();
            selected = GetRawMouseThrottlePreset(
                state.IsValid ? state.Duration : DefaultRawMouseThrottleDuration);
            if (!block.RawMouseThrottleCombo.Items.Contains(selected))
            {
                block.RawMouseThrottleCombo.Items.Add(selected);
            }

            block.RawMouseThrottleCheck.Enabled = true;
            block.RawMouseThrottleCombo.Enabled = true;
            block.RawMouseThrottleCheck.Checked = state.IsValid;
            block.RawMouseThrottleCombo.SelectedItem = selected;
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
            || !SupportsRawMouseThrottleOs()
            || block.RawMouseThrottleCheck is null
            || block.RawMouseThrottleCombo is null)
        {
            return;
        }

        if (block.Device.IsTestDevice)
        {
            int previewDuration = GetSelectedRawMouseThrottleDuration(block);
            bool enabled = block.RawMouseThrottleCheck.Checked;
            WriteLog(
                $"USBPOLL.THROTTLE.TEST: preview only enabled={enabled} duration={previewDuration} " +
                $"(no registry write for test device {block.Device.InstanceId})");
            if (block.RawMouseThrottleStatusLabel is not null)
            {
                block.RawMouseThrottleStatusLabel.Text = enabled
                    ? $"current: preview {GetRawMouseThrottleCapTag(previewDuration)} (DWORD={previewDuration})"
                    : "current: preview off";
                block.RawMouseThrottleStatusLabel.ForeColor = _statusActive;
            }

            UpdateRawMouseThrottleInfoLine(block);
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
        if (!SupportsRawMouseThrottleOs())
        {
            throw new InvalidOperationException("RawMouseThrottleDuration requires Windows 11+.");
        }

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

    private void ApplyAutoRawMouseThrottle(OperationReport? report = null)
    {
        if (!SupportsRawMouseThrottleOs())
        {
            WriteLog("AUTO.THROTTLE: skipped (RawMouseThrottle requires Windows 11+)");
            return;
        }

        try
        {
            SetRawMouseThrottle(enabled: true, DefaultRawMouseThrottleDuration);
            RefreshRawMouseThrottleUi();
            WriteLog($"AUTO.THROTTLE: {RawMouseThrottleValueName}={DefaultRawMouseThrottleDuration} cap=50Hz");
        }
        catch (Exception ex)
        {
            WriteLog($"AUTO.THROTTLE: failed: {ex.Message}");
            report?.AddError("Raw mouse throttle", ex.Message);
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
