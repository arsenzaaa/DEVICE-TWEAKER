namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        if (_testDevicesOnly && _testAutoDryRun
            && _devicesScroll is not null
            && _devicesScroll.Visible
            && !HasActiveInputControl()
            && (keyData is Keys.Home or Keys.End or Keys.PageUp or Keys.PageDown))
        {
            int before = _devicesScroll.Value;
            int page = Math.Max(_devicesScroll.SmallChange, _devicesScroll.ViewportSize - UiScale(40));
            int target = (keyData & Keys.KeyCode) switch
            {
                Keys.Home => 0,
                Keys.End => _devicesScroll.Maximum,
                Keys.PageUp => before - page,
                Keys.PageDown => before + page,
                _ => before
            };

            _devicesScroll.Value = target;
            WriteLog(
                $"TEST.QA.SCROLL: key={keyData & Keys.KeyCode} before={before} after={_devicesScroll.Value} " +
                $"page={page} maximum={_devicesScroll.Maximum} viewport={_devicesScroll.ViewportSize}");
            return true;
        }

        return base.ProcessCmdKey(ref msg, keyData);
    }

    private void OnMainFormKeyDown(object? sender, KeyEventArgs e)
    {
        // Preserve caret movement, popup navigation and text undo in editors.
        // KeyPreview otherwise delivers these keys to the form before the field.
        if (HasActiveInputControl()
            && ((e.Modifiers == Keys.None && e.KeyCode is Keys.Home or Keys.End or Keys.PageUp or Keys.PageDown)
                || (e.Control && !e.Alt && !e.Shift && e.KeyCode == Keys.Z)))
        {
            return;
        }

        if (_devicesScroll is not null
            && _devicesScroll.Visible
            && e.Modifiers == Keys.None
            && e.KeyCode is Keys.Home or Keys.End or Keys.PageUp or Keys.PageDown)
        {
            int before = _devicesScroll.Value;
            int page = Math.Max(_devicesScroll.SmallChange, _devicesScroll.ViewportSize - UiScale(40));
            int target = e.KeyCode switch
            {
                Keys.Home => 0,
                Keys.End => _devicesScroll.Maximum,
                Keys.PageUp => before - page,
                Keys.PageDown => before + page,
                _ => before
            };

            _devicesScroll.Value = target;
            e.Handled = true;
            e.SuppressKeyPress = true;
            if (_testDevicesOnly && _testAutoDryRun)
            {
                WriteLog(
                    $"TEST.QA.SCROLL: key={e.KeyCode} before={before} after={_devicesScroll.Value} " +
                    $"page={page} maximum={_devicesScroll.Maximum} viewport={_devicesScroll.ViewportSize}");
            }
            return;
        }

        if ((e.KeyCode == Keys.F5 && e.Modifiers == Keys.None)
            || (e.Control && !e.Alt && !e.Shift && e.KeyCode == Keys.R))
        {
            e.Handled = true;
            e.SuppressKeyPress = true;
            WriteLog("UI: Shortcut F5 / Ctrl+R triggered REFRESH");
            _btnScanRef?.PerformClick();
            return;
        }

        if (e.Control && !e.Alt && !e.Shift && e.KeyCode == Keys.S)
        {
            e.Handled = true;
            e.SuppressKeyPress = true;
            WriteLog("UI: Shortcut Ctrl+S triggered APPLY");
            _btnApplyRef?.PerformClick();
            return;
        }

        if (e.Control && !e.Alt && !e.Shift && e.KeyCode == Keys.O)
        {
            e.Handled = true;
            e.SuppressKeyPress = true;
            WriteLog("UI: Shortcut Ctrl+O triggered AUTO-OPTIMIZATION");
            _btnAutoRef?.PerformClick();
            return;
        }

        if (e.Control && !e.Alt && !e.Shift && e.KeyCode == Keys.Z)
        {
            e.Handled = true;
            e.SuppressKeyPress = true;
            WriteLog("UI: Shortcut Ctrl+Z triggered RESTORE");
            _btnRestoreRef?.PerformClick();
            return;
        }

        if (e.Control && !e.Alt && !e.Shift && e.KeyCode == Keys.F)
        {
            e.Handled = true;
            e.SuppressKeyPress = true;
            WriteLog("UI: Shortcut Ctrl+F triggered filter search");
            FocusFilterBox();
            return;
        }

        if (e.KeyCode == Keys.Escape && e.Modifiers == Keys.None)
        {
            if (!string.IsNullOrEmpty(_searchFilterText) || (_searchFilterBox?.Inner.Focused == true))
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                WriteLog("UI: Shortcut Escape triggered filter clear");
                ClearFilterBox();
                return;
            }
        }

#if DEVELOPER_QA
        if (e.Control && e.Shift && e.KeyCode == Keys.T)
        {
            e.Handled = true;
            e.SuppressKeyPress = true;
            WriteLog("UI: TEST ADMIN hotkey");
            ShowTestAdminDialog();
        }
#endif
    }

    private bool HasActiveInputControl()
    {
        Control? active = ActiveControl;
        while (active is ContainerControl container && container.ActiveControl is Control child)
            active = child;
        return active is TextBoxBase or UpDownBase or ComboBox or ThemedDropDownPicker
            || active is ThemedTextBox;
    }

    private void UpdateCpuHeaderUi()
    {
        if (_cpuHeaderLabel is not null)
        {
            _cpuHeaderLabel.Text = _cpuHeaderText;
        }

        if (_htPrefixLabel is null || _htStatusLabel is null)
        {
            return;
        }

        static string OnOff(bool enabled) => enabled ? "On" : "Off";

        void SetHeaderFlag(Label? prefixLabel, Label? statusLabel, string name, string value, bool active)
        {
            if (prefixLabel is null || statusLabel is null)
            {
                return;
            }

            prefixLabel.Text = $"{UiLanguage.Text(name)} -";
            statusLabel.Text = UiLanguage.Text(value);
            statusLabel.Visible = true;
            statusLabel.ForeColor = active ? _statusActive : _statusInactive;
        }

        void AddHeaderFlag(Label? prefixLabel, Label? statusLabel)
        {
            if (_cpuFlagsPanel is null || prefixLabel is null || statusLabel is null)
            {
                return;
            }

            if (_cpuFlagsPanel.Controls.Count > 0)
            {
                _cpuFlagsPanel.Controls.Add(new Label
                {
                    Text = "|",
                    AutoSize = true,
                    Font = _htFont,
                    ForeColor = _statusSeparator,
                    Margin = new Padding(UiScale(14), 0, UiScale(14), 0),
                });
            }

            _cpuFlagsPanel.Controls.Add(prefixLabel);
            _cpuFlagsPanel.Controls.Add(statusLabel);
        }

        bool IsReusableHeaderFlagControl(Control control)
        {
            return ReferenceEquals(control, _htPrefixLabel)
                || ReferenceEquals(control, _htStatusLabel)
                || ReferenceEquals(control, _hybridCpuPrefixLabel)
                || ReferenceEquals(control, _hybridCpuStatusLabel)
                || ReferenceEquals(control, _cppcPrefixLabel)
                || ReferenceEquals(control, _cppcStatusLabel)
                || ReferenceEquals(control, _dualCcdPrefixLabel)
                || ReferenceEquals(control, _dualCcdStatusLabel)
                || ReferenceEquals(control, _sandboxPrefixLabel)
                || ReferenceEquals(control, _sandboxStatusLabel);
        }

        bool smtEnabled = _cpuInfo?.Topology.ByCore.Values.Any(g => g.Count > 1) == true;
        if (_smtText.Contains("DISABLED", StringComparison.OrdinalIgnoreCase)
            || _smtText.Contains("OFF", StringComparison.OrdinalIgnoreCase))
        {
            smtEnabled = false;
        }
        else if (_smtText.Contains("ENABLED", StringComparison.OrdinalIgnoreCase)
            || _smtText.Contains("ON", StringComparison.OrdinalIgnoreCase))
        {
            smtEnabled = true;
        }

        string threadingName = _smtText.Contains("Hyper-Threading", StringComparison.OrdinalIgnoreCase)
            ? "Hyper-Threading"
            : "SMT";
        bool hasHybridCpu = HasHybridCpu();
        bool hasDualCcdCpu = HasDualCcdCpu();
        bool sandboxOn = IsSandboxDryRunActive()
            && !string.Equals(Environment.GetEnvironmentVariable("DEVICE_TWEAKER_QA_HIDE_SANDBOX_HEADER"), "1", StringComparison.Ordinal);

        SetHeaderFlag(_htPrefixLabel, _htStatusLabel, threadingName, OnOff(smtEnabled), smtEnabled);
        SetHeaderFlag(_hybridCpuPrefixLabel, _hybridCpuStatusLabel, "Hybrid CPU", OnOff(hasHybridCpu), hasHybridCpu);
        SetHeaderFlag(_cppcPrefixLabel, _cppcStatusLabel, "CPPC", OnOff(_cppcEnabled), _cppcEnabled);
        SetHeaderFlag(_dualCcdPrefixLabel, _dualCcdStatusLabel, "Dual-CCD", hasDualCcdCpu ? "True" : "False", hasDualCcdCpu);
        SetHeaderFlag(
            _sandboxPrefixLabel,
            _sandboxStatusLabel,
            QaRun.IsReadOnly ? "QA" : "Sandbox",
            QaRun.IsReadOnly ? "Read-only" : sandboxOn ? "On" : "Off",
            sandboxOn);

        _cpuFlagsPanel?.SuspendLayout();
        try
        {
            if (_cpuFlagsPanel is not null)
            {
                foreach (Control control in _cpuFlagsPanel.Controls.Cast<Control>().ToArray())
                {
                    if (!IsReusableHeaderFlagControl(control))
                    {
                        control.Dispose();
                    }
                }

                _cpuFlagsPanel.Controls.Clear();
            }

            AddHeaderFlag(_htPrefixLabel, _htStatusLabel);
            AddHeaderFlag(_hybridCpuPrefixLabel, _hybridCpuStatusLabel);
            AddHeaderFlag(_cppcPrefixLabel, _cppcStatusLabel);
            AddHeaderFlag(_dualCcdPrefixLabel, _dualCcdStatusLabel);
            if (sandboxOn)
            {
                AddHeaderFlag(_sandboxPrefixLabel, _sandboxStatusLabel);
            }
        }
        finally
        {
            _cpuFlagsPanel?.ResumeLayout();
        }
    }

    private bool IsSandboxDryRunActive() => _testAutoDryRun;

    private void DemandRealWrite(string operation)
    {
        QaRun.DemandWrite(operation);
        if (!IsSandboxDryRunActive()) return;
#if DEVELOPER_QA
        if (_testSandbox.Active) _testSandbox.BlockedRealWrites++;
#endif
        WriteLog($"SANDBOX.BLOCK: {operation}");
        throw new InvalidOperationException($"Sandbox dry-run blocked: {operation}.");
    }

    private bool TryBlockSandboxHardwareWrite(string operation)
    {
        if (!IsSandboxDryRunActive())
        {
            return false;
        }

        WriteLog($"SANDBOX.BLOCK: {operation}");
#if DEVELOPER_QA
        if (_testSandbox.Active) _testSandbox.BlockedRealWrites++;
#endif
        if (QaRun.TryBlock(operation)) return true;
        ShowThemedInfo(
            "Sandbox dry-run is ON.\nNo registry or driver writes are performed.\nThis developer build cannot apply settings.");
        return true;
    }

}
