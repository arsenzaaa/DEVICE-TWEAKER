
using System.Globalization;
using System.Text;
using System.Text.Json;

namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private sealed class TestCpuConfig
    {
        public int LogicalCount { get; set; }
        public bool SmtEnabled { get; set; }
        public bool UseHyperThreadingLabel { get; set; }
        public HashSet<int> ECoreLps { get; } = new();
        public Dictionary<int, int> CoreMap { get; } = new();
        public Dictionary<int, int>? CcdMap { get; set; }
        public Dictionary<int, int>? CcxMap { get; set; }
        public Dictionary<int, int> CppcRatings { get; } = new();
        public string CpuName { get; set; } = string.Empty;
    }


    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        if (_testDevicesOnly && _testAutoDryRun
            && _devicesScroll is not null
            && _devicesScroll.Visible
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

        if (e.Control && e.Alt && e.Shift && e.KeyCode == Keys.T)
        {
            e.Handled = true;
            e.SuppressKeyPress = true;
            WriteLog("UI: TEST ADMIN hotkey");
            ShowTestAdminDialog();
        }
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
            "Sandbox",
            sandboxOn ? "On" : "Off",
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

    private bool TryBlockSandboxHardwareWrite(string operation)
    {
        if (!IsSandboxDryRunActive())
        {
            return false;
        }

        WriteLog($"SANDBOX.BLOCK: {operation}");
        ShowThemedInfo(
            "Sandbox dry-run is ON.\nNo registry or driver writes are performed.\nUncheck it in TEST ADMIN for real writes.");
        return true;
    }

    private void NotifySandboxModeChanged(string reason)
    {
        WriteLog(
            $"SANDBOX.MODE: reason={reason} dryRun={_testAutoDryRun} testOnly={_testDevicesOnly} " +
            $"testDevicesEnabled={_testDevicesEnabled} testCpu={_testCpuActive}");
        UpdateCpuHeaderUi();
    }

    private void ShowTestAdminDialog()
    {
        using Form dialog = new ThemedDialogForm();
        dialog.Name = "TEST_ADMIN_DIALOG";
        dialog.Text = "TEST ADMIN";
        dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
        dialog.StartPosition = FormStartPosition.CenterParent;
        dialog.MaximizeBox = false;
        dialog.MinimizeBox = false;
        dialog.ShowInTaskbar = false;
        dialog.AutoScaleMode = AutoScaleMode.None;
        StyleThemedDialogSurface(dialog);
        WireThemedTitleBar(dialog);
        dialog.BackColor = _bgPanel;
        dialog.Font = _baseFont;
        dialog.Icon = Icon;
        dialog.ClientSize = new Size(1040, 620);
        using ThemedToolTip adminToolTip = new(showAlways: false, font: _technicalFont)
        {
            InitialDelay = 700,
            ReshowDelay = 120,
            AutoPopDelay = 8000,
            Active = false,
        };

        TableLayoutPanel layout = new()
        {
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Dock = DockStyle.Top,
            ColumnCount = 2,
            RowCount = 20,
            Padding = new Padding(24, 20, 24, 20),
            BackColor = _bgPanel,
        };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 190F));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        layout.RowStyles.Clear();
        for (int i = 0; i < layout.RowCount; i++)
        {
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        }

        Label titleLabel = new()
        {
            Text = "Test CPU Topology",
            AutoSize = true,
            Font = _titleFont,
            ForeColor = _accent,
            Margin = new Padding(0, 0, 0, 6),
        };
        layout.Controls.Add(titleLabel, 0, 0);
        layout.SetColumnSpan(titleLabel, 2);

        Label statusLabel = new()
        {
            Text = _testCpuActive ? "Test CPU mode: ACTIVE" : "Test CPU mode: OFF",
            AutoSize = true,
            ForeColor = _testCpuActive ? _statusActive : _statusInactive,
            Margin = new Padding(0, 0, 0, 12),
        };
        layout.Controls.Add(statusLabel, 0, 1);
        layout.SetColumnSpan(statusLabel, 2);

        string GetCurrentCpuNameForTest()
        {
            if (!string.IsNullOrWhiteSpace(_testCpuName))
            {
                return _testCpuName;
            }

            string text = _cpuHeaderText;
            if (text.StartsWith("CPU:", StringComparison.OrdinalIgnoreCase))
            {
                text = text[4..].Trim();
            }

            if (text.StartsWith("Test Mode", StringComparison.OrdinalIgnoreCase))
            {
                return string.Empty;
            }

            return text;
        }

        bool GetSmtEnabledFallback()
        {
            if (_cpuInfo?.Topology is not null)
            {
                return _cpuInfo.Topology.ByCore.Values.Any(g => g.Count > 1);
            }

            return true;
        }

        bool ResolveSmtEnabled(string text, bool fallback)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return fallback;
            }

            if (text.Contains("DISABLED", StringComparison.OrdinalIgnoreCase)
                || text.Contains("OFF", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            if (text.Contains("ENABLED", StringComparison.OrdinalIgnoreCase)
                || text.Contains("ON", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return fallback;
        }

        string ResolveSmtPrefix(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return string.Empty;
            }

            string[] parts = text.Split(':', 2, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length >= 1)
            {
                return parts[0].Trim();
            }

            return string.Empty;
        }

        bool ResolveUseHyperThreadingLabel(string text)
        {
            string prefix = ResolveSmtPrefix(text);
            if (prefix.Contains("HYPER", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (prefix.Contains("SMT", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            return false;
        }

        Label cpuNameLabel = NewDialogLabel("CPU name:");
        TextBox cpuNameTextBox = NewDialogTextBox(420);
        cpuNameTextBox.Text = GetCurrentCpuNameForTest();

        int currentLogical = Math.Min(MaxAffinityBits, GetCurrentLogicalCount());

        Label cpuPresetLabel = NewDialogLabel("CPU preset:");
        ThemedDropDownPicker cpuPresetCombo = NewDialogCombo(460);
        cpuPresetCombo.MaxDropDownItems = 12;
        cpuPresetCombo.DropDownWidth = 540;
        cpuPresetCombo.Items.AddRange(new object[]
        {
            "Manual (current)",
            "Intel Core i7-10700K/11700K 8C/16T",
            "Intel Core i5-13600K/14600K 6P+8E/20T",
            "Intel Core i9-13900K/14900K 8P+16E/32T",
            "Intel Core Ultra 9 285K 8P+16E/24T",
            "Intel Core Ultra 9 285HX 8P+16E/24T",
            "AMD Ryzen 5 7500F/7600X 6C/12T",
            "AMD Ryzen 7 7700X/9700X 8C/16T",
            "AMD Ryzen 9 7900X/9900X 12C/24T",
            "AMD Ryzen 9 7950X/9950X 16C/32T",
            "AMD Ryzen 5 3500X Zen2 6C/6T SMT off 2 CCX",
            "AMD Ryzen 7 3700X/3800X Zen2 8C/16T 2 CCX",
            "AMD Ryzen 9 3900X Zen2 12C/24T 4 CCX",
            "AMD Ryzen 9 3950X Zen2 16C/32T 4 CCX",
            "AMD Ryzen 7 7800X3D/9800X3D 8C/16T V-Cache",
            "AMD Ryzen 9 7900X3D/9900X3D 12C/24T V-Cache CCD0",
            "AMD Ryzen 9 7950X3D/9950X3D 16C/32T V-Cache CCD0",
            "AMD Ryzen 9 9955HX3D 16C/32T V-Cache CCD0",
            "AMD Ryzen 9 8940HX 16C/32T Dual-CCD",
            "AMD Ryzen 9 9950X3D2 16C/32T dual V-Cache",
        });
        cpuPresetLabel.Text = $"CPU preset ({cpuPresetCombo.Items.Count}):";
        cpuPresetCombo.SelectedIndex = 0;
        Button cpuPresetButton = NewDialogButton("LOAD");
        cpuPresetButton.Size = new Size(84, 26);
        cpuPresetButton.Margin = new Padding(8, 0, 0, 6);
        FlowLayoutPanel cpuPresetPanel = NewRowFlowPanel();
        cpuPresetPanel.Controls.Add(cpuPresetCombo);
        cpuPresetPanel.Controls.Add(cpuPresetButton);

        const string SystemPresetIntel14900K = "Full PC | Z790 | Core i9-14900K | RTX 4090 | Intel I226-V";
        const string SystemPresetIntel14600K = "Full PC | Z790 | Core i5-14600K | RTX 4070 SUPER | Intel I225-V";
        const string SystemPresetIntel285K5090 = "Full PC | Z890 | Core Ultra 9 285K | RTX 5090 | I226-V + BE200";
        const string SystemPresetRyzen7800X3D = "Full PC | B650E | Ryzen 7 7800X3D | RTX 4080 SUPER | I225-V";
        const string SystemPresetRyzen9800X3D5090 = "Full PC | X870E | Ryzen 7 9800X3D | RTX 5090 | Realtek 5G";
        const string SystemPresetRyzen9850X3DClient = "Client PC | B850 | Ryzen 7 9850X3D | RTX 5070 | Killer 2.5G";
        const string SystemPresetRyzen9800X3DRx = "Full PC | X870E | Ryzen 7 9800X3D | RX 9070 XT | Intel I225-V";
        const string SystemPresetRyzen9950X3DNetCx = "Full PC | X870E | Ryzen 9 9950X3D | RTX 5090 | Realtek 5G";
        const string SystemPresetRyzen9950X3DNdis = "Full PC | X870E | Ryzen 9 9950X3D | RTX 5090 | Intel I226-V";
        const string SystemPresetRyzen9950X3D2 = "Full PC | X870E | Ryzen 9 9950X3D2 | RTX 5090 | Realtek 5G";
        const string SystemPresetRyzen3500X = "Full PC | B450/X570 | Ryzen 5 3500X | GTX 1660 SUPER | Realtek 1G";
        const string SystemPresetRyzen3900X = "Full PC | X570 | Ryzen 9 3900X | RTX 3080 | Realtek 2.5G";
        const string SystemPresetRyzen9950X = "Full PC | X670E | Ryzen 9 9950X | RTX 4090 | Intel X550 10G";
        const string SystemPresetLaptopIntel285HX = "Laptop | Core Ultra 9 285HX | RTX 5080 Laptop | BE200 Wi-Fi";
        const string SystemPresetLaptopRyzen9955HX3D = "Laptop | Ryzen 9 9955HX3D | RTX 5090 Laptop | Wi-Fi 7";
        const string SystemPresetLaptopRyzen8940HX = "Laptop | Ryzen 9 8940HX | RTX 5060 Laptop | Realtek 1G";
        const string SystemPresetLaptopRyzen8940HXMouse = "Laptop | Ryzen 9 8940HX | Mouse + Keyboard | RTX 5060 Laptop";
        const string SystemPresetMultiControllerInput = "Full PC | Multi-xHCI | Mouse + Gamepad + Keyboard | RTX 5090";

        Label systemPresetLabel = NewDialogLabel("System preset:");
        ThemedDropDownPicker systemPresetCombo = NewDialogCombo(460);
        systemPresetCombo.MaxDropDownItems = 12;
        systemPresetCombo.DropDownWidth = 560;
        systemPresetCombo.Items.AddRange(new object[]
        {
            "Manual (current)",
            SystemPresetIntel14900K,
            SystemPresetIntel14600K,
            SystemPresetIntel285K5090,
            SystemPresetRyzen7800X3D,
            SystemPresetRyzen9800X3D5090,
            SystemPresetRyzen9850X3DClient,
            SystemPresetRyzen9800X3DRx,
            SystemPresetRyzen9950X3DNetCx,
            SystemPresetRyzen9950X3DNdis,
            SystemPresetRyzen9950X3D2,
            SystemPresetRyzen3500X,
            SystemPresetRyzen3900X,
            SystemPresetRyzen9950X,
            SystemPresetLaptopIntel285HX,
            SystemPresetLaptopRyzen9955HX3D,
            SystemPresetLaptopRyzen8940HX,
            SystemPresetLaptopRyzen8940HXMouse,
            SystemPresetMultiControllerInput,
        });
        systemPresetLabel.Text = $"System preset ({systemPresetCombo.Items.Count}):";
        systemPresetCombo.SelectedIndex = 0;
        Button loadSystemPresetButton = NewDialogButton("LOAD FULL");
        loadSystemPresetButton.Size = new Size(112, 26);
        loadSystemPresetButton.Margin = new Padding(8, 0, 0, 6);
        FlowLayoutPanel systemPresetPanel = NewRowFlowPanel();
        systemPresetPanel.Controls.Add(systemPresetCombo);
        systemPresetPanel.Controls.Add(loadSystemPresetButton);

        Label logicalLabel = NewDialogLabel("Total logical processors:");
        NumericUpDown logicalUpDown = NewNumericUpDown(1, MaxAffinityBits, currentLogical);

        Label smtStateLabel = NewDialogLabel("SMT status:");
        ThemedDropDownPicker smtStateCombo = NewDialogCombo(160);
        smtStateCombo.Items.AddRange(new object[] { "Enabled", "Disabled" });

        Label htStateLabel = NewDialogLabel("Hyper-Threading status:");
        ThemedDropDownPicker htStateCombo = NewDialogCombo(160);
        htStateCombo.Items.AddRange(new object[] { "Enabled", "Disabled" });

        Label cppcRatingsLabel = NewDialogLabel("CPPC ratings:");
        TextBox cppcRatingsBox = NewDialogTextBox(420);

        bool useHyperThreadingLabel = ResolveUseHyperThreadingLabel(_smtText);
        bool suppressSmtSync = false;
        bool smtAutoGenActive = false;

        int coreGroupCount = 1;
        int[] coreAssign = BuildAssignmentsFromGroupsText(GetCurrentCoreGroupsText(), currentLogical, out coreGroupCount);
        int ccdGroupCount = 1;
        int[] ccdAssign = BuildAssignmentsFromGroupsText(GetCurrentCcdGroupsText(), currentLogical, out ccdGroupCount);
        int ccxGroupCount = 1;
        int[] ccxAssign = BuildAssignmentsFromGroupsText(GetCurrentCcxGroupsText(), currentLogical, out ccxGroupCount);
        bool[] eAssign = BuildECoreFlags(GetCurrentECoreText(), currentLogical);

        string GetCurrentCppcRatingsText()
        {
            if (!_cppcEnabled || _cppcRatings.Count == 0)
            {
                return string.Empty;
            }

            return string.Join(
                ", ",
                _cppcRatings
                    .Where(kvp => kvp.Key >= 0 && kvp.Key < (int)logicalUpDown.Value)
                    .OrderBy(kvp => kvp.Key)
                    .Select(kvp => $"{kvp.Key}={kvp.Value}"));
        }

        cppcRatingsBox.Text = GetCurrentCppcRatingsText();

        Label groupCountLabel = NewDialogLabel("Group counts:");
        groupCountLabel.Anchor = AnchorStyles.Top | AnchorStyles.Left;
        groupCountLabel.Margin = new Padding(0, 14, 12, 10);
        Label assignmentsLabel = NewDialogLabel("LP assignments (manual):");

        Label coreGroupCountLabel = NewInlineLabel("Core groups:");
        Label ccdGroupCountLabel = NewInlineLabel("CCD groups:");
        Label ccxGroupCountLabel = NewInlineLabel("CCX groups:");
        NumericUpDown coreGroupCountUpDown = NewNumericUpDown(1, Math.Max(1, currentLogical), coreGroupCount);
        int maxCcdGroupsInit = Math.Min(2, Math.Max(1, currentLogical));
        int maxCcxGroupsInit = Math.Min(8, Math.Max(1, currentLogical));
        NumericUpDown ccdGroupCountUpDown = NewNumericUpDown(1, maxCcdGroupsInit, ccdGroupCount);
        NumericUpDown ccxGroupCountUpDown = NewNumericUpDown(1, maxCcxGroupsInit, ccxGroupCount);
        coreGroupCountUpDown.Size = new Size(88, 24);
        ccdGroupCountUpDown.Size = new Size(88, 24);
        ccxGroupCountUpDown.Size = new Size(88, 24);
        coreGroupCountUpDown.Margin = new Padding(0, 0, 10, 6);
        ccdGroupCountUpDown.Margin = new Padding(0, 0, 10, 6);
        ccxGroupCountUpDown.Margin = new Padding(0, 0, 0, 6);
        coreGroupCountLabel.Margin = new Padding(0, 5, 6, 0);
        ccdGroupCountLabel.Margin = new Padding(16, 5, 6, 0);
        ccxGroupCountLabel.Margin = new Padding(16, 5, 6, 0);

        FlowLayoutPanel groupCountPanel = NewRowFlowPanel();
        groupCountPanel.Margin = new Padding(0, 12, 0, 8);
        groupCountPanel.Padding = new Padding(0, 0, 0, 0);
        groupCountPanel.Controls.Add(coreGroupCountLabel);
        groupCountPanel.Controls.Add(coreGroupCountUpDown);
        groupCountPanel.Controls.Add(ccdGroupCountLabel);
        groupCountPanel.Controls.Add(ccdGroupCountUpDown);
        groupCountPanel.Controls.Add(ccxGroupCountLabel);
        groupCountPanel.Controls.Add(ccxGroupCountUpDown);

        TableLayoutPanel assignmentsTable = new()
        {
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            ColumnCount = 5,
            Dock = DockStyle.Top,
            Margin = new Padding(0),
            Padding = new Padding(4, 2, 4, 2),
        };
        assignmentsTable.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 60F));
        assignmentsTable.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150F));
        assignmentsTable.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 140F));
        assignmentsTable.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 140F));
        assignmentsTable.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 126F));

        Panel assignmentsHost = NewBoxPanel();
        assignmentsHost.AutoSize = true;
        assignmentsHost.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        assignmentsHost.Dock = DockStyle.Top;
        assignmentsHost.Margin = new Padding(0, 4, 0, 10);
        assignmentsHost.Controls.Add(assignmentsTable);

        bool suppressAssignmentEvents = false;
        Action? syncDialogScroll = null;

        void SetSmtState(bool enabled, bool useHyperLabel, bool autoGenerate)
        {
            suppressSmtSync = true;
            int index = enabled ? 0 : 1;
            smtStateCombo.SelectedIndex = index;
            htStateCombo.SelectedIndex = index;
            useHyperThreadingLabel = useHyperLabel;
            suppressSmtSync = false;

            if (autoGenerate)
            {
                AutoGenerateSmtTopology(enabled);
            }
        }

        smtStateCombo.SelectedIndexChanged += (_, _) =>
        {
            if (suppressSmtSync)
            {
                return;
            }

            SetSmtState(smtStateCombo.SelectedIndex == 0, false, true);
        };

        htStateCombo.SelectedIndexChanged += (_, _) =>
        {
            if (suppressSmtSync)
            {
                return;
            }

            SetSmtState(htStateCombo.SelectedIndex == 0, true, true);
        };

        SetSmtState(ResolveSmtEnabled(_smtText, GetSmtEnabledFallback()), useHyperThreadingLabel, false);

        layout.Controls.Add(systemPresetLabel, 0, 2);
        layout.Controls.Add(systemPresetPanel, 1, 2);
        layout.Controls.Add(cpuPresetLabel, 0, 3);
        layout.Controls.Add(cpuPresetPanel, 1, 3);
        layout.Controls.Add(cpuNameLabel, 0, 4);
        layout.Controls.Add(cpuNameTextBox, 1, 4);
        layout.Controls.Add(logicalLabel, 0, 5);
        layout.Controls.Add(logicalUpDown, 1, 5);
        layout.Controls.Add(smtStateLabel, 0, 6);
        layout.Controls.Add(smtStateCombo, 1, 6);
        layout.Controls.Add(htStateLabel, 0, 7);
        layout.Controls.Add(htStateCombo, 1, 7);
        layout.Controls.Add(cppcRatingsLabel, 0, 8);
        layout.Controls.Add(cppcRatingsBox, 1, 8);
        adminToolTip.SetToolTip(cppcRatingsBox, "Optional. Format: 0=120, 1=110, 2=100 or just 120,110,100. Empty = CPPC off in test mode.");
        layout.Controls.Add(groupCountLabel, 0, 9);
        layout.Controls.Add(groupCountPanel, 1, 9);
        layout.Controls.Add(assignmentsLabel, 0, 10);
        layout.Controls.Add(assignmentsHost, 1, 10);

        Label helpLabel = NewHintLabel("How to use: set group counts, then assign each LP to Core/CCD/CCX groups. Tick E-core where needed.");
        helpLabel.Margin = new Padding(0, 10, 0, 4);
        layout.Controls.Add(helpLabel, 0, 11);
        layout.SetColumnSpan(helpLabel, 2);

        Label noteLabel = new()
        {
            Text = $"Note: UI and affinity masks are capped at {MaxAffinityBits} LPs.",
            AutoSize = true,
            ForeColor = _mutedText,
            Margin = new Padding(0, 0, 0, 12),
        };
        layout.Controls.Add(noteLabel, 0, 12);
        layout.SetColumnSpan(noteLabel, 2);

        Label testSectionLabel = new()
        {
            Text = "Test Devices",
            AutoSize = true,
            Font = _titleFont,
            ForeColor = _accent,
            Margin = new Padding(0, 18, 0, 8),
        };
        layout.Controls.Add(testSectionLabel, 0, 13);
        layout.SetColumnSpan(testSectionLabel, 2);

        Panel testDevicesPanel = NewBoxPanel();
        testDevicesPanel.AutoSize = true;
        testDevicesPanel.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        testDevicesPanel.Dock = DockStyle.Top;
        testDevicesPanel.Margin = new Padding(0, 0, 0, 6);

        TableLayoutPanel testDevicesLayout = new()
        {
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            ColumnCount = 2,
            Dock = DockStyle.Top,
            Margin = Padding.Empty,
            Padding = new Padding(4, 2, 4, 2),
        };
        testDevicesLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 140F));
        testDevicesLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        testDevicesLayout.RowCount = 22;
        for (int i = 0; i < testDevicesLayout.RowCount; i++)
        {
            testDevicesLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        }

        // Opening TEST ADMIN arms dry-run so APPLY/AUTO/RESET stay preview-safe
        // until the user explicitly turns it off (blocked while "test only" is on).
        _testAutoDryRun = true;

        CheckBox enableTestDevicesCheck = new()
        {
            Text = "Enable test devices",
            AutoSize = true,
            AutoCheck = false,
            BackColor = _bgPanel,
            ForeColor = _fgMain,
            Checked = _testDevicesEnabled || _testDevices.Count > 0,
            Margin = new Padding(0, 0, 16, 0),
        };
        if (enableTestDevicesCheck.Checked)
        {
            _testDevicesEnabled = true;
        }

        CheckBox testDevicesOnlyCheck = new()
        {
            Text = "Show test devices only",
            AutoSize = true,
            BackColor = _bgPanel,
            ForeColor = _fgMain,
            Checked = _testDevicesOnly,
            Margin = new Padding(0, 0, 16, 0),
        };
        testDevicesOnlyCheck.Enabled = _testDevicesEnabled;

        CheckBox dryRunAutoCheck = new()
        {
            Text = "Sandbox dry-run (no registry writes)",
            AutoSize = true,
            AutoCheck = true,
            BackColor = _bgPanel,
            ForeColor = _fgMain,
            Checked = _testAutoDryRun,
            Margin = new Padding(0, 0, 0, 0),
        };

        void SyncSandboxDryRunLock(string reason)
        {
            bool simulationActive = _testDevicesEnabled || _testDevices.Count > 0;
            if (simulationActive)
            {
                _testAutoDryRun = true;
                dryRunAutoCheck.Checked = true;
                dryRunAutoCheck.AutoCheck = false;
                dryRunAutoCheck.Enabled = true;
                adminToolTip.SetToolTip(dryRunAutoCheck, UiLanguage.Text("Required while simulated devices are active."));
            }
            else
            {
                dryRunAutoCheck.AutoCheck = true;
                dryRunAutoCheck.Enabled = true;
                adminToolTip.SetToolTip(dryRunAutoCheck, string.Empty);
            }

            NotifySandboxModeChanged(reason);
        }

        FlowLayoutPanel testOptionsPanel = NewRowFlowPanel();
        testOptionsPanel.WrapContents = true;
        testOptionsPanel.Margin = new Padding(0, 0, 0, 6);
        testOptionsPanel.Controls.Add(enableTestDevicesCheck);
        testOptionsPanel.Controls.Add(testDevicesOnlyCheck);
        testOptionsPanel.Controls.Add(dryRunAutoCheck);

        Label testOptionsLabel = NewHeaderLabel("Test options");
        testDevicesLayout.Controls.Add(testOptionsLabel, 0, 0);
        testDevicesLayout.SetColumnSpan(testOptionsLabel, 2);
        testDevicesLayout.Controls.Add(testOptionsPanel, 0, 1);
        testDevicesLayout.SetColumnSpan(testOptionsPanel, 2);

        Label addDeviceLabel = NewHeaderLabel("Add fake device");
        addDeviceLabel.Margin = new Padding(0, 6, 12, 4);
        testDevicesLayout.Controls.Add(addDeviceLabel, 0, 2);
        testDevicesLayout.SetColumnSpan(addDeviceLabel, 2);

        Label testPresetLabel = NewDialogLabel("Device preset:");
        ThemedDropDownPicker testDevicePresetCombo = NewDialogCombo(470);
        testDevicePresetCombo.MaxDropDownItems = 12;
        testDevicePresetCombo.DropDownWidth = 560;
        testDevicePresetCombo.Items.AddRange(new object[]
        {
            "Manual (custom)",
            "USB - Intel Z790 PCH xHCI / Mouse 8K",
            "USB - Intel Z790 PCH xHCI / Mouse 1K",
            "USB - Intel Z790 PCH xHCI / Keyboard 8K",
            "USB - Intel Z790 PCH xHCI / Mouse 8K + Keyboard 8K",
            "USB - Intel Z790 PCH xHCI / Input + audio",
            "USB - Intel Z790 PCH xHCI / Gamepad",
            "USB - Intel Z790 PCH xHCI / Audio + microphone",
            "USB - Intel Z790 PCH xHCI / Empty controller",
            "USB - Intel Z890 PCH xHCI / Mouse 8K (CHIP 1)",
            "USB - Intel Meteor Lake CPU TB4 / Mouse 8K (CHIP 0)",
            "USB - AMD AM5 CPU xHCI / Mouse 8K (CHIP 0)",
            "USB - AMD AM5 chipset xHCI / Keyboard 8K (CHIP 1)",
            "USB - AMD X870 chipset xHCI / Keyboard 8K (CHIP 1)",
            "USB - AMD AM5 chipset xHCI / Audio + microphone",
            "USB - ASMedia ASM2142 add-in xHCI / Edge case (CHIP 1)",
            "GPU - NVIDIA GeForce RTX 5090",
            "GPU - NVIDIA GeForce RTX 5080",
            "GPU - NVIDIA GeForce RTX 4090",
            "GPU - NVIDIA GeForce RTX 4080 SUPER",
            "GPU - NVIDIA GeForce RTX 4070 SUPER",
            "GPU - NVIDIA GeForce RTX 3080",
            "GPU - NVIDIA GeForce RTX 5060 Ti",
            "GPU - AMD Radeon RX 9070 XT",
            "GPU - AMD Radeon RX 7900 XTX",
            "GPU - Intel Arc B580",
            "GPU - Intel integrated GPU",
            "NIC - Realtek RTL8125BG 2.5GbE NetAdapterCx",
            "NIC - Realtek RTL8125BG 2.5GbE NDIS",
            "NIC - Realtek RTL8126 5GbE NetAdapterCx",
            "NIC - Intel I225-V NDIS",
            "NIC - Intel I226-V NDIS",
            "NIC - Intel X550 10GbE NDIS",
            "NIC - Intel AX200 Wi-Fi NDIS",
            "NIC - Intel AX210 Wi-Fi NDIS",
            "NIC - Intel BE200 Wi-Fi 7 NDIS",
            "NIC - MediaTek MT7922 Wi-Fi NDIS",
            "Audio - Realtek HDA",
            "Audio - Realtek ALC4080 USB Audio",
            "Audio - USB DAC",
            "Audio - HDMI/DP monitor",
            "Storage - Samsung 990 PRO NVMe",
            "Storage - Crucial T705 PCIe 5.0 NVMe",
            "Storage - SATA AHCI SSD",
        });
        testPresetLabel.Text = $"Device preset ({testDevicePresetCombo.Items.Count}):";
        testDevicePresetCombo.SelectedIndex = 0;
        Button addTestPresetButton = NewDialogButton("ADD PRESET");
        addTestPresetButton.Size = new Size(128, 26);
        addTestPresetButton.Margin = new Padding(8, 0, 0, 6);
        FlowLayoutPanel testPresetPanel = NewRowFlowPanel();
        testPresetPanel.Controls.Add(testDevicePresetCombo);
        testPresetPanel.Controls.Add(addTestPresetButton);
        testDevicesLayout.Controls.Add(testPresetLabel, 0, 3);
        testDevicesLayout.Controls.Add(testPresetPanel, 1, 3);

        Label testNameLabel = NewDialogLabel("Name:");
        TextBox testNameBox = NewDialogTextBox(500);
        testDevicesLayout.Controls.Add(testNameLabel, 0, 4);
        testDevicesLayout.Controls.Add(testNameBox, 1, 4);

        Label testPnpIdLabel = NewDialogLabel("PNP ID:");
        TextBox testPnpIdBox = NewDialogTextBox(500);
        adminToolTip.SetToolTip(testPnpIdBox, @"Optional fake hardware ID. Example: PCI\VEN_10EC&DEV_8125\TEST for NIC ITR profile preview.");
        testDevicesLayout.Controls.Add(testPnpIdLabel, 0, 5);
        testDevicesLayout.Controls.Add(testPnpIdBox, 1, 5);

        Label testKindLabel = NewDialogLabel("Kind:");
        ThemedDropDownPicker testKindCombo = NewDialogCombo(180);
        testKindCombo.Items.Add(DeviceKind.USB);
        testKindCombo.Items.Add(DeviceKind.GPU);
        testKindCombo.Items.Add(DeviceKind.AUDIO);
        testKindCombo.Items.Add(DeviceKind.NET_NDIS);
        testKindCombo.Items.Add(DeviceKind.NET_CX);
        testKindCombo.Items.Add(DeviceKind.STOR);
        testKindCombo.SelectedIndex = 0;
        testDevicesLayout.Controls.Add(testKindLabel, 0, 6);
        testDevicesLayout.Controls.Add(testKindCombo, 1, 6);

        Label testUsbRolesLabel = NewDialogLabel("USB roles:");
        TextBox testUsbRolesBox = NewDialogTextBox(500);
        testDevicesLayout.Controls.Add(testUsbRolesLabel, 0, 7);
        testDevicesLayout.Controls.Add(testUsbRolesBox, 1, 7);

        Label testAudioLabel = NewDialogLabel("Audio endpoints:");
        TextBox testAudioBox = NewDialogTextBox(500);
        testDevicesLayout.Controls.Add(testAudioLabel, 0, 8);
        testDevicesLayout.Controls.Add(testAudioBox, 1, 8);

        Label testStorageLabel = NewDialogLabel("Storage tag:");
        TextBox testStorageBox = NewDialogTextBox(180);
        testDevicesLayout.Controls.Add(testStorageLabel, 0, 9);
        testDevicesLayout.Controls.Add(testStorageBox, 1, 9);

        Label testIrqCountLabel = NewDialogLabel("IRQ count:");
        TextBox testIrqCountBox = NewDialogTextBox(120);
        testIrqCountBox.Text = "Auto";
        adminToolTip.SetToolTip(testIrqCountBox, "Optional fake IRQ count shown in device blocks. Empty/Auto = role-based default.");
        testDevicesLayout.Controls.Add(testIrqCountLabel, 0, 10);
        testDevicesLayout.Controls.Add(testIrqCountBox, 1, 10);

        Label testMsiStatusLabel = NewDialogLabel("MSI status:");
        ThemedDropDownPicker testMsiStatusCombo = NewDialogCombo(160);
        testMsiStatusCombo.Items.AddRange(new object[] { "Auto", "Enabled", "Disabled" });
        testMsiStatusCombo.SelectedIndex = 0;
        testDevicesLayout.Controls.Add(testMsiStatusLabel, 0, 11);
        testDevicesLayout.Controls.Add(testMsiStatusCombo, 1, 11);

        Label testSuspendLabel = NewDialogLabel("USB Selective Suspend:");
        ThemedDropDownPicker testSuspendCombo = NewDialogCombo(160);
        testSuspendCombo.Items.AddRange(new object[] { "off", "on", "unset" });
        testSuspendCombo.SelectedIndex = 0;
        adminToolTip.SetToolTip(testSuspendCombo, "Fake SelectiveSuspendEnabled for USB test devices (Selective Suspend picker).");
        testDevicesLayout.Controls.Add(testSuspendLabel, 0, 12);
        testDevicesLayout.Controls.Add(testSuspendCombo, 1, 12);

        Label testNicPowerLabel = NewDialogLabel("NIC power saving:");
        ThemedDropDownPicker testNicPowerCombo = NewDialogCombo(160);
        testNicPowerCombo.Items.AddRange(new object[] { "on", "off" });
        testNicPowerCombo.SelectedIndex = 0;
        adminToolTip.SetToolTip(
            testNicPowerCombo,
            "Fake Device Manager 'Allow the computer to turn off this device to save power' for wired NIC test devices.\n" +
            "Hidden for Wi‑Fi. on = checkbox checked, off = unchecked.");
        testDevicesLayout.Controls.Add(testNicPowerLabel, 0, 13);
        testDevicesLayout.Controls.Add(testNicPowerCombo, 1, 13);

        CheckBox testWifiCheck = new()
        {
            Text = "WiFi",
            AutoSize = true,
            BackColor = _bgPanel,
            ForeColor = _fgMain,
            Margin = new Padding(0, 0, 12, 0),
        };

        CheckBox testXhciCheck = new()
        {
            Text = "USB XHCI",
            AutoSize = true,
            BackColor = _bgPanel,
            ForeColor = _fgMain,
            Checked = true,
            Margin = new Padding(0, 0, 12, 0),
        };

        CheckBox testHasDevicesCheck = new()
        {
            Text = "USB has devices",
            AutoSize = true,
            BackColor = _bgPanel,
            ForeColor = _fgMain,
            Checked = true,
            Margin = new Padding(0, 0, 0, 0),
        };

        CheckBox testIntegratedGpuCheck = new()
        {
            Text = "Integrated GPU (iGPU)",
            AutoSize = true,
            BackColor = _bgPanel,
            ForeColor = _fgMain,
            Margin = new Padding(0, 0, 12, 0),
        };

        FlowLayoutPanel testDeviceOptionsPanel = NewRowFlowPanel();
        testDeviceOptionsPanel.WrapContents = true;
        testDeviceOptionsPanel.Margin = new Padding(0, 0, 0, 6);
        testDeviceOptionsPanel.Controls.Add(testWifiCheck);
        testDeviceOptionsPanel.Controls.Add(testXhciCheck);
        testDeviceOptionsPanel.Controls.Add(testHasDevicesCheck);
        testDeviceOptionsPanel.Controls.Add(testIntegratedGpuCheck);

        Label testDeviceOptionsLabel = NewDialogLabel("Options:");
        testDevicesLayout.Controls.Add(testDeviceOptionsLabel, 0, 14);
        testDevicesLayout.Controls.Add(testDeviceOptionsPanel, 1, 14);

        Button addTestDeviceButton = NewDialogButton("ADD FAKE DEVICE");
        addTestDeviceButton.Size = new Size(180, 30);
        addTestDeviceButton.Margin = new Padding(0, 4, 12, 10);

        Button updateTestDeviceButton = NewDialogButton("UPDATE SELECTED");
        updateTestDeviceButton.Size = new Size(180, 30);
        updateTestDeviceButton.Margin = new Padding(0, 4, 0, 10);
        updateTestDeviceButton.Enabled = false;

        FlowLayoutPanel addDeviceButtonsPanel = NewRowFlowPanel();
        addDeviceButtonsPanel.WrapContents = true;
        addDeviceButtonsPanel.Controls.Add(addTestDeviceButton);
        addDeviceButtonsPanel.Controls.Add(updateTestDeviceButton);
        testDevicesLayout.Controls.Add(addDeviceButtonsPanel, 0, 15);
        testDevicesLayout.SetColumnSpan(addDeviceButtonsPanel, 2);

        Label testListLabel = NewHeaderLabel($"Current test devices: {_testDevices.Count}");
        testDevicesLayout.Controls.Add(testListLabel, 0, 16);
        testDevicesLayout.SetColumnSpan(testListLabel, 2);

        Panel CreateThemedListHost(ListBox list, int width, int height)
        {
            Panel host = new()
            {
                Size = new Size(width, height),
                BackColor = Color.FromArgb(18, 18, 22),
                Margin = new Padding(0, 0, 0, 6),
                Padding = new Padding(1),
            };
            host.Paint += (_, e) =>
            {
                Rectangle border = host.ClientRectangle;
                border.Width -= 1;
                border.Height -= 1;
                using Pen pen = new(Color.FromArgb(70, 70, 76));
                e.Graphics.DrawRectangle(pen, border);
            };

            ThemedScrollBar scroll = new()
            {
                Dock = DockStyle.None,
                Width = UiScale(12),
                BackColor = Color.FromArgb(18, 18, 22),
                TrackColor = Color.FromArgb(18, 18, 22),
                RailColor = Color.FromArgb(18, 18, 22),
                ThumbColor = _accent,
                ThumbWidth = UiScale(8),
                RailWidth = 0,
                ThumbCornerRadius = UiScale(6),
                SmallChange = 1,
                Visible = false,
            };

            list.BorderStyle = BorderStyle.None;
            list.HorizontalScrollbar = false;
            list.Dock = DockStyle.None;
            list.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            list.Margin = Padding.Empty;

            void LayoutListHost()
            {
                if (host.IsDisposed || list.IsDisposed || scroll.IsDisposed)
                {
                    return;
                }

                int border = UiScale(1);
                int lane = scroll.Visible ? scroll.Width + UiScale(2) : 0;
                list.Bounds = new Rectangle(
                    border,
                    border,
                    Math.Max(1, host.ClientSize.Width - (border * 2) - lane),
                    Math.Max(1, host.ClientSize.Height - (border * 2)));
                scroll.Bounds = new Rectangle(
                    Math.Max(border, host.ClientSize.Width - border - scroll.Width),
                    border,
                    scroll.Width,
                    Math.Max(1, host.ClientSize.Height - (border * 2)));
            }

            bool syncing = false;
            void SyncListScroll()
            {
                if (list.IsDisposed || host.IsDisposed)
                {
                    return;
                }

                if (list.IsHandleCreated)
                {
                    HideNativeScrollBars(list);
                }

                int itemHeight = Math.Max(1, list.ItemHeight);
                int visibleRows = Math.Max(1, list.ClientSize.Height / itemHeight);
                bool needsScroll = list.Items.Count > visibleRows;
                scroll.Visible = needsScroll;
                LayoutListHost();
                visibleRows = Math.Max(1, list.ClientSize.Height / itemHeight);
                scroll.Maximum = Math.Max(1, list.Items.Count);
                scroll.ViewportSize = visibleRows;
                scroll.LargeChange = visibleRows;
                syncing = true;
                scroll.Value = needsScroll && list.Items.Count > 0 ? Math.Max(0, list.TopIndex) : 0;
                syncing = false;
                scroll.BringToFront();
            }

            scroll.ValueChanged += (_, _) =>
            {
                if (syncing || list.Items.Count == 0)
                {
                    return;
                }

                list.TopIndex = Math.Min(list.Items.Count - 1, scroll.Value);
                list.Invalidate();
            };
            list.MouseWheel += (_, _) => BeginInvoke(new Action(SyncListScroll));
            list.SelectedIndexChanged += (_, _) => BeginInvoke(new Action(SyncListScroll));
            list.KeyUp += (_, _) => BeginInvoke(new Action(SyncListScroll));
            list.Paint += (_, _) => SyncListScroll();
            list.HandleCreated += (_, _) => BeginInvoke(new Action(SyncListScroll));
            host.SizeChanged += (_, _) =>
            {
                LayoutListHost();
                BeginInvoke(new Action(SyncListScroll));
            };
            list.MouseMove += (_, e) =>
            {
                int index = list.IndexFromPoint(e.Location);
                string tip = index >= 0 && index < list.Items.Count
                    ? list.Items[index]?.ToString() ?? string.Empty
                    : string.Empty;
                adminToolTip.SetToolTip(list, tip);
            };

            host.Tag = (Action)SyncListScroll;
            host.Controls.Add(list);
            host.Controls.Add(scroll);
            LayoutListHost();
            scroll.BringToFront();
            return host;
        }

        static void SyncThemedListHost(ListBox list)
        {
            if (list.Parent?.Tag is Action sync)
            {
                sync();
            }
        }

        ListBox testDeviceListBox = new()
        {
            Height = 122,
            Width = 700,
            BackColor = Color.FromArgb(18, 18, 22),
            ForeColor = _fgMain,
            IntegralHeight = false,
            SelectionMode = SelectionMode.One,
            HorizontalScrollbar = false,
        };
        Panel testDeviceListHost = CreateThemedListHost(testDeviceListBox, 840, 122);
        testDeviceListHost.Dock = DockStyle.Top;
        testDevicesLayout.Controls.Add(testDeviceListHost, 0, 17);
        testDevicesLayout.SetColumnSpan(testDeviceListHost, 2);

        Button removeTestDeviceButton = NewDialogButton("REMOVE SELECTED");
        removeTestDeviceButton.Size = new Size(180, 28);
        removeTestDeviceButton.Margin = new Padding(0, 0, 12, 0);

        Button clearTestDeviceButton = NewDialogButton("CLEAR ALL");
        clearTestDeviceButton.Size = new Size(140, 28);
        clearTestDeviceButton.Margin = new Padding(0, 0, 0, 0);

        FlowLayoutPanel testDeviceButtonsPanel = NewRowFlowPanel();
        testDeviceButtonsPanel.WrapContents = true;
        testDeviceButtonsPanel.Margin = new Padding(0, 0, 0, 6);
        testDeviceButtonsPanel.Controls.Add(removeTestDeviceButton);
        testDeviceButtonsPanel.Controls.Add(clearTestDeviceButton);
        testDevicesLayout.Controls.Add(testDeviceButtonsPanel, 0, 18);
        testDevicesLayout.SetColumnSpan(testDeviceButtonsPanel, 2);

        Label realVisibilityLabel = NewHeaderLabel("Real device visibility");
        realVisibilityLabel.Margin = new Padding(0, 12, 12, 4);
        testDevicesLayout.Controls.Add(realVisibilityLabel, 0, 19);
        testDevicesLayout.SetColumnSpan(realVisibilityLabel, 2);

        ListBox realDeviceListBox = new()
        {
            Height = 100,
            Width = 430,
            BackColor = Color.FromArgb(18, 18, 22),
            ForeColor = _fgMain,
            IntegralHeight = false,
            SelectionMode = SelectionMode.One,
            HorizontalScrollbar = false,
        };

        ListBox hiddenDeviceListBox = new()
        {
            Height = 100,
            Width = 330,
            BackColor = Color.FromArgb(18, 18, 22),
            ForeColor = _fgMain,
            IntegralHeight = false,
            SelectionMode = SelectionMode.One,
            HorizontalScrollbar = false,
        };
        Panel realDeviceListHost = CreateThemedListHost(realDeviceListBox, 430, 100);
        Panel hiddenDeviceListHost = CreateThemedListHost(hiddenDeviceListBox, 430, 100);
        realDeviceListHost.Dock = DockStyle.Fill;
        hiddenDeviceListHost.Dock = DockStyle.Fill;

        Button hideRealDeviceButton = NewDialogButton("HIDE SELECTED");
        hideRealDeviceButton.Size = new Size(150, 28);
        hideRealDeviceButton.Margin = new Padding(0, 0, 8, 0);

        Button unhideRealDeviceButton = NewDialogButton("UNHIDE SELECTED");
        unhideRealDeviceButton.Size = new Size(170, 28);
        unhideRealDeviceButton.Margin = new Padding(0, 0, 8, 0);

        Button clearHiddenDeviceButton = NewDialogButton("CLEAR HIDDEN");
        clearHiddenDeviceButton.Size = new Size(150, 28);
        clearHiddenDeviceButton.Margin = new Padding(0, 0, 0, 0);

        TableLayoutPanel realVisibilityPanel = new()
        {
            AutoSize = false,
            Height = 170,
            ColumnCount = 2,
            RowCount = 3,
            Dock = DockStyle.Top,
            Margin = new Padding(0, 0, 0, 8),
            Padding = Padding.Empty,
        };
        realVisibilityPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
        realVisibilityPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
        realVisibilityPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 24F));
        realVisibilityPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 108F));
        realVisibilityPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 38F));
        realVisibilityPanel.Controls.Add(NewHeaderLabel("Visible real devices"), 0, 0);
        realVisibilityPanel.Controls.Add(NewHeaderLabel("Hidden real devices"), 1, 0);
        realVisibilityPanel.Controls.Add(realDeviceListHost, 0, 1);
        realVisibilityPanel.Controls.Add(hiddenDeviceListHost, 1, 1);

        FlowLayoutPanel hideButtonsPanel = NewRowFlowPanel();
        hideButtonsPanel.Controls.Add(hideRealDeviceButton);
        FlowLayoutPanel unhideButtonsPanel = NewRowFlowPanel();
        unhideButtonsPanel.Controls.Add(unhideRealDeviceButton);
        unhideButtonsPanel.Controls.Add(clearHiddenDeviceButton);
        realVisibilityPanel.Controls.Add(hideButtonsPanel, 0, 2);
        realVisibilityPanel.Controls.Add(unhideButtonsPanel, 1, 2);

        testDevicesLayout.Controls.Add(realVisibilityPanel, 0, 20);
        testDevicesLayout.SetColumnSpan(realVisibilityPanel, 2);

        Label testHintLabel = NewHintLabel(@"Tip: full PC presets are above. This section is for adding/removing individual fake devices and temporarily hiding real devices.");
        testDevicesLayout.Controls.Add(testHintLabel, 0, 21);
        testDevicesLayout.SetColumnSpan(testHintLabel, 2);

        testDevicesPanel.Controls.Add(testDevicesLayout);
        layout.Controls.Add(testDevicesPanel, 0, 14);
        layout.SetColumnSpan(testDevicesPanel, 2);

        Label scenarioLabLabel = new()
        {
            Text = "Scenario Lab",
            AutoSize = true,
            Font = _titleFont,
            ForeColor = _accent,
            Margin = new Padding(0, 18, 0, 8),
        };
        layout.Controls.Add(scenarioLabLabel, 0, 15);
        layout.SetColumnSpan(scenarioLabLabel, 2);

        Panel scenarioLabPanel = NewBoxPanel();
        scenarioLabPanel.AutoSize = true;
        scenarioLabPanel.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        scenarioLabPanel.Dock = DockStyle.Top;
        scenarioLabPanel.Margin = new Padding(0, 0, 0, 8);

        TableLayoutPanel scenarioLayout = new()
        {
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            ColumnCount = 2,
            Dock = DockStyle.Top,
            Padding = new Padding(4),
            Margin = Padding.Empty,
        };
        scenarioLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150F));
        scenarioLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

        TextBox scenarioNameBox = NewDialogTextBox(500);
        scenarioNameBox.Text = "Release regression";
        ThemedDropDownPicker scenarioOperationCombo = NewDialogCombo(180);
        scenarioOperationCombo.Items.AddRange(Enum.GetValues<TestSandboxOperation>().Cast<object>().ToArray());
        scenarioOperationCombo.SelectedItem = TestSandboxOperation.Auto;
        ThemedDropDownPicker scenarioFaultCombo = NewDialogCombo(310);
        scenarioFaultCombo.Items.AddRange(Enum.GetValues<TestSandboxFault>().Cast<object>().ToArray());
        scenarioFaultCombo.SelectedItem = TestSandboxFault.None;
        CheckBox scenarioImodCheck = new()
        {
            Text = "Include USB IMOD",
            AutoSize = true,
            Checked = true,
            BackColor = _bgPanel,
            ForeColor = _fgMain,
            Margin = new Padding(16, 4, 0, 6),
        };

        ThemedDropDownPicker scenarioDeviceCombo = NewDialogCombo(560);
        scenarioDeviceCombo.DropDownWidth = 690;
        ThemedDropDownPicker stateMsiCombo = NewDialogCombo(130);
        stateMsiCombo.Items.AddRange(["Enabled", "Disabled"]);
        TextBox stateLimitBox = NewDialogTextBox(100);
        ThemedDropDownPicker statePriorityCombo = NewDialogCombo(140);
        statePriorityCombo.Items.AddRange(["Undefined", "Low", "Normal", "High"]);
        ThemedDropDownPicker statePolicyCombo = NewDialogCombo(180);
        statePolicyCombo.Items.AddRange(["MachineDefault", "AllClose", "Single", "All", "SpecCPU", "SpreadMessages"]);
        TextBox stateMaskBox = NewDialogTextBox(140);
        TextBox stateRssBaseBox = NewDialogTextBox(90);
        NumericUpDown stateRssQueuesBox = NewNumericUpDown(1, MaxAffinityBits, 1);
        ThemedDropDownPicker stateNdisModeCombo = NewDialogCombo(110);
        stateNdisModeCombo.Items.AddRange(["RSS", "IRQ", "BOTH"]);
        ThemedDropDownPicker statePowerCombo = NewDialogCombo(140);
        statePowerCombo.Items.AddRange(["preserve", "on", "off"]);
        TextBox stateImodBox = NewDialogTextBox(260);
        TextBox stateNicItrBox = NewDialogTextBox(180);

        FlowLayoutPanel operationRow = NewRowFlowPanel();
        operationRow.Controls.Add(scenarioOperationCombo);
        operationRow.Controls.Add(scenarioImodCheck);
        FlowLayoutPanel stateMsiRow = NewRowFlowPanel();
        stateMsiRow.Controls.Add(stateMsiCombo);
        stateMsiRow.Controls.Add(NewInlineLabel("Limit (absent/0..2048):"));
        stateMsiRow.Controls.Add(stateLimitBox);
        FlowLayoutPanel stateAffinityRow = NewRowFlowPanel();
        stateAffinityRow.Controls.Add(statePolicyCombo);
        stateAffinityRow.Controls.Add(NewInlineLabel("Mask:"));
        stateAffinityRow.Controls.Add(stateMaskBox);
        FlowLayoutPanel stateRssRow = NewRowFlowPanel();
        stateRssRow.Controls.Add(NewInlineLabel("Base (absent/int):"));
        stateRssRow.Controls.Add(stateRssBaseBox);
        stateRssRow.Controls.Add(NewInlineLabel("Queues:"));
        stateRssRow.Controls.Add(stateRssQueuesBox);
        stateRssRow.Controls.Add(NewInlineLabel("Mode:"));
        stateRssRow.Controls.Add(stateNdisModeCombo);
        FlowLayoutPanel stateDriverRow = NewRowFlowPanel();
        stateDriverRow.Controls.Add(NewInlineLabel("IMOD:"));
        stateDriverRow.Controls.Add(stateImodBox);
        stateDriverRow.Controls.Add(NewInlineLabel("NIC ITR:"));
        stateDriverRow.Controls.Add(stateNicItrBox);

        int scenarioRow = 0;
        void AddScenarioRow(string label, Control control)
        {
            scenarioLayout.Controls.Add(NewDialogLabel(label), 0, scenarioRow);
            scenarioLayout.Controls.Add(control, 1, scenarioRow);
            scenarioRow++;
        }
        AddScenarioRow("Scenario name:", scenarioNameBox);
        AddScenarioRow("Operation:", operationRow);
        AddScenarioRow("Failure injection:", scenarioFaultCombo);
        AddScenarioRow("Initial state device:", scenarioDeviceCombo);
        AddScenarioRow("MSI state:", stateMsiRow);
        AddScenarioRow("IRQ Priority:", statePriorityCombo);
        AddScenarioRow("CPU Affinity state:", stateAffinityRow);
        AddScenarioRow("RSS state:", stateRssRow);
        AddScenarioRow("Power Saving:", statePowerCombo);
        AddScenarioRow("Driver state:", stateDriverRow);

        void LoadSelectedInitialState()
        {
            int index = scenarioDeviceCombo.SelectedIndex;
            if (index < 0 || index >= _testDevices.Count)
            {
                return;
            }
            TestDeviceState state = EnsureTestDeviceState(_testDevices[index]);
            stateMsiCombo.SelectedItem = state.MsiEnabled ? "Enabled" : "Disabled";
            stateLimitBox.Text = state.MsiLimit?.ToString(CultureInfo.InvariantCulture) ?? "absent";
            statePriorityCombo.SelectedItem = state.Priority switch { 1 => "Low", 2 => "Normal", 3 => "High", _ => "Undefined" };
            statePolicyCombo.SelectedItem = FormatPolicyValue(state.Policy);
            stateMaskBox.Text = $"0x{state.AffinityMask:X}";
            stateRssBaseBox.Text = state.RssBaseCore?.ToString(CultureInfo.InvariantCulture) ?? "absent";
            stateRssQueuesBox.Value = Math.Max(1, Math.Min(MaxAffinityBits, state.RssQueues));
            stateNdisModeCombo.SelectedItem = state.NdisMode;
            statePowerCombo.SelectedItem = state.PowerSavingEnabled switch { true => "on", false => "off", _ => "preserve" };
            stateImodBox.Text = state.ImodValue;
            stateNicItrBox.Text = state.NicItrValue;
        }

        bool TrySaveSelectedInitialState(out string error)
        {
            error = string.Empty;
            int index = scenarioDeviceCombo.SelectedIndex;
            if (index < 0 || index >= _testDevices.Count)
            {
                error = "Select a test device first.";
                return false;
            }
            string limitText = stateLimitBox.Text.Trim();
            int? limit = null;
            if (!limitText.Equals("absent", StringComparison.OrdinalIgnoreCase)
                && !string.IsNullOrWhiteSpace(limitText)
                && limitText != "0")
            {
                if (!int.TryParse(limitText, out int parsedLimit) || parsedLimit is < 1 or > 2048)
                {
                    error = "MSI Limit must be absent/0 or 1..2048.";
                    return false;
                }
                limit = parsedLimit;
            }
            string maskText = stateMaskBox.Text.Trim();
            if (maskText.StartsWith("0x", StringComparison.OrdinalIgnoreCase)) maskText = maskText[2..];
            if (!ulong.TryParse(maskText, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out ulong mask))
            {
                error = "Affinity mask must be hexadecimal, for example 0x3.";
                return false;
            }
            int? rssBase = null;
            string rssText = stateRssBaseBox.Text.Trim();
            if (!rssText.Equals("absent", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(rssText))
            {
                if (!int.TryParse(rssText, out int parsedBase) || parsedBase < 0 || parsedBase >= MaxAffinityBits)
                {
                    error = $"RSS base must be absent or 0..{MaxAffinityBits - 1}.";
                    return false;
                }
                rssBase = parsedBase;
            }

            DeviceInfo device = _testDevices[index];
            TestDeviceState state = EnsureTestDeviceState(device);
            state.MsiEnabled = string.Equals(stateMsiCombo.SelectedItem?.ToString(), "Enabled", StringComparison.Ordinal);
            state.MsiLimit = limit;
            state.Priority = statePriorityCombo.SelectedItem?.ToString() switch { "Low" => 1, "Normal" => 2, "High" => 3, _ => null };
            state.Policy = MapPolicyText(statePolicyCombo.SelectedItem?.ToString() ?? "MachineDefault") ?? 0;
            state.AffinityMask = mask;
            state.RssBaseCore = rssBase;
            state.RssQueues = (int)stateRssQueuesBox.Value;
            state.NdisMode = stateNdisModeCombo.SelectedItem?.ToString() ?? "RSS";
            state.PowerSavingEnabled = statePowerCombo.SelectedItem?.ToString() switch { "on" => true, "off" => false, _ => null };
            state.ImodValue = stateImodBox.Text.Trim();
            state.NicItrValue = stateNicItrBox.Text.Trim();
            if (device.Kind == DeviceKind.USB && state.PowerSavingEnabled.HasValue) device.UsbSelectiveSuspend = state.PowerSavingEnabled.Value ? "on" : "off";
            if (device.Kind is DeviceKind.NET_NDIS or DeviceKind.NET_CX && !device.Wifi && state.PowerSavingEnabled.HasValue) device.NicPowerSaving = state.PowerSavingEnabled.Value ? "on" : "off";
            DeviceBlock? block = _blocks.FirstOrDefault(candidate => string.Equals(candidate.Device.InstanceId, device.InstanceId, StringComparison.OrdinalIgnoreCase));
            if (block is not null) LoadBlockSettings(block);
            WriteLog($"TEST.SCENARIO.STATE: id={device.InstanceId} initial-state-updated");
            return true;
        }

        Button stateApplyButton = NewDialogButton("APPLY INITIAL STATE");
        stateApplyButton.Dock = DockStyle.Fill;
        stateApplyButton.Click += (_, _) =>
        {
            if (!TrySaveSelectedInitialState(out string error)) ShowThemedInfo(error);
        };
        Button scenarioRunButton = NewDialogButton("RUN SCENARIO");
        scenarioRunButton.Dock = DockStyle.Fill;
        scenarioRunButton.Click += (_, _) =>
        {
            if (!TrySaveSelectedInitialState(out string error))
            {
                ShowThemedInfo(error);
                return;
            }
            TestSandboxScenario scenario = CaptureCurrentTestScenario(
                scenarioNameBox.Text,
                (TestSandboxOperation)(scenarioOperationCombo.SelectedItem ?? TestSandboxOperation.Auto),
                (TestSandboxFault)(scenarioFaultCombo.SelectedItem ?? TestSandboxFault.None),
                scenarioImodCheck.Checked);
            _ = RunTestSandboxScenario(scenario, showResult: true);
        };
        Button scenarioMatrixButton = NewDialogButton("RUN CORE MATRIX");
        scenarioMatrixButton.Dock = DockStyle.Fill;
        bool RunCoreScenarioMatrix(bool showDialog)
        {
            if (showDialog && !TrySaveSelectedInitialState(out string error))
            {
                ShowThemedInfo(error);
                return false;
            }
            (TestSandboxOperation Operation, TestSandboxFault Fault, bool Imod)[] matrix =
            [
                (TestSandboxOperation.Auto, TestSandboxFault.None, true),
                (TestSandboxOperation.Auto, TestSandboxFault.CpuTopologyUnavailable, false),
                (TestSandboxOperation.Auto, TestSandboxFault.CpuMapUnavailable, false),
                (TestSandboxOperation.Apply, TestSandboxFault.None, true),
                (TestSandboxOperation.SafeReset, TestSandboxFault.None, false),
                (TestSandboxOperation.Backup, TestSandboxFault.None, false),
                (TestSandboxOperation.Backup, TestSandboxFault.BackupAccessDenied, false),
                (TestSandboxOperation.Auto, TestSandboxFault.BackupAccessDenied, true),
                (TestSandboxOperation.Auto, TestSandboxFault.RegistryWriteDenied, true),
                (TestSandboxOperation.Auto, TestSandboxFault.UsbPowerWriteFailed, true),
                (TestSandboxOperation.Auto, TestSandboxFault.WmiTimeout, true),
                (TestSandboxOperation.Auto, TestSandboxFault.KduMissing, true),
                (TestSandboxOperation.Auto, TestSandboxFault.KduExitCode, true),
                (TestSandboxOperation.Auto, TestSandboxFault.KduDeviceUnavailable577, true),
                (TestSandboxOperation.Auto, TestSandboxFault.KduTimeout, true),
                (TestSandboxOperation.Auto, TestSandboxFault.ImodReadFailed, true),
                (TestSandboxOperation.Auto, TestSandboxFault.ImodWriteFailed, true),
                (TestSandboxOperation.Auto, TestSandboxFault.ImodVerificationMismatch, true),
                (TestSandboxOperation.Auto, TestSandboxFault.NicItrReadFailed, true),
                (TestSandboxOperation.Auto, TestSandboxFault.NicItrWriteFailed, true),
                (TestSandboxOperation.Auto, TestSandboxFault.RefreshFailed, true),
                (TestSandboxOperation.Restore, TestSandboxFault.None, false),
                (TestSandboxOperation.Restore, TestSandboxFault.CorruptBackup, false),
                (TestSandboxOperation.Restore, TestSandboxFault.RestoreWriteFailed, false),
                (TestSandboxOperation.Restore, TestSandboxFault.RollbackFailed, false),
            ];
            int passed = 0;
            List<string> failed = [];
            Dictionary<string, TestDeviceState> baselineState = CaptureTestSandboxState();
            try
            {
                static bool PipeColumnsMatch(string text)
                {
                    int[] columns = text.Split(["\r\n", "\n"], StringSplitOptions.RemoveEmptyEntries)
                        .Select(line => line.IndexOf(" | ", StringComparison.Ordinal))
                        .Where(column => column >= 0)
                        .ToArray();
                    return columns.Length >= 2 && columns.All(column => column == columns[0]);
                }

                static bool TokenColumnsMatch(string text, string token)
                {
                    int[] columns = text.Split(["\r\n", "\n"], StringSplitOptions.RemoveEmptyEntries)
                        .Select(line => line.IndexOf(token, StringComparison.Ordinal))
                        .Where(column => column >= 0)
                        .ToArray();
                    return columns.Length >= 2 && columns.All(column => column == columns[0]);
                }

                string deviceMap = FormatDeviceFirstImodLines(
                [
                    "Mouse 1K -> intr0=0x0/0 ns",
                    "Keyboard 8K -> intr0=0x0/0 ns",
                    "Audio -> intr0=0xC8/50us",
                    "Microphone -> intr0=0x0/0 ns",
                ]);
                string interrupterMap = FormatVisibleImodInterrupterLines([0, 0xC8, 0xC8, 0xC8, 0, 0xC8, 0xC8, 0]);
                NicItrProfile nicProfile = NicItrProfiles.First(profile => profile.TimingKind == NicItrTimingKind.RealtekIntrMitV2);
                string nicDetails = FormatNicItrQueueDetailText([0x7E600958, 0, 0x7E600938, 0], nicProfile);
                string[] nicRows = nicDetails.Split(["\r\n", "\n"], StringSplitOptions.RemoveEmptyEntries);
                int[] rawColumns = nicRows.Select(row => row.IndexOf("0x", StringComparison.Ordinal)).ToArray();
                int[] rxColumns = nicRows.Where(row => row.Contains("RX ", StringComparison.Ordinal)).Select(row => row.IndexOf("RX ", StringComparison.Ordinal)).ToArray();
                int[] txColumns = nicRows.Where(row => row.Contains("TX ", StringComparison.Ordinal)).Select(row => row.IndexOf("TX ", StringComparison.Ordinal)).ToArray();
                string nicSummary = FormatNicItrQueueSummary([0x7E600958, 0, 0x7E600938, 0], nicProfile)["current: ".Length..];
                string[] summaryCells = nicSummary.Split(" | ", StringSplitOptions.None);
                using ImodMapTextBox renderedImod = new()
                {
                    Text = $"{deviceMap}\r\n\r\n{interrupterMap}",
                    Font = _technicalFont,
                    Size = new Size(UiScale(640), UiScale(220)),
                };
                using NicItrTableLabel renderedNic = new()
                {
                    Text = nicDetails,
                    Font = _technicalFont,
                    Size = new Size(UiScale(640), UiScale(110)),
                };
                using Bitmap imodBitmap = new(renderedImod.Width, renderedImod.Height);
                using Bitmap nicBitmap = new(renderedNic.Width, renderedNic.Height);
                renderedImod.DrawToBitmap(imodBitmap, renderedImod.ClientRectangle);
                renderedNic.DrawToBitmap(nicBitmap, renderedNic.ClientRectangle);
                bool layoutPassed = TokenColumnsMatch(deviceMap, " -> ")
                    && PipeColumnsMatch(interrupterMap)
                    && rawColumns.Distinct().Count() == 1
                    && rxColumns.Length >= 2 && rxColumns.Distinct().Count() == 1
                    && txColumns.Length >= 2 && txColumns.Distinct().Count() == 1
                    && summaryCells.Length == 4
                    && summaryCells.Take(summaryCells.Length - 1).Select(cell => cell.Length).Distinct().Count() == 1
                    && renderedImod.DisplayRowCount == 11
                    && renderedImod.ValidateColumnLayout(renderedImod.ClientSize.Width)
                    && renderedNic.ValidateColumnLayout(renderedNic.ClientSize.Width);
                if (layoutPassed) passed++; else failed.Add("interrupt columns");
                WriteLog($"TEST.LAYOUT.COLUMNS: status={(layoutPassed ? "PASS" : "FAIL")}");
            }
            catch (Exception ex)
            {
                failed.Add("interrupt columns");
                WriteLog($"TEST.LAYOUT.COLUMNS: status=FAIL error=\"{FlattenLogText(ex.ToString())}\"");
            }
            try
            {
                NicItrProfile i210 = NicItrProfiles.First(profile => profile.FamilyName.StartsWith("Intel I210/I211", StringComparison.Ordinal));
                NicItrProfile i350 = NicItrProfiles.First(profile => profile.FamilyName.StartsWith("Intel I350", StringComparison.Ordinal));
                NicItrProfile i82580 = NicItrProfiles.First(profile => profile.FamilyName.StartsWith("Intel 82580", StringComparison.Ordinal));
                ulong i350RequiredLength = i350.BaseOffset + ((ulong)(i350.MaxQueues - 1) * i350.Stride) + (ulong)(i350.ReadWidth / 8);
                bool profilePassed = i210.MaxQueues == 5
                    && i350.MaxQueues == 25
                    && i210.IntelEitrUnitUs == 1
                    && i350.IntelEitrUnitUs == 1
                    && i82580.IntelEitrUnitUs == 2
                    && FormatNicItrTiming(0x190, i210) == "100 us"
                    && FormatNicItrTiming(0x190, i82580) == "200 us"
                    && TryValidateNicItrMemoryRange(i350, i350RequiredLength, out _)
                    && !TryValidateNicItrMemoryRange(i350, i350RequiredLength - 1, out _);
                if (profilePassed) passed++; else failed.Add("NIC ITR profiles");
                WriteLog($"TEST.NIC.ITR.PROFILES: status={(profilePassed ? "PASS" : "FAIL")} i210={i210.MaxQueues}/{i210.IntelEitrUnitUs}us i350={i350.MaxQueues}/{i350.IntelEitrUnitUs}us");
            }
            catch (Exception ex)
            {
                failed.Add("NIC ITR profiles");
                WriteLog($"TEST.NIC.ITR.PROFILES: status=FAIL error=\"{FlattenLogText(ex.ToString())}\"");
            }
            try
            {
                TestSandboxScenario jsonSource = CaptureCurrentTestScenario("JSON round-trip", TestSandboxOperation.Auto, TestSandboxFault.None, true);
                string json = JsonSerializer.Serialize(jsonSource, TestScenarioJsonOptions());
                TestSandboxScenario? jsonRoundTrip = JsonSerializer.Deserialize<TestSandboxScenario>(json, TestScenarioJsonOptions());
                bool jsonPassed = jsonRoundTrip is not null
                    && jsonRoundTrip.Version == jsonSource.Version
                    && jsonRoundTrip.Cpu.LogicalCount == jsonSource.Cpu.LogicalCount
                    && jsonRoundTrip.Devices.Count == jsonSource.Devices.Count
                    && jsonRoundTrip.Devices.Select(device => device.InstanceId).SequenceEqual(jsonSource.Devices.Select(device => device.InstanceId), StringComparer.OrdinalIgnoreCase);
                if (jsonPassed) passed++; else failed.Add("JSON round-trip");
                WriteLog($"TEST.SCENARIO.JSON: status={(jsonPassed ? "PASS" : "FAIL")} bytes={Encoding.UTF8.GetByteCount(json)} devices={jsonSource.Devices.Count}");
            }
            catch (Exception ex)
            {
                failed.Add("JSON round-trip");
                WriteLog($"TEST.SCENARIO.JSON: status=FAIL error=\"{FlattenLogText(ex.ToString())}\"");
            }
            foreach ((TestSandboxOperation operation, TestSandboxFault fault, bool imod) in matrix)
            {
                RestoreTestSandboxState(baselineState);
                _testSandbox.Backup = null;
                foreach (DeviceBlock block in _blocks)
                {
                    LoadBlockSettings(block);
                }
                TestSandboxScenario run = CaptureCurrentTestScenario($"Matrix {operation}/{fault}", operation, fault, imod);
                TestSandboxRunResult matrixResult = RunTestSandboxScenario(run, showResult: false);
                if (matrixResult.Passed) passed++; else failed.Add($"{operation}/{fault}");
                Application.DoEvents();
            }
            RestoreTestSandboxState(baselineState);
            _testSandbox.Backup = null;
            foreach (DeviceBlock block in _blocks)
            {
                LoadBlockSettings(block);
            }
            RefreshTestDeviceList();
            OperationReport matrixReport = new();
            int totalChecks = matrix.Length + 3;
            if (failed.Count == 0) matrixReport.AddSuccess("CORE MATRIX", $"{passed}/{totalChecks} checks passed");
            else matrixReport.AddError("CORE MATRIX", $"{passed}/{totalChecks} checks passed", $"Failed: {string.Join(", ", failed)}");
            WriteLog($"TEST.MATRIX.FINAL: status={(failed.Count == 0 ? "PASS" : "FAIL")} passed={passed} total={totalChecks} failed=\"{string.Join(",", failed)}\"");
            if (showDialog)
            {
                ShowOperationResult(matrixReport, "Sandbox matrix passed.", "One or more sandbox scenarios failed.", "SANDBOX MATRIX");
            }
            return failed.Count == 0;
        }
        scenarioMatrixButton.Click += (_, _) => _ = RunCoreScenarioMatrix(showDialog: true);
        Button scenarioSaveButton = NewDialogButton("SAVE JSON");
        scenarioSaveButton.Dock = DockStyle.Fill;
        scenarioSaveButton.Click += (_, _) =>
        {
            if (!TrySaveSelectedInitialState(out string error)) { ShowThemedInfo(error); return; }
            using SaveFileDialog save = new() { Filter = "DEVICE TWEAKER scenario (*.json)|*.json", FileName = "DeviceTweaker_Scenario.json", AddExtension = true };
            if (save.ShowDialog(dialog) != DialogResult.OK) return;
            TestSandboxScenario scenario = CaptureCurrentTestScenario(scenarioNameBox.Text, (TestSandboxOperation)(scenarioOperationCombo.SelectedItem ?? TestSandboxOperation.Auto), (TestSandboxFault)(scenarioFaultCombo.SelectedItem ?? TestSandboxFault.None), scenarioImodCheck.Checked);
            File.WriteAllText(save.FileName, JsonSerializer.Serialize(scenario, TestScenarioJsonOptions()), new UTF8Encoding(false));
            WriteLog($"TEST.SCENARIO.SAVE: path={save.FileName}");
        };
        Button scenarioLoadButton = NewDialogButton("LOAD JSON");
        scenarioLoadButton.Dock = DockStyle.Fill;
        scenarioLoadButton.Click += (_, _) =>
        {
            using OpenFileDialog open = new() { Filter = "DEVICE TWEAKER scenario (*.json)|*.json" };
            if (open.ShowDialog(dialog) != DialogResult.OK) return;
            try
            {
                FileInfo scenarioFile = new(open.FileName);
                if (scenarioFile.Length > 2 * 1024 * 1024)
                {
                    throw new InvalidOperationException("Scenario file is larger than 2 MB.");
                }
                TestSandboxScenario? scenario = JsonSerializer.Deserialize<TestSandboxScenario>(File.ReadAllText(open.FileName, Encoding.UTF8), TestScenarioJsonOptions());
                if (scenario is null) throw new InvalidOperationException("Scenario file is empty.");
                LoadTestSandboxScenario(scenario);
                scenarioNameBox.Text = scenario.Name;
                scenarioOperationCombo.SelectedItem = scenario.Operation;
                scenarioFaultCombo.SelectedItem = scenario.Fault;
                scenarioImodCheck.Checked = scenario.OptimizeUsbImod;
                RefreshTestDeviceList();
            }
            catch (Exception ex)
            {
                ShowThemedInfo($"Scenario could not be loaded.\n{ex.Message}");
                WriteLog($"TEST.SCENARIO.LOAD.ERROR: {FlattenLogText(ex.ToString())}");
            }
        };

        TableLayoutPanel scenarioButtons = new()
        {
            Size = new Size(690, 108),
            ColumnCount = 2,
            RowCount = 3,
            Margin = new Padding(0, 6, 0, 2),
            Padding = Padding.Empty,
        };
        scenarioButtons.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
        scenarioButtons.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
        scenarioButtons.RowStyles.Add(new RowStyle(SizeType.Percent, 33.333F));
        scenarioButtons.RowStyles.Add(new RowStyle(SizeType.Percent, 33.333F));
        scenarioButtons.RowStyles.Add(new RowStyle(SizeType.Percent, 33.334F));
        foreach (Button button in new[] { stateApplyButton, scenarioRunButton, scenarioMatrixButton, scenarioSaveButton, scenarioLoadButton })
        {
            button.Margin = new Padding(0, 0, 10, 6);
        }
        scenarioButtons.Controls.Add(stateApplyButton, 0, 0);
        scenarioButtons.Controls.Add(scenarioRunButton, 1, 0);
        scenarioButtons.Controls.Add(scenarioSaveButton, 0, 1);
        scenarioButtons.Controls.Add(scenarioLoadButton, 1, 1);
        scenarioButtons.Controls.Add(scenarioMatrixButton, 0, 2);
        scenarioButtons.SetColumnSpan(scenarioMatrixButton, 2);
        scenarioLayout.Controls.Add(scenarioButtons, 0, scenarioRow);
        scenarioLayout.SetColumnSpan(scenarioButtons, 2);
        scenarioLabPanel.Controls.Add(scenarioLayout);
        layout.Controls.Add(scenarioLabPanel, 0, 16);
        layout.SetColumnSpan(scenarioLabPanel, 2);

        Label resultGalleryLabel = new()
        {
            Text = "Result Dialog Gallery",
            AutoSize = true,
            Font = _titleFont,
            ForeColor = _accent,
            Margin = new Padding(0, 18, 0, 8),
        };
        layout.Controls.Add(resultGalleryLabel, 0, 17);
        layout.SetColumnSpan(resultGalleryLabel, 2);

        Label resultGalleryHint = NewHintLabel(
            "Preview the production result dialogs without running APPLY or AUTO-OPTIMIZATION. Partial also exercises the Vulnerable Driver Blocklist confirm when the blocklist is enabled.");
        resultGalleryHint.Margin = new Padding(0, 0, 0, 12);
        layout.Controls.Add(resultGalleryHint, 0, 18);
        layout.SetColumnSpan(resultGalleryHint, 2);

        FlowLayoutPanel resultGalleryPanel = NewRowFlowPanel();
        resultGalleryPanel.WrapContents = true;
        resultGalleryPanel.Margin = new Padding(0, 0, 0, 10);

        void AttachPreviewBackupPath(OperationReport preview)
        {
            // Prefer an existing managed folder; otherwise point at EXE\Backups so BACKUPS
            // always resolves the same way as production result dialogs.
            string backupDir = EnumerateBackupDirectories().FirstOrDefault(Directory.Exists)
                ?? GetBackupDirectory(BackupLocation.Local);
            preview.SetBackupPath(Path.Combine(backupDir, "DeviceTweakerBackup_ORIGINAL.json"));
        }

        Button resultSuccessButton = NewDialogButton("RESULT: SUCCESS");
        resultSuccessButton.Size = new Size(168, 32);
        resultSuccessButton.Margin = new Padding(0, 0, 10, 8);
        resultSuccessButton.Click += (_, _) =>
        {
            CloseDevicesBusyOverlay();
            OperationReport preview = new();
            preview.AddSuccess("DEVICE SETTINGS", "7 processed");
            preview.AddSuccess("CPU AFFINITY", "4 targets");
            preview.AddSuccess("USB POWER", "Applied");
            preview.AddSuccess("USB IMOD", "Applied");
            AttachPreviewBackupPath(preview);
            ShowOperationResult(
                preview,
                "Please reboot your PC to finish applying the changes.",
                string.Empty,
                operationName: "AUTO-OPTIMIZATION");
        };

        Button resultWarningsButton = NewDialogButton("RESULT: WARNINGS");
        resultWarningsButton.Size = new Size(168, 32);
        resultWarningsButton.Margin = new Padding(0, 0, 10, 8);
        resultWarningsButton.Click += (_, _) =>
        {
            CloseDevicesBusyOverlay();
            OperationReport preview = new();
            preview.AddSuccess("DEVICE SETTINGS", "10 processed");
            preview.AddWarning("HPET TIMER", "Platform timer was already disabled.");
            preview.AddWarning("NETWORK STACK", "QoS packet scheduler restart deferred until next reboot.");
            AttachPreviewBackupPath(preview);
            ShowOperationResult(
                preview,
                "Applied with warnings. Core settings were successfully configured.",
                string.Empty,
                operationName: "APPLY");
        };

        Button resultPartialButton = NewDialogButton("RESULT: PARTIAL");
        resultPartialButton.Size = new Size(168, 32);
        resultPartialButton.Margin = new Padding(0, 0, 10, 8);
        resultPartialButton.Click += (_, _) =>
        {
            CloseDevicesBusyOverlay();
            OperationReport preview = new();
            preview.AddSuccess("DEVICE SETTINGS", "7 processed");
            preview.AddSuccess("CPU AFFINITY", "4 targets");
            preview.AddSuccess("USB POWER", "Applied");
            preview.AddError(
                "USB IMOD",
                FormatImodUnavailableUserMessage(
                    "Windows cannot verify the digital signature for this file. (code 577) (kernel CI blocked; enabled=True testSign=False)",
                    includeNotChanged: true),
                "KDU exit code: 0\r\nDTIMOD device: unavailable\r\nService fallback: Windows error 577\r\nKernel CI: enabled\r\nVulnerable Driver Blocklist: likely enabled\r\nHVCI/VBS: disabled\r\n\r\nThis is simulated TEST ADMIN diagnostic text.");
            AttachPreviewBackupPath(preview);
            ShowOperationResult(
                preview,
                "Core optimization was saved.",
                "Applied steps were saved. One optional step was not completed.",
                operationName: "AUTO-OPTIMIZATION");
            MaybeOfferVulnerableDriverBlocklistDisable(preview);
        };

        Button resultPartialMultiButton = NewDialogButton("RESULT: PARTIAL (3x)");
        resultPartialMultiButton.Size = new Size(168, 32);
        resultPartialMultiButton.Margin = new Padding(0, 0, 10, 8);
        resultPartialMultiButton.Click += (_, _) =>
        {
            CloseDevicesBusyOverlay();
            OperationReport preview = new();
            preview.AddSuccess("DEVICE SETTINGS", "5 processed");
            preview.AddWarning("CPU AFFINITY", "SMT thread balancing kept at default.");
            preview.AddError("USB IMOD", "Kernel CI blocked driver execution (code 577).", "Simulated driver block details.");
            preview.AddError("NETWORK TCP", "Registry access denied for TcpAckFrequency.", "Access is denied (5).");
            AttachPreviewBackupPath(preview);
            ShowOperationResult(
                preview,
                "Applied steps were saved.",
                "Partially applied (5/8 successful, 3 skipped or failed).",
                operationName: "AUTO-OPTIMIZATION");
        };

        Button resultFailedButton = NewDialogButton("RESULT: FAILED");
        resultFailedButton.Size = new Size(168, 32);
        resultFailedButton.Margin = new Padding(0, 0, 10, 8);
        resultFailedButton.Click += (_, _) =>
        {
            CloseDevicesBusyOverlay();
            OperationReport preview = new();
            preview.MarkNoChangesMade();
            preview.AddError(
                "AUTOMATIC BACKUP",
                "The backup could not be created. No changes were made.",
                "UnauthorizedAccessException: simulated access denied while creating the pre-operation backup.");
            AttachPreviewBackupPath(preview);
            ShowOperationResult(
                preview,
                string.Empty,
                "The operation stopped before any changes were made.",
                operationName: "APPLY");
        };

        Button resultStressButton = NewDialogButton("RESULT: STRESS");
        resultStressButton.Size = new Size(168, 32);
        resultStressButton.Margin = new Padding(0, 0, 10, 8);
        resultStressButton.Click += (_, _) =>
        {
            CloseDevicesBusyOverlay();
            OperationReport preview = new();
            preview.AddSuccess("DEVICE SETTINGS", "12 processed");
            for (int i = 1; i <= 14; i++)
            {
                preview.AddError(
                    $"TEST COMPONENT {i:D2}",
                    i == 1 ? "A deliberately long summary verifies wrapping without exposing raw diagnostics in the main result." : null,
                    $"Simulated technical detail #{i:D2}\r\n" + new string((char)('A' + ((i - 1) % 26)), 420));
            }
            AttachPreviewBackupPath(preview);
            ShowOperationResult(
                preview,
                "Some test steps completed.",
                "Stress preview: verify clipping, wrapping, scrolling and DETAILS.",
                operationName: "GUI STRESS TEST");
        };

        Button promptImodButton = NewDialogButton("PROMPT: IMOD");
        promptImodButton.Size = new Size(168, 32);
        promptImodButton.Margin = new Padding(0, 0, 10, 8);
        promptImodButton.Click += (_, _) =>
        {
            CloseDevicesBusyOverlay();
            ShowThemedConfirm(
                "USB IMOD tuning is available for detected XHCI controller(s).\n\nDTIMOD can be blocked by Vulnerable Driver Blocklist, Windows driver signature protection, antivirus, or anti-cheats.\n\nApply it during AUTO-OPTIMIZATION?",
                "USB IMOD TUNING",
                "APPLY",
                "SKIP");
        };

        Button promptConfirmButton = NewDialogButton("PROMPT: CONFIRM");
        promptConfirmButton.Size = new Size(168, 32);
        promptConfirmButton.Margin = new Padding(0, 0, 10, 8);
        promptConfirmButton.Click += (_, _) =>
        {
            CloseDevicesBusyOverlay();
            ShowThemedConfirm(
                "Reset all active network tweaks and restore default Windows adapter parameters?\n\nThis will re-enable interrupt moderation and standard RSS ring buffers.",
                "RESET ADAPTER SETTINGS",
                "RESET",
                "CANCEL");
        };

        Button promptBackupButton = NewDialogButton("PROMPT: BACKUP");
        promptBackupButton.Size = new Size(168, 32);
        promptBackupButton.Margin = new Padding(0, 0, 10, 8);
        promptBackupButton.Click += (_, _) =>
        {
            CloseDevicesBusyOverlay();
            ShowAutoBackupChoiceDialog();
        };

        Button promptRestoreButton = NewDialogButton("PROMPT: RESTORE");
        promptRestoreButton.Size = new Size(168, 32);
        promptRestoreButton.Margin = new Padding(0, 0, 10, 8);
        promptRestoreButton.Click += (_, _) =>
        {
            CloseDevicesBusyOverlay();
            string backupDir = EnumerateBackupDirectories().FirstOrDefault(Directory.Exists)
                ?? GetBackupDirectory(BackupLocation.Local);
            List<BackupSnapshotInfo> testBackups =
            [
                new BackupSnapshotInfo
                {
                    Path = Path.Combine(backupDir, "DeviceTweakerBackup_20260920_180000.json"),
                    Location = "EXE",
                    LastWriteUtc = DateTime.UtcNow.AddHours(-2),
                    CreatedAt = DateTime.Now.AddHours(-2),
                    Reason = "pre-auto",
                    IsOriginal = false,
                },
                new BackupSnapshotInfo
                {
                    Path = Path.Combine(backupDir, "DeviceTweakerBackup_ORIGINAL.json"),
                    Location = "EXE",
                    LastWriteUtc = DateTime.UtcNow.AddDays(-7),
                    CreatedAt = DateTime.Now.AddDays(-7),
                    Reason = "original-state",
                    IsOriginal = true,
                },
            ];
            ShowRestoreChoiceDialog(testBackups, out _);
        };

        Button promptRestoreEmptyButton = NewDialogButton("PROMPT: RESTORE (0)");
        promptRestoreEmptyButton.Size = new Size(168, 32);
        promptRestoreEmptyButton.Margin = new Padding(0, 0, 10, 8);
        promptRestoreEmptyButton.Click += (_, _) =>
        {
            CloseDevicesBusyOverlay();
            List<BackupSnapshotInfo> emptyBackups = [];
            ShowRestoreChoiceDialog(emptyBackups, out _);
        };

        Button promptInfoButton = NewDialogButton("PROMPT: INFO");
        promptInfoButton.Size = new Size(168, 32);
        promptInfoButton.Margin = new Padding(0, 0, 0, 8);
        promptInfoButton.Click += (_, _) =>
        {
            CloseDevicesBusyOverlay();
            ShowThemedInfo(
                "DEVICE TWEAKER is running in TEST ADMIN mode.\n\nAll hardware adjustments and registry modifications are simulated and completely safe.",
                "TEST MODE INFO");
        };

        resultGalleryPanel.Controls.Add(resultSuccessButton);
        resultGalleryPanel.Controls.Add(resultWarningsButton);
        resultGalleryPanel.Controls.Add(resultPartialButton);
        resultGalleryPanel.SetFlowBreak(resultPartialButton, true);
        resultGalleryPanel.Controls.Add(resultPartialMultiButton);
        resultGalleryPanel.Controls.Add(resultFailedButton);
        resultGalleryPanel.Controls.Add(resultStressButton);
        resultGalleryPanel.SetFlowBreak(resultStressButton, true);
        resultGalleryPanel.Controls.Add(promptImodButton);
        resultGalleryPanel.Controls.Add(promptConfirmButton);
        resultGalleryPanel.Controls.Add(promptBackupButton);
        resultGalleryPanel.SetFlowBreak(promptBackupButton, true);
        resultGalleryPanel.Controls.Add(promptRestoreButton);
        resultGalleryPanel.Controls.Add(promptRestoreEmptyButton);
        resultGalleryPanel.Controls.Add(promptInfoButton);
        layout.Controls.Add(resultGalleryPanel, 0, 19);
        layout.SetColumnSpan(resultGalleryPanel, 2);

        bool suppressTestDeviceToggle = false;
        bool suppressTestDeviceSelection = false;
        List<DeviceInfo> realVisibleDevices = [];
        List<string> hiddenDeviceKeys = [];

        void RefreshTestDeviceList(int selectIndex = -1)
        {
            int previousIndex = testDeviceListBox.SelectedIndex;
            int targetIndex = selectIndex >= 0 ? selectIndex : previousIndex;
            suppressTestDeviceSelection = true;
            testDeviceListBox.BeginUpdate();
            testDeviceListBox.Items.Clear();
            foreach (DeviceInfo device in _testDevices)
            {
                testDeviceListBox.Items.Add(FormatTestDeviceLabel(device));
            }
            testDeviceListBox.EndUpdate();
            SyncThemedListHost(testDeviceListBox);
            if (targetIndex >= 0 && targetIndex < _testDevices.Count)
            {
                testDeviceListBox.SelectedIndex = targetIndex;
            }
            else
            {
                testDeviceListBox.ClearSelected();
            }
            suppressTestDeviceSelection = false;
            updateTestDeviceButton.Enabled = testDeviceListBox.SelectedIndex >= 0;
            if (selectIndex >= 0 && testDeviceListBox.SelectedIndex >= 0 && testDeviceListBox.SelectedIndex < _testDevices.Count)
            {
                LoadTestDeviceToFields(_testDevices[testDeviceListBox.SelectedIndex]);
            }
            testListLabel.Text = $"Current test devices: {_testDevices.Count}";

            int previousScenarioIndex = scenarioDeviceCombo.SelectedIndex;
            scenarioDeviceCombo.Items.Clear();
            foreach (DeviceInfo device in _testDevices)
            {
                scenarioDeviceCombo.Items.Add($"{device.Kind}: {device.Name} | {device.InstanceId}");
            }
            int scenarioIndex = selectIndex >= 0 ? selectIndex : previousScenarioIndex;
            if (_testDevices.Count > 0)
            {
                scenarioDeviceCombo.SelectedIndex = Math.Max(0, Math.Min(_testDevices.Count - 1, scenarioIndex));
            }
        }

        scenarioDeviceCombo.SelectedIndexChanged += (_, _) => LoadSelectedInitialState();

        void RefreshRealDeviceVisibilityLists()
        {
            realVisibleDevices = _blocks
                .Where(block => !block.Device.IsTestDevice)
                .Select(block => block.Device)
                .GroupBy(device => NormalizeInstanceId(device.InstanceId), StringComparer.OrdinalIgnoreCase)
                .Where(group => !string.IsNullOrWhiteSpace(group.Key))
                .Select(group => group.First())
                .OrderBy(device => device.Kind)
                .ThenBy(device => device.Name, StringComparer.OrdinalIgnoreCase)
                .ToList();

            realDeviceListBox.BeginUpdate();
            realDeviceListBox.Items.Clear();
            foreach (DeviceInfo device in realVisibleDevices)
            {
                string label = $"{device.Kind}: {BuildDeviceBlockTitle(device)}";
                realDeviceListBox.Items.Add(label);
            }
            realDeviceListBox.EndUpdate();
            SyncThemedListHost(realDeviceListBox);

            hiddenDeviceKeys = _testHiddenDeviceIds
                .OrderBy(key => _testHiddenDeviceLabels.TryGetValue(key, out string? label) ? label : key, StringComparer.OrdinalIgnoreCase)
                .ToList();

            hiddenDeviceListBox.BeginUpdate();
            hiddenDeviceListBox.Items.Clear();
            foreach (string key in hiddenDeviceKeys)
            {
                string label = _testHiddenDeviceLabels.TryGetValue(key, out string? value) ? value : key;
                hiddenDeviceListBox.Items.Add(label);
            }
            hiddenDeviceListBox.EndUpdate();
            SyncThemedListHost(hiddenDeviceListBox);

            realVisibilityLabel.Text = $"Real device visibility: visible={realVisibleDevices.Count}, hidden={hiddenDeviceKeys.Count}";
        }

        void RefreshAfterRealDeviceVisibilityChange()
        {
            _initialDeviceViewportHeightAdjusted = false;
            RefreshBlocks();
            RefreshRealDeviceVisibilityLists();
        }

        void ResetAdminPanelToRealState()
        {
            _testCpuActive = false;
            _testCpuName = string.Empty;
            _testDevices.Clear();
            _testHiddenDeviceIds.Clear();
            _testHiddenDeviceLabels.Clear();
            _testDevicesEnabled = false;
            _testDevicesOnly = false;
            _testAutoDryRun = false;
            _testDeviceSequence = 0;

            suppressTestDeviceToggle = true;
            enableTestDevicesCheck.Checked = false;
            testDevicesOnlyCheck.Checked = false;
            testDevicesOnlyCheck.Enabled = false;
            dryRunAutoCheck.Checked = false;
            dryRunAutoCheck.AutoCheck = true;
            dryRunAutoCheck.Enabled = true;
            suppressTestDeviceToggle = false;

            systemPresetCombo.SelectedIndex = 0;
            cpuPresetCombo.SelectedIndex = 0;
            testDevicePresetCombo.SelectedIndex = 0;
            SetTestDeviceFields(DeviceKind.USB, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
            RefreshTestDeviceList();

            statusLabel.Text = "Test CPU mode: OFF";
            statusLabel.ForeColor = _statusInactive;
            InitializeCpu();
            logicalUpDown.Value = Math.Min(MaxAffinityBits, GetCurrentLogicalCount());
            LoadAssignmentsFromCurrentCpu();
            SyncSmtStateFromCurrent();
            cppcRatingsBox.Text = GetCurrentCppcRatingsText();
            NotifySandboxModeChanged("reset-to-real");
            UpdateCpuHeaderUi();

            _initialDeviceViewportHeightAdjusted = false;
            RefreshBlocks();
            RefreshRealDeviceVisibilityLists();
            WriteLog("TEST.RESET.REAL: full admin test state cleared");
        }

        void UpdateTestDeviceFieldState()
        {
            if (testKindCombo.SelectedItem is not DeviceKind kind)
            {
                return;
            }

            bool isUsb = kind == DeviceKind.USB;
            bool isAudio = kind == DeviceKind.AUDIO;
            bool isNet = kind is DeviceKind.NET_NDIS or DeviceKind.NET_CX;
            bool isStor = kind == DeviceKind.STOR;
            bool isGpu = kind == DeviceKind.GPU;
            bool hasOptions = isUsb || isNet || isGpu;

            testUsbRolesBox.Enabled = isUsb;
            testAudioBox.Enabled = isAudio;
            testStorageBox.Enabled = isStor;

            testDeviceOptionsLabel.Visible = hasOptions;
            testDeviceOptionsPanel.Visible = hasOptions;
            testWifiCheck.Visible = isNet;
            testWifiCheck.Enabled = isNet;
            testWifiCheck.TabStop = isNet;
            testXhciCheck.Visible = isUsb;
            testXhciCheck.Enabled = isUsb;
            testXhciCheck.TabStop = isUsb;
            testHasDevicesCheck.Visible = isUsb;
            testHasDevicesCheck.Enabled = isUsb;
            testHasDevicesCheck.TabStop = isUsb;
            testSuspendLabel.Visible = isUsb;
            testSuspendCombo.Visible = isUsb;
            testSuspendCombo.Enabled = isUsb;
            testSuspendCombo.TabStop = isUsb;
            bool showNicPower = isNet && !testWifiCheck.Checked;
            testNicPowerLabel.Visible = showNicPower;
            testNicPowerCombo.Visible = showNicPower;
            testNicPowerCombo.Enabled = showNicPower;
            testNicPowerCombo.TabStop = showNicPower;
            testIntegratedGpuCheck.Visible = isGpu;
            testIntegratedGpuCheck.Enabled = isGpu;
            testIntegratedGpuCheck.TabStop = isGpu;

            if (!isNet)
            {
                testWifiCheck.Checked = false;
            }

            if (!isGpu)
            {
                testIntegratedGpuCheck.Checked = false;
            }

            if (isUsb && string.IsNullOrWhiteSpace(testUsbRolesBox.Text))
            {
                testUsbRolesBox.Text = "Microphone";
            }

            if (isAudio && string.IsNullOrWhiteSpace(testAudioBox.Text))
            {
                testAudioBox.Text = "Speakers";
            }

            if (isStor && string.IsNullOrWhiteSpace(testStorageBox.Text))
            {
                testStorageBox.Text = "SSD";
            }
        }

        void SetTestDeviceFields(
            DeviceKind kind,
            string name,
            string pnpId,
            string usbRoles = "",
            string audioEndpoints = "",
            string storageTag = "",
            bool wifi = false,
            bool usbIsXhci = true,
            bool usbHasDevices = true,
            bool integratedGpu = false,
            int? testIrqCount = null,
            string testMsiStatus = "Auto",
            string? usbSelectiveSuspend = "off",
            string? nicPowerSaving = "on")
        {
            testKindCombo.SelectedItem = kind;
            testNameBox.Text = name;
            testPnpIdBox.Text = pnpId;
            testUsbRolesBox.Text = usbRoles;
            testAudioBox.Text = audioEndpoints;
            testStorageBox.Text = storageTag;
            testIrqCountBox.Text = testIrqCount.HasValue ? testIrqCount.Value.ToString(CultureInfo.InvariantCulture) : "Auto";
            testMsiStatusCombo.SelectedItem = string.IsNullOrWhiteSpace(testMsiStatus) ? "Auto" : testMsiStatus;
            if (testMsiStatusCombo.SelectedItem is null)
            {
                testMsiStatusCombo.SelectedItem = "Auto";
            }

            string suspendItem = usbSelectiveSuspend switch
            {
                "on" => "on",
                "off" => "off",
                _ => "unset",
            };
            testSuspendCombo.SelectedItem = suspendItem;
            if (testSuspendCombo.SelectedItem is null)
            {
                testSuspendCombo.SelectedItem = "off";
            }

            testNicPowerCombo.SelectedItem = string.Equals(nicPowerSaving, "off", StringComparison.OrdinalIgnoreCase)
                ? "off"
                : "on";
            if (testNicPowerCombo.SelectedItem is null)
            {
                testNicPowerCombo.SelectedItem = "on";
            }

            testWifiCheck.Checked = wifi;
            testXhciCheck.Checked = kind == DeviceKind.USB && usbIsXhci;
            testHasDevicesCheck.Checked = kind == DeviceKind.USB && usbHasDevices;
            testIntegratedGpuCheck.Checked = kind == DeviceKind.GPU && integratedGpu;
            UpdateTestDeviceFieldState();
        }

        void LoadTestDeviceToFields(DeviceInfo device)
        {
            SetTestDeviceFields(
                device.Kind,
                device.Name,
                device.InstanceId,
                device.UsbRoles,
                device.AudioEndpoints,
                device.StorageTag,
                device.Wifi,
                device.UsbIsXhci,
                device.UsbHasDevices,
                device.IsIntegratedGpu,
                device.TestIrqCount,
                device.TestMsiStatus,
                device.UsbSelectiveSuspend,
                device.NicPowerSaving);
        }

        void AddSystemPresetDevice(
            DeviceKind kind,
            string name,
            string pnpId,
            string usbRoles = "",
            string audioEndpoints = "",
            string storageTag = "",
            bool wifi = false,
            bool usbIsXhci = true,
            bool usbHasDevices = true,
            bool integratedGpu = false,
            int? testIrqCount = null,
            string testMsiStatus = "Auto",
            string? nicPowerSaving = "on")
        {
            DeviceInfo testDevice = CreateTestDevice(
                kind,
                name,
                pnpId,
                usbRoles,
                audioEndpoints,
                storageTag,
                wifi,
                usbIsXhci,
                usbHasDevices,
                integratedGpu,
                testIrqCount,
                testMsiStatus,
                usbSelectiveSuspend: "off",
                nicPowerSaving: wifi ? null : nicPowerSaving);
            _testDevices.Add(testDevice);
            WriteLog($"TEST.SYSTEM.DEV: {testDevice.InstanceId} Kind={kind} Name=\"{testDevice.Name}\"");
            if (testDevice.Kind == DeviceKind.USB && testDevice.UsbChipPath is UsbChipPathInfo chip)
            {
                WriteLog(
                    $"TEST.USB.CHIP: {testDevice.InstanceId} {chip.CompactTag} origin={chip.Origin} " +
                    $"platform=\"{chip.Platform}\" usb=\"{chip.UsbSpec}\"");
                WarnTestUsbChipMismatch(testDevice.InstanceId, chip);
            }
        }

        void AddSystemStorage(string name, string pnpId, string storageTag = "SSD")
        {
            AddSystemPresetDevice(DeviceKind.STOR, name, pnpId, storageTag: storageTag);
        }

        void AddSystemAudio(string name, string pnpId, string endpoints)
        {
            AddSystemPresetDevice(DeviceKind.AUDIO, name, pnpId, audioEndpoints: endpoints);
        }

        void AddNvidiaDisplayAudio(string suffix)
        {
            AddSystemAudio("NVIDIA High Definition Audio", $@"HDAUDIO\FUNC_01&VEN_10DE&DEV_00A1\{suffix}", "Monitor DisplayPort");
        }

        void AddAmdDisplayAudio(string suffix)
        {
            AddSystemAudio("AMD High Definition Audio Device", $@"HDAUDIO\FUNC_01&VEN_1002&DEV_AAF0\{suffix}", "Monitor DisplayPort");
        }

        void AddIntelDisplayAudio(string suffix)
        {
            AddSystemAudio("Intel(R) Display Audio", $@"HDAUDIO\FUNC_01&VEN_8086&DEV_280F\{suffix}", "Monitor DisplayPort");
        }

        void EnableSystemPresetTestMode()
        {
            suppressTestDeviceToggle = true;
            _testDevicesEnabled = true;
            _testDevicesOnly = true;
            _testAutoDryRun = true;
            enableTestDevicesCheck.Checked = true;
            testDevicesOnlyCheck.Checked = true;
            testDevicesOnlyCheck.Enabled = true;
            dryRunAutoCheck.Checked = true;
            dryRunAutoCheck.AutoCheck = false;
            dryRunAutoCheck.Enabled = true;
            suppressTestDeviceToggle = false;
            NotifySandboxModeChanged("system-preset");
        }

        void ApplySystemCpuPreset(string cpuPresetName)
        {
            int previousIndex = cpuPresetCombo.SelectedIndex;
            cpuPresetCombo.SelectedItem = cpuPresetName;
            if (!string.Equals(cpuPresetCombo.SelectedItem?.ToString(), cpuPresetName, StringComparison.Ordinal))
            {
                WriteLog($"TEST.SYSTEM.PRESET: CPU preset not found name=\"{cpuPresetName}\"");
                return;
            }

            if (cpuPresetCombo.SelectedIndex == previousIndex)
            {
                ApplySelectedCpuPreset();
            }
        }

        Dictionary<string, (string Msi, string Limit, string Priority, string Policy, ulong Mask)> SeedAutoPolicyRegressionState()
        {
            Dictionary<string, (string Msi, string Limit, string Priority, string Policy, ulong Mask)> preservedExpected =
                new(StringComparer.OrdinalIgnoreCase);

            foreach (DeviceBlock block in _blocks)
            {
                if (IsAutoDisplayAudioMsiOnly(block))
                {
                    block.MsiCombo.SelectedItem = "Disabled";
                    block.LimitBox.Text = "17";
                    block.PrioCombo.SelectedItem = "Normal";
                    block.PolicyCombo.SelectedItem = "SpecCPU";
                    block.SuppressCpuEvents++;
                    try
                    {
                        for (int i = 0; i < block.CpuBoxes.Count; i++)
                        {
                            block.CpuBoxes[i].Checked = i == 0;
                        }
                    }
                    finally
                    {
                        block.SuppressCpuEvents--;
                    }

                    RecalcAffinityMask(block);
                    preservedExpected[block.Device.InstanceId] =
                        ("Enabled", block.LimitBox.Text, block.PrioCombo.SelectedItem?.ToString() ?? string.Empty,
                            block.PolicyCombo.SelectedItem?.ToString() ?? string.Empty, block.AffinityMask);
                }
                else if (block.Device.Wifi)
                {
                    block.MsiCombo.SelectedItem = "Disabled";
                    block.LimitBox.Text = "8";
                    block.PrioCombo.SelectedItem = "Low";
                    block.RssBaseCore = 3;
                    if (block.RssQueueBox is not null)
                    {
                        block.RssQueueBox.Value = Math.Min(2, block.RssQueueBox.Maximum);
                    }
                    if (block.NdisModeCombo is not null)
                    {
                        block.NdisModeCombo.SelectedItem = "BOTH";
                    }
                    preservedExpected[block.Device.InstanceId] =
                        ("Disabled", "8", "Low", block.PolicyCombo.SelectedItem?.ToString() ?? string.Empty, block.AffinityMask);
                }
                else if (block.Kind == DeviceKind.STOR)
                {
                    block.MsiCombo.SelectedItem = "Disabled";
                    block.LimitBox.Text = "2048";
                    block.PrioCombo.SelectedItem = "Normal";
                    block.PolicyCombo.SelectedItem = "SpreadMessages";
                }
            }

            return preservedExpected;
        }

        void ValidateAutoPolicyRegressionState(
            IReadOnlyDictionary<string, (string Msi, string Limit, string Priority, string Policy, ulong Mask)> preservedExpected)
        {
            int failures = 0;
            foreach (DeviceBlock block in _blocks.Where(IsAutoDisplayAudioMsiOnly))
            {
                (string msi, string limit, string priority, string policy, ulong mask) = preservedExpected[block.Device.InstanceId];
                bool passed = string.Equals(block.MsiCombo.SelectedItem?.ToString(), msi, StringComparison.Ordinal)
                    && string.Equals(block.LimitBox.Text, limit, StringComparison.Ordinal)
                    && string.Equals(block.PrioCombo.SelectedItem?.ToString(), priority, StringComparison.Ordinal)
                    && string.Equals(block.PolicyCombo.SelectedItem?.ToString(), policy, StringComparison.Ordinal)
                    && block.AffinityMask == mask;
                failures += passed ? 0 : 1;
                WriteLog(
                    $"TEST.AUTO.POLICY.DISPLAY-AUDIO: status={(passed ? "PASS" : "FAIL")} id={block.Device.InstanceId} " +
                    $"MSI={block.MsiCombo.SelectedItem} limit={block.LimitBox.Text}/{limit} " +
                    $"priority={block.PrioCombo.SelectedItem}/{priority} policy={block.PolicyCombo.SelectedItem}/{policy} " +
                    $"mask=0x{block.AffinityMask:X}/0x{mask:X}");
            }

            foreach (DeviceBlock block in _blocks.Where(block => block.Device.Wifi))
            {
                (string msi, string limit, string priority, string policy, ulong mask) = preservedExpected[block.Device.InstanceId];
                bool passed = string.Equals(block.MsiCombo.SelectedItem?.ToString(), msi, StringComparison.Ordinal)
                    && string.Equals(block.LimitBox.Text, limit, StringComparison.Ordinal)
                    && string.Equals(block.PrioCombo.SelectedItem?.ToString(), priority, StringComparison.Ordinal)
                    && string.Equals(block.PolicyCombo.SelectedItem?.ToString(), policy, StringComparison.Ordinal)
                    && block.AffinityMask == mask
                    && block.RssBaseCore == 3
                    && block.RssQueueBox?.Value == 2
                    && string.Equals(block.NdisModeCombo?.SelectedItem?.ToString(), "BOTH", StringComparison.Ordinal);
                failures += passed ? 0 : 1;
                WriteLog(
                    $"TEST.AUTO.POLICY.WIFI: status={(passed ? "PASS" : "FAIL")} id={block.Device.InstanceId} " +
                    $"MSI={block.MsiCombo.SelectedItem}/{msi} limit={block.LimitBox.Text}/{limit} " +
                    $"priority={block.PrioCombo.SelectedItem}/{priority} policy={block.PolicyCombo.SelectedItem}/{policy} " +
                    $"mask=0x{block.AffinityMask:X}/0x{mask:X} rssBase={block.RssBaseCore}/3 " +
                    $"rssQueues={block.RssQueueBox?.Value}/2 mode={block.NdisModeCombo?.SelectedItem}/BOTH");
            }

            foreach (DeviceBlock block in _blocks.Where(block => block.Kind == DeviceKind.STOR))
            {
                bool passed = string.Equals(block.MsiCombo.SelectedItem?.ToString(), "Enabled", StringComparison.Ordinal)
                    && string.Equals(block.LimitBox.Text, "0", StringComparison.Ordinal)
                    && string.Equals(block.PrioCombo.SelectedItem?.ToString(), "High", StringComparison.Ordinal)
                    && block.AffinityMask == 0;
                failures += passed ? 0 : 1;
                WriteLog(
                    $"TEST.AUTO.POLICY.STORAGE: status={(passed ? "PASS" : "FAIL")} id={block.Device.InstanceId} " +
                    $"MSI={block.MsiCombo.SelectedItem} limit={block.LimitBox.Text} priority={block.PrioCombo.SelectedItem} " +
                    $"mask=0x{block.AffinityMask:X}");
            }

            bool imodStatesPassed =
                FormatAutoImodRequestState(optimizeUsbImod: false, hasUsbImodTarget: true) == "skipped: declined by user"
                && FormatAutoImodRequestState(optimizeUsbImod: false, hasUsbImodTarget: false) == "skipped: no eligible XHCI controller"
                && GetAutoImodSkipReason(hasUsbImodTarget: true) == "user-declined"
                && GetAutoImodSkipReason(hasUsbImodTarget: false) == "no-eligible-xhci-controller";
            failures += imodStatesPassed ? 0 : 1;
            WriteLog($"TEST.AUTO.POLICY.IMOD-STATES: status={(imodStatesPassed ? "PASS" : "FAIL")}");

            (string Text, int Value)[] affinityPolicies =
            [
                ("MachineDefault", 0),
                ("AllClose", 1),
                ("Single", 2),
                ("All", 3),
                ("SpecCPU", 4),
                ("SpreadMessages", 5),
            ];
            bool affinityPoliciesPassed = affinityPolicies.All(policy =>
                MapPolicyText(policy.Text) == policy.Value
                && string.Equals(FormatPolicyValue(policy.Value), policy.Text, StringComparison.Ordinal));
            failures += affinityPoliciesPassed ? 0 : 1;
            WriteLog($"TEST.AUTO.POLICY.IRQ-MAP: status={(affinityPoliciesPassed ? "PASS" : "FAIL")}");

            (string Text, bool Valid, bool Unlocked, int Value)[] msiLimits =
            [
                ("", true, true, 0),
                ("0", true, true, 0),
                ("unlocked", true, true, 0),
                ("unlimited", true, true, 0),
                ("1", true, false, 1),
                ("16", true, false, 16),
                ("2048", true, false, 2048),
                ("-1", false, false, 0),
                ("2049", false, false, 0),
                ("1.5", false, false, 0),
                ("invalid", false, false, 0),
            ];
            bool msiLimitsPassed = msiLimits.All(test =>
            {
                bool valid = TryParseMsiLimitInput(test.Text, out bool unlocked, out int value);
                return valid == test.Valid
                    && (!valid || (unlocked == test.Unlocked && value == test.Value));
            });
            failures += msiLimitsPassed ? 0 : 1;
            WriteLog($"TEST.AUTO.POLICY.MSI-LIMIT: status={(msiLimitsPassed ? "PASS" : "FAIL")}");
            WriteLog($"TEST.AUTO.POLICY.FINAL: status={(failures == 0 ? "PASS" : "FAIL")} failures={failures}");
        }

        bool LoadSystemPreset()
        {
            string preset = systemPresetCombo.SelectedItem?.ToString() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(preset)
                || string.Equals(preset, "Manual (current)", StringComparison.Ordinal)
                || string.Equals(preset, "Manual / current", StringComparison.Ordinal))
            {
                return false;
            }

            string cpuPreset = preset switch
            {
                SystemPresetIntel14900K => "Intel Core i9-13900K/14900K 8P+16E/32T",
                SystemPresetIntel14600K => "Intel Core i5-13600K/14600K 6P+8E/20T",
                SystemPresetIntel285K5090 => "Intel Core Ultra 9 285K 8P+16E/24T",
                SystemPresetLaptopIntel285HX => "Intel Core Ultra 9 285HX 8P+16E/24T",
                SystemPresetRyzen7800X3D => "AMD Ryzen 7 7800X3D/9800X3D 8C/16T V-Cache",
                SystemPresetRyzen9800X3D5090 => "AMD Ryzen 7 7800X3D/9800X3D 8C/16T V-Cache",
                SystemPresetRyzen9850X3DClient => "AMD Ryzen 7 7800X3D/9800X3D 8C/16T V-Cache",
                SystemPresetRyzen9800X3DRx => "AMD Ryzen 7 7800X3D/9800X3D 8C/16T V-Cache",
                SystemPresetRyzen9950X3DNetCx => "AMD Ryzen 9 7950X3D/9950X3D 16C/32T V-Cache CCD0",
                SystemPresetRyzen9950X3DNdis => "AMD Ryzen 9 7950X3D/9950X3D 16C/32T V-Cache CCD0",
                SystemPresetRyzen9950X3D2 => "AMD Ryzen 9 9950X3D2 16C/32T dual V-Cache",
                SystemPresetLaptopRyzen9955HX3D => "AMD Ryzen 9 9955HX3D 16C/32T V-Cache CCD0",
                SystemPresetLaptopRyzen8940HX => "AMD Ryzen 9 8940HX 16C/32T Dual-CCD",
                SystemPresetLaptopRyzen8940HXMouse => "AMD Ryzen 9 8940HX 16C/32T Dual-CCD",
                SystemPresetMultiControllerInput => "AMD Ryzen 9 7950X/9950X 16C/32T",
                SystemPresetRyzen3500X => "AMD Ryzen 5 3500X Zen2 6C/6T SMT off 2 CCX",
                SystemPresetRyzen3900X => "AMD Ryzen 9 3900X Zen2 12C/24T 4 CCX",
                SystemPresetRyzen9950X => "AMD Ryzen 9 7950X/9950X 16C/32T",
                _ => string.Empty,
            };

            string cpuDisplayName = preset switch
            {
                SystemPresetIntel14900K => "Intel Core i9-14900K",
                SystemPresetIntel14600K => "Intel Core i5-14600K",
                SystemPresetIntel285K5090 => "Intel Core Ultra 9 285K",
                SystemPresetLaptopIntel285HX => "Intel Core Ultra 9 285HX",
                SystemPresetRyzen7800X3D => "AMD Ryzen 7 7800X3D",
                SystemPresetRyzen9800X3D5090 => "AMD Ryzen 7 9800X3D",
                SystemPresetRyzen9850X3DClient => "AMD Ryzen 7 9850X3D",
                SystemPresetRyzen9800X3DRx => "AMD Ryzen 7 9800X3D",
                SystemPresetRyzen9950X3DNetCx => "AMD Ryzen 9 9950X3D",
                SystemPresetRyzen9950X3DNdis => "AMD Ryzen 9 9950X3D",
                SystemPresetRyzen9950X3D2 => "AMD Ryzen 9 9950X3D2",
                SystemPresetLaptopRyzen9955HX3D => "AMD Ryzen 9 9955HX3D",
                SystemPresetLaptopRyzen8940HX => "AMD Ryzen 9 8940HX with Radeon Graphics",
                SystemPresetLaptopRyzen8940HXMouse => "AMD Ryzen 9 8940HX with Radeon Graphics",
                SystemPresetMultiControllerInput => "AMD Ryzen 9 9950X",
                SystemPresetRyzen3500X => "AMD Ryzen 5 3500X 6-Core Processor",
                SystemPresetRyzen3900X => "AMD Ryzen 9 3900X",
                SystemPresetRyzen9950X => "AMD Ryzen 9 9950X",
                _ => string.Empty,
            };

            if (string.IsNullOrWhiteSpace(cpuPreset))
            {
                return false;
            }

            _testDevices.Clear();
            EnableSystemPresetTestMode();
            ApplySystemCpuPreset(cpuPreset);
            if (!string.IsNullOrWhiteSpace(cpuDisplayName))
            {
                cpuNameTextBox.Text = cpuDisplayName;
            }

            if (!TryParseTestCppcRatings(cppcRatingsBox.Text, (int)logicalUpDown.Value, out Dictionary<int, int> cppcRatings, out string cppcError))
            {
                WriteLog($"TEST.SYSTEM.PRESET: CPPC parse failed name=\"{preset}\" error=\"{cppcError}\"");
                return false;
            }

            ApplyTestCpuConfig(BuildConfigFromAssignments(cppcRatings));
            statusLabel.Text = "Test CPU mode: ACTIVE";
            statusLabel.ForeColor = _statusActive;

            switch (preset)
            {
                case SystemPresetIntel14900K:
                    AddSystemPresetDevice(DeviceKind.USB, "Intel(R) USB 3.20 eXtensible Host Controller - 1.20", @"PCI\VEN_8086&DEV_7AE0\TEST_Z790_XHCI_FULL", "Mouse 8K, Keyboard 8K, Audio, Microphone, Gamepad");
                    AddSystemAudio("Realtek ALC4080 USB Audio", @"USB\VID_0BDA&PID_402E\SYS_Z790_ALC4080", "Speakers, Microphone");
                    AddSystemPresetDevice(DeviceKind.GPU, "NVIDIA GeForce RTX 4090", @"PCI\VEN_10DE&DEV_2684\SYS_RTX4090");
                    AddNvidiaDisplayAudio("SYS_RTX4090_AUDIO");
                    AddSystemPresetDevice(DeviceKind.NET_NDIS, "Intel Ethernet Controller I226-V", @"PCI\VEN_8086&DEV_125C\SYS_Z790_I226_V");
                    AddSystemStorage("Samsung 990 PRO NVMe Controller", @"PCI\VEN_144D&DEV_A80C\SYS_Z790_990PRO");
                    break;
                case SystemPresetIntel14600K:
                    AddSystemPresetDevice(DeviceKind.USB, "Intel(R) USB 3.20 eXtensible Host Controller - 1.20", @"PCI\VEN_8086&DEV_7AE0\SYS_INTEL_Z790_XHCI_MOUSE_8K", "Mouse 8K");
                    AddSystemPresetDevice(DeviceKind.USB, "Intel(R) USB 3.10 eXtensible Host Controller - 1.20", @"PCI\VEN_8086&DEV_7AE1\SYS_INTEL_Z790_XHCI_KEYBOARD_AUDIO", "Keyboard 8K, Gamepad, Audio, Microphone");
                    AddSystemAudio("Realtek ALC897 High Definition Audio", @"HDAUDIO\FUNC_01&VEN_10EC&DEV_0897\SYS_Z790_ALC897", "Speakers, Microphone");
                    AddSystemPresetDevice(DeviceKind.GPU, "NVIDIA GeForce RTX 4070 SUPER", @"PCI\VEN_10DE&DEV_2783\SYS_RTX4070_SUPER");
                    AddNvidiaDisplayAudio("SYS_RTX4070S_AUDIO");
                    AddSystemPresetDevice(DeviceKind.NET_NDIS, "Intel Ethernet Controller I225-V", @"PCI\VEN_8086&DEV_15F3\SYS_Z790_I225_V");
                    AddSystemStorage("Standard NVM Express Controller", @"PCI\VEN_144D&DEV_A808\SYS_Z790_NVME");
                    break;
                case SystemPresetIntel285K5090:
                    AddSystemPresetDevice(DeviceKind.USB, "Intel(R) USB 3.20 eXtensible Host Controller - 1.20", @"PCI\VEN_8086&DEV_7F6E\SYS_INTEL_Z890_XHCI", "Mouse 8K, Keyboard 8K, Audio, Microphone");
                    AddSystemAudio("Realtek ALC4080 USB Audio", @"USB\VID_0BDA&PID_402E\SYS_Z890_ALC4080", "Speakers, Microphone");
                    AddSystemPresetDevice(DeviceKind.GPU, "NVIDIA GeForce RTX 5090", @"PCI\VEN_10DE&DEV_2B85\SYS_RTX5090");
                    AddNvidiaDisplayAudio("SYS_RTX5090_AUDIO");
                    AddSystemPresetDevice(DeviceKind.NET_NDIS, "Intel Ethernet Controller I226-V", @"PCI\VEN_8086&DEV_125C\SYS_Z890_I226_V");
                    AddSystemPresetDevice(DeviceKind.NET_NDIS, "Intel(R) Wi-Fi 7 BE200 320MHz", @"PCI\VEN_8086&DEV_272B\SYS_BE200_WIFI", wifi: true);
                    AddSystemStorage("Crucial T705 PCIe 5.0 NVMe Controller", @"PCI\VEN_C0A9&DEV_540A\SYS_Z890_T705");
                    break;
                case SystemPresetRyzen7800X3D:
                    AddSystemPresetDevice(DeviceKind.USB, "AMD USB 3.10 eXtensible Host Controller - 1.20", @"PCI\VEN_1022&DEV_15B6\SYS_AMD_B650E_CPU_XHCI_MOUSE_8K", "Mouse 8K");
                    AddSystemPresetDevice(DeviceKind.USB, "AMD USB 2.0 eXtensible Host Controller - 1.20", @"PCI\VEN_1022&DEV_43F7\SYS_AMD_B650E_CHIPSET_XHCI_KEYBOARD_AUDIO", "Keyboard 8K, Audio, Microphone");
                    AddSystemAudio("Realtek ALC1220 High Definition Audio", @"HDAUDIO\FUNC_01&VEN_10EC&DEV_1220\SYS_B650E_ALC1220", "Speakers, Microphone");
                    AddSystemPresetDevice(DeviceKind.GPU, "NVIDIA GeForce RTX 4080 SUPER", @"PCI\VEN_10DE&DEV_2702\SYS_RTX4080_SUPER");
                    AddNvidiaDisplayAudio("SYS_RTX4080S_AUDIO");
                    AddSystemPresetDevice(DeviceKind.NET_NDIS, "Intel Ethernet Controller I225-V", @"PCI\VEN_8086&DEV_15F3\SYS_B650E_I225_V");
                    AddSystemStorage("Samsung 990 PRO NVMe Controller", @"PCI\VEN_144D&DEV_A80C\SYS_B650E_990PRO");
                    break;
                case SystemPresetRyzen9800X3D5090:
                    AddSystemPresetDevice(DeviceKind.USB, "AMD USB 3.20 eXtensible Host Controller - 1.10", @"PCI\VEN_1022&DEV_15B6\SYS_X870E_CPU_USB4_MOUSE_8K", "Mouse 8K");
                    AddSystemPresetDevice(DeviceKind.USB, "AMD USB 3.10 eXtensible Host Controller - 1.20", @"PCI\VEN_1022&DEV_43FC\SYS_X870E_CHIPSET_XHCI_KEYBOARD_8K", "Keyboard 8K");
                    AddSystemPresetDevice(DeviceKind.USB, "AMD USB 2.0 eXtensible Host Controller - 1.20", @"PCI\VEN_1022&DEV_43FC\SYS_X870E_CHIPSET_XHCI_AUDIO", "Audio, Microphone");
                    AddSystemAudio("Realtek ALC4080 USB Audio", @"USB\VID_0BDA&PID_402E\SYS_X870E_ALC4080", "Speakers, Microphone");
                    AddSystemPresetDevice(DeviceKind.GPU, "NVIDIA GeForce RTX 5090", @"PCI\VEN_10DE&DEV_2B85\SYS_RTX5090");
                    AddNvidiaDisplayAudio("SYS_X870E_RTX5090_AUDIO");
                    AddSystemPresetDevice(DeviceKind.NET_CX, "Realtek RTL8126 5GbE Controller", @"PCI\VEN_10EC&DEV_8126\SYS_REALTEK_RTL8126");
                    AddSystemStorage("Crucial T705 PCIe 5.0 NVMe Controller", @"PCI\VEN_C0A9&DEV_540A\SYS_X870E_T705");
                    break;
                case SystemPresetRyzen9850X3DClient:
                    AddSystemPresetDevice(DeviceKind.USB, "AMD USB 3.10 eXtensible Host Controller - 1.20", @"PCI\VEN_1022&DEV_15B6\SYS_B850_CPU_XHCI_AUDIO", "Audio");
                    AddSystemPresetDevice(DeviceKind.USB, "AMD USB 3.10 eXtensible Host Controller - 1.20", @"PCI\VEN_1022&DEV_15B7\SYS_B850_CPU_XHCI_MOUSE_8K", "Mouse 8K");
                    AddSystemPresetDevice(DeviceKind.USB, "AMD USB 3.20 eXtensible Host Controller - 1.10", @"PCI\VEN_1022&DEV_43FC\SYS_B850_CHIPSET_XHCI_KEYBOARD_1K", "Audio, Keyboard 1K");
                    AddSystemPresetDevice(DeviceKind.GPU, "NVIDIA GeForce RTX 5070", @"PCI\VEN_10DE&DEV_2F04\SYS_RTX5070");
                    AddNvidiaDisplayAudio("SYS_B850_RTX5070_AUDIO");
                    AddSystemPresetDevice(DeviceKind.NET_NDIS, "Killer 2.5 Gigabit Ethernet Controller", @"PCI\VEN_10EC&DEV_3000\SYS_B850_KILLER_25G");
                    AddSystemStorage("Standard NVM Express Controller", @"PCI\VEN_15B7&DEV_5030\SYS_B850_WD_SN850X");
                    break;
                case SystemPresetRyzen9800X3DRx:
                    AddSystemPresetDevice(DeviceKind.USB, "AMD USB 3.20 eXtensible Host Controller - 1.10", @"PCI\VEN_1022&DEV_15B6\SYS_AMD_X870E_CPU_USB4_MOUSE_8K", "Mouse 8K");
                    AddSystemPresetDevice(DeviceKind.USB, "AMD USB 3.10 eXtensible Host Controller - 1.20", @"PCI\VEN_1022&DEV_43FC\SYS_AMD_X870E_CHIPSET_XHCI_KEYBOARD_8K", "Keyboard 8K");
                    AddSystemAudio("Realtek ALC4080 USB Audio", @"USB\VID_0BDA&PID_402E\SYS_X870E_RX_ALC4080", "Speakers, Microphone");
                    AddSystemPresetDevice(DeviceKind.GPU, "AMD Radeon RX 9070 XT", @"PCI\VEN_1002&DEV_7550\SYS_RX9070_XT");
                    AddAmdDisplayAudio("SYS_RX9070XT_AUDIO");
                    AddSystemPresetDevice(DeviceKind.NET_NDIS, "Intel Ethernet Controller I225-V", @"PCI\VEN_8086&DEV_15F3\SYS_I225_V");
                    AddSystemStorage("Samsung 990 PRO NVMe Controller", @"PCI\VEN_144D&DEV_A80C\SYS_X870E_RX_990PRO");
                    break;
                case SystemPresetRyzen9950X3DNetCx:
                    AddSystemPresetDevice(DeviceKind.USB, "AMD USB 3.20 eXtensible Host Controller - 1.10", @"PCI\VEN_1022&DEV_15B6\SYS_AMD_X870E_CPU_USB4_MOUSE_1K", "Mouse 1K");
                    AddSystemPresetDevice(DeviceKind.USB, "AMD USB 3.10 eXtensible Host Controller - 1.20", @"PCI\VEN_1022&DEV_43FC\SYS_AMD_X870E_CHIPSET_XHCI_KEYBOARD_8K", "Keyboard 8K");
                    AddSystemPresetDevice(DeviceKind.USB, "AMD USB 2.0 eXtensible Host Controller - 1.20", @"PCI\VEN_1022&DEV_43FC\SYS_AMD_X870E_CHIPSET_XHCI_AUDIO", "Audio, Microphone");
                    AddSystemAudio("Realtek ALC4080 USB Audio", @"USB\VID_0BDA&PID_402E\SYS_9950X3D_ALC4080", "Speakers, Microphone");
                    AddSystemPresetDevice(DeviceKind.GPU, "NVIDIA GeForce RTX 5090", @"PCI\VEN_10DE&DEV_2B85\SYS_RTX5090");
                    AddNvidiaDisplayAudio("SYS_9950X3D_RTX5090_AUDIO");
                    AddSystemPresetDevice(DeviceKind.NET_CX, "Realtek RTL8126 5GbE Controller", @"PCI\VEN_10EC&DEV_8126\SYS_REALTEK_RTL8126");
                    AddSystemStorage("Crucial T705 PCIe 5.0 NVMe Controller", @"PCI\VEN_C0A9&DEV_540A\SYS_9950X3D_T705");
                    break;
                case SystemPresetRyzen9950X3DNdis:
                    AddSystemPresetDevice(DeviceKind.USB, "AMD USB 3.20 eXtensible Host Controller - 1.10", @"PCI\VEN_1022&DEV_15B6\SYS_AMD_X870E_CPU_USB4_MOUSE_1K_NDIS", "Mouse 1K");
                    AddSystemPresetDevice(DeviceKind.USB, "AMD USB 3.10 eXtensible Host Controller - 1.20", @"PCI\VEN_1022&DEV_43FC\SYS_AMD_X870E_CHIPSET_XHCI_KEYBOARD_8K_NDIS", "Keyboard 8K");
                    AddSystemPresetDevice(DeviceKind.USB, "AMD USB 2.0 eXtensible Host Controller - 1.20", @"PCI\VEN_1022&DEV_43FC\SYS_AMD_X870E_CHIPSET_XHCI_AUDIO_NDIS", "Audio, Microphone");
                    AddSystemAudio("Realtek ALC4080 USB Audio", @"USB\VID_0BDA&PID_402E\SYS_9950X3D_NDIS_ALC4080", "Speakers, Microphone");
                    AddSystemPresetDevice(DeviceKind.GPU, "NVIDIA GeForce RTX 5090", @"PCI\VEN_10DE&DEV_2B85\SYS_RTX5090_NDIS");
                    AddNvidiaDisplayAudio("SYS_9950X3D_NDIS_AUDIO");
                    AddSystemPresetDevice(DeviceKind.NET_NDIS, "Intel Ethernet Controller I226-V", @"PCI\VEN_8086&DEV_125C\SYS_X870E_I226_V");
                    AddSystemStorage("Samsung 990 PRO NVMe Controller", @"PCI\VEN_144D&DEV_A80C\SYS_9950X3D_990PRO");
                    break;
                case SystemPresetRyzen9950X3D2:
                    AddSystemPresetDevice(DeviceKind.USB, "AMD USB 3.20 eXtensible Host Controller - 1.10", @"PCI\VEN_1022&DEV_15B6\SYS_9950X3D2_CPU_USB4_MOUSE_1K", "Mouse 1K");
                    AddSystemPresetDevice(DeviceKind.USB, "AMD USB 3.10 eXtensible Host Controller - 1.20", @"PCI\VEN_1022&DEV_43FC\SYS_9950X3D2_CHIPSET_XHCI_KEYBOARD_8K", "Keyboard 8K");
                    AddSystemPresetDevice(DeviceKind.USB, "AMD USB 2.0 eXtensible Host Controller - 1.20", @"PCI\VEN_1022&DEV_43FC\SYS_9950X3D2_CHIPSET_XHCI_AUDIO", "Audio, Microphone");
                    AddSystemAudio("Realtek ALC4080 USB Audio", @"USB\VID_0BDA&PID_402E\SYS_9950X3D2_ALC4080", "Speakers, Microphone");
                    AddSystemPresetDevice(DeviceKind.GPU, "NVIDIA GeForce RTX 5090", @"PCI\VEN_10DE&DEV_2B85\SYS_9950X3D2_RTX5090");
                    AddNvidiaDisplayAudio("SYS_9950X3D2_RTX5090_AUDIO");
                    AddSystemPresetDevice(DeviceKind.NET_CX, "Realtek RTL8126 5GbE Controller", @"PCI\VEN_10EC&DEV_8126\SYS_9950X3D2_RTL8126");
                    AddSystemStorage("Crucial T705 PCIe 5.0 NVMe Controller", @"PCI\VEN_C0A9&DEV_540A\SYS_9950X3D2_T705");
                    break;
                case SystemPresetRyzen3500X:
                    // Field log 2026-08-15: Gigabyte AM4, SMT off, VXE 1K mouse on CPU xHCI.
                    AddSystemPresetDevice(
                        DeviceKind.USB,
                        "AMD USB 3.10 eXtensible Host Controller - 1.10",
                        @"PCI\VEN_1022&DEV_149C\SYS_3500X_CPU_XHCI_MOUSE_1K",
                        "Mouse 1K",
                        testIrqCount: 7,
                        testMsiStatus: "Enabled");
                    AddSystemPresetDevice(
                        DeviceKind.AUDIO,
                        "High Definition Audio Controller",
                        @"PCI\VEN_1022&DEV_1487\SYS_3500X_HDA_SPEAKERS",
                        audioEndpoints: "Speakers",
                        testIrqCount: 1,
                        testMsiStatus: "Disabled");
                    AddSystemPresetDevice(
                        DeviceKind.GPU,
                        "NVIDIA GeForce GTX 1660 SUPER",
                        @"PCI\VEN_10DE&DEV_21C4\SYS_3500X_GTX1660S",
                        testIrqCount: 1,
                        testMsiStatus: "Enabled");
                    AddSystemPresetDevice(
                        DeviceKind.AUDIO,
                        "High Definition Audio Controller",
                        @"PCI\VEN_10DE&DEV_1AEB\SYS_3500X_HDMI_MUCAI",
                        audioEndpoints: "Monitor - MUCAI",
                        testIrqCount: 1,
                        testMsiStatus: "Disabled");
                    AddSystemPresetDevice(
                        DeviceKind.NET_CX,
                        "Realtek PCIe GbE Family Controller",
                        @"PCI\VEN_10EC&DEV_8168\SYS_3500X_RTL8168",
                        testIrqCount: 1,
                        testMsiStatus: "Enabled",
                        nicPowerSaving: "on");
                    AddSystemPresetDevice(
                        DeviceKind.STOR,
                        "Standard NVM Express Controller",
                        @"PCI\VEN_126F&DEV_2263\SYS_3500X_NVME",
                        storageTag: "SSD",
                        testIrqCount: 7,
                        testMsiStatus: "Enabled");
                    AddSystemPresetDevice(
                        DeviceKind.STOR,
                        "Standard SATA AHCI Controller",
                        @"PCI\VEN_1022&DEV_43EB\SYS_3500X_AHCI",
                        storageTag: string.Empty,
                        testIrqCount: 1,
                        testMsiStatus: "Enabled");
                    break;
                case SystemPresetRyzen3900X:
                    AddSystemPresetDevice(DeviceKind.USB, "AMD USB 3.10 eXtensible Host Controller - 1.10", @"PCI\VEN_1022&DEV_149C\SYS_AMD_X570_CPU_XHCI_MOUSE_1K", "Mouse 1K");
                    AddSystemPresetDevice(DeviceKind.USB, "AMD USB 3.10 eXtensible Host Controller - 1.20", @"PCI\VEN_1022&DEV_43D5\SYS_AMD_X570_CHIPSET_XHCI_KEYBOARD_AUDIO", "Keyboard 8K, Audio, Microphone");
                    AddSystemAudio("Realtek ALC1220 High Definition Audio", @"HDAUDIO\FUNC_01&VEN_10EC&DEV_1220\SYS_X570_ALC1220", "Speakers, Microphone");
                    AddSystemPresetDevice(DeviceKind.GPU, "NVIDIA GeForce RTX 3080", @"PCI\VEN_10DE&DEV_2206\SYS_RTX3080");
                    AddNvidiaDisplayAudio("SYS_RTX3080_AUDIO");
                    AddSystemPresetDevice(DeviceKind.NET_NDIS, "Realtek RTL8125BG 2.5GbE Controller", @"PCI\VEN_10EC&DEV_8125\SYS_X570_RTL8125_NDIS");
                    AddSystemStorage("Standard NVM Express Controller", @"PCI\VEN_144D&DEV_A808\SYS_X570_NVME");
                    break;
                case SystemPresetRyzen9950X:
                    AddSystemPresetDevice(DeviceKind.USB, "AMD USB 3.20 eXtensible Host Controller - 1.10", @"PCI\VEN_1022&DEV_15B6\SYS_AMD_X670E_CPU_XHCI_MOUSE_1K", "Mouse 1K");
                    AddSystemPresetDevice(DeviceKind.USB, "AMD USB 3.10 eXtensible Host Controller - 1.20", @"PCI\VEN_1022&DEV_43F7\SYS_AMD_X670E_CHIPSET_XHCI_KEYBOARD_8K", "Keyboard 8K");
                    AddSystemPresetDevice(DeviceKind.USB, "AMD USB 2.0 eXtensible Host Controller - 1.20", @"PCI\VEN_1022&DEV_43F7\SYS_AMD_X670E_CHIPSET_XHCI_AUDIO_GAMEPAD", "Audio, Microphone, Gamepad");
                    AddSystemAudio("Realtek ALC4080 USB Audio", @"USB\VID_0BDA&PID_402E\SYS_X670E_ALC4080", "Speakers, Microphone");
                    AddSystemPresetDevice(DeviceKind.GPU, "NVIDIA GeForce RTX 4090", @"PCI\VEN_10DE&DEV_2684\SYS_WORKSTATION_RTX4090");
                    AddNvidiaDisplayAudio("SYS_WORKSTATION_RTX4090_AUDIO");
                    AddSystemPresetDevice(DeviceKind.NET_NDIS, "Intel Ethernet Controller X550-T2", @"PCI\VEN_8086&DEV_1563\SYS_X550_T2");
                    AddSystemStorage("Crucial T705 PCIe 5.0 NVMe Controller", @"PCI\VEN_C0A9&DEV_540A\SYS_X670E_T705");
                    break;
                case SystemPresetLaptopIntel285HX:
                    AddSystemPresetDevice(DeviceKind.USB, "Intel(R) USB 3.20 eXtensible Host Controller - 1.20", @"PCI\VEN_8086&DEV_7EC0\SYS_INTEL_285HX_XHCI_INPUT", "Mouse 1K, Keyboard 8K, Gamepad");
                    AddSystemPresetDevice(DeviceKind.USB, "Intel(R) USB 3.10 eXtensible Host Controller - 1.20", @"PCI\VEN_8086&DEV_7E7D\SYS_INTEL_285HX_XHCI_AUDIO", "Audio, Microphone");
                    AddSystemAudio("Realtek ALC287 High Definition Audio", @"HDAUDIO\FUNC_01&VEN_10EC&DEV_0287\SYS_285HX_ALC287", "Speakers, Microphone");
                    AddSystemPresetDevice(DeviceKind.GPU, "NVIDIA GeForce RTX 5080 Laptop GPU", @"PCI\VEN_10DE&DEV_2C02\SYS_RTX5080_LAPTOP");
                    AddNvidiaDisplayAudio("SYS_RTX5080_LAPTOP_AUDIO");
                    AddIntelDisplayAudio("SYS_285HX_DISPLAY_AUDIO");
                    AddSystemPresetDevice(DeviceKind.NET_NDIS, "Intel(R) Wi-Fi 7 BE200 320MHz", @"PCI\VEN_8086&DEV_272B\SYS_285HX_BE200_WIFI", wifi: true);
                    AddSystemStorage("Samsung PM9E1 NVMe Controller", @"PCI\VEN_144D&DEV_A80D\SYS_285HX_PM9E1");
                    break;
                case SystemPresetLaptopRyzen9955HX3D:
                    AddSystemPresetDevice(DeviceKind.USB, "AMD USB 3.20 eXtensible Host Controller - 1.10", @"PCI\VEN_1022&DEV_1587\SYS_9955HX3D_XHCI_INPUT", "Mouse 1K, Keyboard 8K, Gamepad");
                    AddSystemPresetDevice(DeviceKind.USB, "AMD USB 3.10 eXtensible Host Controller - 1.20", @"PCI\VEN_1022&DEV_1588\SYS_9955HX3D_XHCI_AUDIO", "Audio, Microphone");
                    AddSystemAudio("Realtek ALC3306 High Definition Audio", @"HDAUDIO\FUNC_01&VEN_10EC&DEV_0330\SYS_9955HX3D_ALC3306", "Speakers, Microphone");
                    AddSystemPresetDevice(DeviceKind.GPU, "NVIDIA GeForce RTX 5090 Laptop GPU", @"PCI\VEN_10DE&DEV_2B85\SYS_RTX5090_LAPTOP");
                    AddNvidiaDisplayAudio("SYS_RTX5090_LAPTOP_AUDIO");
                    AddSystemPresetDevice(DeviceKind.NET_NDIS, "Qualcomm FastConnect 7800 Wi-Fi 7 Adapter", @"PCI\VEN_17CB&DEV_1107\SYS_9955HX3D_WIFI7", wifi: true);
                    AddSystemStorage("WD Black SN8100 NVMe Controller", @"PCI\VEN_15B7&DEV_5041\SYS_9955HX3D_SN8100");
                    break;
                case SystemPresetLaptopRyzen8940HX:
                    // Field log 2026-09-10: Dual-CCD 16C/32T, Radeon 610M iGPU + RTX 5060 dGPU.
                    AddSystemPresetDevice(
                        DeviceKind.USB,
                        "AMD USB 2.0 eXtensible Host Controller - 1.20",
                        @"PCI\VEN_1022&DEV_15B8&SUBSYS_15B61022&REV_00\4&2C288D56&0&0043",
                        "Webcam",
                        testIrqCount: 8,
                        testMsiStatus: "Enabled");
                    AddSystemPresetDevice(
                        DeviceKind.USB,
                        "AMD USB 3.10 eXtensible Host Controller - 1.20",
                        @"PCI\VEN_1022&DEV_15B7&SUBSYS_201F1043&REV_00\4&16012499&0&0441",
                        "Keyboard 100Hz",
                        testIrqCount: 8,
                        testMsiStatus: "Enabled");
                    AddSystemAudio(
                        "High Definition Audio Controller",
                        @"PCI\VEN_1022&DEV_15E3&SUBSYS_11C41043&REV_00\4&16012499&0&0641",
                        "Microphone");
                    AddSystemPresetDevice(
                        DeviceKind.GPU,
                        "AMD Radeon(TM) 610M",
                        @"PCI\VEN_1002&DEV_164E&SUBSYS_11C41043&REV_D8\4&16012499&0&0041",
                        integratedGpu: true,
                        testIrqCount: 1,
                        testMsiStatus: "Enabled");
                    AddSystemPresetDevice(
                        DeviceKind.GPU,
                        "NVIDIA GeForce RTX 5060 Laptop GPU",
                        @"PCI\VEN_10DE&DEV_2D59&SUBSYS_11C41043&REV_A1\612BE8CAB32DB04800",
                        testIrqCount: 1,
                        testMsiStatus: "Enabled");
                    AddNvidiaDisplayAudio("SYS_8940HX_RTX5060_LAPTOP_AUDIO");
                    AddSystemPresetDevice(
                        DeviceKind.NET_CX,
                        "Realtek PCIe GbE Family Controller",
                        @"PCI\VEN_10EC&DEV_8168&SUBSYS_208F1043&REV_15\01000000684CE00000",
                        testIrqCount: 1,
                        testMsiStatus: "Enabled");
                    AddSystemStorage(
                        "Standard NVM Express Controller",
                        @"PCI\VEN_15B7&DEV_5036&SUBSYS_503615B7&REV_00\4&218BD16B&0&000A");
                    break;
                case SystemPresetLaptopRyzen8940HXMouse:
                    AddSystemPresetDevice(
                        DeviceKind.USB,
                        "AMD USB 2.0 eXtensible Host Controller - 1.20",
                        @"PCI\VEN_1022&DEV_15B8&SUBSYS_15B61022&REV_00\4&2C288D56&0&0043",
                        "Webcam",
                        testIrqCount: 8,
                        testMsiStatus: "Enabled");
                    AddSystemPresetDevice(
                        DeviceKind.USB,
                        "AMD USB 3.10 eXtensible Host Controller - 1.20",
                        @"PCI\VEN_1022&DEV_15B7&SUBSYS_201F1043&REV_00\4&16012499&0&0441",
                        "Mouse 1K, Keyboard 100Hz",
                        testIrqCount: 8,
                        testMsiStatus: "Enabled");
                    AddSystemAudio(
                        "High Definition Audio Controller",
                        @"PCI\VEN_1022&DEV_15E3&SUBSYS_11C41043&REV_00\4&16012499&0&0641",
                        "Microphone");
                    AddSystemPresetDevice(
                        DeviceKind.GPU,
                        "AMD Radeon(TM) 610M",
                        @"PCI\VEN_1002&DEV_164E&SUBSYS_11C41043&REV_D8\4&16012499&0&0041",
                        integratedGpu: true,
                        testIrqCount: 1,
                        testMsiStatus: "Enabled");
                    AddSystemPresetDevice(
                        DeviceKind.GPU,
                        "NVIDIA GeForce RTX 5060 Laptop GPU",
                        @"PCI\VEN_10DE&DEV_2D59&SUBSYS_11C41043&REV_A1\612BE8CAB32DB04800",
                        testIrqCount: 1,
                        testMsiStatus: "Enabled");
                    AddNvidiaDisplayAudio("SYS_8940HX_RTX5060_LAPTOP_AUDIO");
                    AddSystemPresetDevice(
                        DeviceKind.NET_CX,
                        "Realtek PCIe GbE Family Controller",
                        @"PCI\VEN_10EC&DEV_8168&SUBSYS_208F1043&REV_15\01000000684CE00000",
                        testIrqCount: 1,
                        testMsiStatus: "Enabled");
                    AddSystemStorage(
                        "Standard NVM Express Controller",
                        @"PCI\VEN_15B7&DEV_5036&SUBSYS_503615B7&REV_00\4&218BD16B&0&000A");
                    break;
                case SystemPresetMultiControllerInput:
                    AddSystemPresetDevice(
                        DeviceKind.USB,
                        "Intel(R) USB 3.20 eXtensible Host Controller - 1.20 (CPU)",
                        @"PCI\VEN_8086&DEV_7F6E\SYS_MULTI_XHCI_MOUSE",
                        "Mouse 8K",
                        testIrqCount: 8,
                        testMsiStatus: "Enabled");
                    AddSystemPresetDevice(
                        DeviceKind.USB,
                        "Intel(R) USB 3.10 eXtensible Host Controller - 1.20 (PCH)",
                        @"PCI\VEN_8086&DEV_7AE0\SYS_MULTI_XHCI_GAMEPAD",
                        "Gamepad",
                        testIrqCount: 8,
                        testMsiStatus: "Enabled");
                    AddSystemPresetDevice(
                        DeviceKind.USB,
                        "ASMedia USB3.1 eXtensible Host Controller (PCIe)",
                        @"PCI\VEN_1B21&DEV_2142\SYS_MULTI_XHCI_KEYBOARD",
                        "Keyboard 8K",
                        testIrqCount: 8,
                        testMsiStatus: "Enabled");
                    AddSystemPresetDevice(
                        DeviceKind.GPU,
                        "NVIDIA GeForce RTX 5090",
                        @"PCI\VEN_10DE&DEV_2B85\SYS_RTX5090",
                        testIrqCount: 1,
                        testMsiStatus: "Enabled");
                    AddNvidiaDisplayAudio("SYS_RTX5090_AUDIO");
                    AddSystemPresetDevice(
                        DeviceKind.NET_NDIS,
                        "Intel Ethernet Controller I226-V",
                        @"PCI\VEN_8086&DEV_125C\SYS_Z890_I226_V",
                        testIrqCount: 1,
                        testMsiStatus: "Enabled");
                    AddSystemStorage(
                        "Crucial T705 PCIe 5.0 NVMe Controller",
                        @"PCI\VEN_C0A9&DEV_540A\SYS_Z890_T705");
                    break;
            }

            RefreshTestDeviceList();
            _initialDeviceViewportHeightAdjusted = false;
            RefreshBlocks();
            WriteLog($"TEST.SYSTEM.PRESET: loaded name=\"{preset}\" cpu=\"{cpuPreset}\" devices={_testDevices.Count} replacement=1");
            bool runAutoPolicyRegression = string.Equals(preset, SystemPresetRyzen3500X, StringComparison.Ordinal)
                || string.Equals(preset, SystemPresetLaptopIntel285HX, StringComparison.Ordinal);
            Dictionary<string, (string Msi, string Limit, string Priority, string Policy, ulong Mask)> preservedExpected =
                runAutoPolicyRegression
                    ? SeedAutoPolicyRegressionState()
                    : new Dictionary<string, (string Msi, string Limit, string Priority, string Policy, ulong Mask)>(StringComparer.OrdinalIgnoreCase);
            bool hasUsbImodTarget = _blocks.Any(block => IsUsbImodTarget(block.Device));
            InvokeAutoOptimization(optimizeUsbImod: hasUsbImodTarget, hasUsbImodTarget: hasUsbImodTarget);
            if (runAutoPolicyRegression)
            {
                ValidateAutoPolicyRegressionState(preservedExpected);
            }
            WriteLog($"TEST.SYSTEM.PRESET.AUTO: dry-run preview complete blocks={_blocks.Count} optimizeUsbImod=1");
            LogGuiSnapshot("system-preset-auto");
            return true;
        }

        bool ValidateQaAffinity8940HxLayout()
        {
            List<string> issues = [];
            int ccdCount = _cpuInfo?.CcdMap.Values.Distinct().Count() ?? 0;
            if (_maxLogical < 32 || ccdCount < 2)
            {
                issues.Add($"topology maxLogical={_maxLogical} ccd={ccdCount}");
            }

            int usbIndex = 0;
            foreach (DeviceBlock block in _blocks.Where(static b => b.Kind == DeviceKind.USB))
            {
                if (block.CpuPanel is ScrollableControl scrollable
                    && scrollable.AutoScroll
                    && scrollable.HorizontalScroll.Visible)
                {
                    issues.Add($"usb[{usbIndex}] cpuHorizontalScroll=1");
                }

                if (block.CpuPanel is not null && block.SettingsPanel is not null)
                {
                    Rectangle cpu = block.CpuPanel.Bounds;
                    Rectangle settings = block.SettingsPanel.Bounds;
                    if (settings.IntersectsWith(cpu))
                    {
                        bool stacked = settings.Top >= cpu.Bottom - UiScale(4);
                        if (!stacked)
                        {
                            issues.Add($"usb[{usbIndex}] settingsOverlap=1");
                        }
                    }
                }

                if (ShouldSkipBuiltInUsbAffinity(block.Device.UsbRoles) && block.AffinityMask != 0)
                {
                    issues.Add($"usb[{usbIndex}] builtInAffinity=0x{block.AffinityMask:X}");
                }

                usbIndex++;
            }

            DeviceBlock? igpu = _blocks.FirstOrDefault(b => b.Kind == DeviceKind.GPU && b.Device.IsIntegratedGpu);
            DeviceBlock? dgpu = _blocks.FirstOrDefault(b => b.Kind == DeviceKind.GPU && !b.Device.IsIntegratedGpu);
            if (igpu is null)
            {
                issues.Add("igpu-missing");
            }
            else if (igpu.AffinityMask != 0)
            {
                issues.Add($"igpuAffinity=0x{igpu.AffinityMask:X}");
            }

            if (dgpu is null)
            {
                issues.Add("dgpu-missing");
            }
            else if (dgpu.AffinityMask == 0)
            {
                issues.Add("dgpuAffinity=0");
            }
            else if (dgpu.AffinityMask != 0x5000UL)
            {
                // Field-proven irq-safe tail pair on 8940HX CCD0 is [12,14] / 0x5000.
                issues.Add($"dgpuAffinity=0x{dgpu.AffinityMask:X} expected=0x5000");
            }

            int usbBlocks = _blocks.Count(static b => b.Kind == DeviceKind.USB);
            bool pass = issues.Count == 0;
            WriteLog(
                pass
                    ? $"TEST.QA.AFFINITY.8940HX: status=PASS maxLogical={_maxLogical} ccd={ccdCount} usbBlocks={usbBlocks} igpu=1 dgpu=1"
                    : $"TEST.QA.AFFINITY.8940HX: status=FAIL maxLogical={_maxLogical} ccd={ccdCount} usbBlocks={usbBlocks} issues=\"{SanitizeLogValue(string.Join("; ", issues))}\"");
            return pass;
        }

        bool ValidateQaMultiControllerAffinity()
        {
            List<string> issues = [];
            DeviceBlock? mouse = _blocks.FirstOrDefault(static b => b.Kind == DeviceKind.USB && HasRoleText(b.Device.UsbRoles, "Mouse"));
            DeviceBlock? gamepad = _blocks.FirstOrDefault(static b => b.Kind == DeviceKind.USB && HasRoleText(b.Device.UsbRoles, "Gamepad"));
            DeviceBlock? keyboard = _blocks.FirstOrDefault(static b => b.Kind == DeviceKind.USB && HasRoleText(b.Device.UsbRoles, "Keyboard"));

            if (mouse is null || mouse.AffinityMask == 0)
            {
                issues.Add("mouse-affinity-missing");
            }

            if (gamepad is null || gamepad.AffinityMask == 0)
            {
                issues.Add("gamepad-affinity-missing");
            }

            if (keyboard is null || keyboard.AffinityMask == 0)
            {
                issues.Add("keyboard-affinity-missing");
            }

            bool pass = issues.Count == 0;
            WriteLog(
                pass
                    ? $"TEST.QA.MULTI_CONTROLLER: status=PASS mouseMask=0x{mouse?.AffinityMask:X} gamepadMask=0x{gamepad?.AffinityMask:X} keyboardMask=0x{keyboard?.AffinityMask:X}"
                    : $"TEST.QA.MULTI_CONTROLLER: status=FAIL issues=\"{SanitizeLogValue(string.Join("; ", issues))}\"");
            return pass;
        }

        bool LoadTestDevicePresetToFields()
        {
            string preset = testDevicePresetCombo.SelectedItem?.ToString() ?? string.Empty;
            switch (preset)
            {
                case "USB - Intel Z790 PCH xHCI / Mouse 8K":
                    SetTestDeviceFields(DeviceKind.USB, "Intel(R) USB 3.20 eXtensible Host Controller - 1.20", @"PCI\VEN_8086&DEV_7AE0\TEST_Z790_XHCI_MOUSE_8K", "Mouse 8K");
                    return true;
                case "USB - Intel Z790 PCH xHCI / Mouse 1K":
                    SetTestDeviceFields(DeviceKind.USB, "Intel(R) USB 3.20 eXtensible Host Controller - 1.20", @"PCI\VEN_8086&DEV_7AE0\TEST_Z790_XHCI_MOUSE_1K", "Mouse 1K");
                    return true;
                case "USB - Intel Z790 PCH xHCI / Keyboard 8K":
                    SetTestDeviceFields(DeviceKind.USB, "Intel(R) USB 3.20 eXtensible Host Controller - 1.20", @"PCI\VEN_8086&DEV_7AE0\TEST_Z790_XHCI_KEYBOARD_8K", "Keyboard 8K");
                    return true;
                case "USB - Intel Z790 PCH xHCI / Mouse 8K + Keyboard 8K":
                    SetTestDeviceFields(DeviceKind.USB, "Intel(R) USB 3.20 eXtensible Host Controller - 1.20", @"PCI\VEN_8086&DEV_7AE0\TEST_Z790_XHCI_INPUT_8K", "Mouse 8K, Keyboard 8K");
                    return true;
                case "USB - Intel Z790 PCH xHCI / Input + audio":
                    SetTestDeviceFields(DeviceKind.USB, "Intel(R) USB 3.20 eXtensible Host Controller - 1.20", @"PCI\VEN_8086&DEV_7AE0\TEST_Z790_XHCI_INPUT_AUDIO", "Mouse 1K, Keyboard 8K, Audio, Microphone");
                    return true;
                case "USB - Intel Z790 PCH xHCI / Gamepad":
                    SetTestDeviceFields(DeviceKind.USB, "Intel(R) USB 3.20 eXtensible Host Controller - 1.20", @"PCI\VEN_8086&DEV_7AE0\TEST_Z790_XHCI_GAMEPAD", "Gamepad");
                    return true;
                case "USB - Intel Z790 PCH xHCI / Audio + microphone":
                    SetTestDeviceFields(DeviceKind.USB, "Intel(R) USB 3.20 eXtensible Host Controller - 1.20", @"PCI\VEN_8086&DEV_7AE0\TEST_Z790_XHCI_AUDIO", "Audio, Microphone");
                    return true;
                case "USB - Intel Z790 PCH xHCI / Empty controller":
                    SetTestDeviceFields(DeviceKind.USB, "Intel(R) USB 3.20 eXtensible Host Controller - 1.20", @"PCI\VEN_8086&DEV_7AE0\TEST_Z790_XHCI_EMPTY", usbHasDevices: false);
                    return true;
                case "USB - Intel Z890 PCH xHCI / Mouse 8K (CHIP 1)":
                    SetTestDeviceFields(DeviceKind.USB, "Intel(R) USB 3.20 eXtensible Host Controller - 1.20", @"PCI\VEN_8086&DEV_7F6E\TEST_Z890_XHCI_MOUSE_8K", "Mouse 8K");
                    return true;
                case "USB - Intel Meteor Lake CPU TB4 / Mouse 8K (CHIP 0)":
                    SetTestDeviceFields(DeviceKind.USB, "Intel(R) Thunderbolt(TM) USB 4 Host Controller", @"PCI\VEN_8086&DEV_7EC0\TEST_MTL_TB4_MOUSE_8K", "Mouse 8K");
                    return true;
                case "USB - AMD AM5 CPU xHCI / Mouse 8K (CHIP 0)":
                    SetTestDeviceFields(DeviceKind.USB, "AMD USB 3.20 eXtensible Host Controller - 1.10", @"PCI\VEN_1022&DEV_15B6\TEST_AMD_AM5_CPU_XHCI_MOUSE_8K", "Mouse 8K");
                    return true;
                case "USB - AMD AM5 chipset xHCI / Keyboard 8K (CHIP 1)":
                    SetTestDeviceFields(DeviceKind.USB, "AMD USB 3.10 eXtensible Host Controller - 1.20", @"PCI\VEN_1022&DEV_43F7\TEST_AMD_AM5_CHIPSET_XHCI_KEYBOARD_8K", "Keyboard 8K");
                    return true;
                case "USB - AMD X870 chipset xHCI / Keyboard 8K (CHIP 1)":
                    SetTestDeviceFields(DeviceKind.USB, "AMD USB 3.10 eXtensible Host Controller - 1.20", @"PCI\VEN_1022&DEV_43FC\TEST_AMD_X870_CHIPSET_XHCI_KEYBOARD_8K", "Keyboard 8K");
                    return true;
                case "USB - AMD AM5 chipset xHCI / Audio + microphone":
                    SetTestDeviceFields(DeviceKind.USB, "AMD USB 2.0 eXtensible Host Controller - 1.20", @"PCI\VEN_1022&DEV_43F7\TEST_AMD_AM5_CHIPSET_XHCI_AUDIO", "Audio, Microphone");
                    return true;
                case "USB - ASMedia ASM2142 add-in xHCI / Edge case (CHIP 1)":
                    SetTestDeviceFields(DeviceKind.USB, "ASMedia USB3.1 eXtensible Host Controller", @"PCI\VEN_1B21&DEV_2142\TEST_ASMEDIA_XHCI_EDGE", "Gamepad, Audio");
                    return true;
                case "GPU - NVIDIA GeForce RTX 5090":
                    SetTestDeviceFields(DeviceKind.GPU, "NVIDIA GeForce RTX 5090", @"PCI\VEN_10DE&DEV_2B85\TEST_RTX5090");
                    return true;
                case "GPU - NVIDIA GeForce RTX 5080":
                    SetTestDeviceFields(DeviceKind.GPU, "NVIDIA GeForce RTX 5080", @"PCI\VEN_10DE&DEV_2C02\TEST_RTX5080");
                    return true;
                case "GPU - NVIDIA GeForce RTX 4090":
                    SetTestDeviceFields(DeviceKind.GPU, "NVIDIA GeForce RTX 4090", @"PCI\VEN_10DE&DEV_2684\TEST_RTX4090");
                    return true;
                case "GPU - NVIDIA GeForce RTX 4080 SUPER":
                    SetTestDeviceFields(DeviceKind.GPU, "NVIDIA GeForce RTX 4080 SUPER", @"PCI\VEN_10DE&DEV_2702\TEST_RTX4080_SUPER");
                    return true;
                case "GPU - NVIDIA GeForce RTX 4070 SUPER":
                    SetTestDeviceFields(DeviceKind.GPU, "NVIDIA GeForce RTX 4070 SUPER", @"PCI\VEN_10DE&DEV_2783\TEST_RTX4070_SUPER");
                    return true;
                case "GPU - NVIDIA GeForce RTX 3080":
                    SetTestDeviceFields(DeviceKind.GPU, "NVIDIA GeForce RTX 3080", @"PCI\VEN_10DE&DEV_2206\TEST_RTX3080");
                    return true;
                case "GPU - NVIDIA GeForce RTX 5060 Ti":
                    SetTestDeviceFields(DeviceKind.GPU, "NVIDIA GeForce RTX 5060 Ti", @"PCI\VEN_10DE&DEV_2D04\TEST_RTX5060_TI");
                    return true;
                case "GPU - AMD Radeon RX 9070 XT":
                    SetTestDeviceFields(DeviceKind.GPU, "AMD Radeon RX 9070 XT", @"PCI\VEN_1002&DEV_7550\TEST_RX9070_XT");
                    return true;
                case "GPU - AMD Radeon RX 7900 XTX":
                    SetTestDeviceFields(DeviceKind.GPU, "AMD Radeon RX 7900 XTX", @"PCI\VEN_1002&DEV_744C\TEST_RX7900_XTX");
                    return true;
                case "GPU - Intel Arc B580":
                    SetTestDeviceFields(DeviceKind.GPU, "Intel Arc B580 Graphics", @"PCI\VEN_8086&DEV_E20B\TEST_ARC_B580");
                    return true;
                case "GPU - Intel integrated GPU":
                    SetTestDeviceFields(DeviceKind.GPU, "Intel Arc Graphics iGPU", @"PCI\VEN_8086&DEV_7D55\TEST_INTEL_IGPU", integratedGpu: true);
                    return true;
                case "NIC - Realtek RTL8125BG 2.5GbE NetAdapterCx":
                    SetTestDeviceFields(DeviceKind.NET_CX, "Realtek RTL8125BG 2.5GbE Controller", @"PCI\VEN_10EC&DEV_8125\TEST_NETCX");
                    return true;
                case "NIC - Realtek RTL8125BG 2.5GbE NDIS":
                    SetTestDeviceFields(DeviceKind.NET_NDIS, "Realtek RTL8125BG 2.5GbE Controller", @"PCI\VEN_10EC&DEV_8125\TEST_RTL8125_NDIS");
                    return true;
                case "NIC - Realtek RTL8126 5GbE NetAdapterCx":
                    SetTestDeviceFields(DeviceKind.NET_CX, "Realtek RTL8126 5GbE Controller", @"PCI\VEN_10EC&DEV_8126\TEST_REALTEK_5G_NETCX");
                    return true;
                case "NIC - Intel I225-V NDIS":
                    SetTestDeviceFields(DeviceKind.NET_NDIS, "Intel Ethernet Controller I225-V", @"PCI\VEN_8086&DEV_15F3\TEST_I225_NDIS");
                    return true;
                case "NIC - Intel I226-V NDIS":
                    SetTestDeviceFields(DeviceKind.NET_NDIS, "Intel Ethernet Controller I226-V", @"PCI\VEN_8086&DEV_125C\TEST_I226_NDIS");
                    return true;
                case "NIC - Intel X550 10GbE NDIS":
                    SetTestDeviceFields(DeviceKind.NET_NDIS, "Intel Ethernet Controller X550-T2", @"PCI\VEN_8086&DEV_1563\TEST_X550_NDIS");
                    return true;
                case "NIC - Intel AX200 Wi-Fi NDIS":
                    SetTestDeviceFields(DeviceKind.NET_NDIS, "Intel(R) Wi-Fi 6 AX200 160MHz", @"PCI\VEN_8086&DEV_2723\TEST_AX200_WIFI", wifi: true);
                    return true;
                case "NIC - Intel AX210 Wi-Fi NDIS":
                    SetTestDeviceFields(DeviceKind.NET_NDIS, "Intel(R) Wi-Fi 6E AX210 160MHz", @"PCI\VEN_8086&DEV_2725\TEST_AX210_WIFI", wifi: true);
                    return true;
                case "NIC - Intel BE200 Wi-Fi 7 NDIS":
                    SetTestDeviceFields(DeviceKind.NET_NDIS, "Intel(R) Wi-Fi 7 BE200 320MHz", @"PCI\VEN_8086&DEV_272B\TEST_BE200_WIFI", wifi: true);
                    return true;
                case "NIC - MediaTek MT7922 Wi-Fi NDIS":
                    SetTestDeviceFields(DeviceKind.NET_NDIS, "MediaTek Wi-Fi 6E MT7922 Wireless LAN Card", @"PCI\VEN_14C3&DEV_7961\TEST_MEDIATEK_WIFI", wifi: true);
                    return true;
                case "Audio - Realtek HDA":
                    SetTestDeviceFields(DeviceKind.AUDIO, "Realtek ALC897 High Definition Audio", @"HDAUDIO\FUNC_01&VEN_10EC&DEV_0897\TEST_REALTEK_HDA", audioEndpoints: "Speakers, Microphone");
                    return true;
                case "Audio - Realtek ALC4080 USB Audio":
                    SetTestDeviceFields(DeviceKind.AUDIO, "Realtek ALC4080 USB Audio", @"USB\VID_0BDA&PID_402E\TEST_ALC4080_USB_AUDIO", audioEndpoints: "Speakers, Microphone");
                    return true;
                case "Audio - USB DAC":
                    SetTestDeviceFields(DeviceKind.AUDIO, "Focusrite Scarlett USB Audio", @"USB\VID_1235&PID_8211\TEST_USB_DAC_AUDIO", audioEndpoints: "USB DAC, Microphone");
                    return true;
                case "Audio - HDMI/DP monitor":
                    SetTestDeviceFields(DeviceKind.AUDIO, "NVIDIA High Definition Audio", @"HDAUDIO\FUNC_01&VEN_10DE&DEV_00A1\TEST_DISPLAY_AUDIO", audioEndpoints: "Monitor DisplayPort");
                    return true;
                case "Storage - Samsung 990 PRO NVMe":
                    SetTestDeviceFields(DeviceKind.STOR, "Samsung 990 PRO NVMe Controller", @"PCI\VEN_144D&DEV_A80C\TEST_990PRO_NVME", storageTag: "NVMe");
                    return true;
                case "Storage - Crucial T705 PCIe 5.0 NVMe":
                    SetTestDeviceFields(DeviceKind.STOR, "Crucial T705 PCIe 5.0 NVMe Controller", @"PCI\VEN_C0A9&DEV_540A\TEST_T705_NVME", storageTag: "NVMe");
                    return true;
                case "Storage - SATA AHCI SSD":
                    SetTestDeviceFields(DeviceKind.STOR, "Standard SATA AHCI Controller", @"PCI\VEN_8086&DEV_7AE2\TEST_SATA_AHCI", storageTag: "SSD");
                    return true;
                default:
                    return false;
            }
        }

        bool TryCreateTestDeviceFromFields(string fallbackPnpId, out DeviceInfo testDevice)
        {
            testDevice = null!;
            if (testKindCombo.SelectedItem is not DeviceKind kind)
            {
                return false;
            }

            string name = testNameBox.Text?.Trim() ?? string.Empty;
            string pnpIdOverride = testPnpIdBox.Text?.Trim() ?? string.Empty;
            string usbRoles = testUsbRolesBox.Text?.Trim() ?? string.Empty;
            string audioEndpoints = testAudioBox.Text?.Trim() ?? string.Empty;
            string storageTag = testStorageBox.Text?.Trim() ?? string.Empty;
            string irqText = testIrqCountBox.Text?.Trim() ?? string.Empty;
            int? testIrqCount = null;
            if (!string.IsNullOrWhiteSpace(irqText) && !irqText.Equals("Auto", StringComparison.OrdinalIgnoreCase))
            {
                if (!int.TryParse(irqText, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsedIrqCount) || parsedIrqCount < 0 || parsedIrqCount > 64)
                {
                    ShowThemedInfo("IRQ count must be Auto or a number from 0 to 64.");
                    return false;
                }

                testIrqCount = parsedIrqCount;
            }

            string testMsiStatus = testMsiStatusCombo.SelectedItem?.ToString() ?? "Auto";
            string? usbSelectiveSuspend = testSuspendCombo.SelectedItem?.ToString() switch
            {
                "on" => "on",
                "off" => "off",
                _ => null,
            };
            string? nicPowerSaving = testWifiCheck.Checked
                ? null
                : testNicPowerCombo.SelectedItem?.ToString() switch
                {
                    "off" => "off",
                    _ => "on",
                };

            if (kind == DeviceKind.USB && string.IsNullOrWhiteSpace(usbRoles))
            {
                usbRoles = "Microphone";
            }

            if (kind == DeviceKind.AUDIO && string.IsNullOrWhiteSpace(audioEndpoints))
            {
                audioEndpoints = "Speakers";
            }

            if (kind == DeviceKind.STOR && string.IsNullOrWhiteSpace(storageTag))
            {
                storageTag = "SSD";
            }

            if (string.IsNullOrWhiteSpace(pnpIdOverride) && !string.IsNullOrWhiteSpace(fallbackPnpId))
            {
                pnpIdOverride = fallbackPnpId;
            }

            testDevice = CreateTestDevice(
                kind,
                name,
                pnpIdOverride,
                usbRoles,
                audioEndpoints,
                storageTag,
                testWifiCheck.Checked,
                testXhciCheck.Checked,
                testHasDevicesCheck.Checked,
                testIntegratedGpuCheck.Checked,
                testIrqCount,
                testMsiStatus,
                usbSelectiveSuspend,
                nicPowerSaving);
            return true;
        }

        bool AddTestDeviceFromFields()
        {
            if (!TryCreateTestDeviceFromFields(string.Empty, out DeviceInfo testDevice))
            {
                return false;
            }

            _testDevices.Add(testDevice);
            WriteLog($"TEST.DEV.ADD: {testDevice.InstanceId} Kind={testDevice.Kind} Name=\"{testDevice.Name}\"");

            RefreshTestDeviceList(_testDevices.Count - 1);
            bool shouldRefresh = false;
            if (!_testDevicesEnabled)
            {
                suppressTestDeviceToggle = true;
                _testDevicesEnabled = true;
                enableTestDevicesCheck.Checked = true;
                testDevicesOnlyCheck.Enabled = true;
                suppressTestDeviceToggle = false;
                shouldRefresh = true;
            }

            if (_testDevicesEnabled || shouldRefresh)
            {
                _initialDeviceViewportHeightAdjusted = false;
                RefreshBlocks();
            }

            return true;
        }

        bool UpdateSelectedTestDeviceFromFields()
        {
            int index = testDeviceListBox.SelectedIndex;
            if (index < 0 || index >= _testDevices.Count)
            {
                return false;
            }

            DeviceInfo previous = _testDevices[index];
            if (!TryCreateTestDeviceFromFields(previous.InstanceId, out DeviceInfo updated))
            {
                return false;
            }

            updated.TestState = previous.TestState?.Clone() ?? EnsureTestDeviceState(previous).Clone();
            _testDevices[index] = updated;
            WriteLog($"TEST.DEV.UPDATE: old={previous.InstanceId} new={updated.InstanceId} Kind={updated.Kind} Name=\"{updated.Name}\"");
            RefreshTestDeviceList(index);
            if (_testDevicesEnabled)
            {
                _initialDeviceViewportHeightAdjusted = false;
                RefreshBlocks();
            }

            return true;
        }

        enableTestDevicesCheck.CheckedChanged += (_, _) =>
        {
            if (suppressTestDeviceToggle)
            {
                return;
            }

            suppressTestDeviceToggle = true;
            _testDevicesEnabled = true;
            enableTestDevicesCheck.Checked = true;
            testDevicesOnlyCheck.Enabled = true;
            suppressTestDeviceToggle = false;

            WriteLog($"TEST.DEVICES: enabled={_testDevicesEnabled}");
            SyncSandboxDryRunLock("test-devices-enabled");
            _initialDeviceViewportHeightAdjusted = false;
            RefreshBlocks();
        };

        testDevicesOnlyCheck.CheckedChanged += (_, _) =>
        {
            if (suppressTestDeviceToggle)
            {
                return;
            }

            suppressTestDeviceToggle = true;
            if (testDevicesOnlyCheck.Checked)
            {
                if (!_testDevicesEnabled)
                {
                    enableTestDevicesCheck.Checked = true;
                    _testDevicesEnabled = true;
                }
                _testDevicesOnly = true;
            }
            else
            {
                _testDevicesOnly = false;
            }
            suppressTestDeviceToggle = false;

            WriteLog($"TEST.DEVICES: only={_testDevicesOnly}");
            SyncSandboxDryRunLock(_testDevicesOnly ? "test-only-on" : "test-only-off");
            _initialDeviceViewportHeightAdjusted = false;
            RefreshBlocks();
        };

        dryRunAutoCheck.CheckedChanged += (_, _) =>
        {
            if (_testDevicesEnabled || _testDevices.Count > 0)
            {
                // Any simulated device keeps the whole session write-safe.
                if (!dryRunAutoCheck.Checked || !_testAutoDryRun)
                {
                    dryRunAutoCheck.Checked = true;
                    _testAutoDryRun = true;
                }

                dryRunAutoCheck.AutoCheck = false;
                dryRunAutoCheck.Enabled = true;
                WriteLog("TEST.AUTO.DRYRUN: forced enabled while simulated devices are active");
                NotifySandboxModeChanged("dry-run-locked");
                return;
            }

            _testAutoDryRun = dryRunAutoCheck.Checked;
            WriteLog($"TEST.AUTO.DRYRUN: {(_testAutoDryRun ? "enabled" : "disabled")}");
            NotifySandboxModeChanged("dry-run-toggle");
        };

        testKindCombo.SelectedIndexChanged += (_, _) => UpdateTestDeviceFieldState();
        testWifiCheck.CheckedChanged += (_, _) => UpdateTestDeviceFieldState();
        testDevicePresetCombo.SelectedIndexChanged += (_, _) =>
        {
            if (testDevicePresetCombo.SelectedIndex > 0)
            {
                LoadTestDevicePresetToFields();
            }
        };
        UpdateTestDeviceFieldState();
        RefreshTestDeviceList();
        RefreshRealDeviceVisibilityLists();

        loadSystemPresetButton.Click += (_, _) => LoadSystemPreset();
        systemPresetCombo.SelectedIndexChanged += (_, _) =>
        {
            if (systemPresetCombo.SelectedIndex > 0)
            {
                LoadSystemPreset();
            }
        };

        addTestPresetButton.Click += (_, _) =>
        {
            if (LoadTestDevicePresetToFields())
            {
                AddTestDeviceFromFields();
            }
        };

        addTestDeviceButton.Click += (_, _) => AddTestDeviceFromFields();
        updateTestDeviceButton.Click += (_, _) => UpdateSelectedTestDeviceFromFields();
        testDeviceListBox.SelectedIndexChanged += (_, _) =>
        {
            updateTestDeviceButton.Enabled = testDeviceListBox.SelectedIndex >= 0;
            if (suppressTestDeviceSelection)
            {
                return;
            }

            int index = testDeviceListBox.SelectedIndex;
            if (index >= 0 && index < _testDevices.Count)
            {
                LoadTestDeviceToFields(_testDevices[index]);
            }
        };
        testDeviceListBox.DoubleClick += (_, _) => UpdateSelectedTestDeviceFromFields();

        hideRealDeviceButton.Click += (_, _) =>
        {
            int index = realDeviceListBox.SelectedIndex;
            if (index < 0 || index >= realVisibleDevices.Count)
            {
                return;
            }

            DeviceInfo device = realVisibleDevices[index];
            string key = NormalizeInstanceId(device.InstanceId);
            if (string.IsNullOrWhiteSpace(key))
            {
                return;
            }

            string label = BuildDeviceBlockTitle(device);
            _testHiddenDeviceIds.Add(key);
            _testHiddenDeviceLabels[key] = label;
            WriteLog($"TEST.HIDE.ADD: {device.InstanceId} Kind={device.Kind} Name=\"{label}\"");
            RefreshAfterRealDeviceVisibilityChange();
        };

        unhideRealDeviceButton.Click += (_, _) =>
        {
            int index = hiddenDeviceListBox.SelectedIndex;
            if (index < 0 || index >= hiddenDeviceKeys.Count)
            {
                return;
            }

            string key = hiddenDeviceKeys[index];
            string label = _testHiddenDeviceLabels.TryGetValue(key, out string? value) ? value : key;
            _testHiddenDeviceIds.Remove(key);
            _testHiddenDeviceLabels.Remove(key);
            WriteLog($"TEST.HIDE.REMOVE: {key} Name=\"{label}\"");
            RefreshAfterRealDeviceVisibilityChange();
        };

        clearHiddenDeviceButton.Click += (_, _) =>
        {
            if (_testHiddenDeviceIds.Count == 0)
            {
                return;
            }

            int count = _testHiddenDeviceIds.Count;
            _testHiddenDeviceIds.Clear();
            _testHiddenDeviceLabels.Clear();
            WriteLog($"TEST.HIDE.CLEAR: count={count}");
            RefreshAfterRealDeviceVisibilityChange();
        };

        removeTestDeviceButton.Click += (_, _) =>
        {
            int index = testDeviceListBox.SelectedIndex;
            if (index < 0 || index >= _testDevices.Count)
            {
                return;
            }

            DeviceInfo removed = _testDevices[index];
            _testDevices.RemoveAt(index);
            WriteLog($"TEST.DEV.REMOVE: {removed.InstanceId} Kind={removed.Kind} Name=\"{removed.Name}\"");

            RefreshTestDeviceList(Math.Min(index, _testDevices.Count - 1));
            if (_testDevicesEnabled)
            {
                _initialDeviceViewportHeightAdjusted = false;
                RefreshBlocks();
            }
        };

        clearTestDeviceButton.Click += (_, _) =>
        {
            if (_testDevices.Count == 0)
            {
                return;
            }

            _testDevices.Clear();
            WriteLog("TEST.DEV.CLEAR: all test devices removed");

            RefreshTestDeviceList();
            if (_testDevicesEnabled)
            {
                _initialDeviceViewportHeightAdjusted = false;
                RefreshBlocks();
            }
        };

        void SyncSmtStateFromCurrent()
        {
            bool enabled = ResolveSmtEnabled(_smtText, GetSmtEnabledFallback());
            bool useHyperLabel = ResolveUseHyperThreadingLabel(_smtText);
            SetSmtState(enabled, useHyperLabel, false);
        }

        Label NewHeaderLabel(string text)
        {
            return new Label
            {
                Text = text,
                AutoSize = true,
                ForeColor = _mutedText,
                Margin = new Padding(0, 0, 12, 4),
            };
        }

        int[] BuildAssignmentsFromGroupsText(string text, int logicalCount, out int groupCount)
        {
            groupCount = 1;
            int[] assign = new int[logicalCount];
            if (TryParseGroups(text, logicalCount, out List<List<int>> groups, out _))
            {
                if (groups.Count > 0)
                {
                    groupCount = Math.Max(1, groups.Count);
                    for (int g = 0; g < groups.Count; g++)
                    {
                        foreach (int lp in groups[g])
                        {
                            if (lp >= 0 && lp < logicalCount)
                            {
                                assign[lp] = g;
                            }
                        }
                    }
                }
            }

            return assign;
        }

        bool[] BuildECoreFlags(string text, int logicalCount)
        {
            bool[] flags = new bool[logicalCount];
            HashSet<int> set = ParseIndexSet(text);
            foreach (int lp in set)
            {
                if (lp >= 0 && lp < logicalCount)
                {
                    flags[lp] = true;
                }
            }

            return flags;
        }

        int[] ResizeAssignments(int[] current, int logicalCount)
        {
            int[] next = new int[logicalCount];
            int copy = Math.Min(current.Length, logicalCount);
            if (copy > 0)
            {
                Array.Copy(current, next, copy);
            }

            return next;
        }

        bool[] ResizeFlags(bool[] current, int logicalCount)
        {
            bool[] next = new bool[logicalCount];
            int copy = Math.Min(current.Length, logicalCount);
            if (copy > 0)
            {
                Array.Copy(current, next, copy);
            }

            return next;
        }

        void ClampAssignments(int[] assignments, int groupCount)
        {
            int maxIndex = Math.Max(1, groupCount) - 1;
            for (int i = 0; i < assignments.Length; i++)
            {
                int value = assignments[i];
                if (value < 0)
                {
                    assignments[i] = 0;
                }
                else if (value > maxIndex)
                {
                    assignments[i] = maxIndex;
                }
            }
        }

        bool IsSingleGroupAssignment(int[] assignments, int logicalCount)
        {
            if (logicalCount <= 0 || assignments.Length == 0)
            {
                return true;
            }

            int count = Math.Min(logicalCount, assignments.Length);
            int first = assignments[0];
            for (int i = 1; i < count; i++)
            {
                if (assignments[i] != first)
                {
                    return false;
                }
            }

            return true;
        }

        void AutoSplitCcdAssignments(int logicalCount, int ccdCount)
        {
            if (logicalCount <= 0 || ccdAssign.Length == 0)
            {
                return;
            }

            if (ccdCount <= 1)
            {
                for (int i = 0; i < logicalCount && i < ccdAssign.Length; i++)
                {
                    ccdAssign[i] = 0;
                }
                return;
            }

            int coreCount = (int)coreGroupCountUpDown.Value;
            List<List<int>> coreGroups = [];
            for (int core = 0; core < coreCount; core++)
            {
                List<int> lps = [];
                for (int lp = 0; lp < logicalCount && lp < coreAssign.Length; lp++)
                {
                    if (coreAssign[lp] == core)
                    {
                        lps.Add(lp);
                    }
                }

                if (lps.Count > 0)
                {
                    coreGroups.Add(lps);
                }
            }

            if (coreGroups.Count == 0)
            {
                for (int lp = 0; lp < logicalCount && lp < ccdAssign.Length; lp++)
                {
                    ccdAssign[lp] = lp % ccdCount;
                }
                return;
            }

            int totalGroups = coreGroups.Count;
            int baseCount = totalGroups / ccdCount;
            int extra = totalGroups % ccdCount;
            int groupIndex = 0;

            for (int ccd = 0; ccd < ccdCount; ccd++)
            {
                int take = baseCount + (ccd < extra ? 1 : 0);
                for (int i = 0; i < take; i++)
                {
                    if (groupIndex >= totalGroups)
                    {
                        break;
                    }

                    foreach (int lp in coreGroups[groupIndex])
                    {
                        if (lp >= 0 && lp < ccdAssign.Length)
                        {
                            ccdAssign[lp] = ccd;
                        }
                    }

                    groupIndex++;
                }
            }
        }

        void AutoSplitCcxAssignments(int logicalCount, int ccxCount)
        {
            if (logicalCount <= 0 || ccxAssign.Length == 0)
            {
                return;
            }

            if (ccxCount <= 1)
            {
                for (int i = 0; i < logicalCount && i < ccxAssign.Length; i++)
                {
                    ccxAssign[i] = 0;
                }
                return;
            }

            int coreCount = (int)coreGroupCountUpDown.Value;
            List<List<int>> coreGroups = [];
            for (int core = 0; core < coreCount; core++)
            {
                List<int> lps = [];
                for (int lp = 0; lp < logicalCount && lp < coreAssign.Length; lp++)
                {
                    if (coreAssign[lp] == core)
                    {
                        lps.Add(lp);
                    }
                }

                if (lps.Count > 0)
                {
                    coreGroups.Add(lps);
                }
            }

            if (coreGroups.Count == 0)
            {
                for (int lp = 0; lp < logicalCount && lp < ccxAssign.Length; lp++)
                {
                    ccxAssign[lp] = lp % ccxCount;
                }
                return;
            }

            int totalGroups = coreGroups.Count;
            int baseCount = totalGroups / ccxCount;
            int extra = totalGroups % ccxCount;
            int groupIndex = 0;

            for (int ccx = 0; ccx < ccxCount; ccx++)
            {
                int take = baseCount + (ccx < extra ? 1 : 0);
                for (int i = 0; i < take; i++)
                {
                    if (groupIndex >= totalGroups)
                    {
                        break;
                    }

                    foreach (int lp in coreGroups[groupIndex])
                    {
                        if (lp >= 0 && lp < ccxAssign.Length)
                        {
                            ccxAssign[lp] = ccx;
                        }
                    }

                    groupIndex++;
                }
            }
        }

        void AutoGenerateSmtTopology(bool enabled)
        {
            if (smtAutoGenActive)
            {
                return;
            }

            smtAutoGenActive = true;
            try
            {
                int coreCount = (int)coreGroupCountUpDown.Value;
                if (coreCount <= 0)
                {
                    return;
                }

                int ccdCount = Math.Min(2, Math.Max(1, (int)ccdGroupCountUpDown.Value));
                int ccxCount = Math.Min(8, Math.Max(1, (int)ccxGroupCountUpDown.Value));
                int logicalCount = (int)logicalUpDown.Value;
                if (logicalCount <= 0)
                {
                    logicalCount = coreCount;
                }

                bool[] coreIsE = new bool[coreCount];
                int[] coreCcd = new int[coreCount];
                int[] coreCcx = new int[coreCount];
                for (int i = 0; i < coreCount; i++)
                {
                    coreCcd[i] = -1;
                    coreCcx[i] = -1;
                }

                int lpMax = Math.Min(logicalCount, coreAssign.Length);
                for (int lp = 0; lp < lpMax; lp++)
                {
                    int core = coreAssign[lp];
                    if (core < 0 || core >= coreCount)
                    {
                        continue;
                    }

                    if (lp < eAssign.Length && eAssign[lp])
                    {
                        coreIsE[core] = true;
                    }

                    if (coreCcd[core] < 0 && lp < ccdAssign.Length)
                    {
                        int ccd = ccdAssign[lp];
                        if (ccd < 0 || ccd >= ccdCount)
                        {
                            ccd = 0;
                        }

                        coreCcd[core] = ccd;
                    }

                    if (coreCcx[core] < 0 && lp < ccxAssign.Length)
                    {
                        int ccx = ccxAssign[lp];
                        if (ccx < 0 || ccx >= ccxCount)
                        {
                            ccx = 0;
                        }

                        coreCcx[core] = ccx;
                    }
                }

                for (int core = 0; core < coreCount; core++)
                {
                    if (coreCcd[core] < 0)
                    {
                        coreCcd[core] = core % ccdCount;
                    }

                    if (coreCcx[core] < 0)
                    {
                        coreCcx[core] = core % ccxCount;
                    }
                }

                int newLogical = 0;
                if (enabled)
                {
                    for (int core = 0; core < coreCount; core++)
                    {
                        newLogical += coreIsE[core] ? 1 : 2;
                    }
                }
                else
                {
                    newLogical = coreCount;
                }

                if (newLogical < 1)
                {
                    newLogical = 1;
                }

                bool truncated = newLogical > MaxAffinityBits;
                int maxLogical = Math.Min(MaxAffinityBits, newLogical);
                int[] newCoreAssign = new int[maxLogical];
                int[] newCcdAssign = new int[maxLogical];
                int[] newCcxAssign = new int[maxLogical];
                bool[] newEAssign = new bool[maxLogical];

                int index = 0;
                for (int core = 0; core < coreCount && index < maxLogical; core++)
                {
                    int threads = enabled && !coreIsE[core] ? 2 : 1;
                    for (int t = 0; t < threads && index < maxLogical; t++)
                    {
                        newCoreAssign[index] = core;
                        newCcdAssign[index] = coreCcd[core];
                        newCcxAssign[index] = coreCcx[core];
                        newEAssign[index] = coreIsE[core];
                        index++;
                    }
                }

                for (int lp = index; lp < maxLogical; lp++)
                {
                    int core = lp % coreCount;
                    newCoreAssign[lp] = core;
                    newCcdAssign[lp] = core % ccdCount;
                    newCcxAssign[lp] = core % ccxCount;
                    newEAssign[lp] = false;
                }

                suppressAssignmentEvents = true;
                logicalUpDown.Value = maxLogical;
                coreAssign = newCoreAssign;
                ccdAssign = newCcdAssign;
                ccxAssign = newCcxAssign;
                eAssign = newEAssign;

                int maxGroups = Math.Max(1, maxLogical);
                int maxCcdGroups = Math.Min(2, maxGroups);
                int maxCcxGroups = Math.Min(8, maxGroups);
                coreGroupCountUpDown.Maximum = maxGroups;
                ccdGroupCountUpDown.Maximum = maxCcdGroups;
                ccxGroupCountUpDown.Maximum = maxCcxGroups;
                if (coreGroupCountUpDown.Value > maxGroups)
                {
                    coreGroupCountUpDown.Value = maxGroups;
                }

                if (ccdGroupCountUpDown.Value > maxCcdGroups)
                {
                    ccdGroupCountUpDown.Value = maxCcdGroups;
                }

                if (ccxGroupCountUpDown.Value > maxCcxGroups)
                {
                    ccxGroupCountUpDown.Value = maxCcxGroups;
                }

                suppressAssignmentEvents = false;

                BuildAssignmentRows();
                syncDialogScroll?.Invoke();

                int eCount = coreIsE.Count(v => v);
                string note = truncated ? $" (capped to {MaxAffinityBits} LP)" : string.Empty;
                WriteLog($"TESTCPU.AUTO: SMT={(enabled ? "Enabled" : "Disabled")} logical={maxLogical} cores={coreCount} ccd={ccdCount} ccx={ccxCount} eCores={eCount}{note}");
            }
            finally
            {
                smtAutoGenActive = false;
            }
        }

        string FormatPresetCppc(Dictionary<int, int> ratings)
        {
            return string.Join(
                ", ",
                ratings
                    .OrderBy(kvp => kvp.Key)
                    .Select(kvp => $"{kvp.Key}={kvp.Value}"));
        }

        void LoadSyntheticPreset(
            string name,
            int logicalCount,
            int physicalCoreCount,
            int ccdCount,
            int ccxCount,
            bool smtEnabled,
            bool useHyperLabel,
            int[] coreMap,
            int[] ccdMap,
            int[] ccxMap,
            bool[] eCoreMap,
            Dictionary<int, int> cppcRatings)
        {
            if (logicalCount <= 0 || logicalCount > MaxAffinityBits)
            {
                return;
            }

            suppressAssignmentEvents = true;
            try
            {
                logicalUpDown.Value = logicalCount;
                coreGroupCountUpDown.Maximum = Math.Max(1, logicalCount);
                ccdGroupCountUpDown.Maximum = Math.Min(2, Math.Max(1, logicalCount));
                ccxGroupCountUpDown.Maximum = Math.Min(8, Math.Max(1, logicalCount));
                coreGroupCountUpDown.Value = Math.Max(1, Math.Min(physicalCoreCount, (int)coreGroupCountUpDown.Maximum));
                ccdGroupCountUpDown.Value = Math.Max(1, Math.Min(ccdCount, (int)ccdGroupCountUpDown.Maximum));
                ccxGroupCountUpDown.Value = Math.Max(1, Math.Min(ccxCount, (int)ccxGroupCountUpDown.Maximum));

                coreAssign = ResizeAssignments(coreMap, logicalCount);
                ccdAssign = ResizeAssignments(ccdMap, logicalCount);
                ccxAssign = ResizeAssignments(ccxMap, logicalCount);
                eAssign = ResizeFlags(eCoreMap, logicalCount);
                ClampAssignments(coreAssign, (int)coreGroupCountUpDown.Value);
                ClampAssignments(ccdAssign, (int)ccdGroupCountUpDown.Value);
                ClampAssignments(ccxAssign, (int)ccxGroupCountUpDown.Value);
            }
            finally
            {
                suppressAssignmentEvents = false;
            }

            SetSmtState(smtEnabled, useHyperLabel, false);
            cpuNameTextBox.Text = name;
            cppcRatingsBox.Text = FormatPresetCppc(cppcRatings);
            BuildAssignmentRows();
            syncDialogScroll?.Invoke();
            WriteLog($"TESTCPU.PRESET: loaded name=\"{name}\" logical={logicalCount} physical={physicalCoreCount} ccd={ccdCount} ccx={ccxCount} smt={smtEnabled} cppcCount={cppcRatings.Count}");
        }

        void LoadIntelHybridPreset(string name, int pCores, int eCores, bool pCoreHt, bool performanceRatings)
        {
            int logicalCount = (pCores * (pCoreHt ? 2 : 1)) + eCores;
            if (logicalCount > MaxAffinityBits)
            {
                logicalCount = MaxAffinityBits;
            }

            int physicalCoreCount = pCores + eCores;
            int[] coreMap = new int[logicalCount];
            int[] ccdMap = new int[logicalCount];
            int[] ccxMap = new int[logicalCount];
            bool[] eCoreMap = new bool[logicalCount];
            Dictionary<int, int> cppc = [];

            int lp = 0;
            for (int core = 0; core < pCores && lp < logicalCount; core++)
            {
                int threads = pCoreHt ? 2 : 1;
                int rating = core < 2 ? 140 - (core * 5) : 120 - Math.Min(12, core);
                for (int t = 0; t < threads && lp < logicalCount; t++)
                {
                    coreMap[lp] = core;
                    ccdMap[lp] = 0;
                    ccxMap[lp] = 0;
                    eCoreMap[lp] = false;
                    if (performanceRatings)
                    {
                        cppc[lp] = rating;
                    }
                    lp++;
                }
            }

            for (int e = 0; e < eCores && lp < logicalCount; e++)
            {
                int core = pCores + e;
                coreMap[lp] = core;
                ccdMap[lp] = 0;
                ccxMap[lp] = 0;
                eCoreMap[lp] = true;
                if (performanceRatings)
                {
                    cppc[lp] = 70 - Math.Min(20, e);
                }
                lp++;
            }

            LoadSyntheticPreset(name, logicalCount, physicalCoreCount, 1, 1, pCoreHt, true, coreMap, ccdMap, ccxMap, eCoreMap, cppc);
        }

        void LoadAmdPreset(string name, int physicalCores, int ccdCount, string cppcProfile, int ccxPerCcd = 1, bool smtEnabled = true)
        {
            int threadsPerCore = smtEnabled ? 2 : 1;
            int logicalCount = Math.Min(MaxAffinityBits, physicalCores * threadsPerCore);
            int[] coreMap = new int[logicalCount];
            int[] ccdMap = new int[logicalCount];
            int[] ccxMap = new int[logicalCount];
            bool[] eCoreMap = new bool[logicalCount];
            Dictionary<int, int> cppc = [];

            int coresPerCcd = Math.Max(1, (int)Math.Ceiling(physicalCores / (double)Math.Max(1, ccdCount)));
            int safeCcxPerCcd = Math.Max(1, ccxPerCcd);
            int ccxCount = Math.Max(1, ccdCount * safeCcxPerCcd);
            int coresPerCcx = Math.Max(1, (int)Math.Ceiling(coresPerCcd / (double)safeCcxPerCcd));
            int lp = 0;
            for (int core = 0; core < physicalCores && lp < logicalCount; core++)
            {
                int ccd = Math.Min(ccdCount - 1, core / coresPerCcd);
                int coreInCcd = core - (coresPerCcd * ccd);
                int ccxInCcd = Math.Min(safeCcxPerCcd - 1, coreInCcd / coresPerCcx);
                int coreInCcx = coreInCcd - (coresPerCcx * ccxInCcd);
                int ccx = (ccd * safeCcxPerCcd) + ccxInCcd;
                int rating = cppcProfile switch
                {
                    "x3d-cache" => ccd == 0
                        ? 140 - Math.Min(12, coreInCcd)
                        : 112 - Math.Min(12, coreInCcd),
                    "x3d-dual-cache" => 136 - Math.Min(16, coreInCcd) - Math.Min(2, ccd),
                    "zen2" => 122 - Math.Min(10, coreInCcx) - Math.Min(4, ccxInCcd * 2) - Math.Min(4, ccd * 2),
                    "standard" => 128 - Math.Min(16, coreInCcd) - Math.Min(4, ccd),
                    _ => 120 - Math.Min(20, coreInCcd),
                };

                for (int t = 0; t < threadsPerCore && lp < logicalCount; t++)
                {
                    coreMap[lp] = core;
                    ccdMap[lp] = ccd;
                    ccxMap[lp] = ccx;
                    eCoreMap[lp] = false;
                    cppc[lp] = rating;

                    lp++;
                }
            }

            LoadSyntheticPreset(name, logicalCount, physicalCores, ccdCount, ccxCount, smtEnabled, false, coreMap, ccdMap, ccxMap, eCoreMap, cppc);
        }

        void ApplySelectedCpuPreset()
        {
            string preset = cpuPresetCombo.SelectedItem?.ToString() ?? string.Empty;
            switch (preset)
            {
                case "Manual (current)":
                case "Manual / current":
                    return;
                case "Intel Core i7-10700K/11700K 8C/16T":
                    LoadIntelHybridPreset("Intel Core i7-10700K/11700K 8C/16T", pCores: 8, eCores: 0, pCoreHt: true, performanceRatings: false);
                    break;
                case "Intel Core i5-13600K/14600K 6P+8E/20T":
                    LoadIntelHybridPreset("Intel Core i5-13600K/14600K 6P+8E/20T", pCores: 6, eCores: 8, pCoreHt: true, performanceRatings: true);
                    break;
                case "Intel Core i9-13900K/14900K 8P+16E/32T":
                    LoadIntelHybridPreset("Intel Core i9-13900K/14900K 8P+16E/32T", pCores: 8, eCores: 16, pCoreHt: true, performanceRatings: true);
                    break;
                case "Intel Core Ultra 9 285K 8P+16E/24T":
                    LoadIntelHybridPreset("Intel Core Ultra 9 285K 8P+16E/24T", pCores: 8, eCores: 16, pCoreHt: false, performanceRatings: true);
                    break;
                case "Intel Core Ultra 9 285HX 8P+16E/24T":
                    LoadIntelHybridPreset("Intel Core Ultra 9 285HX 8P+16E/24T", pCores: 8, eCores: 16, pCoreHt: false, performanceRatings: true);
                    break;
                case "AMD Ryzen 5 7500F/7600X 6C/12T":
                    LoadAmdPreset("AMD Ryzen 5 7500F/7600X 6C/12T", physicalCores: 6, ccdCount: 1, cppcProfile: "standard");
                    break;
                case "AMD Ryzen 7 7700X/9700X 8C/16T":
                    LoadAmdPreset("AMD Ryzen 7 7700X/9700X 8C/16T", physicalCores: 8, ccdCount: 1, cppcProfile: "standard");
                    break;
                case "AMD Ryzen 9 7900X/9900X 12C/24T":
                    LoadAmdPreset("AMD Ryzen 9 7900X/9900X 12C/24T", physicalCores: 12, ccdCount: 2, cppcProfile: "standard");
                    break;
                case "AMD Ryzen 9 7950X/9950X 16C/32T":
                    LoadAmdPreset("AMD Ryzen 9 7950X/9950X 16C/32T", physicalCores: 16, ccdCount: 2, cppcProfile: "standard");
                    break;
                case "AMD Ryzen 5 3500X Zen2 6C/6T SMT off 2 CCX":
                    // CPPC/CCX mirrored from DeviceTweaker_20260815_064357_362.log (Ryzen 5 3500X, SMT off).
                    LoadSyntheticPreset(
                        "AMD Ryzen 5 3500X 6-Core Processor",
                        logicalCount: 6,
                        physicalCoreCount: 6,
                        ccdCount: 1,
                        ccxCount: 2,
                        smtEnabled: false,
                        useHyperLabel: false,
                        coreMap: [0, 1, 2, 3, 4, 5],
                        ccdMap: [0, 0, 0, 0, 0, 0],
                        ccxMap: [0, 0, 0, 1, 1, 1],
                        eCoreMap: [false, false, false, false, false, false],
                        cppcRatings: new Dictionary<int, int>
                        {
                            [0] = 121,
                            [1] = 117,
                            [2] = 114,
                            [3] = 124,
                            [4] = 128,
                            [5] = 128,
                        });
                    break;
                case "AMD Ryzen 7 3700X/3800X Zen2 8C/16T 2 CCX":
                    LoadAmdPreset("AMD Ryzen 7 3700X/3800X Zen2 8C/16T 2 CCX", physicalCores: 8, ccdCount: 1, cppcProfile: "zen2", ccxPerCcd: 2);
                    break;
                case "AMD Ryzen 9 3900X Zen2 12C/24T 4 CCX":
                    LoadAmdPreset("AMD Ryzen 9 3900X Zen2 12C/24T 4 CCX", physicalCores: 12, ccdCount: 2, cppcProfile: "zen2", ccxPerCcd: 2);
                    break;
                case "AMD Ryzen 9 3950X Zen2 16C/32T 4 CCX":
                    LoadAmdPreset("AMD Ryzen 9 3950X Zen2 16C/32T 4 CCX", physicalCores: 16, ccdCount: 2, cppcProfile: "zen2", ccxPerCcd: 2);
                    break;
                case "AMD Ryzen 7 7800X3D/9800X3D 8C/16T V-Cache":
                    LoadAmdPreset("AMD Ryzen 7 7800X3D/9800X3D 8C/16T V-Cache", physicalCores: 8, ccdCount: 1, cppcProfile: "x3d-cache");
                    break;
                case "AMD Ryzen 9 7900X3D/9900X3D 12C/24T V-Cache CCD0":
                    LoadAmdPreset("AMD Ryzen 9 7900X3D/9900X3D 12C/24T V-Cache CCD0", physicalCores: 12, ccdCount: 2, cppcProfile: "x3d-cache");
                    break;
                case "AMD Ryzen 9 7950X3D/9950X3D 16C/32T V-Cache CCD0":
                    LoadAmdPreset("AMD Ryzen 9 7950X3D/9950X3D 16C/32T V-Cache CCD0", physicalCores: 16, ccdCount: 2, cppcProfile: "x3d-cache");
                    break;
                case "AMD Ryzen 9 9955HX3D 16C/32T V-Cache CCD0":
                    LoadAmdPreset("AMD Ryzen 9 9955HX3D 16C/32T V-Cache CCD0", physicalCores: 16, ccdCount: 2, cppcProfile: "x3d-cache");
                    break;
                case "AMD Ryzen 9 8940HX 16C/32T Dual-CCD":
                    // CPPC/CCD mirrored from DeviceTweaker_20260909_182154_688.log (Ryzen 9 8940HX Dual-CCD).
                    LoadSyntheticPreset(
                        "AMD Ryzen 9 8940HX with Radeon Graphics",
                        logicalCount: 32,
                        physicalCoreCount: 16,
                        ccdCount: 2,
                        ccxCount: 2,
                        smtEnabled: true,
                        useHyperLabel: false,
                        coreMap:
                        [
                            0, 0, 2, 2, 4, 4, 6, 6, 8, 8, 10, 10, 12, 12, 14, 14,
                            16, 16, 18, 18, 20, 20, 22, 22, 24, 24, 26, 26, 28, 28, 30, 30,
                        ],
                        ccdMap:
                        [
                            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
                            1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1,
                        ],
                        ccxMap:
                        [
                            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
                            1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1,
                        ],
                        eCoreMap: new bool[32],
                        cppcRatings: new Dictionary<int, int>
                        {
                            [0] = 318, [1] = 318, [2] = 318, [3] = 318,
                            [4] = 291, [5] = 291, [6] = 298, [7] = 298,
                            [8] = 285, [9] = 285, [10] = 305, [11] = 305,
                            [12] = 278, [13] = 278, [14] = 312, [15] = 312,
                            [16] = 264, [17] = 264, [18] = 271, [19] = 271,
                            [20] = 244, [21] = 244, [22] = 258, [23] = 258,
                            [24] = 231, [25] = 231, [26] = 251, [27] = 251,
                            [28] = 224, [29] = 224, [30] = 237, [31] = 237,
                        });
                    break;
                case "AMD Ryzen 9 9950X3D2 16C/32T dual V-Cache":
                    LoadAmdPreset("AMD Ryzen 9 9950X3D2 16C/32T dual V-Cache", physicalCores: 16, ccdCount: 2, cppcProfile: "x3d-dual-cache");
                    break;
            }
        }

        void FillGroupCombo(ThemedDropDownPicker combo, int groupCount)
        {
            combo.Items.Clear();
            for (int i = 0; i < groupCount; i++)
            {
                combo.Items.Add(UiLanguage.IsRussian ? $"Группа {i + 1}" : $"Group {i + 1}");
            }
            combo.Invalidate();
        }

        void BuildAssignmentRows()
        {
            int logicalCount = (int)logicalUpDown.Value;
            int coreCount = (int)coreGroupCountUpDown.Value;
            int ccdCount = (int)ccdGroupCountUpDown.Value;
            int ccxCount = (int)ccxGroupCountUpDown.Value;

            suppressAssignmentEvents = true;
            assignmentsTable.SuspendLayout();
            foreach (Control control in assignmentsTable.Controls.Cast<Control>().ToArray())
            {
                control.Dispose();
            }

            assignmentsTable.Controls.Clear();
            assignmentsTable.RowStyles.Clear();

            assignmentsTable.RowCount = logicalCount + 1;
            assignmentsTable.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            assignmentsTable.Controls.Add(NewHeaderLabel("LP"), 0, 0);
            assignmentsTable.Controls.Add(NewHeaderLabel("Core group"), 1, 0);
            assignmentsTable.Controls.Add(NewHeaderLabel("CCD group"), 2, 0);
            assignmentsTable.Controls.Add(NewHeaderLabel("CCX group"), 3, 0);
            assignmentsTable.Controls.Add(NewHeaderLabel("E-core"), 4, 0);

            for (int i = 0; i < logicalCount; i++)
            {
                assignmentsTable.RowStyles.Add(new RowStyle(SizeType.AutoSize));

                Label lpLabel = NewInlineLabel($"LP {i}");
                lpLabel.ForeColor = _fgMain;
                lpLabel.Margin = new Padding(0, 4, 12, 4);

                ThemedDropDownPicker coreCombo = NewDialogCombo(150);
                coreCombo.Margin = new Padding(0, 2, 12, 4);
                coreCombo.Dock = DockStyle.Top;
                FillGroupCombo(coreCombo, coreCount);
                int coreIndex = i < coreAssign.Length ? coreAssign[i] : 0;
                coreIndex = Math.Clamp(coreIndex, 0, coreCount - 1);
                coreCombo.SelectedIndex = coreIndex;
                coreCombo.Tag = i;
                coreCombo.SelectedIndexChanged += (_, _) =>
                {
                    if (suppressAssignmentEvents)
                    {
                        return;
                    }

                    if (coreCombo.Tag is int lp)
                    {
                        coreAssign[lp] = coreCombo.SelectedIndex;
                    }
                };

                ThemedDropDownPicker ccdCombo = NewDialogCombo(150);
                ccdCombo.Margin = new Padding(0, 2, 12, 4);
                ccdCombo.Dock = DockStyle.Top;
                FillGroupCombo(ccdCombo, ccdCount);
                int ccdIndex = i < ccdAssign.Length ? ccdAssign[i] : 0;
                ccdIndex = Math.Clamp(ccdIndex, 0, ccdCount - 1);
                ccdCombo.SelectedIndex = ccdIndex;
                ccdCombo.Tag = i;
                ccdCombo.SelectedIndexChanged += (_, _) =>
                {
                    if (suppressAssignmentEvents)
                    {
                        return;
                    }

                    if (ccdCombo.Tag is int lp)
                    {
                        ccdAssign[lp] = ccdCombo.SelectedIndex;
                    }
                };

                ThemedDropDownPicker ccxCombo = NewDialogCombo(120);
                ccxCombo.Margin = new Padding(0, 2, 12, 4);
                ccxCombo.Dock = DockStyle.Top;
                FillGroupCombo(ccxCombo, ccxCount);
                int ccxIndex = i < ccxAssign.Length ? ccxAssign[i] : 0;
                ccxIndex = Math.Clamp(ccxIndex, 0, ccxCount - 1);
                ccxCombo.SelectedIndex = ccxIndex;
                ccxCombo.Tag = i;
                ccxCombo.SelectedIndexChanged += (_, _) =>
                {
                    if (suppressAssignmentEvents)
                    {
                        return;
                    }

                    if (ccxCombo.Tag is int lp)
                    {
                        ccxAssign[lp] = ccxCombo.SelectedIndex;
                    }
                };

                CheckBox eCheck = new()
                {
                    Text = "E-Core",
                    AutoSize = true,
                    BackColor = _bgPanel,
                    ForeColor = _fgMain,
                    Margin = new Padding(8, 2, 0, 4),
                    Tag = i,
                };
                eCheck.Checked = i < eAssign.Length && eAssign[i];
                eCheck.CheckedChanged += (_, _) =>
                {
                    if (suppressAssignmentEvents)
                    {
                        return;
                    }

                    if (eCheck.Tag is int lp && lp < eAssign.Length)
                    {
                        eAssign[lp] = eCheck.Checked;
                    }
                };

                assignmentsTable.Controls.Add(lpLabel, 0, i + 1);
                assignmentsTable.Controls.Add(coreCombo, 1, i + 1);
                assignmentsTable.Controls.Add(ccdCombo, 2, i + 1);
                assignmentsTable.Controls.Add(ccxCombo, 3, i + 1);
                assignmentsTable.Controls.Add(eCheck, 4, i + 1);
            }

            assignmentsTable.ResumeLayout();
            suppressAssignmentEvents = false;
        }

        void RefreshAssignmentUi(bool autoSplitCcd, bool autoSplitCcx = false)
        {
            if (suppressAssignmentEvents)
            {
                return;
            }

            suppressAssignmentEvents = true;
            int logicalCount = (int)logicalUpDown.Value;
            coreAssign = ResizeAssignments(coreAssign, logicalCount);
            ccdAssign = ResizeAssignments(ccdAssign, logicalCount);
            ccxAssign = ResizeAssignments(ccxAssign, logicalCount);
            eAssign = ResizeFlags(eAssign, logicalCount);

            int maxGroups = Math.Max(1, logicalCount);
            coreGroupCountUpDown.Maximum = maxGroups;
            int maxCcdGroups = Math.Min(2, maxGroups);
            int maxCcxGroups = Math.Min(8, maxGroups);
            ccdGroupCountUpDown.Maximum = maxCcdGroups;
            ccxGroupCountUpDown.Maximum = maxCcxGroups;

            if (coreGroupCountUpDown.Value > maxGroups)
            {
                coreGroupCountUpDown.Value = maxGroups;
            }

            if (ccdGroupCountUpDown.Value > maxCcdGroups)
            {
                ccdGroupCountUpDown.Value = maxCcdGroups;
            }

            if (ccxGroupCountUpDown.Value > maxCcxGroups)
            {
                ccxGroupCountUpDown.Value = maxCcxGroups;
            }

            ClampAssignments(coreAssign, (int)coreGroupCountUpDown.Value);
            ClampAssignments(ccdAssign, (int)ccdGroupCountUpDown.Value);
            ClampAssignments(ccxAssign, (int)ccxGroupCountUpDown.Value);
            if (autoSplitCcd && (int)ccdGroupCountUpDown.Value > 1 && IsSingleGroupAssignment(ccdAssign, logicalCount))
            {
                AutoSplitCcdAssignments(logicalCount, (int)ccdGroupCountUpDown.Value);
            }
            if (autoSplitCcx && (int)ccxGroupCountUpDown.Value > 1 && IsSingleGroupAssignment(ccxAssign, logicalCount))
            {
                AutoSplitCcxAssignments(logicalCount, (int)ccxGroupCountUpDown.Value);
            }
            suppressAssignmentEvents = false;

            BuildAssignmentRows();
            syncDialogScroll?.Invoke();
        }

        void LoadAssignmentsFromCurrentCpu()
        {
            int logicalCount = (int)logicalUpDown.Value;
            coreAssign = BuildAssignmentsFromGroupsText(GetCurrentCoreGroupsText(), logicalCount, out int coreCount);
            ccdAssign = BuildAssignmentsFromGroupsText(GetCurrentCcdGroupsText(), logicalCount, out int ccdCount);
            ccxAssign = BuildAssignmentsFromGroupsText(GetCurrentCcxGroupsText(), logicalCount, out int ccxCount);
            eAssign = BuildECoreFlags(GetCurrentECoreText(), logicalCount);

            suppressAssignmentEvents = true;
            coreGroupCountUpDown.Maximum = Math.Max(1, logicalCount);
            ccdGroupCountUpDown.Maximum = Math.Min(2, Math.Max(1, logicalCount));
            ccxGroupCountUpDown.Maximum = Math.Min(8, Math.Max(1, logicalCount));
            coreGroupCountUpDown.Value = Math.Max(1, Math.Min(coreCount, (int)coreGroupCountUpDown.Maximum));
            ccdGroupCountUpDown.Value = Math.Max(1, Math.Min(ccdCount, (int)ccdGroupCountUpDown.Maximum));
            ccxGroupCountUpDown.Value = Math.Max(1, Math.Min(ccxCount, (int)ccxGroupCountUpDown.Maximum));
            suppressAssignmentEvents = false;

            ClampAssignments(coreAssign, (int)coreGroupCountUpDown.Value);
            ClampAssignments(ccdAssign, (int)ccdGroupCountUpDown.Value);
            ClampAssignments(ccxAssign, (int)ccxGroupCountUpDown.Value);
            BuildAssignmentRows();
            syncDialogScroll?.Invoke();
        }

        bool TryParseTestCppcRatings(string text, int logicalCount, out Dictionary<int, int> ratings, out string error)
        {
            ratings = [];
            error = string.Empty;
            if (string.IsNullOrWhiteSpace(text))
            {
                return true;
            }

            string[] parts = text
                .Split([',', ';', ' ', '\t', '\r', '\n'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            int sequentialLp = 0;
            foreach (string part in parts)
            {
                string token = part.Trim();
                if (token.Length == 0)
                {
                    continue;
                }

                int lp;
                string ratingText;
                int sep = token.IndexOf('=');
                if (sep < 0)
                {
                    sep = token.IndexOf(':');
                }

                if (sep >= 0)
                {
                    string lpText = token[..sep].Trim();
                    ratingText = token[(sep + 1)..].Trim();
                    if (!int.TryParse(lpText, NumberStyles.Integer, CultureInfo.InvariantCulture, out lp))
                    {
                        error = $"Bad CPPC LP index: {lpText}";
                        return false;
                    }
                }
                else
                {
                    lp = sequentialLp;
                    ratingText = token;
                }

                sequentialLp = Math.Max(sequentialLp + 1, lp + 1);
                if (lp < 0 || lp >= logicalCount)
                {
                    error = $"CPPC LP index out of range: {lp}";
                    return false;
                }

                if (!int.TryParse(ratingText, NumberStyles.Integer, CultureInfo.InvariantCulture, out int rating) || rating < 0)
                {
                    error = $"Bad CPPC rating for LP {lp}: {ratingText}";
                    return false;
                }

                ratings[lp] = rating;
            }

            return true;
        }

        TestCpuConfig BuildConfigFromAssignments(IReadOnlyDictionary<int, int> cppcRatings)
        {
            int logicalCount = (int)logicalUpDown.Value;
            int coreCount = (int)coreGroupCountUpDown.Value;
            int ccdCount = (int)ccdGroupCountUpDown.Value;
            int ccxCount = (int)ccxGroupCountUpDown.Value;

            TestCpuConfig config = new()
            {
                LogicalCount = logicalCount,
                SmtEnabled = smtStateCombo.SelectedIndex == 0,
                UseHyperThreadingLabel = useHyperThreadingLabel,
                CcdMap = new Dictionary<int, int>(),
                CcxMap = new Dictionary<int, int>(),
                CpuName = cpuNameTextBox.Text,
            };

            for (int lp = 0; lp < logicalCount; lp++)
            {
                int coreGroup = lp < coreAssign.Length ? coreAssign[lp] : 0;
                int ccdGroup = lp < ccdAssign.Length ? ccdAssign[lp] : 0;
                int ccxGroup = lp < ccxAssign.Length ? ccxAssign[lp] : 0;
                if (coreGroup < 0 || coreGroup >= coreCount)
                {
                    coreGroup = 0;
                }

                if (ccdGroup < 0 || ccdGroup >= ccdCount)
                {
                    ccdGroup = 0;
                }

                if (ccxGroup < 0 || ccxGroup >= ccxCount)
                {
                    ccxGroup = 0;
                }

                config.CoreMap[lp] = coreGroup;
                config.CcdMap[lp] = ccdGroup;
                config.CcxMap[lp] = ccxGroup;
                if (lp < eAssign.Length && eAssign[lp])
                {
                    config.ECoreLps.Add(lp);
                }
            }

            foreach (KeyValuePair<int, int> pair in cppcRatings)
            {
                config.CppcRatings[pair.Key] = pair.Value;
            }

            return config;
        }

        logicalUpDown.ValueChanged += (_, _) => RefreshAssignmentUi(false);
        coreGroupCountUpDown.ValueChanged += (_, _) => RefreshAssignmentUi(false);
        ccdGroupCountUpDown.ValueChanged += (_, _) => RefreshAssignmentUi(true);
        ccxGroupCountUpDown.ValueChanged += (_, _) => RefreshAssignmentUi(false, true);
        cpuPresetButton.Click += (_, _) => ApplySelectedCpuPreset();
        cpuPresetCombo.SelectedIndexChanged += (_, _) =>
        {
            if (cpuPresetCombo.SelectedIndex > 0)
            {
                ApplySelectedCpuPreset();
            }
        };

        BuildAssignmentRows();

        FlowLayoutPanel buttons = new()
        {
            Dock = DockStyle.Bottom,
            Height = 62,
            WrapContents = false,
            FlowDirection = FlowDirection.LeftToRight,
            Padding = new Padding(24, 14, 24, 16),
            BackColor = _bgPanel,
        };
        buttons.Paint += (_, e) =>
        {
            using Pen pen = new(Color.FromArgb(48, 48, 54));
            e.Graphics.DrawLine(pen, 24, 0, Math.Max(24, buttons.ClientSize.Width - 24), 0);
        };

        void RecenterFooterButtons()
        {
            Control[] visible = buttons.Controls.Cast<Control>().Where(control => control.Visible).ToArray();
            int totalWidth = visible.Sum(control => control.Width + control.Margin.Horizontal);
            int left = Math.Max(24, (buttons.ClientSize.Width - totalWidth) / 2);
            buttons.Padding = new Padding(left, 14, 24, 16);
        }

        buttons.Resize += (_, _) => RecenterFooterButtons();

        Button applyButton = NewDialogButton("APPLY CPU TOPOLOGY");
        applyButton.Size = new Size(230, 32);
        applyButton.Margin = new Padding(0, 0, 12, 0);
        applyButton.FlatAppearance.BorderColor = _accent;
        applyButton.VisibleChanged += (_, _) => RecenterFooterButtons();
        applyButton.Click += (_, _) =>
        {
            if (!TryParseTestCppcRatings(cppcRatingsBox.Text, (int)logicalUpDown.Value, out Dictionary<int, int> cppcRatings, out string cppcError))
            {
                ShowThemedInfo($"CPPC ratings are invalid.\n{cppcError}");
                return;
            }

            TestCpuConfig config = BuildConfigFromAssignments(cppcRatings);
            _testDevicesEnabled = true;
            enableTestDevicesCheck.Checked = true;
            _testAutoDryRun = dryRunAutoCheck.Checked || _testDevicesOnly;
            dryRunAutoCheck.Checked = _testAutoDryRun;
            ApplyTestCpuConfig(config);
            SyncSandboxDryRunLock("apply-test");
            statusLabel.Text = "Test CPU mode: ACTIVE";
            statusLabel.ForeColor = _statusActive;
        };

        Button resetButton = NewDialogButton("DISABLE TEST MODE");
        resetButton.Name = "TEST_ADMIN_DISABLE_MODE";
        resetButton.Size = new Size(220, 32);
        resetButton.Margin = new Padding(0, 0, 12, 0);
        resetButton.Click += (_, _) =>
        {
            ResetAdminPanelToRealState();
        };

        Button closeButton = NewDialogButton("CLOSE");
        closeButton.Name = "TEST_ADMIN_CLOSE";
        closeButton.Size = new Size(170, 32);
        closeButton.Margin = new Padding(0, 0, 0, 0);
        closeButton.DialogResult = DialogResult.Cancel;

        buttons.Controls.Add(applyButton);
        buttons.Controls.Add(resetButton);
        buttons.Controls.Add(closeButton);
        RecenterFooterButtons();

        TableLayoutPanel navigation = new()
        {
            Dock = DockStyle.Top,
            Height = 48,
            ColumnCount = 5,
            RowCount = 1,
            Padding = new Padding(24, 8, 24, 8),
            BackColor = _bgPanel,
            TabStop = false,
        };
        for (int column = 0; column < 4; column++)
        {
            navigation.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F));
        }
        navigation.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, UiScale(102)));
        navigation.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
        navigation.Paint += (_, e) =>
        {
            using Pen pen = new(Color.FromArgb(48, 48, 54));
            int y = Math.Max(0, navigation.ClientSize.Height - 1);
            e.Graphics.DrawLine(pen, 24, y, Math.Max(24, navigation.ClientSize.Width - 24), y);
        };

        Button NewNavigationButton(string name, string text)
        {
            Button button = NewDialogButton(text);
            button.Name = name;
            button.Dock = DockStyle.Fill;
            button.Margin = new Padding(4, 0, 4, 0);
            button.MinimumSize = new Size(0, 30);
            button.FlatAppearance.BorderColor = Color.FromArgb(80, 80, 86);
            button.AccessibleDescription = UiLanguage.Text("Jump to section");
            return button;
        }

        Button cpuNavigationButton = NewNavigationButton("TEST_ADMIN_NAV_CPU", "CPU TOPOLOGY");
        Button devicesNavigationButton = NewNavigationButton("TEST_ADMIN_NAV_DEVICES", "DEVICES");
        Button scenarioNavigationButton = NewNavigationButton("TEST_ADMIN_NAV_SCENARIO", "SCENARIO LAB");
        Button resultsNavigationButton = NewNavigationButton("TEST_ADMIN_NAV_RESULTS", "RESULT DIALOGS");
        Button[] navigationButtons =
        [
            cpuNavigationButton,
            devicesNavigationButton,
            scenarioNavigationButton,
            resultsNavigationButton,
        ];
        navigation.Controls.Add(cpuNavigationButton, 0, 0);
        navigation.Controls.Add(devicesNavigationButton, 1, 0);
        navigation.Controls.Add(scenarioNavigationButton, 2, 0);
        navigation.Controls.Add(resultsNavigationButton, 3, 0);

        Panel langPanel = new()
        {
            Dock = DockStyle.Fill,
            BackColor = _bgPanel,
            Margin = new Padding(UiScale(6), 0, 0, 0),
            TabStop = false,
            Tag = LocalizationIgnoreTag,
        };

        Size langBtnSize = new(UiScale(44), UiScale(28));
        int btnY = (UiScale(32) - langBtnSize.Height) / 2;
        LanguageSelectorButton adminLangEn = new()
        {
            Name = "TEST_ADMIN_LANG_EN",
            Text = "EN",
            AutoSize = false,
            Location = new Point(0, btnY),
            Size = langBtnSize,
            MinimumSize = langBtnSize,
            MaximumSize = langBtnSize,
            BackColor = _bgPanel,
            ForeColor = _mutedText,
            Font = _baseFont,
            Cursor = Cursors.Hand,
            TabStop = true,
            Tag = LocalizationIgnoreTag,
            UseMnemonic = false,
        };

        LanguageSelectorButton adminLangRu = new()
        {
            Name = "TEST_ADMIN_LANG_RU",
            Text = "RU",
            AutoSize = false,
            Location = new Point(langBtnSize.Width + UiScale(6), btnY),
            Size = langBtnSize,
            MinimumSize = langBtnSize,
            MaximumSize = langBtnSize,
            BackColor = _bgPanel,
            ForeColor = _mutedText,
            Font = _baseFont,
            Cursor = Cursors.Hand,
            TabStop = true,
            Tag = LocalizationIgnoreTag,
            UseMnemonic = false,
        };

        void UpdateAdminLangStyle()
        {
            StyleLanguageButton(adminLangEn, UiLanguage.Current == UiLanguageCode.English, "English");
            StyleLanguageButton(adminLangRu, UiLanguage.Current == UiLanguageCode.Russian, "Русский");
        }

        adminLangEn.Click += (_, _) =>
        {
            SetUiLanguage(UiLanguageCode.English);
            UpdateAdminLangStyle();
            LocalizeControlTree(dialog);
            syncDialogScroll?.Invoke();
            dialog.Invalidate(true);
        };

        adminLangRu.Click += (_, _) =>
        {
            SetUiLanguage(UiLanguageCode.Russian);
            UpdateAdminLangStyle();
            LocalizeControlTree(dialog);
            syncDialogScroll?.Invoke();
            dialog.Invalidate(true);
        };

        EventHandler langSyncHandler = (_, _) =>
        {
            if (dialog.IsDisposed) return;
            UpdateAdminLangStyle();
            LocalizeControlTree(dialog);
            syncDialogScroll?.Invoke();
            dialog.Invalidate(true);
        };
        UiLanguage.Changed += langSyncHandler;
        dialog.FormClosed += (_, _) => UiLanguage.Changed -= langSyncHandler;

        UpdateAdminLangStyle();
        langPanel.Controls.Add(adminLangEn);
        langPanel.Controls.Add(adminLangRu);
        navigation.Controls.Add(langPanel, 4, 0);

        const int dialogScrollWidth = 12;
        Panel contentHost = new()
        {
            Dock = DockStyle.Fill,
            BackColor = _bgPanel,
            Padding = Padding.Empty,
        };

        Panel contentPanel = new()
        {
            Dock = DockStyle.None,
            BackColor = _bgPanel,
            AutoScroll = false,
            Padding = Padding.Empty,
        };
        contentPanel.Location = new Point(0, 0);
        contentPanel.Controls.Add(layout);

        ThemedScrollBar dialogScroll = new()
        {
            Width = dialogScrollWidth,
            BackColor = _bgPanel,
            TrackColor = _bgPanel,
            RailColor = _bgPanel,
            ThumbColor = _accent,
            ThumbWidth = 9,
            RailWidth = 0,
            ThumbCornerRadius = 7,
            Visible = false,
            Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Right,
        };

        bool syncingDialogScroll = false;
        bool switchingAdminSection = false;
        int activeAdminSection = 0;

        void UpdateDialogScrollLayout()
        {
            dialogScroll.Location = new Point(Math.Max(0, contentHost.ClientSize.Width - dialogScroll.Width), 0);
            dialogScroll.Height = contentHost.ClientSize.Height;
            dialogScroll.BringToFront();
        }

        void UpdateDialogHostLayout()
        {
            // Reserve a permanent lane for the scrollbar. Overlaying it on
            // the content made the rightmost fields look clipped and changed
            // their width when scrolling became necessary.
            contentPanel.Width = Math.Max(UiScale(320), contentHost.ClientSize.Width - dialogScroll.Width - UiScale(4));
            if (contentPanel.Left != 0)
            {
                contentPanel.Left = 0;
            }

            int maxOffset = Math.Max(0, contentPanel.Height - contentHost.ClientSize.Height);
            int offset = Math.Max(0, -contentPanel.Top);
            if (offset > maxOffset)
            {
                contentPanel.Top = -maxOffset;
            }
        }

        void SetDialogScrollOffset(int offset)
        {
            int maxOffset = Math.Max(0, contentPanel.Height - contentHost.ClientSize.Height);
            int next = Math.Max(0, Math.Min(maxOffset, offset));
            contentPanel.Location = new Point(0, -next);
        }

        int GetDialogContentY(Control control)
        {
            Point screenPoint = control.Parent?.PointToScreen(control.Location)
                ?? control.PointToScreen(Point.Empty);
            return contentPanel.PointToClient(screenPoint).Y;
        }

        void SetActiveNavigation(Button activeButton)
        {
            foreach (Button button in navigationButtons)
            {
                bool active = ReferenceEquals(button, activeButton);
                button.FlatAppearance.BorderColor = active ? _accent : Color.FromArgb(80, 80, 86);
                button.FlatAppearance.BorderSize = active ? 2 : 1;
                button.AccessibleDescription = UiLanguage.Text(active ? "Current section" : "Jump to section");
            }
        }

        Button GetNavigationButton(int section)
        {
            return section switch
            {
                1 => devicesNavigationButton,
                2 => scenarioNavigationButton,
                3 => resultsNavigationButton,
                _ => cpuNavigationButton,
            };
        }

        void ShowAdminSection(int section, string reason, bool validate = true)
        {
            section = Math.Max(0, Math.Min(3, section));
            if (switchingAdminSection)
            {
                return;
            }

            switchingAdminSection = true;
            layout.SuspendLayout();
            foreach (Control control in layout.Controls)
            {
                int row = layout.GetRow(control);
                control.Visible = section switch
                {
                    0 => row is >= 0 and <= 12,
                    1 => row is >= 13 and <= 14,
                    2 => row is >= 15 and <= 16,
                    3 => row is >= 17 and <= 19,
                    _ => false,
                };
            }
            applyButton.Visible = section == 0;
            activeAdminSection = section;
            contentPanel.Location = Point.Empty;
            layout.ResumeLayout(true);
            switchingAdminSection = false;

            SyncDialogScrollBar();
            syncingDialogScroll = true;
            dialogScroll.Value = 0;
            syncingDialogScroll = false;
            SetActiveNavigation(GetNavigationButton(section));
            WriteLog($"TEST.ADMIN.NAV: section={section} reason={reason} offset=0");
            if (validate && dialog.IsHandleCreated)
            {
                dialog.BeginInvoke(new Action(() => ValidateTestAdminLayout($"navigation-{section}-{reason}")));
            }
        }

        cpuNavigationButton.Click += (_, _) => ShowAdminSection(0, "click");
        devicesNavigationButton.Click += (_, _) => ShowAdminSection(1, "click");
        scenarioNavigationButton.Click += (_, _) => ShowAdminSection(2, "click");
        resultsNavigationButton.Click += (_, _) => ShowAdminSection(3, "click");

        void SyncDialogScrollBar()
        {
            if (switchingAdminSection)
            {
                return;
            }

            contentPanel.Width = Math.Max(UiScale(320), contentHost.ClientSize.Width - dialogScroll.Width - UiScale(4));
            if (contentPanel.Left != 0)
            {
                contentPanel.Left = 0;
            }

            layout.PerformLayout();
            contentPanel.Height = layout.PreferredSize.Height;

            UpdateDialogScrollLayout();
            UpdateDialogHostLayout();

            int contentHeight = contentPanel.Height;
            int viewportHeight = contentHost.ClientSize.Height;
            int maxOffset = Math.Max(0, contentHeight - viewportHeight);
            int offset = Math.Max(0, Math.Min(maxOffset, -contentPanel.Top));
            bool needsScroll = contentHeight > viewportHeight + 1;

            dialogScroll.Visible = needsScroll;
            contentPanel.Location = needsScroll
                ? new Point(0, -offset)
                : new Point(0, 0);

            syncingDialogScroll = true;
            dialogScroll.Maximum = Math.Max(contentHeight, 1);
            dialogScroll.ViewportSize = Math.Max(viewportHeight, 1);
            dialogScroll.Value = needsScroll ? offset : 0;
            syncingDialogScroll = false;
            SetActiveNavigation(GetNavigationButton(activeAdminSection));
        }

        dialogScroll.ValueChanged += (_, _) =>
        {
            if (syncingDialogScroll)
            {
                return;
            }

            SetDialogScrollOffset(dialogScroll.Value);
        };

        layout.SizeChanged += (_, _) => SyncDialogScrollBar();
        contentHost.SizeChanged += (_, _) => SyncDialogScrollBar();
        contentHost.MouseEnter += (_, _) => contentHost.Focus();
        contentHost.MouseWheel += (_, e) =>
        {
            if (!dialogScroll.Visible)
            {
                return;
            }

            int delta = e.Delta > 0 ? -dialogScroll.SmallChange : dialogScroll.SmallChange;
            dialogScroll.Value += delta;
            if (e is HandledMouseEventArgs handled)
            {
                handled.Handled = true;
            }
        };

        contentHost.Controls.Add(contentPanel);
        contentHost.Controls.Add(dialogScroll);

        dialog.Controls.Add(contentHost);
        dialog.Controls.Add(navigation);
        dialog.Controls.Add(buttons);
        dialog.AcceptButton = null;
        dialog.CancelButton = closeButton;
        dialog.KeyPreview = true;
        dialog.KeyDown += (_, e) =>
        {
            if (!e.Control)
            {
                return;
            }

            switch (e.KeyCode)
            {
                case Keys.D1:
                case Keys.NumPad1:
                    ShowAdminSection(0, "keyboard");
                    break;
                case Keys.D2:
                case Keys.NumPad2:
                    ShowAdminSection(1, "keyboard");
                    break;
                case Keys.D3:
                case Keys.NumPad3:
                    ShowAdminSection(2, "keyboard");
                    break;
                case Keys.D4:
                case Keys.NumPad4:
                    ShowAdminSection(3, "keyboard");
                    break;
                default:
                    return;
            }

            e.Handled = true;
            e.SuppressKeyPress = true;
        };
        dialog.HandleCreated += (_, _) => ApplyTitleBarTheme(dialog);

        void ValidateTestAdminLayout(string reason)
        {
            List<string> layoutIssues = [];
            int inspectedControls = 0;
            int buttonCount = 0;

            void InspectControls(Control parent)
            {
                Control[] visibleChildren = parent.Controls.Cast<Control>()
                    .Where(control => control.Visible)
                    .ToArray();
                for (int first = 0; first < visibleChildren.Length; first++)
                {
                    for (int second = first + 1; second < visibleChildren.Length; second++)
                    {
                        Control left = visibleChildren[first];
                        Control right = visibleChildren[second];
                        if (left.Bounds.IntersectsWith(right.Bounds))
                        {
                            layoutIssues.Add(
                                $"overlap:{left.GetType().Name}/{GetSourceControlText(left)}+" +
                                $"{right.GetType().Name}/{GetSourceControlText(right)}");
                        }
                    }
                }

                foreach (Control child in visibleChildren)
                {
                    inspectedControls++;
                    int tolerance = UiScale(2);
                    bool scrollCanvas = ReferenceEquals(child, contentPanel);
                    if (!scrollCanvas
                        && (child.Left < -tolerance
                            || child.Top < -tolerance
                            || child.Right > parent.ClientSize.Width + tolerance
                            || child.Bottom > parent.ClientSize.Height + tolerance))
                    {
                        layoutIssues.Add(
                            $"bounds:{child.GetType().Name}/{GetSourceControlText(child)}=" +
                            $"{child.Left},{child.Top},{child.Width}x{child.Height}>" +
                            $"{parent.GetType().Name}:{parent.ClientSize.Width}x{parent.ClientSize.Height}");
                    }

                    if (child is Button button)
                    {
                        buttonCount++;
                        int textWidth = TextRenderer.MeasureText(
                            button.Text,
                            button.Font,
                            Size.Empty,
                            TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix | TextFormatFlags.SingleLine).Width;
                        if (textWidth + UiScale(24) > button.ClientSize.Width)
                        {
                            layoutIssues.Add($"button-text:{GetSourceControlText(button)}={button.ClientSize.Width}<{textWidth + UiScale(24)}");
                        }
                    }
                    else if (child is Label label
                             && !label.AutoSize
                             && child is not ImodMapTextBox
                             && child is not NicItrTableLabel
                             && !string.IsNullOrWhiteSpace(label.Text)
                             && !label.Text.Contains('\n'))
                    {
                        int textWidth = TextRenderer.MeasureText(
                            label.Text,
                            label.Font,
                            Size.Empty,
                            TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix | TextFormatFlags.SingleLine).Width;
                        if (textWidth > label.ClientSize.Width + tolerance)
                        {
                            layoutIssues.Add($"label-text:{GetSourceControlText(label)}={label.ClientSize.Width}<{textWidth}");
                        }
                    }

                    if (child is ListBox list && list.HorizontalScrollbar)
                    {
                        layoutIssues.Add("native-horizontal-list-scrollbar");
                    }

                    InspectControls(child);
                }
            }

            dialog.PerformLayout();
            SyncDialogScrollBar();
            InspectControls(dialog);
            int operationGap = scenarioImodCheck.Left - scenarioOperationCombo.Right;
            if (operationGap < UiScale(12))
            {
                layoutIssues.Add($"scenario-operation-gap={operationGap}");
            }

            WriteLog(layoutIssues.Count == 0
                ? $"TEST.ADMIN.LAYOUT: status=PASS reason={reason} language={UiLanguage.Current} controls={inspectedControls} buttons={buttonCount}"
                : $"TEST.ADMIN.LAYOUT: status=FAIL reason={reason} language={UiLanguage.Current} issues=\"{SanitizeLogValue(string.Join("; ", layoutIssues.Take(30)))}\" total={layoutIssues.Count}");
        }

        dialog.Shown += (_, _) =>
        {
            ApplyTitleBarTheme(dialog);
            SyncDialogScrollBar();
            dialog.BeginInvoke(new Action(() =>
            {
                adminToolTip.Hide(dialog);
                adminToolTip.Active = true;

                ValidateTestAdminLayout("shown");
            }));

            if (string.Equals(
                    Environment.GetEnvironmentVariable("DEVICE_TWEAKER_QA_SANDBOX"),
                    "1",
                    StringComparison.Ordinal))
            {
                dialog.BeginInvoke(new Action(() =>
                {
                    try
                    {
                        // Exercise a long multi-role IMOD map and the Wi-Fi
                        // preservation contract first, then leave the compact
                        // Ryzen preset loaded for the visible QA pass.
                        systemPresetCombo.SelectedItem = SystemPresetIntel14900K;
                        WriteLog("TEST.QA.SANDBOX: multi-role IMOD layout preset completed");
                        bool matrixPassed = RunCoreScenarioMatrix(showDialog: false);
                        WriteLog($"TEST.QA.MATRIX: status={(matrixPassed ? "PASS" : "FAIL")}");
                        bool localizationPassed = UiLanguage.ValidateRussianTranslationContract(out string localizationError);
                        WriteLog(
                            localizationPassed
                                ? "TEST.QA.LOCALIZATION: status=PASS scope=imod-structured-text"
                                : $"TEST.QA.LOCALIZATION.FAIL: status=FAIL detail=\"{SanitizeLogValue(localizationError)}\"");
                        bool startupPassed = ValidateImodStartupPersistenceContract(out string startupError);
                        WriteLog(
                            startupPassed
                                ? "TEST.QA.IMOD.STARTUP: status=PASS mechanism=hkcu-run-explicit-powershell"
                                : $"TEST.QA.IMOD.STARTUP.FAIL: status=FAIL detail=\"{SanitizeLogValue(startupError)}\"");
                        bool backupPersistencePassed = ValidateImodBackupPersistenceContract(out string backupPersistenceError);
                        WriteLog(
                            backupPersistencePassed
                                ? "TEST.QA.BACKUP.IMOD: status=PASS scope=script-and-hkcu-run"
                                : $"TEST.QA.BACKUP.IMOD.FAIL: status=FAIL detail=\"{SanitizeLogValue(backupPersistenceError)}\"");
                        systemPresetCombo.SelectedItem = SystemPresetLaptopIntel285HX;
                        WriteLog("TEST.QA.SANDBOX: Wi-Fi preservation preset completed");
                        systemPresetCombo.SelectedItem = SystemPresetLaptopRyzen8940HX;
                        ValidateQaAffinity8940HxLayout();
                        WriteLog("TEST.QA.SANDBOX: 8940HX Dual-CCD affinity preset completed");
                        systemPresetCombo.SelectedItem = SystemPresetMultiControllerInput;
                        ValidateQaMultiControllerAffinity();
                        WriteLog("TEST.QA.SANDBOX: Multi-controller affinity preset completed");
                        string? targetFinalPreset = Environment.GetEnvironmentVariable("DEVICE_TWEAKER_QA_FINAL_PRESET");
                        string selectedFinalPreset = targetFinalPreset switch
                        {
                            "Intel14900K" => SystemPresetIntel14900K,
                            "Intel14600K" => SystemPresetIntel14600K,
                            "Intel285K" => SystemPresetIntel285K5090,
                            "Ryzen9950X3D" => SystemPresetRyzen9950X3DNdis,
                            "Ryzen9950X3DNetCx" => SystemPresetRyzen9950X3DNetCx,
                            "Ryzen9950X" => SystemPresetRyzen9950X,
                            "Ryzen7800X3D" => SystemPresetRyzen7800X3D,
                            _ => SystemPresetRyzen3500X
                        };
                        systemPresetCombo.SelectedItem = selectedFinalPreset;
                        if (!string.Equals(
                                systemPresetCombo.SelectedItem?.ToString(),
                                selectedFinalPreset,
                                StringComparison.Ordinal))
                        {
                            WriteLog($"TEST.QA.SANDBOX: failed to select {selectedFinalPreset} regression preset");
                            return;
                        }

                        // SelectedIndexChanged loads the preset; force if still empty.
                        if (_testDevices.Count == 0)
                        {
                            _ = LoadSystemPreset();
                        }

                        suppressTestDeviceToggle = true;
                        _testDevicesEnabled = true;
                        _testDevicesOnly = true;
                        _testAutoDryRun = true;
                        enableTestDevicesCheck.Checked = true;
                        testDevicesOnlyCheck.Checked = true;
                        testDevicesOnlyCheck.Enabled = true;
                        dryRunAutoCheck.Checked = true;
                        dryRunAutoCheck.AutoCheck = false;
                        dryRunAutoCheck.Enabled = true;
                        suppressTestDeviceToggle = false;

                        _initialDeviceViewportHeightAdjusted = false;
                        RefreshBlocks();
                        string? envFilter = Environment.GetEnvironmentVariable("DEVICE_TWEAKER_CATEGORY_FILTER");
                        if (!string.IsNullOrWhiteSpace(envFilter))
                        {
                            SetCategoryFilter(envFilter);
                        }
                        NotifySandboxModeChanged("qa-sandbox");
                        WriteLog(
                            $"TEST.QA.SANDBOX: ready testDevices={_testDevices.Count} only={_testDevicesOnly} dryRun={_testAutoDryRun}");
                        ShowAdminSection(1, "qa-lists", validate: false);
                        int listsTargetOffset = Math.Max(0, GetDialogContentY(testListLabel) - UiScale(18));
                        SetDialogScrollOffset(listsTargetOffset);
                        SyncDialogScrollBar();
                        ValidateTestAdminLayout("lists-ready");
                        WriteLog($"TEST.QA.LISTS.VIEW: target={listsTargetOffset} actual={Math.Max(0, -contentPanel.Top)}");

                        // Leave the device lists visible long enough for the
                        // external QA harness to capture their custom scrollbars,
                        // then move to Scenario Lab for its independent capture.
                        dialog.BeginInvoke(new Action(() =>
                        {
                            System.Windows.Forms.Timer scenarioTimer = new() { Interval = 8000 };
                            scenarioTimer.Tick += (_, _) =>
                            {
                                scenarioTimer.Stop();
                                scenarioTimer.Dispose();
                                ShowAdminSection(2, "qa-scenario", validate: false);
                                SyncDialogScrollBar();
                                int targetOffset = Math.Max(0, scenarioLabLabel.Top - UiScale(18));
                                SetDialogScrollOffset(targetOffset);
                                SyncDialogScrollBar();
                                ValidateTestAdminLayout("scenario-ready");
                                WriteLog($"TEST.QA.SCENARIO.VIEW: target={targetOffset} actual={Math.Max(0, -contentPanel.Top)}");
                            };
                            scenarioTimer.Start();
                        }));
                    }
                    catch (Exception ex)
                    {
                        WriteLog($"TEST.QA.SANDBOX: failed: {ex.Message}");
                    }
                }));
            }
        };

        // This dialog is constructed from explicit pixel coordinates. Scale
        // the complete tree once, after construction, so every test control
        // follows the same per-monitor DPI factor as the main window.
        ShowAdminSection(0, "initial", validate: false);
        if (Math.Abs(_uiScale - 1f) > 0.01f)
        {
            dialog.Scale(new SizeF(_uiScale, _uiScale));
        }

        dialog.PerformLayout();
        int desiredHeight = layout.PreferredSize.Height + navigation.Height + buttons.Height + UiScale(12);
        int maxHeight = Math.Min(UiScale(760), Screen.FromControl(dialog).WorkingArea.Height - UiScale(140));
        int targetHeight = Math.Min(maxHeight, Math.Max(UiScale(520), desiredHeight));
        dialog.ClientSize = new Size(dialog.ClientSize.Width, targetHeight);
        syncDialogScroll = SyncDialogScrollBar;
        SyncSandboxDryRunLock("dialog-open");

        WireThemedTitleBar(dialog);
        ShowDialogDimmed(dialog);
    }

    private Button NewDialogButton(string text)
    {
        Button btn = new()
        {
            Text = text,
            Size = new Size(150, 32),
            Margin = new Padding(0, 0, 10, 0),
            FlatStyle = FlatStyle.Flat,
            Font = _buttonFont,
            UseVisualStyleBackColor = false,
            Cursor = Cursors.Hand,
            BackColor = _bgPanel,
            ForeColor = _fgMain,
        };
        btn.FlatAppearance.BorderSize = 1;
        btn.FlatAppearance.BorderColor = Color.FromArgb(80, 80, 86);
        btn.MouseEnter += (_, _) =>
        {
            btn.BackColor = _accent;
            btn.ForeColor = Color.FromArgb(15, 15, 15);
        };
        btn.MouseLeave += (_, _) =>
        {
            btn.BackColor = _bgPanel;
            btn.ForeColor = _fgMain;
        };
        return btn;
    }

    private Label NewDialogLabel(string text)
    {
        return new Label
        {
            Text = text,
            AutoSize = true,
            ForeColor = _mutedText,
            TextAlign = ContentAlignment.MiddleLeft,
            Anchor = AnchorStyles.Left,
            Margin = new Padding(0, 4, 12, 6),
        };
    }

    private Label NewInlineLabel(string text)
    {
        return new Label
        {
            Text = text,
            AutoSize = true,
            ForeColor = _mutedText,
            TextAlign = ContentAlignment.MiddleLeft,
            Margin = new Padding(10, 4, 6, 0),
        };
    }

    private Label NewHintLabel(string text)
    {
        return new Label
        {
            Text = text,
            AutoSize = true,
            ForeColor = _mutedText,
            Margin = new Padding(0, 4, 0, 8),
            MaximumSize = new Size(780, 0),
        };
    }

    private Panel NewBoxPanel()
    {
        Panel panel = new()
        {
            BackColor = Color.FromArgb(16, 16, 19),
            ForeColor = _fgMain,
            Padding = new Padding(10, 8, 10, 8),
            Margin = new Padding(0),
        };
        panel.Paint += (_, e) =>
        {
            Rectangle rect = panel.ClientRectangle;
            rect.Width -= 1;
            rect.Height -= 1;
            using Pen pen = new(Color.FromArgb(70, 70, 76));
            e.Graphics.DrawRectangle(pen, rect);
        };
        return panel;
    }

    private NumericUpDown NewNumericUpDown(int min, int max, int value)
    {
        int clamped = Math.Min(max, Math.Max(min, value));
        return new ThemedNumericUpDown
        {
            Minimum = min,
            Maximum = max,
            Value = clamped,
            TextAlign = HorizontalAlignment.Center,
            BackColor = Color.FromArgb(18, 18, 22),
            ForeColor = _fgMain,
            BorderStyle = BorderStyle.FixedSingle,
            Size = new Size(120, 24),
            Margin = new Padding(0, 0, 0, 6),
            ButtonBackColor = Color.FromArgb(14, 14, 17),
            ButtonHoverColor = Color.FromArgb(40, 40, 48),
            ArrowColor = _fgMain,
        };
    }

    private ThemedDropDownPicker NewDialogCombo(int width)
    {
        ThemedDropDownPicker picker = new()
        {
            BackColor = Color.FromArgb(18, 18, 22),
            ForeColor = _fgMain,
            Size = new Size(width, 26),
            Margin = new Padding(0, 0, 0, 6),
            DropDownWidth = width,
        };
        StyleDarkDropDownPicker(picker);
        return picker;
    }

    private FlowLayoutPanel NewRowFlowPanel()
    {
        return new FlowLayoutPanel
        {
            AutoSize = true,
            WrapContents = false,
            FlowDirection = FlowDirection.LeftToRight,
            Margin = new Padding(0, 0, 0, 6),
            Padding = Padding.Empty,
        };
    }

    private TextBox NewDialogTextBox(int width)
    {
        TextBox textBox = new()
        {
            BackColor = Color.FromArgb(18, 18, 22),
            ForeColor = _fgMain,
            BorderStyle = BorderStyle.FixedSingle,
            Size = new Size(width, 24),
            Margin = new Padding(0, 0, 0, 6),
        };
        textBox.TextChanged += (_, _) =>
        {
            if (textBox.Focused || textBox.TextLength == 0 || !textBox.IsHandleCreated)
            {
                return;
            }

            textBox.BeginInvoke(new Action(() =>
            {
                if (!textBox.IsDisposed && !textBox.Focused && textBox.TextLength > 0)
                {
                    textBox.Select(0, 0);
                    textBox.ScrollToCaret();
                }
            }));
        };
        return textBox;
    }

    private DeviceInfo CreateTestDevice(
        DeviceKind kind,
        string name,
        string pnpIdOverride,
        string usbRoles,
        string audioEndpoints,
        string storageTag,
        bool wifi,
        bool usbIsXhci,
        bool usbHasDevices,
        bool integratedGpu,
        int? testIrqCount = null,
        string testMsiStatus = "Auto",
        string? usbSelectiveSuspend = "off",
        string? nicPowerSaving = "on")
    {
        _testDeviceSequence++;
        int seq = _testDeviceSequence;
        string id = string.IsNullOrWhiteSpace(pnpIdOverride)
            ? $"TEST\\{kind}\\{seq:D4}"
            : pnpIdOverride.Trim().Replace('/', '\\');
        string displayName = string.IsNullOrWhiteSpace(name) ? $"Test {kind} {seq:D2}" : name.Trim();
        string className = kind switch
        {
            DeviceKind.USB => "USB",
            DeviceKind.GPU => "Display",
            DeviceKind.AUDIO => "MEDIA",
            DeviceKind.NET_NDIS => "Net",
            DeviceKind.NET_CX => "Net",
            DeviceKind.STOR => "SCSIAdapter",
            _ => "System",
        };

        bool isUsb = kind == DeviceKind.USB;
        bool isAudio = kind == DeviceKind.AUDIO;
        bool isNet = kind is DeviceKind.NET_NDIS or DeviceKind.NET_CX;
        bool isStor = kind == DeviceKind.STOR;
        string derivedPolling = isUsb ? DeriveTestUsbPollingRates(usbRoles) : string.Empty;
        string? suspend = isUsb
            ? (usbHasDevices ? usbSelectiveSuspend : null)
            : null;
        string? nicPower = isNet && !wifi
            ? (string.Equals(nicPowerSaving, "off", StringComparison.OrdinalIgnoreCase) ? "off" : "on")
            : null;

        DeviceInfo device = new()
        {
            Name = displayName,
            InstanceId = id,
            Class = className,
            RegBase = $@"SYSTEM\CurrentControlSet\Enum\{id}",
            Kind = kind,
            UsbRoles = isUsb ? usbRoles : string.Empty,
            UsbPollingRates = derivedPolling,
            AudioEndpoints = isAudio ? audioEndpoints : string.Empty,
            StorageTag = isStor ? storageTag : string.Empty,
            IsIntegratedGpu = kind == DeviceKind.GPU && integratedGpu,
            Wifi = isNet && wifi,
            UsbIsXhci = isUsb && usbIsXhci,
            UsbHasDevices = isUsb && usbHasDevices,
            UsbChipPath = isUsb ? UsbChipPath.Classify(id) : null,
            UsbSelectiveSuspend = suspend,
            NicPowerSaving = nicPower,
            IsTestDevice = true,
            TestIrqCount = testIrqCount,
            TestMsiStatus = string.IsNullOrWhiteSpace(testMsiStatus) ? "Auto" : testMsiStatus,
        };
        device.TestState = new TestDeviceState
        {
            MsiEnabled = !string.Equals(testMsiStatus, "Disabled", StringComparison.OrdinalIgnoreCase),
            PowerSavingEnabled = wifi ? null : isUsb
                ? !string.Equals(suspend, "off", StringComparison.OrdinalIgnoreCase)
                : isNet
                    ? !string.Equals(nicPower, "off", StringComparison.OrdinalIgnoreCase)
                    : null,
            ImodValue = isUsb ? "0x0,0xC8,0xC8,0xC8,0xC8,0xC8,0xC8,0xC8" : "0x0",
            NicItrValue = isNet && id.Contains("VEN_10EC&DEV_8125", StringComparison.OrdinalIgnoreCase)
                ? "0x7E600958,0x0,0x7E600938,0x0"
                : "default",
        };
        return device;
    }

    private void WarnTestUsbChipMismatch(string instanceId, UsbChipPathInfo chip)
    {
        bool nameCpu = instanceId.Contains("_CPU_", StringComparison.OrdinalIgnoreCase)
            || instanceId.Contains("_TB4_", StringComparison.OrdinalIgnoreCase)
            || instanceId.Contains("MTL_TB4", StringComparison.OrdinalIgnoreCase);
        bool nameChipset = instanceId.Contains("CHIPSET", StringComparison.OrdinalIgnoreCase)
            || instanceId.Contains("_PCH_", StringComparison.OrdinalIgnoreCase)
            || instanceId.Contains("Z790_XHCI", StringComparison.OrdinalIgnoreCase)
            || instanceId.Contains("Z890_XHCI", StringComparison.OrdinalIgnoreCase);
        bool nameAddon = instanceId.Contains("ASMEDIA", StringComparison.OrdinalIgnoreCase);

        if (nameAddon && chip.Origin != UsbChipOrigin.Addon)
        {
            WriteLog($"TEST.USB.CHIP.WARN: expected add-in for {instanceId} got {chip.CompactTag}/{chip.Origin}");
        }
        else if (nameCpu && !nameChipset && chip.BaseChipCount != 0)
        {
            WriteLog($"TEST.USB.CHIP.WARN: expected CHIP 0 for {instanceId} got {chip.CompactTag}/{chip.Origin}");
        }
        else if (nameChipset && !nameCpu && chip.BaseChipCount == 0)
        {
            WriteLog($"TEST.USB.CHIP.WARN: expected CHIP 1 for {instanceId} got {chip.CompactTag}/{chip.Origin}");
        }
    }

    private static string DeriveTestUsbPollingRates(string usbRoles)
    {
        if (string.IsNullOrWhiteSpace(usbRoles))
        {
            return string.Empty;
        }

        List<string> parts = [];
        foreach (string raw in usbRoles.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            bool isMouse = raw.Contains("Mouse", StringComparison.OrdinalIgnoreCase);
            bool isKeyboard = raw.Contains("Keyboard", StringComparison.OrdinalIgnoreCase);
            if (!isMouse && !isKeyboard)
            {
                continue;
            }

            string rate = raw.Contains("8K", StringComparison.OrdinalIgnoreCase)
                ? "8K"
                : raw.Contains("1K", StringComparison.OrdinalIgnoreCase)
                    ? "1K"
                    : "n/a";
            parts.Add(isMouse ? $"Mouse: {rate}" : $"Keyboard: {rate}");
        }

        return string.Join(", ", parts);
    }

    private static string FormatTestDeviceLabel(DeviceInfo device)
    {
        string label = $"{device.Kind}: {device.Name}";
        if (device.Kind == DeviceKind.USB && !string.IsNullOrWhiteSpace(device.UsbRoles))
        {
            string chip = device.UsbChipPath is { } path ? $" | {path.CompactTag}" : string.Empty;
            label += $" [{device.UsbRoles}{chip}]";
        }
        else if (device.Kind == DeviceKind.USB && device.UsbChipPath is { } path)
        {
            label += $" [{path.CompactTag}]";
        }
        else if (device.Kind == DeviceKind.GPU && device.IsIntegratedGpu)
        {
            label += " [iGPU]";
        }
        else if (device.Kind == DeviceKind.AUDIO && !string.IsNullOrWhiteSpace(device.AudioEndpoints))
        {
            label += $" [{device.AudioEndpoints}]";
        }
        else if (device.Kind == DeviceKind.STOR && !string.IsNullOrWhiteSpace(device.StorageTag))
        {
            label += $" [{device.StorageTag}]";
        }
        else if ((device.Kind == DeviceKind.NET_NDIS || device.Kind == DeviceKind.NET_CX) && device.Wifi)
        {
            label += " [WiFi]";
        }

        if (device.Kind == DeviceKind.USB
            && string.Equals(device.UsbSelectiveSuspend, "off", StringComparison.OrdinalIgnoreCase))
        {
            label += " [PowerSaving off]";
        }

        if ((device.Kind == DeviceKind.NET_NDIS || device.Kind == DeviceKind.NET_CX)
            && !device.Wifi
            && string.Equals(device.NicPowerSaving, "off", StringComparison.OrdinalIgnoreCase))
        {
            label += " [PowerSaving off]";
        }

        if (device.TestIrqCount.HasValue || !device.TestMsiStatus.Equals("Auto", StringComparison.OrdinalIgnoreCase))
        {
            string irq = device.TestIrqCount.HasValue ? device.TestIrqCount.Value.ToString(CultureInfo.InvariantCulture) : "Auto";
            label += $" [IRQ {irq}, MSI {device.TestMsiStatus}]";
        }

        return label;
    }

    private static HashSet<int> ParseIndexSet(string text)
    {
        if (TryParseIndexList(text, int.MaxValue, out List<int> indices, out _))
        {
            return indices.ToHashSet();
        }

        return [];
    }

    private void ApplyTestCpuConfig(TestCpuConfig config)
    {
        if (config.LogicalCount <= 0)
        {
            return;
        }

        List<CpuLpInfo> entries = [];
        for (int lp = 0; lp < config.LogicalCount; lp++)
        {
            int core = config.CoreMap.TryGetValue(lp, out int coreIndex) ? coreIndex : lp;
            // Match the native Windows meaning: higher EfficiencyClass means
            // more performance and less efficiency.
            int eff = config.ECoreLps.Contains(lp) ? 0 : 1;
            int llc = 0;
            if (config.CcxMap is not null && config.CcxMap.TryGetValue(lp, out int ccx))
            {
                llc = ccx;
            }
            else if (config.CcdMap is not null && config.CcdMap.TryGetValue(lp, out int ccd))
            {
                llc = ccd;
            }

            entries.Add(new CpuLpInfo(
                Group: 0,
                LP: lp,
                Core: core,
                LLC: llc,
                NUMA: 0,
                EffClass: eff,
                LocalIndex: lp,
                CpuSetId: lp));
        }

        CpuTopology topo = new(entries.OrderBy(x => x.LP).ToList());
        Dictionary<int, int> ccdMap = config.CcdMap ?? BuildCcdMap(topo);
        Dictionary<int, int> ccxMap = config.CcxMap ?? BuildCcxMap(topo);

        _cpuInfo = new CpuInfo
        {
            Topology = topo,
            CcdMap = ccdMap,
            CcxMap = ccxMap,
        };
        UpdateEfficiencyClassMap(topo);
        _cppcRatings.Clear();
        _cppcRanks.Clear();
        _cppcEnabled = false;
        if (config.CppcRatings.Count > 0)
        {
            Dictionary<int, int> collected = config.CppcRatings
                .Where(kvp => kvp.Key >= 0 && kvp.Key < topo.Logical)
                .ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
            List<int> uniqueRatings = collected.Values.Distinct().OrderByDescending(v => v).ToList();
            if (collected.Count != topo.Logical)
            {
                WriteLog($"TESTCPU.CPPC: disabled, incomplete test data parsed={collected.Count} required={topo.Logical}");
            }
            else if (uniqueRatings.Count > 1)
            {
                int rank = 1;
                foreach (int rating in uniqueRatings)
                {
                    foreach (KeyValuePair<int, int> item in collected.Where(kvp => kvp.Value == rating).OrderBy(kvp => kvp.Key))
                    {
                        _cppcRatings[item.Key] = item.Value;
                        _cppcRanks[item.Key] = rank;
                    }

                    rank++;
                }

                _cppcEnabled = _cppcRanks.Count > 0;
                string ratingsText = string.Join(
                    " ",
                    _cppcRatings
                        .OrderBy(kvp => kvp.Key)
                        .Select(kvp => $"CPU{kvp.Key}=R{kvp.Value}/#{_cppcRanks[kvp.Key]}"));
                WriteLog($"TESTCPU.CPPC: enabled count={_cppcRanks.Count} {ratingsText}");
            }
            else
            {
                WriteLog($"TESTCPU.CPPC: disabled, all test ratings share rating={uniqueRatings.FirstOrDefault()} count={collected.Count}");
            }
        }
        else
        {
            WriteLog("TESTCPU.CPPC: disabled (not specified)");
        }

        _cpuGroupCount = 1;
        _cpuLpByIndex.Clear();
        _cpuSetIdByIndex.Clear();
        _cpuIndexByCpuSetId.Clear();
        foreach (CpuLpInfo lp in topo.LPs)
        {
            _cpuLpByIndex[lp.LP] = lp;
            int cpuSetId = lp.CpuSetId >= 0 ? lp.CpuSetId : lp.LP;
            _cpuSetIdByIndex[lp.LP] = cpuSetId;
            _cpuIndexByCpuSetId.TryAdd(cpuSetId, lp.LP);
        }

        _maxLogical = Math.Min(topo.Logical, MaxAffinityBits);
        _grpHeight = UiScale(280);

        string smtPrefix = config.UseHyperThreadingLabel ? "Hyper-Threading" : "SMT";
        _smtText = config.SmtEnabled
            ? $"{smtPrefix}: Enabled (Test)"
            : $"{smtPrefix}: Disabled (Test)";
        _testCpuName = config.CpuName?.Trim() ?? string.Empty;
        _cpuHeaderText = BuildTestCpuHeaderText(_testCpuName, topo.Logical);

        _testCpuActive = true;

        string eText = FormatIndexList(config.ECoreLps);
        string coreGroups = GetCurrentCoreGroupsText();
        string ccdGroups = GetCurrentCcdGroupsText();
        string ccxGroups = GetCurrentCcxGroupsText();
        WriteLog($"TESTCPU: enabled logical={topo.Logical} eCores=[{eText}] cores=[{coreGroups}] ccd=[{ccdGroups}] ccx=[{ccxGroups}]");

        UpdateCpuHeaderUi();
        _initialDeviceViewportHeightAdjusted = false;
        RefreshBlocks();
    }

    private void DisableTestCpuMode()
    {
        if (!_testCpuActive)
        {
            InitializeCpu();
            UpdateCpuHeaderUi();
            _initialDeviceViewportHeightAdjusted = false;
            RefreshBlocks();
            return;
        }

        _testCpuActive = false;
        InitializeCpu();
        UpdateCpuHeaderUi();
        _initialDeviceViewportHeightAdjusted = false;
        RefreshBlocks();
        WriteLog("TESTCPU: disabled (restored real CPU)");
    }

    private int GetCurrentLogicalCount()
    {
        return _cpuInfo?.Topology.Logical ?? Environment.ProcessorCount;
    }

    private static string BuildTestCpuHeaderText(string? cpuName, int logical)
    {
        string name = cpuName?.Trim() ?? string.Empty;
        if (name.Length == 0)
        {
            return $"CPU: Test Mode ({logical} LP)";
        }

        if (name.StartsWith("CPU:", StringComparison.OrdinalIgnoreCase))
        {
            name = name[4..].Trim();
            if (name.Length == 0)
            {
                return $"CPU: Test Mode ({logical} LP)";
            }
        }

        return $"CPU: {name}";
    }

    private string GetCurrentECoreText()
    {
        if (_cpuInfo is null)
        {
            return string.Empty;
        }

        List<int> eLps = _cpuInfo.Topology.LPs
            .Where(lp => IsEfficiencyCore(lp))
            .Select(lp => lp.LP)
            .OrderBy(x => x)
            .ToList();
        return FormatIndexList(eLps);
    }

    private string GetCurrentCoreGroupsText()
    {
        if (_cpuInfo is null)
        {
            return string.Empty;
        }

        List<List<int>> groups = _cpuInfo.Topology.ByCore
            .OrderBy(kvp => kvp.Key)
            .Select(kvp => kvp.Value.Select(lp => lp.LP).OrderBy(x => x).ToList())
            .ToList();
        return FormatGroups(groups);
    }

    private string GetCurrentCcdGroupsText()
    {
        if (_cpuInfo is null)
        {
            return string.Empty;
        }

        List<List<int>> groups = _cpuInfo.CcdMap
            .GroupBy(kvp => kvp.Value)
            .OrderBy(g => g.Key)
            .Select(g => g.Select(kvp => kvp.Key).OrderBy(x => x).ToList())
            .ToList();
        return FormatGroups(groups);
    }

    private string GetCurrentCcxGroupsText()
    {
        if (_cpuInfo is null)
        {
            return string.Empty;
        }

        List<List<int>> groups = _cpuInfo.CcxMap
            .GroupBy(kvp => kvp.Value)
            .OrderBy(g => g.Key)
            .Select(g => g.Select(kvp => kvp.Key).OrderBy(x => x).ToList())
            .ToList();
        return FormatGroups(groups);
    }

    private static bool TryParseIndexList(string input, int maxExclusive, out List<int> indices, out string error)
    {
        indices = [];
        error = string.Empty;
        if (string.IsNullOrWhiteSpace(input))
        {
            return true;
        }

        string[] tokens = input.Split([',', ';', ' ', '\t', '\r', '\n'], StringSplitOptions.RemoveEmptyEntries);
        HashSet<int> set = [];

        foreach (string token in tokens)
        {
            string trimmed = token.Trim();
            if (trimmed.Length == 0)
            {
                continue;
            }

            int dash = trimmed.IndexOf('-', StringComparison.Ordinal);
            if (dash >= 0)
            {
                string left = trimmed[..dash];
                string right = trimmed[(dash + 1)..];
                if (!int.TryParse(left, NumberStyles.Integer, CultureInfo.InvariantCulture, out int start)
                    || !int.TryParse(right, NumberStyles.Integer, CultureInfo.InvariantCulture, out int end))
                {
                    error = $"Invalid range \"{trimmed}\".";
                    return false;
                }

                if (start > end)
                {
                    (start, end) = (end, start);
                }

                for (int i = start; i <= end; i++)
                {
                    if (i < 0 || i >= maxExclusive)
                    {
                        error = $"LP {i} is out of range (0-{maxExclusive - 1}).";
                        return false;
                    }
                    set.Add(i);
                }
            }
            else
            {
                if (!int.TryParse(trimmed, NumberStyles.Integer, CultureInfo.InvariantCulture, out int value))
                {
                    error = $"Invalid index \"{trimmed}\".";
                    return false;
                }

                if (value < 0 || value >= maxExclusive)
                {
                    error = $"LP {value} is out of range (0-{maxExclusive - 1}).";
                    return false;
                }

                set.Add(value);
            }
        }

        indices = set.OrderBy(x => x).ToList();
        return true;
    }

    private static bool TryParseGroups(string input, int maxExclusive, out List<List<int>> groups, out string error)
    {
        groups = [];
        error = string.Empty;

        if (string.IsNullOrWhiteSpace(input))
        {
            return true;
        }

        string[] groupTokens = input.Split('|', StringSplitOptions.RemoveEmptyEntries);
        foreach (string group in groupTokens)
        {
            if (!TryParseIndexList(group, maxExclusive, out List<int> indices, out error))
            {
                error = $"Group \"{group.Trim()}\": {error}";
                return false;
            }

            if (indices.Count == 0)
            {
                error = "Empty group is not allowed.";
                return false;
            }

            groups.Add(indices);
        }

        HashSet<int> seen = [];
        foreach (List<int> group in groups)
        {
            foreach (int lp in group)
            {
                if (!seen.Add(lp))
                {
                    error = $"LP {lp} appears in multiple groups.";
                    return false;
                }
            }
        }

        return true;
    }

    private static string FormatIndexList(IEnumerable<int> indices)
    {
        if (indices is null)
        {
            return string.Empty;
        }

        List<int> list = indices.Distinct().OrderBy(x => x).ToList();
        if (list.Count == 0)
        {
            return string.Empty;
        }

        StringBuilder sb = new();
        int start = list[0];
        int prev = list[0];

        void AppendRange(int rangeStart, int rangeEnd)
        {
            if (sb.Length > 0)
            {
                sb.Append(',');
            }

            if (rangeStart == rangeEnd)
            {
                sb.Append(rangeStart);
            }
            else
            {
                sb.Append(rangeStart).Append('-').Append(rangeEnd);
            }
        }

        for (int i = 1; i < list.Count; i++)
        {
            int current = list[i];
            if (current == prev + 1)
            {
                prev = current;
                continue;
            }

            AppendRange(start, prev);
            start = current;
            prev = current;
        }

        AppendRange(start, prev);
        return sb.ToString();
    }

    private static string FormatGroups(IEnumerable<List<int>> groups)
    {
        if (groups is null)
        {
            return string.Empty;
        }

        List<string> parts = groups
            .Select(FormatIndexList)
            .Where(s => !string.IsNullOrWhiteSpace(s))
            .ToList();

        return string.Join("|", parts);
    }
}
