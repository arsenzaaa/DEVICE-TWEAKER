using System.Globalization;
using System.Text;

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


    private void OnMainFormKeyDown(object? sender, KeyEventArgs e)
    {
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

            prefixLabel.Text = $"{name} -";
            statusLabel.Text = value;
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
                || ReferenceEquals(control, _dualCcdStatusLabel);
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

        SetHeaderFlag(_htPrefixLabel, _htStatusLabel, threadingName, OnOff(smtEnabled), smtEnabled);
        SetHeaderFlag(_hybridCpuPrefixLabel, _hybridCpuStatusLabel, "Hybrid CPU", OnOff(hasHybridCpu), hasHybridCpu);
        SetHeaderFlag(_cppcPrefixLabel, _cppcStatusLabel, "CPPC", OnOff(_cppcEnabled), _cppcEnabled);
        SetHeaderFlag(_dualCcdPrefixLabel, _dualCcdStatusLabel, "Dual-CCD", hasDualCcdCpu ? "True" : "False", hasDualCcdCpu);

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
        }
        finally
        {
            _cpuFlagsPanel?.ResumeLayout();
        }
    }

    private void ShowTestAdminDialog()
    {
        using Form dialog = new();
        dialog.Text = "TEST ADMIN";
        dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
        dialog.StartPosition = FormStartPosition.CenterParent;
        dialog.MaximizeBox = false;
        dialog.MinimizeBox = false;
        dialog.ShowInTaskbar = false;
        dialog.AutoScaleMode = AutoScaleMode.None;
        dialog.BackColor = _bgForm;
        dialog.ForeColor = _fgMain;
        dialog.Font = _baseFont;
        dialog.Icon = Icon;
        dialog.ClientSize = new Size(860, 540);
        using ToolTip adminToolTip = new()
        {
            UseFading = true,
            UseAnimation = true,
            IsBalloon = false,
            ShowAlways = false,
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
            RowCount = 16,
            Padding = new Padding(20, 26, 20, 20),
            BackColor = _bgForm,
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
        ComboBox cpuPresetCombo = NewDialogCombo(400);
        cpuPresetCombo.Items.AddRange(new object[]
        {
            "Manual / current",
            "Intel 8P/16T classic",
            "Intel 6P+8E/20T hybrid",
            "Intel 8P+16E/32T hybrid",
            "Intel Core Ultra 9 285K 8P+16E/24T",
            "AMD Ryzen Zen2 8C/16T 1 CCD / 2 CCX",
            "AMD Ryzen Zen2 12C/24T 2 CCD / 4 CCX",
            "AMD Ryzen 8C/16T 1 CCD",
            "AMD Ryzen 12C/24T 2 CCD",
            "AMD Ryzen 16C/32T 2 CCD",
            "AMD Ryzen X3D 16C/32T 2 CCD boost-preferred",
        });
        cpuPresetCombo.SelectedIndex = 0;
        Button cpuPresetButton = NewDialogButton("LOAD");
        cpuPresetButton.Size = new Size(84, 26);
        cpuPresetButton.Margin = new Padding(8, 0, 0, 6);
        FlowLayoutPanel cpuPresetPanel = NewRowFlowPanel();
        cpuPresetPanel.Controls.Add(cpuPresetCombo);
        cpuPresetPanel.Controls.Add(cpuPresetButton);

        Label logicalLabel = NewDialogLabel("Total logical processors:");
        NumericUpDown logicalUpDown = NewNumericUpDown(1, MaxAffinityBits, currentLogical);

        Label smtStateLabel = NewDialogLabel("SMT status:");
        ComboBox smtStateCombo = NewDialogCombo(160);
        smtStateCombo.Items.AddRange(new object[] { "Enabled", "Disabled" });

        Label htStateLabel = NewDialogLabel("Hyper-Threading status:");
        ComboBox htStateCombo = NewDialogCombo(160);
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
        assignmentsTable.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 116F));

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

        layout.Controls.Add(cpuPresetLabel, 0, 2);
        layout.Controls.Add(cpuPresetPanel, 1, 2);
        layout.Controls.Add(cpuNameLabel, 0, 3);
        layout.Controls.Add(cpuNameTextBox, 1, 3);
        layout.Controls.Add(logicalLabel, 0, 4);
        layout.Controls.Add(logicalUpDown, 1, 4);
        layout.Controls.Add(smtStateLabel, 0, 5);
        layout.Controls.Add(smtStateCombo, 1, 5);
        layout.Controls.Add(htStateLabel, 0, 6);
        layout.Controls.Add(htStateCombo, 1, 6);
        layout.Controls.Add(cppcRatingsLabel, 0, 7);
        layout.Controls.Add(cppcRatingsBox, 1, 7);
        adminToolTip.SetToolTip(cppcRatingsBox, "Optional. Format: 0=120, 1=110, 2=100 or just 120,110,100. Empty = CPPC off in test mode.");
        adminToolTip.SetToolTip(cpuPresetCombo, "Loads a synthetic topology for AUTO planner testing. You can still edit LP/core/CCD/CCX/E-core/CPPC values after loading.");
        layout.Controls.Add(groupCountLabel, 0, 8);
        layout.Controls.Add(groupCountPanel, 1, 8);
        layout.Controls.Add(assignmentsLabel, 0, 9);
        layout.Controls.Add(assignmentsHost, 1, 9);

        Label helpLabel = NewHintLabel("How to use: set group counts, then assign each LP to Core/CCD/CCX groups. Tick E-core where needed.");
        helpLabel.Margin = new Padding(0, 10, 0, 4);
        layout.Controls.Add(helpLabel, 0, 10);
        layout.SetColumnSpan(helpLabel, 2);

        Label noteLabel = new()
        {
            Text = $"Note: UI and affinity masks are capped at {MaxAffinityBits} LPs.",
            AutoSize = true,
            ForeColor = _mutedText,
            Margin = new Padding(0, 0, 0, 12),
        };
        layout.Controls.Add(noteLabel, 0, 11);
        layout.SetColumnSpan(noteLabel, 2);

        Label testSectionLabel = new()
        {
            Text = "Test Devices",
            AutoSize = true,
            Font = _titleFont,
            ForeColor = _accent,
            Margin = new Padding(0, 8, 0, 6),
        };
        layout.Controls.Add(testSectionLabel, 0, 12);
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
        testDevicesLayout.RowCount = 15;
        for (int i = 0; i < testDevicesLayout.RowCount; i++)
        {
            testDevicesLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        }

        CheckBox enableTestDevicesCheck = new()
        {
            Text = "Enable test devices",
            AutoSize = true,
            BackColor = _bgForm,
            ForeColor = _fgMain,
            Checked = _testDevicesEnabled,
            Margin = new Padding(0, 0, 16, 0),
        };

        CheckBox testDevicesOnlyCheck = new()
        {
            Text = "Show test devices only",
            AutoSize = true,
            BackColor = _bgForm,
            ForeColor = _fgMain,
            Checked = _testDevicesOnly,
            Margin = new Padding(0, 0, 16, 0),
        };
        testDevicesOnlyCheck.Enabled = _testDevicesEnabled;

        CheckBox dryRunAutoCheck = new()
        {
            Text = "Auto-optimization dry-run (no registry writes)",
            AutoSize = true,
            BackColor = _bgForm,
            ForeColor = _fgMain,
            Checked = _testAutoDryRun,
            Margin = new Padding(0, 0, 0, 0),
        };

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

        Label testNameLabel = NewDialogLabel("Name:");
        TextBox testNameBox = NewDialogTextBox(360);
        testDevicesLayout.Controls.Add(testNameLabel, 0, 3);
        testDevicesLayout.Controls.Add(testNameBox, 1, 3);

        Label testPnpIdLabel = NewDialogLabel("PNP ID:");
        TextBox testPnpIdBox = NewDialogTextBox(360);
        adminToolTip.SetToolTip(testPnpIdBox, @"Optional fake hardware ID. Example: PCI\VEN_10EC&DEV_8125\TEST for NIC ITR profile preview.");
        testDevicesLayout.Controls.Add(testPnpIdLabel, 0, 4);
        testDevicesLayout.Controls.Add(testPnpIdBox, 1, 4);

        Label testKindLabel = NewDialogLabel("Kind:");
        ComboBox testKindCombo = NewDialogCombo(180);
        testKindCombo.Items.Add(DeviceKind.USB);
        testKindCombo.Items.Add(DeviceKind.GPU);
        testKindCombo.Items.Add(DeviceKind.AUDIO);
        testKindCombo.Items.Add(DeviceKind.NET_NDIS);
        testKindCombo.Items.Add(DeviceKind.NET_CX);
        testKindCombo.Items.Add(DeviceKind.STOR);
        testKindCombo.SelectedIndex = 0;
        testDevicesLayout.Controls.Add(testKindLabel, 0, 5);
        testDevicesLayout.Controls.Add(testKindCombo, 1, 5);

        Label testUsbRolesLabel = NewDialogLabel("USB roles:");
        TextBox testUsbRolesBox = NewDialogTextBox(360);
        testDevicesLayout.Controls.Add(testUsbRolesLabel, 0, 6);
        testDevicesLayout.Controls.Add(testUsbRolesBox, 1, 6);

        Label testAudioLabel = NewDialogLabel("Audio endpoints:");
        TextBox testAudioBox = NewDialogTextBox(360);
        testDevicesLayout.Controls.Add(testAudioLabel, 0, 7);
        testDevicesLayout.Controls.Add(testAudioBox, 1, 7);

        Label testStorageLabel = NewDialogLabel("Storage tag:");
        TextBox testStorageBox = NewDialogTextBox(180);
        testDevicesLayout.Controls.Add(testStorageLabel, 0, 8);
        testDevicesLayout.Controls.Add(testStorageBox, 1, 8);

        CheckBox testWifiCheck = new()
        {
            Text = "WiFi",
            AutoSize = true,
            BackColor = _bgForm,
            ForeColor = _fgMain,
            Margin = new Padding(0, 0, 12, 0),
        };

        CheckBox testXhciCheck = new()
        {
            Text = "USB XHCI",
            AutoSize = true,
            BackColor = _bgForm,
            ForeColor = _fgMain,
            Checked = true,
            Margin = new Padding(0, 0, 12, 0),
        };

        CheckBox testHasDevicesCheck = new()
        {
            Text = "USB has devices",
            AutoSize = true,
            BackColor = _bgForm,
            ForeColor = _fgMain,
            Checked = true,
            Margin = new Padding(0, 0, 0, 0),
        };

        CheckBox testIntegratedGpuCheck = new()
        {
            Text = "Integrated GPU (iGPU)",
            AutoSize = true,
            BackColor = _bgForm,
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
        testDevicesLayout.Controls.Add(testDeviceOptionsLabel, 0, 9);
        testDevicesLayout.Controls.Add(testDeviceOptionsPanel, 1, 9);

        Button addTestDeviceButton = NewDialogButton("ADD FAKE DEVICE");
        addTestDeviceButton.Size = new Size(180, 30);
        addTestDeviceButton.Anchor = AnchorStyles.Left;
        addTestDeviceButton.Margin = new Padding(0, 4, 0, 10);
        testDevicesLayout.Controls.Add(addTestDeviceButton, 0, 10);
        testDevicesLayout.SetColumnSpan(addTestDeviceButton, 2);

        Label testListLabel = NewHeaderLabel($"Current test devices: {_testDevices.Count}");
        testDevicesLayout.Controls.Add(testListLabel, 0, 11);
        testDevicesLayout.SetColumnSpan(testListLabel, 2);

        ListBox testDeviceListBox = new()
        {
            Height = 140,
            BorderStyle = BorderStyle.FixedSingle,
            BackColor = Color.FromArgb(18, 18, 22),
            ForeColor = _fgMain,
            IntegralHeight = false,
            SelectionMode = SelectionMode.One,
            Margin = new Padding(0, 0, 0, 6),
        };
        testDevicesLayout.Controls.Add(testDeviceListBox, 0, 12);
        testDevicesLayout.SetColumnSpan(testDeviceListBox, 2);

        Button removeTestDeviceButton = NewDialogButton("REMOVE SELECTED");
        removeTestDeviceButton.Size = new Size(180, 30);
        removeTestDeviceButton.Margin = new Padding(0, 0, 12, 0);

        Button clearTestDeviceButton = NewDialogButton("CLEAR ALL");
        clearTestDeviceButton.Size = new Size(140, 30);
        clearTestDeviceButton.Margin = new Padding(0, 0, 0, 0);

        FlowLayoutPanel testDeviceButtonsPanel = NewRowFlowPanel();
        testDeviceButtonsPanel.WrapContents = true;
        testDeviceButtonsPanel.Margin = new Padding(0, 0, 0, 6);
        testDeviceButtonsPanel.Controls.Add(removeTestDeviceButton);
        testDeviceButtonsPanel.Controls.Add(clearTestDeviceButton);
        testDevicesLayout.Controls.Add(testDeviceButtonsPanel, 0, 13);
        testDevicesLayout.SetColumnSpan(testDeviceButtonsPanel, 2);

        Label testHintLabel = NewHintLabel(@"Tip: use roles like Mouse, Keyboard, Gamepad, Webcam, Microphone. For NIC ITR preview use PNP ID like PCI\VEN_10EC&DEV_8125\TEST.");
        testDevicesLayout.Controls.Add(testHintLabel, 0, 14);
        testDevicesLayout.SetColumnSpan(testHintLabel, 2);

        testDevicesPanel.Controls.Add(testDevicesLayout);
        layout.Controls.Add(testDevicesPanel, 0, 13);
        layout.SetColumnSpan(testDevicesPanel, 2);

        bool suppressTestDeviceToggle = false;
        void RefreshTestDeviceList()
        {
            testDeviceListBox.BeginUpdate();
            testDeviceListBox.Items.Clear();
            foreach (DeviceInfo device in _testDevices)
            {
                testDeviceListBox.Items.Add(FormatTestDeviceLabel(device));
            }
            testDeviceListBox.EndUpdate();
            testListLabel.Text = $"Current test devices: {_testDevices.Count}";
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

        enableTestDevicesCheck.CheckedChanged += (_, _) =>
        {
            if (suppressTestDeviceToggle)
            {
                return;
            }

            suppressTestDeviceToggle = true;
            _testDevicesEnabled = enableTestDevicesCheck.Checked;
            if (!_testDevicesEnabled)
            {
                _testDevicesOnly = false;
                testDevicesOnlyCheck.Checked = false;
            }
            else if (!dryRunAutoCheck.Checked)
            {
                dryRunAutoCheck.Checked = true;
            }
            testDevicesOnlyCheck.Enabled = _testDevicesEnabled;
            suppressTestDeviceToggle = false;

            WriteLog($"TEST.DEVICES: enabled={_testDevicesEnabled}");
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
            _initialDeviceViewportHeightAdjusted = false;
            RefreshBlocks();
        };

        dryRunAutoCheck.CheckedChanged += (_, _) =>
        {
            _testAutoDryRun = dryRunAutoCheck.Checked;
            WriteLog($"TEST.AUTO.DRYRUN: {(_testAutoDryRun ? "enabled" : "disabled")}");
        };

        testKindCombo.SelectedIndexChanged += (_, _) => UpdateTestDeviceFieldState();
        UpdateTestDeviceFieldState();
        RefreshTestDeviceList();

        addTestDeviceButton.Click += (_, _) =>
        {
            if (testKindCombo.SelectedItem is not DeviceKind kind)
            {
                return;
            }

            string name = testNameBox.Text?.Trim() ?? string.Empty;
            string pnpIdOverride = testPnpIdBox.Text?.Trim() ?? string.Empty;
            string usbRoles = testUsbRolesBox.Text?.Trim() ?? string.Empty;
            string audioEndpoints = testAudioBox.Text?.Trim() ?? string.Empty;
            string storageTag = testStorageBox.Text?.Trim() ?? string.Empty;

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

            DeviceInfo testDevice = CreateTestDevice(kind, name, pnpIdOverride, usbRoles, audioEndpoints, storageTag, testWifiCheck.Checked, testXhciCheck.Checked, testHasDevicesCheck.Checked, testIntegratedGpuCheck.Checked);
            _testDevices.Add(testDevice);
            WriteLog($"TEST.DEV.ADD: {testDevice.InstanceId} Kind={kind} Name=\"{testDevice.Name}\"");

            if (!dryRunAutoCheck.Checked)
            {
                dryRunAutoCheck.Checked = true;
            }

            RefreshTestDeviceList();
            bool shouldRefresh = false;
            if (!_testDevicesEnabled || _testDevicesOnly)
            {
                suppressTestDeviceToggle = true;
                _testDevicesEnabled = true;
                _testDevicesOnly = false;
                enableTestDevicesCheck.Checked = true;
                testDevicesOnlyCheck.Checked = false;
                testDevicesOnlyCheck.Enabled = true;
                suppressTestDeviceToggle = false;
                shouldRefresh = true;
            }

            if (_testDevicesEnabled || shouldRefresh)
            {
                _initialDeviceViewportHeightAdjusted = false;
                RefreshBlocks();
            }
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

            RefreshTestDeviceList();
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

        void LoadIntelHybridPreset(string name, int pCores, int eCores, bool pCoreHt)
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
                    cppc[lp] = rating;
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
                cppc[lp] = 70 - Math.Min(20, e);
                lp++;
            }

            LoadSyntheticPreset(name, logicalCount, physicalCoreCount, 1, 1, pCoreHt, true, coreMap, ccdMap, ccxMap, eCoreMap, cppc);
        }

        void LoadAmdPreset(string name, int physicalCores, int ccdCount, bool boostCcdPreferred, int ccxPerCcd = 1)
        {
            int logicalCount = Math.Min(MaxAffinityBits, physicalCores * 2);
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
                int ccx = (ccd * safeCcxPerCcd) + ccxInCcd;
                int rating = boostCcdPreferred
                    ? ccd == ccdCount - 1
                        ? 130 - Math.Min(12, coreInCcd)
                        : 100 - Math.Min(10, core)
                    : 120 - Math.Min(20, coreInCcd);

                for (int t = 0; t < 2 && lp < logicalCount; t++)
                {
                    coreMap[lp] = core;
                    ccdMap[lp] = ccd;
                    ccxMap[lp] = ccx;
                    eCoreMap[lp] = false;
                    cppc[lp] = rating;

                    lp++;
                }
            }

            LoadSyntheticPreset(name, logicalCount, physicalCores, ccdCount, ccxCount, true, false, coreMap, ccdMap, ccxMap, eCoreMap, cppc);
        }

        void ApplySelectedCpuPreset()
        {
            string preset = cpuPresetCombo.SelectedItem?.ToString() ?? string.Empty;
            switch (preset)
            {
                case "Manual / current":
                    return;
                case "Intel 8P/16T classic":
                    LoadIntelHybridPreset("Intel 8P/16T classic", pCores: 8, eCores: 0, pCoreHt: true);
                    break;
                case "Intel 6P+8E/20T hybrid":
                    LoadIntelHybridPreset("Intel 6P+8E/20T hybrid", pCores: 6, eCores: 8, pCoreHt: true);
                    break;
                case "Intel 8P+16E/32T hybrid":
                    LoadIntelHybridPreset("Intel 8P+16E/32T hybrid", pCores: 8, eCores: 16, pCoreHt: true);
                    break;
                case "Intel Core Ultra 9 285K 8P+16E/24T":
                    LoadIntelHybridPreset("Intel Core Ultra 9 285K 8P+16E/24T", pCores: 8, eCores: 16, pCoreHt: false);
                    break;
                case "AMD Ryzen Zen2 8C/16T 1 CCD / 2 CCX":
                    LoadAmdPreset("AMD Ryzen Zen2 8C/16T 1 CCD / 2 CCX", physicalCores: 8, ccdCount: 1, boostCcdPreferred: false, ccxPerCcd: 2);
                    break;
                case "AMD Ryzen Zen2 12C/24T 2 CCD / 4 CCX":
                    LoadAmdPreset("AMD Ryzen Zen2 12C/24T 2 CCD / 4 CCX", physicalCores: 12, ccdCount: 2, boostCcdPreferred: false, ccxPerCcd: 2);
                    break;
                case "AMD Ryzen 8C/16T 1 CCD":
                    LoadAmdPreset("AMD Ryzen 8C/16T 1 CCD", physicalCores: 8, ccdCount: 1, boostCcdPreferred: false);
                    break;
                case "AMD Ryzen 12C/24T 2 CCD":
                    LoadAmdPreset("AMD Ryzen 12C/24T 2 CCD", physicalCores: 12, ccdCount: 2, boostCcdPreferred: false);
                    break;
                case "AMD Ryzen 16C/32T 2 CCD":
                    LoadAmdPreset("AMD Ryzen 16C/32T 2 CCD", physicalCores: 16, ccdCount: 2, boostCcdPreferred: false);
                    break;
                case "AMD Ryzen X3D 16C/32T 2 CCD boost-preferred":
                    LoadAmdPreset("AMD Ryzen X3D 16C/32T 2 CCD boost-preferred", physicalCores: 16, ccdCount: 2, boostCcdPreferred: true);
                    break;
            }
        }

        void FillGroupCombo(ComboBox combo, int groupCount)
        {
            combo.BeginUpdate();
            combo.Items.Clear();
            for (int i = 0; i < groupCount; i++)
            {
                combo.Items.Add($"Group {i + 1}");
            }
            combo.EndUpdate();
        }

        void BuildAssignmentRows()
        {
            int logicalCount = (int)logicalUpDown.Value;
            int coreCount = (int)coreGroupCountUpDown.Value;
            int ccdCount = (int)ccdGroupCountUpDown.Value;
            int ccxCount = (int)ccxGroupCountUpDown.Value;

            suppressAssignmentEvents = true;
            assignmentsTable.SuspendLayout();
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

                ComboBox coreCombo = NewDialogCombo(150);
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

                ComboBox ccdCombo = NewDialogCombo(150);
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

                ComboBox ccxCombo = NewDialogCombo(120);
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
                    BackColor = _bgForm,
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

        TableLayoutPanel buttons = new()
        {
            Dock = DockStyle.Bottom,
            ColumnCount = 3,
            RowCount = 1,
            Padding = new Padding(20, 10, 20, 14),
            BackColor = _bgForm,
        };
        buttons.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33F));
        buttons.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33F));
        buttons.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33F));
        buttons.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        Button applyButton = NewDialogButton("APPLY TEST");
        applyButton.Size = new Size(170, 32);
        applyButton.Anchor = AnchorStyles.None;
        applyButton.Margin = new Padding(6, 0, 6, 0);
        applyButton.Click += (_, _) =>
        {
            if (!TryParseTestCppcRatings(cppcRatingsBox.Text, (int)logicalUpDown.Value, out Dictionary<int, int> cppcRatings, out string cppcError))
            {
                ShowThemedInfo($"CPPC ratings are invalid.\n{cppcError}");
                return;
            }

            TestCpuConfig config = BuildConfigFromAssignments(cppcRatings);
            ApplyTestCpuConfig(config);
            statusLabel.Text = "Test CPU mode: ACTIVE";
            statusLabel.ForeColor = _statusActive;
        };

        Button resetButton = NewDialogButton("RESET TO REAL");
        resetButton.Size = new Size(170, 32);
        resetButton.Anchor = AnchorStyles.None;
        resetButton.Margin = new Padding(6, 0, 6, 0);
        resetButton.Click += (_, _) =>
        {
            DisableTestCpuMode();
            statusLabel.Text = "Test CPU mode: OFF";
            statusLabel.ForeColor = _statusInactive;
            logicalUpDown.Value = Math.Min(MaxAffinityBits, GetCurrentLogicalCount());
            LoadAssignmentsFromCurrentCpu();
            SyncSmtStateFromCurrent();
            cppcRatingsBox.Text = GetCurrentCppcRatingsText();
        };

        Button closeButton = NewDialogButton("CLOSE");
        closeButton.Size = new Size(170, 32);
        closeButton.Anchor = AnchorStyles.None;
        closeButton.Margin = new Padding(6, 0, 6, 0);
        closeButton.DialogResult = DialogResult.Cancel;

        buttons.Controls.Add(applyButton, 0, 0);
        buttons.Controls.Add(resetButton, 1, 0);
        buttons.Controls.Add(closeButton, 2, 0);

        const int dialogScrollWidth = 13;
        Panel contentHost = new()
        {
            Dock = DockStyle.Fill,
            BackColor = _bgForm,
            Padding = Padding.Empty,
        };

        Panel contentPanel = new()
        {
            Dock = DockStyle.None,
            BackColor = _bgForm,
            AutoScroll = false,
            Padding = new Padding(0, 0, dialogScrollWidth + 8, 0),
        };
        contentPanel.Location = new Point(0, 0);
        contentPanel.Controls.Add(layout);

        ThemedScrollBar dialogScroll = new()
        {
            Width = dialogScrollWidth,
            BackColor = _bgForm,
            TrackColor = _bgForm,
            RailColor = _bgForm,
            ThumbColor = _accent,
            ThumbWidth = 9,
            RailWidth = 0,
            ThumbCornerRadius = 7,
            Visible = false,
            Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Right,
        };

        bool syncingDialogScroll = false;

        void UpdateDialogScrollLayout()
        {
            dialogScroll.Location = new Point(Math.Max(0, contentHost.ClientSize.Width - dialogScroll.Width), 0);
            dialogScroll.Height = contentHost.ClientSize.Height;
            dialogScroll.BringToFront();
        }

        void UpdateDialogHostLayout()
        {
            contentPanel.Width = contentHost.ClientSize.Width;
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

        void SyncDialogScrollBar()
        {
            contentPanel.Width = contentHost.ClientSize.Width;
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
        dialog.Controls.Add(buttons);
        dialog.AcceptButton = applyButton;
        dialog.CancelButton = closeButton;
        dialog.Shown += (_, _) =>
        {
            ApplyTitleBarTheme(dialog);
            SyncDialogScrollBar();
            dialog.BeginInvoke(new Action(() =>
            {
                adminToolTip.Hide(dialog);
                adminToolTip.Active = true;
            }));
        };

        dialog.PerformLayout();
        int desiredHeight = layout.PreferredSize.Height + buttons.Height + 12;
        int maxHeight = Math.Min(760, Screen.FromControl(dialog).WorkingArea.Height - 140);
        int targetHeight = Math.Min(maxHeight, Math.Max(520, desiredHeight));
        dialog.ClientSize = new Size(dialog.ClientSize.Width, targetHeight);
        syncDialogScroll = SyncDialogScrollBar;

        dialog.ShowDialog(this);
    }

    private Button NewDialogButton(string text)
    {
        Button btn = new()
        {
            Text = text,
            Size = new Size(150, 32),
            Margin = new Padding(8, 0, 0, 0),
            FlatStyle = FlatStyle.Flat,
            Font = _buttonFont,
            UseVisualStyleBackColor = false,
            Cursor = Cursors.Hand,
            BackColor = _bgForm,
            ForeColor = _fgMain,
        };
        btn.FlatAppearance.BorderSize = 1;
        btn.FlatAppearance.BorderColor = _accent;
        btn.MouseEnter += (_, _) =>
        {
            btn.BackColor = _accent;
            btn.ForeColor = Color.FromArgb(15, 15, 15);
        };
        btn.MouseLeave += (_, _) =>
        {
            btn.BackColor = _bgForm;
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
            ForeColor = _fgMain,
            TextAlign = ContentAlignment.MiddleLeft,
            Anchor = AnchorStyles.Left,
            Margin = new Padding(0, 2, 12, 6),
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
            Margin = new Padding(8, 4, 6, 0),
        };
    }

    private Label NewHintLabel(string text)
    {
        return new Label
        {
            Text = text,
            AutoSize = true,
            ForeColor = _mutedText,
            Margin = new Padding(0, 0, 0, 6),
            MaximumSize = new Size(560, 0),
        };
    }

    private Panel NewBoxPanel()
    {
        Panel panel = new()
        {
            BackColor = _bgForm,
            ForeColor = _fgMain,
            Padding = new Padding(6),
            Margin = new Padding(0),
        };
        panel.Paint += (_, e) =>
        {
            Rectangle rect = panel.ClientRectangle;
            rect.Width -= 1;
            rect.Height -= 1;
            using Pen pen = new(_border);
            e.Graphics.DrawRectangle(pen, rect);
        };
        return panel;
    }

    private NumericUpDown NewNumericUpDown(int min, int max, int value)
    {
        int clamped = Math.Min(max, Math.Max(min, value));
        return new NumericUpDown
        {
            Minimum = min,
            Maximum = max,
            Value = clamped,
            TextAlign = HorizontalAlignment.Center,
            BackColor = Color.FromArgb(18, 18, 22),
            ForeColor = _fgMain,
            Size = new Size(120, 24),
            Margin = new Padding(0, 0, 0, 6),
        };
    }

    private ComboBox NewDialogCombo(int width)
    {
        return new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(18, 18, 22),
            ForeColor = _fgMain,
            Size = new Size(width, 26),
            Margin = new Padding(0, 0, 0, 6),
        };
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
        return new TextBox
        {
            BackColor = Color.FromArgb(18, 18, 22),
            ForeColor = _fgMain,
            BorderStyle = BorderStyle.FixedSingle,
            Size = new Size(width, 24),
            Margin = new Padding(0, 0, 0, 6),
        };
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
        bool integratedGpu)
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

        return new DeviceInfo
        {
            Name = displayName,
            InstanceId = id,
            Class = className,
            RegBase = $@"SYSTEM\CurrentControlSet\Enum\{id}",
            Kind = kind,
            UsbRoles = isUsb ? usbRoles : string.Empty,
            AudioEndpoints = isAudio ? audioEndpoints : string.Empty,
            StorageTag = isStor ? storageTag : string.Empty,
            IsIntegratedGpu = kind == DeviceKind.GPU && integratedGpu,
            Wifi = isNet && wifi,
            UsbIsXhci = isUsb && usbIsXhci,
            UsbHasDevices = isUsb && usbHasDevices,
            IsTestDevice = true,
        };
    }

    private static string FormatTestDeviceLabel(DeviceInfo device)
    {
        string label = $"{device.Kind}: {device.Name}";
        if (device.Kind == DeviceKind.USB && !string.IsNullOrWhiteSpace(device.UsbRoles))
        {
            label += $" [{device.UsbRoles}]";
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
            int eff = config.ECoreLps.Contains(lp) ? 1 : 0;
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
            if (uniqueRatings.Count > 1)
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
        _grpHeight = UiScale(120) + (_maxLogical * UiScale(24)) + UiScale(160);

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
