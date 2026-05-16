namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private const string StorageAffinityNoteText = "(affinity masks not supported on SSD/HDD)";
    private static readonly string[] ImodDeviceEditorRoleOrder = ["Mouse", "Keyboard", "Audio", "Gamepad"];

    private static IReadOnlyList<string> GetImodDeviceEditorRoles(DeviceInfo device)
    {
        HashSet<string> roles = new(StringComparer.OrdinalIgnoreCase);
        AddImodDeviceEditorRoles(device.UsbRoles, roles);
        AddImodDeviceEditorRoles(device.AudioEndpoints, roles);

        return ImodDeviceEditorRoleOrder
            .Where(roles.Contains)
            .ToArray();
    }

    private static void AddImodDeviceEditorRoles(string? text, HashSet<string> roles)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return;
        }

        foreach (string part in text.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            if (TryGetAdaptiveRoleBase(part, out string role)
                && ImodDeviceEditorRoleOrder.Contains(role, StringComparer.OrdinalIgnoreCase))
            {
                roles.Add(role);
            }
        }
    }

    private static uint GetImodDeviceEditorDefaultValue(string role)
    {
        return role.Equals("Mouse", StringComparison.OrdinalIgnoreCase)
            ? 0x0
            : role.Equals("Audio", StringComparison.OrdinalIgnoreCase)
                ? 0xFA0
                : ImodDefaultInterval;
    }

    private int GetDevicesScrollGutterWidth()
    {
        if (_devicesScroll is null)
        {
            return 0;
        }

        return _devicesScroll.Width + UiScale(2);
    }

    private int GetDevicesViewportWidth()
    {
        int width = _devicesHost.ClientSize.Width;
        if (width <= 0)
        {
            width = _devicesPanel.ClientSize.Width;
        }

        width -= GetDevicesScrollGutterWidth();

        return Math.Max(UiScale(360) + (UiScale(24) * 2), width);
    }

    private void FixRssPolicyLabelOverlap(DeviceBlock block)
    {
        int desiredLeft = block.PolicyLabel.Right + UiScale(8);
        if (desiredLeft <= block.PolicyCombo.Left)
        {
            return;
        }

        int rightPadding = UiScale(24);
        int parentLeft = block.PolicyCombo.Parent?.Left ?? 0;
        int maxWidth = block.Group.ClientSize.Width - rightPadding - parentLeft - desiredLeft;
        int newWidth = block.PolicyCombo.Width;
        if (maxWidth < newWidth)
        {
            newWidth = maxWidth;
        }

        if (newWidth < UiScale(90))
        {
            newWidth = UiScale(90);
        }

        block.PolicyCombo.Location = new Point(desiredLeft, block.PolicyCombo.Top);
        block.PolicyCombo.Width = newWidth;
    }

    private string BuildDeviceBlockTitle(DeviceInfo device, string? usbRolesOverride = null)
    {
        string title = device.Name;
        string usbRoles = usbRolesOverride ?? device.UsbRoles;
        if (device.Kind == DeviceKind.GPU && device.IsIntegratedGpu)
        {
            title = $"{device.Name} [iGPU]";
        }
        else if (device.Kind == DeviceKind.USB && !string.IsNullOrWhiteSpace(usbRoles))
        {
            title = $"{device.Name} [{usbRoles}]";
        }
        else if (device.Kind == DeviceKind.USB)
        {
            title = $"{device.Name} [No HID roles]";
        }
        else if (device.Kind == DeviceKind.NET_NDIS)
        {
            title = $"{device.Name} [NDIS]";
        }
        else if (device.Kind == DeviceKind.NET_CX)
        {
            title = $"{device.Name} [NetAdapterCx]";
        }
        else if (device.Kind == DeviceKind.STOR && !string.IsNullOrWhiteSpace(device.StorageTag))
        {
            title = $"{device.Name} [{device.StorageTag}]";
        }
        else if (device.Kind == DeviceKind.AUDIO && !string.IsNullOrWhiteSpace(device.AudioEndpoints))
        {
            title = $"{device.Name} [{device.AudioEndpoints}]";
        }

        if (device.IsTestDevice)
        {
            title = $"[TEST] {title}";
        }

        return title;
    }

    private void NewDeviceBlock(
        DeviceInfo device,
        int index,
        IReadOnlyDictionary<string, string>? priorImodStatuses = null)
    {
        DeviceCardPanel grp = new()
        {
            BorderColor = _border,
        };

        string title = BuildDeviceBlockTitle(device);
        string logTitle = device.Kind == DeviceKind.STOR ? $"{title} {StorageAffinityNoteText}" : title;

        WriteLog(
            $"UI.BLOCK: idx={index} title=\"{logTitle}\" kind={device.Kind} id={device.InstanceId} roles=\"{device.UsbRoles}\" audio=\"{device.AudioEndpoints}\" storage=\"{device.StorageTag}\"");

        grp.Width = GetDevicesViewportWidth() - UiScale(40);
        grp.Height = _grpHeight;
        grp.BackColor = _bgGroup;
        grp.ForeColor = _fgMain;
        grp.Font = _blockFont;
        grp.Margin = new Padding(0);
        grp.Padding = new Padding(UiScale(12), UiScale(16), UiScale(12), UiScale(16));

        FlowLayoutPanel headerPanel = new()
        {
            AutoSize = false,
            Size = new Size(grp.Width - UiScale(40), UiScale(24)),
            Location = new Point(UiScale(18), UiScale(8)),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            WrapContents = false,
            Margin = Padding.Empty,
            Padding = Padding.Empty,
            BackColor = Color.Transparent,
        };

        Label headerLabel = new HighlightLabel
        {
            Text = title,
            Font = _blockTitleFont,
            ForeColor = _fgMain,
            AutoSize = true,
            Margin = Padding.Empty,
            HighlightText = "Mouse scanning",
            HighlightColor = Color.FromArgb(120, 120, 120),
        };
        headerPanel.Controls.Add(headerLabel);

        Label? headerNote = null;
        if (device.Kind == DeviceKind.STOR)
        {
            headerNote = new Label
            {
                Text = StorageAffinityNoteText,
                Font = _blockFont,
                ForeColor = _mutedWarn,
                AutoSize = true,
                Margin = new Padding(UiScale(6), UiScale(2), 0, 0),
            };
            headerPanel.Controls.Add(headerNote);
        }

        Panel divider = new()
        {
            BackColor = _border,
            Size = new Size(grp.Width - UiScale(32), UiScale(1)),
            Location = new Point(UiScale(16), UiScale(36)),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
        };

        int contentTop = UiScale(48);
        Label cpuLabel = new()
        {
            Text = "CPU Affinity",
            AutoSize = true,
            ForeColor = _fgMain,
            Location = new Point(UiScale(18), contentTop),
        };
        if (device.Kind == DeviceKind.STOR)
        {
            cpuLabel.ForeColor = _mutedText;
        }

        int cpuPanelTop = cpuLabel.Bottom + UiScale(6);
        int cpuPanelHeight = UiScale(150);
        int settingsMinimumWidth = Math.Min(UiScale(420), Math.Max(UiScale(320), grp.Width - UiScale(48)));
        int settingsSideMinimumWidth = UiScale(600);
        int cpuPanelMinimumWidth = UiScale(308);
        int settingsSideGap = UiScale(40);
        int cpuPanelFullMaximumWidth = Math.Max(
            cpuPanelMinimumWidth,
            grp.Width - UiScale(16) - UiScale(24));
        int cpuPanelSideMaximumWidth = Math.Max(
            cpuPanelMinimumWidth,
            grp.Width - UiScale(16) - settingsSideGap - settingsMinimumWidth - UiScale(24));
        int cpuPanelMaximumWidth = cpuPanelSideMaximumWidth;
        int cpuPanelWidth = cpuPanelMinimumWidth;
        Panel cpuPanel = new()
        {
            Location = new Point(UiScale(16), cpuPanelTop),
            Size = new Size(cpuPanelWidth, cpuPanelHeight),
            BackColor = _bgForm,
            Padding = new Padding(UiScale(8), UiScale(6), UiScale(8), UiScale(6)),
        };
        cpuPanel.Paint += (_, e) =>
        {
            Rectangle rect = cpuPanel.ClientRectangle;
            rect.Width -= 1;
            rect.Height -= 1;
            using Pen pen = new(_border);
            e.Graphics.DrawRectangle(pen, rect);
        };

        List<CheckBox> cpuBoxes = [];
        List<(int Lp, CheckBox Control, int Ccd, int Eff)> lpMeta = [];
        int checkSpacing = UiScale(22);

        for (int i = 0; i < _maxLogical; i++)
        {
            CheckBox cb = new()
            {
                Text = $"CPU {i}",
                AutoSize = true,
                ForeColor = _fgMain,
                BackColor = _bgForm,
                FlatStyle = FlatStyle.Flat,
            };
            StyleCpuCheckbox(cb, i);
            cb.Tag = i;
            if (device.Kind == DeviceKind.STOR)
            {
                cb.AutoCheck = false;
                cb.TabStop = false;
                cb.Cursor = Cursors.No;
                cb.ForeColor = _border;
            }

            cpuBoxes.Add(cb);

            int ccdId = _cpuInfo?.CcdMap.TryGetValue(i, out int cid) == true ? cid : 0;
            int eff = _cpuLpByIndex.TryGetValue(i, out CpuLpInfo? lpInfoLocal) ? lpInfoLocal.EffClass : -1;
            lpMeta.Add((i, cb, ccdId, eff));
            cpuPanel.Controls.Add(cb);
        }

        List<int> ccdKeys = lpMeta.Select(m => m.Ccd).Distinct().OrderBy(x => x).ToList();
        if (ccdKeys.Count == 0)
        {
            ccdKeys.Add(0);
        }

        List<List<(int Lp, CheckBox Control, int Ccd, int Eff)>> columns = [];
        foreach (int cid in ccdKeys)
        {
            List<(int Lp, CheckBox Control, int Ccd, int Eff)> items = lpMeta.Where(m => m.Ccd == cid).ToList();
            List<(int Lp, CheckBox Control, int Ccd, int Eff)> pItems = items.Where(m => !IsEfficiencyClass(m.Eff)).ToList();
            List<(int Lp, CheckBox Control, int Ccd, int Eff)> eItems = items.Where(m => IsEfficiencyClass(m.Eff)).ToList();
            List<(int Lp, CheckBox Control, int Ccd, int Eff)> other = items.Except(pItems).Except(eItems).ToList();
            List<(int Lp, CheckBox Control, int Ccd, int Eff)> ordered = [.. pItems, .. eItems, .. other];
            columns.Add(ordered);
        }

        int columnGap = UiScale(16);
        int startX = UiScale(10);
        int minColumnWidth = UiScale(120);
        int checkboxTextSafety = UiScale(30);

        int runningX = startX;
        int maxColumnCount = 0;
        List<int> columnWidths = [];
        foreach (List<(int Lp, CheckBox Control, int Ccd, int Eff)> ordered in columns)
        {
            if (ordered.Count > maxColumnCount)
            {
                maxColumnCount = ordered.Count;
            }

            int maxWidth = minColumnWidth;
            if (ordered.Count > 0)
            {
                int w = ordered.Max(o =>
                {
                    int preferred = o.Control.PreferredSize.Width;
                    int measured = TextRenderer.MeasureText(o.Control.Text, o.Control.Font).Width + checkboxTextSafety;
                    return Math.Max(preferred, measured);
                });
                if (w > 0)
                {
                    maxWidth = Math.Max(minColumnWidth, w + UiScale(14));
                }
            }

            columnWidths.Add(maxWidth);
        }

        for (int i = 0; i < columns.Count; i++)
        {
            List<(int Lp, CheckBox Control, int Ccd, int Eff)> ordered = columns[i];
            int maxWidth = columnWidths[i];
            int cellWidth = Math.Max(minColumnWidth, maxWidth - UiScale(4));
            int cellHeight = Math.Max(UiScale(18), checkSpacing - UiScale(2));
            int y = UiScale(4);
            foreach ((int _, CheckBox control, int _, int _) in ordered)
            {
                control.AutoSize = false;
                control.Size = new Size(cellWidth, cellHeight);
                control.TextAlign = ContentAlignment.MiddleLeft;
                control.Location = new Point(runningX, y);
                y += checkSpacing;
            }

            runningX += maxWidth + columnGap;
        }

        int requiredWidth = columns.Count == 0 ? cpuPanel.Width : runningX - columnGap + startX + UiScale(18);
        int cpuPanelHorizontalSlack = UiScale(30);
        int sideSettingsX = 0;
        int sideSettingsWidth = 0;
        int settingsX = UiScale(18);
        int availableSettingsWidth = Math.Max(UiScale(260), grp.Width - settingsX - UiScale(24));
        bool allowWindowAutoExpand = true;

        void UpdateResponsivePlacement()
        {
            settingsSideMinimumWidth = Math.Min(UiScale(600), Math.Max(UiScale(440), grp.Width - UiScale(48)));
            settingsMinimumWidth = Math.Min(UiScale(420), Math.Max(UiScale(320), grp.Width - UiScale(48)));
            cpuPanelFullMaximumWidth = Math.Max(cpuPanelMinimumWidth, grp.Width - cpuPanel.Left - UiScale(24));
            int desiredSideCpuPanelWidth = Math.Max(cpuPanelMinimumWidth, requiredWidth + cpuPanelHorizontalSlack);
            int desiredSideGroupWidth = cpuPanel.Left + desiredSideCpuPanelWidth + settingsSideGap + settingsSideMinimumWidth + UiScale(24);
            if (allowWindowAutoExpand && grp.Width < desiredSideGroupWidth)
            {
                TryExpandWindowForSideSettings(desiredSideGroupWidth);
            }

            cpuPanelFullMaximumWidth = Math.Max(cpuPanelMinimumWidth, grp.Width - cpuPanel.Left - UiScale(24));
            cpuPanelSideMaximumWidth = Math.Max(
                cpuPanelMinimumWidth,
                grp.Width - cpuPanel.Left - settingsSideGap - settingsSideMinimumWidth - UiScale(24));

            cpuPanelMaximumWidth = cpuPanelSideMaximumWidth;
            int targetCpuPanelWidth = Math.Min(
                cpuPanelMaximumWidth,
                Math.Max(cpuPanelMinimumWidth, requiredWidth + cpuPanelHorizontalSlack));
            if (cpuPanel.Width != targetCpuPanelWidth)
            {
                cpuPanel.Width = targetCpuPanelWidth;
            }

            if (requiredWidth > cpuPanel.ClientSize.Width + UiScale(2)
                && requiredWidth > cpuPanelFullMaximumWidth)
            {
                cpuPanel.AutoScroll = true;
                cpuPanel.AutoScrollMinSize = new Size(requiredWidth + cpuPanelHorizontalSlack, cpuPanel.Height);
            }
            else
            {
                cpuPanel.AutoScroll = false;
                cpuPanel.AutoScrollMinSize = Size.Empty;
            }

            sideSettingsX = cpuPanel.Right + settingsSideGap;
            sideSettingsWidth = grp.Width - sideSettingsX - UiScale(24);
            settingsX = sideSettingsX;
            availableSettingsWidth = Math.Max(UiScale(260), grp.Width - settingsX - UiScale(24));
        }

        UpdateResponsivePlacement();

        int desiredHeight = Math.Max((maxColumnCount * checkSpacing) + UiScale(18), UiScale(150));
        if (cpuPanel.Height != desiredHeight)
        {
            cpuPanel.Height = desiredHeight;
        }

        int maskY = cpuPanel.Bottom + UiScale(10);
        Label lblMask = new()
        {
            Text = "Affinity Mask: 0x0",
            AutoSize = true,
            ForeColor = _accent,
            Location = new Point(UiScale(18), maskY),
        };
        if (device.Kind == DeviceKind.STOR)
        {
            lblMask.ForeColor = _mutedText;
        }

        Label lblIrq = new()
        {
            Text = "IRQ Count: [Click CALCULATE IRQ COUNTS]",
            AutoSize = true,
            ForeColor = _mutedText,
            Location = new Point(UiScale(18), maskY + UiScale(20)),
        };

        int valueX = UiScale(132);
        int rowGap = UiScale(12);
        int rowTop = 0;
        int labelOffset = UiScale(4);

        bool TryExpandWindowForSideSettings(int desiredGroupWidth)
        {
            int desiredViewportWidth = desiredGroupWidth + UiScale(36);
            if (!TryExpandMainWindowForViewportWidth(desiredViewportWidth))
            {
                return false;
            }

            int expandedGroupWidth = Math.Max(grp.Width, GetDevicesViewportWidth() - UiScale(36));
            if (expandedGroupWidth > grp.Width)
            {
                grp.Width = expandedGroupWidth;
            }

            return grp.Width >= desiredGroupWidth - UiScale(2);
        }

        Panel settingsPanel = new()
        {
            AutoSize = false,
            BackColor = Color.Transparent,
        };

        Label lblMsi = new()
        {
            Text = "MSI Mode:",
            AutoSize = true,
            Location = new Point(0, rowTop + labelOffset),
        };

        ThemedDropDownPicker cmbMsi = new()
        {
            Location = new Point(valueX, rowTop),
            Size = UiScale(150, 26),
        };
        StyleDarkDropDownPicker(cmbMsi);
        cmbMsi.Items.AddRange(new object[] { "Disabled", "Enabled" });
        cmbMsi.DropDownWidth = cmbMsi.Width;
        cmbMsi.MaxDropDownItems = 2;

        settingsPanel.Controls.AddRange([lblMsi, cmbMsi]);
        rowTop = cmbMsi.Bottom + rowGap;

        Label lblLimit = new()
        {
            Text = "MSI Limit:",
            AutoSize = true,
            Location = new Point(0, rowTop + labelOffset),
        };

        TextBox txtLimit = new()
        {
            Location = new Point(valueX, rowTop),
            Size = UiScale(100, 24),
            BackColor = Color.FromArgb(18, 18, 22),
            BorderStyle = BorderStyle.FixedSingle,
            ForeColor = _fgMain,
            TextAlign = HorizontalAlignment.Center,
            Text = "0",
        };

        Label lblLimitHint = new()
        {
            Text = "(0 = unlimited)",
            AutoSize = true,
            ForeColor = _mutedText,
            Location = new Point(txtLimit.Right + UiScale(8), txtLimit.Top + UiScale(4)),
        };

        settingsPanel.Controls.AddRange([lblLimit, txtLimit, lblLimitHint]);
        rowTop = txtLimit.Bottom + rowGap;

        Label lblPrio = new()
        {
            Text = "IRQ Priority:",
            AutoSize = true,
            Location = new Point(0, rowTop + labelOffset),
        };

        ThemedDropDownPicker cmbPrio = new()
        {
            Location = new Point(valueX, rowTop),
            Size = UiScale(150, 26),
        };
        StyleDarkDropDownPicker(cmbPrio);
        cmbPrio.Items.AddRange(new object[] { "Undefined", "Low", "Normal", "High" });
        cmbPrio.DropDownWidth = cmbPrio.Width;
        cmbPrio.MaxDropDownItems = 4;

        settingsPanel.Controls.AddRange([lblPrio, cmbPrio]);
        rowTop = cmbPrio.Bottom + rowGap;

        Label lblPolicy = new()
        {
            Text = "Policy:",
            AutoSize = true,
            Location = new Point(0, rowTop + labelOffset),
        };

        ThemedDropDownPicker cmbPolicy = new()
        {
            Location = new Point(valueX, rowTop),
            Size = UiScale(170, 26),
        };
        StyleDarkDropDownPicker(cmbPolicy);
        cmbPolicy.DropDownWidth = cmbPolicy.Width;
        cmbPolicy.MaxDropDownItems = 6;
        if (device.Kind == DeviceKind.STOR)
        {
            lblPolicy.ForeColor = _mutedText;
            cmbPolicy.Enabled = false;
        }

        if (device.Kind != DeviceKind.NET_NDIS)
        {
            settingsPanel.Controls.AddRange([lblPolicy, cmbPolicy]);
            rowTop = cmbPolicy.Bottom + rowGap;
        }
        else
        {
            lblPolicy.Visible = false;
            cmbPolicy.Visible = false;
        }

        Label lblNdisMode = new()
        {
            Text = "NDIS Mode:",
            AutoSize = true,
            Location = new Point(0, rowTop + labelOffset),
            ForeColor = _fgMain,
            Visible = device.Kind == DeviceKind.NET_NDIS,
        };

        ThemedDropDownPicker cmbNdisMode = new()
        {
            Location = new Point(valueX, rowTop),
            Size = UiScale(94, 26),
            BackColor = Color.FromArgb(18, 18, 22),
            ForeColor = _fgMain,
            Font = _blockFont,
            BorderColor = _border,
            ButtonColor = Color.FromArgb(14, 14, 17),
            SelectedBackColor = Color.FromArgb(48, 48, 58),
            SelectedForeColor = _fgMain,
            ArrowColor = _fgMain,
            ItemHeight = UiScale(18),
            Visible = device.Kind == DeviceKind.NET_NDIS,
        };
        cmbNdisMode.Items.AddRange(["RSS", "IRQ", "BOTH"]);
        cmbNdisMode.SelectedIndex = 0;
        cmbNdisMode.DropDownWidth = cmbNdisMode.Width;
        cmbNdisMode.MaxDropDownItems = 3;

        if (device.Kind == DeviceKind.NET_NDIS)
        {
            settingsPanel.Controls.AddRange([lblNdisMode, cmbNdisMode]);
            _copyToolTip.SetToolTip(cmbNdisMode, "RSS writes network receive queue CPU placement. IRQ writes interrupt affinity policy. BOTH writes both when the adapter supports RSS.");
            rowTop = cmbNdisMode.Bottom + rowGap;
        }

        Label lblRssQueues = new()
        {
            Text = "RSS Queues:",
            AutoSize = true,
            Location = new Point(0, rowTop + labelOffset),
            ForeColor = _fgMain,
            Visible = device.Kind == DeviceKind.NET_NDIS,
        };

        NumericUpDown nudRssQueues = new()
        {
            Location = new Point(valueX, rowTop),
            Size = UiScale(70, 26),
            Minimum = 1,
            Maximum = Math.Max(1, _maxLogical),
            Value = 1,
            BackColor = Color.FromArgb(18, 18, 22),
            ForeColor = _fgMain,
            BorderStyle = BorderStyle.FixedSingle,
            Visible = device.Kind == DeviceKind.NET_NDIS,
        };

        if (device.Kind == DeviceKind.NET_NDIS)
        {
            settingsPanel.Controls.AddRange([lblRssQueues, nudRssQueues]);
            _copyToolTip.SetToolTip(nudRssQueues, "Number of RSS receive processors/queues to pin from the selected base CPU.");
            rowTop = nudRssQueues.Bottom + rowGap;
        }

        NicItrProfile? nicItrProfile = device.Kind is DeviceKind.NET_NDIS or DeviceKind.NET_CX
            ? TryGetNicItrProfile(device.InstanceId)
            : null;
        bool showNicItr = nicItrProfile is not null;
        int nicSetButtonWidth = UiScale(50);
        int nicSaveButtonWidth = UiScale(58);
        int nicButtonGap = UiScale(6);
        int nicInlineGap = UiScale(8);
        int nicInputAvailableWidth = Math.Max(UiScale(100), availableSettingsWidth - valueX - UiScale(8));
        int nicInputMinBase = nicItrProfile is { MaxQueues: > 1 } ? UiScale(260) : UiScale(130);
        int nicInputMinWidth = Math.Min(nicInputMinBase, nicInputAvailableWidth);
        int nicInputDesiredWidth = nicItrProfile is { MaxQueues: > 1 } ? UiScale(330) : UiScale(160);
        int nicInlineMaxWidth = availableSettingsWidth - valueX - nicInlineGap - nicSetButtonWidth - nicButtonGap - nicSaveButtonWidth - UiScale(8);
        bool nicButtonsInline = nicInlineMaxWidth >= nicInputMinWidth;
        int nicInputWidth = nicButtonsInline
            ? Math.Min(nicInputDesiredWidth, Math.Max(nicInputMinWidth, nicInlineMaxWidth))
            : Math.Min(nicInputDesiredWidth, Math.Max(nicInputMinWidth, nicInputAvailableWidth));
        int nicStatusWidth = Math.Max(UiScale(120), nicInputAvailableWidth);
        int nicDetailRows = nicItrProfile is { MaxQueues: > 1 and <= 4 } ? nicItrProfile.MaxQueues : 2;
        int nicDetailHeight = Math.Max(UiScale(42), UiScale((nicDetailRows * 17) + 8));
        Label lblNicItr = new()
        {
            Text = "NIC ITR:",
            AutoSize = true,
            Location = new Point(0, rowTop + labelOffset),
            ForeColor = _fgMain,
            Visible = showNicItr,
        };

        TextBox txtNicItr = new()
        {
            Location = new Point(valueX, rowTop),
            Size = new Size(nicInputWidth, UiScale(24)),
            BackColor = Color.FromArgb(18, 18, 22),
            BorderStyle = BorderStyle.FixedSingle,
            ForeColor = _fgMain,
            TextAlign = HorizontalAlignment.Left,
            Text = "0x0",
            Visible = showNicItr,
        };

        Button btnNicItr = new()
        {
            Text = "SET",
            Size = new Size(nicSetButtonWidth, UiScale(24)),
            Location = nicButtonsInline
                ? new Point(txtNicItr.Right + nicInlineGap, txtNicItr.Top - UiScale(1))
                : new Point(valueX, txtNicItr.Bottom + UiScale(6)),
            FlatStyle = FlatStyle.Flat,
            Font = _blockFont,
            UseVisualStyleBackColor = false,
            Cursor = Cursors.Hand,
            Visible = showNicItr,
        };
        SetTopButtonBaseStyle(btnNicItr);
        btnNicItr.MouseEnter += (_, _) => SetTopButtonHoverStyle(btnNicItr);
        btnNicItr.MouseLeave += (_, _) => SetTopButtonBaseStyle(btnNicItr);

        Button btnNicItrSave = new()
        {
            Text = "SAVE",
            Size = new Size(nicSaveButtonWidth, UiScale(24)),
            Location = new Point(btnNicItr.Right + nicButtonGap, btnNicItr.Top),
            FlatStyle = FlatStyle.Flat,
            Font = _blockFont,
            UseVisualStyleBackColor = false,
            Cursor = Cursors.Hand,
            Visible = showNicItr,
        };
        SetTopButtonBaseStyle(btnNicItrSave);
        btnNicItrSave.MouseEnter += (_, _) => SetTopButtonHoverStyle(btnNicItrSave);
        btnNicItrSave.MouseLeave += (_, _) => SetTopButtonBaseStyle(btnNicItrSave);

        Label lblNicItrStatus = new HighlightLabel()
        {
            Text = "current: reading...",
            HighlightText = "current:",
            HighlightColor = _statusPrefix,
            AutoSize = false,
            Size = new Size(nicStatusWidth, UiScale(22)),
            ForeColor = _statusInactive,
            Location = new Point(valueX, Math.Max(txtNicItr.Bottom, btnNicItr.Bottom) + UiScale(4)),
            Visible = showNicItr,
            UseMnemonic = false,
        };

        Label lblNicItrTime = new()
        {
            Text = "time: reading...",
            AutoSize = false,
            Size = new Size(nicStatusWidth, nicDetailHeight),
            ForeColor = _statusInactive,
            Location = new Point(valueX, lblNicItrStatus.Bottom + UiScale(2)),
            Visible = showNicItr,
            UseMnemonic = false,
        };

        if (showNicItr)
        {
            settingsPanel.Controls.AddRange([lblNicItr, txtNicItr, btnNicItr, btnNicItrSave, lblNicItrStatus, lblNicItrTime]);
            rowTop = lblNicItrTime.Bottom + rowGap;
        }

        bool showRawMouseThrottle = HasMouseThrottleContext(device);
        Label lblRawMouseThrottle = new()
        {
            Text = "Mouse Throttle:",
            AutoSize = true,
            Location = new Point(0, rowTop + labelOffset),
            ForeColor = _fgMain,
            Visible = showRawMouseThrottle,
        };

        ThemedCheckBox chkRawMouseThrottle = new()
        {
            Text = "Enabled",
            AutoSize = false,
            Location = new Point(valueX, rowTop + UiScale(2)),
            Size = UiScale(104, 22),
            BackColor = _bgGroup,
            ForeColor = _fgMain,
            Cursor = Cursors.Hand,
            Visible = showRawMouseThrottle,
            BorderColor = _border,
            BoxBackColor = _bgGroup,
            CheckedBackColor = Color.FromArgb(12, 12, 15),
            HoverBackColor = Color.FromArgb(22, 22, 26),
            PressedBackColor = Color.FromArgb(34, 34, 40),
            CheckColor = _fgMain,
        };

        ThemedDropDownPicker cmbRawMouseThrottle = new()
        {
            Location = new Point(chkRawMouseThrottle.Right + UiScale(10), rowTop),
            Size = UiScale(86, 26),
            BackColor = Color.FromArgb(18, 18, 22),
            ForeColor = _fgMain,
            Font = _blockFont,
            BorderColor = _border,
            ButtonColor = Color.FromArgb(14, 14, 17),
            SelectedBackColor = Color.FromArgb(48, 48, 58),
            SelectedForeColor = _fgMain,
            ArrowColor = _fgMain,
            ItemHeight = UiScale(18),
            Visible = showRawMouseThrottle,
        };
        cmbRawMouseThrottle.DropDownWidth = cmbRawMouseThrottle.Width;
        cmbRawMouseThrottle.MaxDropDownItems = 7;

        Label lblRawMouseThrottleStatus = new HighlightLabel()
        {
            Text = "current: off",
            HighlightText = "current:",
            HighlightColor = _statusPrefix,
            AutoSize = true,
            ForeColor = _statusInactive,
            Location = new Point(cmbRawMouseThrottle.Right + UiScale(8), rowTop + labelOffset),
            Visible = showRawMouseThrottle,
        };

        if (showRawMouseThrottle)
        {
            settingsPanel.Controls.AddRange([lblRawMouseThrottle, chkRawMouseThrottle, cmbRawMouseThrottle, lblRawMouseThrottleStatus]);
            rowTop = cmbRawMouseThrottle.Bottom + rowGap;
        }

        int imodCheckSize = UiScale(14);
        int imodCheckGap = UiScale(4);

        CheckBox chkImod = new()
        {
            AutoSize = false,
            Size = new Size(imodCheckSize, imodCheckSize),
            Checked = false,
            BackColor = _bgGroup,
            ForeColor = _fgMain,
            Cursor = Cursors.Hand,
            UseVisualStyleBackColor = false,
        };

        Label lblImod = new()
        {
            Text = "IMOD Value:",
            AutoSize = true,
            ForeColor = _fgMain,
        };

        Label lblImodMode = new()
        {
            Text = "Mode:",
            AutoSize = true,
            ForeColor = _fgMain,
        };

        ThemedDropDownPicker cmbImodMode = new()
        {
            Size = UiScale(156, 26),
            BackColor = Color.FromArgb(18, 18, 22),
            ForeColor = _fgMain,
            Font = _blockFont,
            BorderColor = _border,
            ButtonColor = Color.FromArgb(14, 14, 17),
            SelectedBackColor = Color.FromArgb(48, 48, 58),
            SelectedForeColor = _fgMain,
            ArrowColor = _fgMain,
            ItemHeight = UiScale(18),
        };
        cmbImodMode.Items.AddRange([ImodModeSingle, ImodModeVector, ImodModeRoles]);
        cmbImodMode.SelectedIndex = 0;
        cmbImodMode.DropDownWidth = cmbImodMode.Width;
        cmbImodMode.MaxDropDownItems = 3;

        Label lblImodModeHint = new()
        {
            Text = "single value for controller",
            AutoSize = false,
            AutoEllipsis = true,
            ForeColor = _statusInactive,
        };

        TextBox txtImod = new()
        {
            Size = UiScale(360, 24),
            BackColor = Color.FromArgb(18, 18, 22),
            BorderStyle = BorderStyle.FixedSingle,
            ForeColor = _fgMain,
            TextAlign = HorizontalAlignment.Left,
            Text = "0x0",
        };

        Label lblImodDefault = new()
        {
            Text = "default: 0x0",
            AutoSize = false,
            AutoEllipsis = true,
            ForeColor = _mutedText,
            Cursor = Cursors.Hand,
        };
        lblImodDefault.Click += (_, _) =>
        {
            if (lblImodDefault.Tag is string value && !string.IsNullOrWhiteSpace(value))
            {
                txtImod.Text = value;
            }
        };

        string? priorImodCurrent = null;
        if (priorImodStatuses is not null
            && !string.IsNullOrWhiteSpace(device.InstanceId)
            && priorImodStatuses.TryGetValue(NormalizeInstanceId(device.InstanceId), out string? cachedImodStatus)
            && !string.IsNullOrWhiteSpace(cachedImodStatus)
            && !cachedImodStatus.Equals("current: reading...", StringComparison.OrdinalIgnoreCase))
        {
            priorImodCurrent = cachedImodStatus;
        }

        Label lblImodCurrent = new HighlightLabel()
        {
            Text = priorImodCurrent ?? "current: -",
            HighlightText = "current:",
            HighlightColor = _statusPrefix,
            AutoSize = false,
            AutoEllipsis = true,
            ForeColor = _statusInactive,
        };

        ImodMapTextBox lblImodMap = new()
        {
            Text = "devices: -",
            BackColor = _bgGroup,
            ForeColor = _mutedText,
            PrefixColor = _statusPrefix,
            RoleColor = _statusActive,
            ValueColor = _fgMain,
            TabStop = false,
            Cursor = Cursors.IBeam,
        };

        lblImodMap.DoubleClick += (_, _) =>
        {
            string text = lblImodMap.Tag as string ?? lblImodMap.Text;
            if (!string.IsNullOrWhiteSpace(text))
            {
                Clipboard.SetText(text);
                ShowCopiedToolTip(lblImodMap);
            }
        };

        const string imodRoleTemplate = "Mouse=0x0, Keyboard=0xC8, Audio=0xFA0, Gamepad=0xC8";
        Label lblImodHelp = new()
        {
            Text = "input: hex / list / roles",
            AutoSize = false,
            AutoEllipsis = true,
            ForeColor = _mutedText,
        };

        IReadOnlyList<string> imodDeviceEditorRoles = GetImodDeviceEditorRoles(device);
        Dictionary<string, TextBox> imodDeviceEditorBoxes = new(StringComparer.OrdinalIgnoreCase);
        bool suppressImodDeviceEditor = false;
        Panel imodDeviceEditorPanel = new()
        {
            BackColor = _bgGroup,
            Visible = imodDeviceEditorRoles.Count > 0,
        };

        void SyncImodDeviceEditorFromText()
        {
            if (imodDeviceEditorBoxes.Count == 0 || suppressImodDeviceEditor)
            {
                return;
            }

            bool hasRoleValues = TryParseImodRoleIntervals(txtImod.Text ?? string.Empty, out Dictionary<string, uint> roleValues)
                && roleValues.Count > 0;
            uint? singleValue = null;
            string imodText = txtImod.Text ?? string.Empty;
            if (!hasRoleValues
                && !string.IsNullOrWhiteSpace(imodText)
                && TryParseImodInterval(imodText, ImodDefaultInterval, out uint parsedSingle))
            {
                singleValue = parsedSingle;
            }

            suppressImodDeviceEditor = true;
            try
            {
                foreach (KeyValuePair<string, TextBox> pair in imodDeviceEditorBoxes)
                {
                    uint value = hasRoleValues && roleValues.TryGetValue(pair.Key, out uint roleValue)
                        ? roleValue
                        : singleValue ?? GetImodDeviceEditorDefaultValue(pair.Key);
                    pair.Value.Text = FormatImodValue(value);
                    pair.Value.ForeColor = _fgMain;
                }
            }
            finally
            {
                suppressImodDeviceEditor = false;
            }
        }

        void UpdateImodTextFromDeviceEditor()
        {
            if (suppressImodDeviceEditor || imodDeviceEditorBoxes.Count == 0)
            {
                return;
            }

            Dictionary<string, uint> roleValues = new(StringComparer.OrdinalIgnoreCase);
            bool valid = true;
            foreach (string role in imodDeviceEditorRoles)
            {
                if (!imodDeviceEditorBoxes.TryGetValue(role, out TextBox? box))
                {
                    continue;
                }

                string text = box.Text.Trim();
                if (text.Length == 0)
                {
                    continue;
                }

                if (!TryParseImodInterval(text, ImodDefaultInterval, out uint value))
                {
                    box.ForeColor = Color.FromArgb(255, 110, 110);
                    valid = false;
                    continue;
                }

                box.ForeColor = _fgMain;
                roleValues[role] = value;
            }

            if (!valid || roleValues.Count == 0)
            {
                return;
            }

            suppressImodDeviceEditor = true;
            try
            {
                txtImod.Text = FormatImodRoleIntervals(roleValues);
                chkImod.Checked = true;
                cmbImodMode.SelectedItem = ImodModeRoles;
            }
            finally
            {
                suppressImodDeviceEditor = false;
            }
        }

        void NormalizeImodDeviceEditorBox(TextBox box)
        {
            if (TryParseImodInterval(box.Text.Trim(), ImodDefaultInterval, out uint value))
            {
                suppressImodDeviceEditor = true;
                try
                {
                    box.Text = FormatImodValue(value);
                    box.ForeColor = _fgMain;
                }
                finally
                {
                    suppressImodDeviceEditor = false;
                }

                UpdateImodTextFromDeviceEditor();
            }
        }

        if (imodDeviceEditorRoles.Count > 0)
        {
            int editorLabelWidth = UiScale(72);
            int editorBoxWidth = UiScale(78);
            int editorColumnGap = UiScale(18);
            int editorRowHeight = UiScale(28);
            int editorX = 0;
            int editorY = 0;
            for (int i = 0; i < imodDeviceEditorRoles.Count; i++)
            {
                if (i == 2)
                {
                    editorX = 0;
                    editorY += editorRowHeight;
                }

                string role = imodDeviceEditorRoles[i];
                Label roleLabel = new()
                {
                    Text = role + ":",
                    AutoSize = false,
                    Size = new Size(editorLabelWidth, UiScale(22)),
                    Location = new Point(editorX, editorY + UiScale(3)),
                    ForeColor = _fgMain,
                };

                TextBox roleBox = new()
                {
                    Size = new Size(editorBoxWidth, UiScale(24)),
                    Location = new Point(roleLabel.Right + UiScale(4), editorY),
                    BackColor = Color.FromArgb(18, 18, 22),
                    BorderStyle = BorderStyle.FixedSingle,
                    ForeColor = _fgMain,
                    Text = FormatImodValue(GetImodDeviceEditorDefaultValue(role)),
                    TextAlign = HorizontalAlignment.Left,
                };
                roleBox.TextChanged += (_, _) => UpdateImodTextFromDeviceEditor();
                roleBox.Leave += (_, _) => NormalizeImodDeviceEditorBox(roleBox);
                roleBox.KeyDown += (_, e) =>
                {
                    if (e.KeyCode == Keys.Enter)
                    {
                        NormalizeImodDeviceEditorBox(roleBox);
                        e.SuppressKeyPress = true;
                    }
                };

                imodDeviceEditorBoxes[role] = roleBox;
                imodDeviceEditorPanel.Controls.AddRange([roleLabel, roleBox]);
                editorX = roleBox.Right + editorColumnGap;
            }
        }

        txtImod.TextChanged += (_, _) => SyncImodDeviceEditorFromText();
        SyncImodDeviceEditorFromText();

        Button btnImodApply = new()
        {
            Text = "SET",
            Size = UiScale(54, 24),
            FlatStyle = FlatStyle.Flat,
            Font = _blockFont,
            UseVisualStyleBackColor = false,
            Cursor = Cursors.Hand,
        };
        SetTopButtonBaseStyle(btnImodApply);
        btnImodApply.MouseEnter += (_, _) => SetTopButtonHoverStyle(btnImodApply);
        btnImodApply.MouseLeave += (_, _) => SetTopButtonBaseStyle(btnImodApply);
        btnImodApply.Click += (_, _) =>
        {
            WriteLog("UI: SET IMOD button clicked");
            CreateDeviceTweakerBackup("pre-imod", showDialog: false);
            _ = ApplyImodSettings(out string? note);
            RefreshImodCurrentValues(reason: "imod-set");
            if (!string.IsNullOrWhiteSpace(note))
            {
                ShowThemedInfo(note);
            }
        };

        Button btnImodDelete = new()
        {
            Text = "DELETE",
            Size = UiScale(76, 24),
            FlatStyle = FlatStyle.Flat,
            Font = _blockFont,
            UseVisualStyleBackColor = false,
            Cursor = Cursors.Hand,
        };
        SetTopButtonBaseStyle(btnImodDelete);
        btnImodDelete.MouseEnter += (_, _) => SetTopButtonHoverStyle(btnImodDelete);
        btnImodDelete.MouseLeave += (_, _) => SetTopButtonBaseStyle(btnImodDelete);
        btnImodDelete.Click += (_, _) =>
        {
            WriteLog("UI: DELETE IMOD button clicked");
            ResetImodIntervalsToDefault("imod-delete");
            ShowThemedInfo("IMOD reset to defaults.\nStartup script and DTIMOD.sys removed.");
        };

        void UpdateImodDetailsVisibility()
        {
            bool showDetails = chkImod.Visible
                && (chkImod.Checked || IsImodAttentionStatus(lblImodCurrent.Text));
            lblImodCurrent.Visible = showDetails;
            lblImodDefault.Visible = showDetails;
            lblImodMap.Visible = showDetails;
        }

        void UpdateImodModeHint()
        {
            string mode = cmbImodMode.SelectedItem?.ToString() ?? ImodModeSingle;
            lblImodModeHint.Text = mode switch
            {
                ImodModeRoles => "detected USB device roles",
                ImodModeVector => "values by interrupter index",
                _ => "single value for controller",
            };
        }

        bool showImod = ShouldShowImod(device);
        if (showImod)
        {
            lblImodMode.Visible = true;
            cmbImodMode.Visible = true;
            lblImodModeHint.Visible = true;

            lblImodMode.Location = new Point(0, rowTop + labelOffset);
            cmbImodMode.Location = new Point(valueX, rowTop);
            int modeHintLeft = cmbImodMode.Right + UiScale(14);
            lblImodModeHint.Location = new Point(modeHintLeft, rowTop + labelOffset);
            lblImodModeHint.Size = new Size(
                Math.Max(UiScale(110), availableSettingsWidth - modeHintLeft - UiScale(8)),
                UiScale(18));
            UpdateImodModeHint();
            rowTop = cmbImodMode.Bottom + rowGap;

            Size imodLabelSize = TextRenderer.MeasureText(lblImod.Text, _baseFont);
            int checkY = rowTop + labelOffset + Math.Max(0, (imodLabelSize.Height - imodCheckSize) / 2);
            chkImod.Location = new Point(0, checkY);
            lblImod.Location = new Point(imodCheckSize + imodCheckGap, rowTop + labelOffset);
            int buttonGap = UiScale(6);
            int inlineGap = UiScale(8);
            int buttonTotalWidth = btnImodApply.Width + btnImodDelete.Width + buttonGap;
            int inlineInputWidth = availableSettingsWidth - valueX - buttonTotalWidth - inlineGap;
            bool useInlineImodButtons = inlineInputWidth >= UiScale(180);
            int imodInputWidth = useInlineImodButtons
                ? Math.Min(UiScale(300), inlineInputWidth)
                : Math.Min(UiScale(300), Math.Max(UiScale(120), availableSettingsWidth - valueX - UiScale(8)));
            txtImod.Width = imodInputWidth;
            txtImod.Location = new Point(valueX, rowTop);
            lblImodHelp.Visible = false;
            lblImodHelp.Location = new Point(txtImod.Left, txtImod.Bottom + UiScale(4));
            lblImodHelp.Size = new Size(UiScale(180), UiScale(18));
            int imodButtonTop = useInlineImodButtons
                ? txtImod.Top + Math.Max(0, (txtImod.Height - btnImodApply.Height) / 2)
                : txtImod.Bottom + UiScale(6);
            int imodButtonLeft = useInlineImodButtons ? txtImod.Right + inlineGap : 0;
            btnImodApply.Location = new Point(imodButtonLeft, imodButtonTop);
            btnImodDelete.Location = new Point(btnImodApply.Right + buttonGap, imodButtonTop);
            int statusTop = Math.Max(txtImod.Bottom, btnImodApply.Bottom) + UiScale(6);
            if (imodDeviceEditorPanel.Visible)
            {
                imodDeviceEditorPanel.Location = new Point(valueX, statusTop - UiScale(1));
                int editorRows = imodDeviceEditorRoles.Count > 2 ? 2 : 1;
                imodDeviceEditorPanel.Size = new Size(
                    Math.Max(UiScale(80), availableSettingsWidth - valueX - UiScale(8)),
                    UiScale(editorRows * 26));
                statusTop = imodDeviceEditorPanel.Bottom + UiScale(3);
            }

            int imodDetailsX = UiScale(28);
            int imodStatusTop = statusTop + UiScale(4);
            lblImodCurrent.Location = new Point(imodDetailsX, imodStatusTop);
            int defaultStatusWidth = UiScale(120);
            int statusGap = UiScale(14);
            int currentStatusWidth = Math.Max(
                UiScale(90),
                Math.Min(UiScale(260), availableSettingsWidth - imodDetailsX - defaultStatusWidth - statusGap));
            lblImodCurrent.Size = new Size(currentStatusWidth, UiScale(18));
            lblImodDefault.Location = new Point(lblImodCurrent.Right + UiScale(14), imodStatusTop);
            lblImodDefault.Size = new Size(defaultStatusWidth, UiScale(18));
            int imodMapRows = imodDeviceEditorRoles.Contains("Gamepad", StringComparer.OrdinalIgnoreCase) ? 7 : 6;
            lblImodMap.Location = new Point(imodDetailsX, lblImodCurrent.Bottom + UiScale(14));
            lblImodMap.Size = new Size(Math.Max(UiScale(120), availableSettingsWidth - imodDetailsX), UiScale(imodMapRows * 14));
            settingsPanel.Controls.AddRange([lblImodMode, cmbImodMode, lblImodModeHint, chkImod, lblImod, txtImod, btnImodApply, btnImodDelete, imodDeviceEditorPanel, lblImodCurrent, lblImodDefault, lblImodMap]);
            _copyToolTip.SetToolTip(txtImod, $"Supported IMOD input:\n0xC8\n0xC8, 0xFA0\n{imodRoleTemplate}");
            _copyToolTip.SetToolTip(cmbImodMode, "XHCI writes one value to all interrupters on this USB host controller. Devices maps detected devices to their interrupter. Interrupters writes values by index.");
            _copyToolTip.SetToolTip(lblImodModeHint, "Shows how the selected IMOD mode interprets IMOD Value.");
            _copyToolTip.SetToolTip(btnImodApply, "Apply the current IMOD configuration and read back hardware values.");
            _copyToolTip.SetToolTip(lblImodHelp, "Role mode applies values to the detected interrupter for each USB role.");
            _copyToolTip.SetToolTip(lblImodDefault, "Click to copy the default IMOD interval into the input field.");
            _copyToolTip.SetToolTip(lblImodMap, string.Empty);
            UpdateImodDetailsVisibility();
            rowTop = Math.Max(lblImodMap.Bottom, lblImodCurrent.Bottom) + rowGap;
        }
        else
        {
            chkImod.Visible = false;
            lblImod.Visible = false;
            txtImod.Visible = false;
            lblImodMode.Visible = false;
            cmbImodMode.Visible = false;
            lblImodModeHint.Visible = false;
            lblImodDefault.Visible = false;
            lblImodCurrent.Visible = false;
            btnImodApply.Visible = false;
            btnImodDelete.Visible = false;
        }

        int settingsContentRight = 0;
        int settingsContentBottom = 0;
        foreach (Control child in settingsPanel.Controls)
        {
            if (!child.Visible)
            {
                continue;
            }

            settingsContentRight = Math.Max(settingsContentRight, child.Right);
            settingsContentBottom = Math.Max(settingsContentBottom, child.Bottom);
        }

        UpdateResponsivePlacement();

        int settingsMinWidth = Math.Min(UiScale(420), availableSettingsWidth);
        Size settingsSize = new(
            Math.Min(Math.Max(settingsContentRight + UiScale(8), settingsMinWidth), availableSettingsWidth),
            Math.Max(settingsContentBottom + UiScale(8), UiScale(24)));
        settingsPanel.Size = settingsSize;

        int settingsTop = cpuPanel.Top + Math.Max(0, (cpuPanel.Height - settingsSize.Height) / 2);
        settingsPanel.Location = new Point(settingsX, settingsTop);

        int infoY = Math.Max(lblIrq.Bottom + UiScale(14), cpuPanel.Bottom + UiScale(18));
        infoY = Math.Max(infoY, settingsPanel.Bottom + UiScale(10));
        InfoTextBox lblInfo = new()
        {
            Text = "PNP ID: -",
            Location = new Point(UiScale(18), infoY),
            Size = new Size(grp.Width - UiScale(40), UiScale(70)),
            Cursor = Cursors.Hand,
            BackColor = _bgGroup,
            ForeColor = _mutedText,
            PrefixColor = _statusPrefix,
            ValueColor = _fgMain,
            SeparatorColor = _statusSeparator,
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            TabStop = false,
        };
        lblInfo.Click += (_, _) =>
        {
            if (lblInfo.Tag is string txt && !string.IsNullOrWhiteSpace(txt))
            {
                Clipboard.SetText(txt);
                ShowCopiedToolTip(lblInfo);
            }
        };

        void RelayoutDeviceBlockChrome()
        {
            UpdateResponsivePlacement();

            int visibleSettingsRight = 0;
            int visibleSettingsBottom = 0;
            foreach (Control child in settingsPanel.Controls)
            {
                if (!child.Visible)
                {
                    continue;
                }

                visibleSettingsRight = Math.Max(visibleSettingsRight, child.Right);
                visibleSettingsBottom = Math.Max(visibleSettingsBottom, child.Bottom);
            }

            UpdateResponsivePlacement();

            int currentAvailableSettingsWidth = Math.Max(UiScale(260), grp.Width - settingsX - UiScale(24));
            int currentSettingsMinWidth = Math.Min(UiScale(420), currentAvailableSettingsWidth);
            Size currentSettingsSize = new(
                Math.Min(Math.Max(visibleSettingsRight + UiScale(8), currentSettingsMinWidth), currentAvailableSettingsWidth),
                Math.Max(visibleSettingsBottom + UiScale(8), UiScale(24)));
            settingsPanel.Size = currentSettingsSize;

            int currentSettingsTop = cpuPanel.Top + Math.Max(0, (cpuPanel.Height - currentSettingsSize.Height) / 2);
            settingsPanel.Location = new Point(settingsX, currentSettingsTop);

            int currentInfoY = Math.Max(lblIrq.Bottom + UiScale(14), cpuPanel.Bottom + UiScale(18));
            currentInfoY = Math.Max(currentInfoY, settingsPanel.Bottom + UiScale(10));
            lblInfo.Location = new Point(lblInfo.Left, currentInfoY);
            lblInfo.Size = new Size(grp.Width - UiScale(40), lblInfo.Height);

            grp.Height = Math.Max(
                Math.Max(cpuPanel.Bottom + UiScale(110), settingsPanel.Bottom + UiScale(20)),
                lblInfo.Bottom + UiScale(20));
        }

        grp.Controls.AddRange(
        [
            headerPanel,
            divider,
            cpuLabel,
            cpuPanel,
            lblMask,
            lblIrq,
            settingsPanel,
            lblInfo,
        ]);
        WireDevicesMouseWheelForwarding(grp);

        RelayoutDeviceBlockChrome();

        DeviceBlock block = new()
        {
            Device = device,
            Kind = device.Kind,
            Group = grp,
            HeaderPanel = headerPanel,
            Divider = divider,
            CpuTitleLabel = cpuLabel,
            CpuPanel = cpuPanel,
            SettingsPanel = settingsPanel,
            TitleLabel = headerLabel,
            CpuBoxes = cpuBoxes,
            AffinityLabel = lblMask,
            IrqLabel = lblIrq,
            MsiCombo = cmbMsi,
            LimitBox = txtLimit,
            PrioCombo = cmbPrio,
            PolicyCombo = cmbPolicy,
            PolicyLabel = lblPolicy,
            NdisModeLabel = device.Kind == DeviceKind.NET_NDIS ? lblNdisMode : null,
            NdisModeCombo = device.Kind == DeviceKind.NET_NDIS ? cmbNdisMode : null,
            RssQueueBox = device.Kind == DeviceKind.NET_NDIS ? nudRssQueues : null,
            NicItrBox = showNicItr ? txtNicItr : null,
            NicItrStatusLabel = showNicItr ? lblNicItrStatus : null,
            NicItrTimeLabel = showNicItr ? lblNicItrTime : null,
            NicItrApplyButton = showNicItr ? btnNicItr : null,
            NicItrSaveButton = showNicItr ? btnNicItrSave : null,
            ImodAutoCheck = chkImod,
            ImodModeCombo = showImod ? cmbImodMode : null,
            ImodBox = txtImod,
            ImodDefaultLabel = lblImodDefault,
            ImodCurrentLabel = lblImodCurrent,
            ImodMapLabel = lblImodMap,
            RawMouseThrottleCheck = showRawMouseThrottle ? chkRawMouseThrottle : null,
            RawMouseThrottleCombo = showRawMouseThrottle ? cmbRawMouseThrottle : null,
            RawMouseThrottleStatusLabel = showRawMouseThrottle ? lblRawMouseThrottleStatus : null,
            InfoLabel = lblInfo,
            RelayoutAction = RelayoutDeviceBlockChrome,
            AffinityMask = 0,
            IrqCount = null,
        };

        foreach (CheckBox cb in cpuBoxes)
        {
            cb.CheckedChanged += (_, _) =>
            {
                if (block.SuppressCpuEvents == 0)
                {
                    if (block.Kind == DeviceKind.NET_NDIS)
                    {
                        HandleNdisCheckboxChanged(block, cb);
                    }
                    else
                    {
                        RecalcAffinityMask(block);
                    }
                }
            };
        }

        if (block.Kind == DeviceKind.NET_NDIS && block.RssQueueBox is not null)
        {
            block.RssQueueBox.ValueChanged += (_, _) =>
            {
                if (block.SuppressCpuEvents > 0)
                {
                    return;
                }

                int baseCore = block.RssBaseCore ?? GetFirstCheckedCore(block) ?? 0;
                int queues = ClampRssQueueCount((int)block.RssQueueBox.Value);
                ApplyNdisSelection(block, baseCore, queues);
            };
        }

        if (block.NdisModeCombo is not null)
        {
            block.NdisModeCombo.SelectedIndexChanged += (_, _) => UpdateBlockInfoText(block);
        }

        if (block.RawMouseThrottleCheck is not null && block.RawMouseThrottleCombo is not null)
        {
            block.RawMouseThrottleCheck.CheckedChanged += (_, _) => HandleRawMouseThrottleChanged(block);
            block.RawMouseThrottleCombo.SelectedIndexChanged += (_, _) => HandleRawMouseThrottleChanged(block);
        }

        if (block.NicItrApplyButton is not null)
        {
            block.NicItrApplyButton.Click += (_, _) => ApplyNicItrFromBlock(block);
        }

        if (block.NicItrSaveButton is not null)
        {
            block.NicItrSaveButton.Click += (_, _) => SaveNicItrPersistenceFromBlock(block);
        }

        if (block.NicItrBox is not null)
        {
            block.NicItrBox.TextChanged += (_, _) => UpdateNicItrInputTimeLabel(block);
        }

        if (showImod)
        {
            InitializeImodSelectors(block);
            block.ImodAutoCheck.CheckedChanged += (_, _) =>
            {
                UpdateImodDetailsVisibility();
                RelayoutDeviceBlockChrome();
                LayoutBlocks();
            };
            block.ImodModeCombo!.SelectedIndexChanged += (_, _) =>
            {
                HandleImodModeChanged(block);
                UpdateImodModeHint();
            };
            block.ImodBox.TextChanged += (_, _) =>
            {
                if (block.SuppressImodEvents == 0)
                {
                    UpdateImodSelectorsFromText(block);
                }
            };
        }

        LoadBlockSettings(block);
        if (showImod)
        {
            UpdateImodDetailsVisibility();
            RelayoutDeviceBlockChrome();
        }

        allowWindowAutoExpand = false;
        _devicesPanel.Controls.Add(grp);
        _blocks.Add(block);
    }

    private void LayoutBlocks()
    {
        int paddingX = UiScale(24);
        int gapY = UiScale(18);
        int y = UiScale(12);
        bool firstPlaced = true;

        Panel? reserved = _reservedCpuPanel;
        bool reservedInserted = false;

        DeviceBlock? lastStorBlock = null;
        if (reserved is not null && _blocks.Count > 0)
        {
            List<DeviceBlock> storBlocks = _blocks.Where(b => b.Kind == DeviceKind.STOR).ToList();
            if (storBlocks.Count > 0)
            {
                lastStorBlock = storBlocks[^1];
            }
        }

        DeviceBlock? lastBlock = _blocks.Count > 0 ? _blocks[^1] : null;

        void PlaceControl(Control control)
        {
            if (!firstPlaced)
            {
                y += gapY;
            }

            control.Location = new Point(paddingX, y);
            y += control.Height;
            firstPlaced = false;
        }

        foreach (DeviceBlock b in _blocks)
        {
            int width = GetDevicesViewportWidth() - paddingX - UiScale(12);
            if (width < UiScale(360))
            {
                width = UiScale(360);
            }

            b.Group.Width = width;
            RelayoutDeviceBlockChrome(b);
            if (firstPlaced && _devicesHost is not null)
            {
                int maxFirstY = _devicesHost.ClientSize.Height - b.Group.Height - UiScale(2);
                if (maxFirstY < y)
                {
                    y = Math.Max(UiScale(6), maxFirstY);
                }
            }

            int currentHeight = b.InfoLabel.Height > 0 ? b.InfoLabel.Height : UiScale(60);
            int infoWidth = Math.Max(UiScale(140), b.Group.Width - b.InfoLabel.Left - UiScale(24));
            b.InfoLabel.Size = new Size(infoWidth, currentHeight);

            PlaceControl(b.Group);

            if (reserved is not null && !reservedInserted)
            {
                bool placeHere = false;
                if (lastStorBlock is not null && ReferenceEquals(b, lastStorBlock))
                {
                    placeHere = true;
                }
                else if (lastStorBlock is null && lastBlock is not null && ReferenceEquals(b, lastBlock))
                {
                    placeHere = true;
                }

                if (placeHere)
                {
                    UpdateReservedCpuSetsPanelLayout(reserved, width);
                    PlaceControl(reserved);
                    reservedInserted = true;
                }
            }
        }

        if (reserved is not null && !reservedInserted)
        {
            int width = GetDevicesViewportWidth() - paddingX - UiScale(12);
            if (width < UiScale(360))
            {
                width = UiScale(360);
            }

            UpdateReservedCpuSetsPanelLayout(reserved, width);
            PlaceControl(reserved);
        }

        int bottomPadding = UiScale(32);
        int contentHeight = y + bottomPadding;
        if (_devicesPanel.Height != contentHeight)
        {
            _devicesPanel.Height = contentHeight;
        }

        SyncDevicesScrollBar();
    }

    private void RefreshBlocks(bool includeImodReadback = true)
    {
        InvalidateImodCache();
        _ndisRssRuntimeCache.Clear();
        Dictionary<string, string> priorImodStatuses = [];
        foreach (DeviceBlock block in _blocks)
        {
            if (block.ImodCurrentLabel is null || string.IsNullOrWhiteSpace(block.Device.InstanceId))
            {
                continue;
            }

            string statusText = block.ImodCurrentLabel.Text ?? string.Empty;
            if (statusText.Equals("current: reading...", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            priorImodStatuses[NormalizeInstanceId(block.Device.InstanceId)] = statusText;
        }

        _devicesPanel.SuspendLayout();
        try
        {
            _devicesPanel.Controls.Clear();
            _devicesPanel.Location = new Point(0, 0);
            if (_devicesScroll is not null)
            {
                _devicesScroll.Value = 0;
            }
            _blocks.Clear();
            _reservedCpuPanel = null;

            List<DeviceInfo> devs = GetDeviceList();
            WarnIfMissingGpuDriver(devs);
            int index = 0;
            foreach (DeviceInfo d in devs)
            {
                NewDeviceBlock(d, index, priorImodStatuses);
                index++;
            }

            _reservedCpuPanel = NewReservedCpuSetsPanel();
            if (_reservedCpuPanel is not null)
            {
                _devicesPanel.Controls.Add(_reservedCpuPanel);
            }

            LayoutBlocks();
        }
        finally
        {
            _devicesPanel.ResumeLayout();
        }

        AdjustInitialDeviceViewportHeight();
        _devicesPanel.Invalidate(true);
        _devicesHost.Invalidate(true);

        if (includeImodReadback)
        {
            RefreshImodCurrentValues(reason: "refresh-blocks");
        }
        LogGuiSnapshot("refresh");
    }

    private void WarnIfMissingGpuDriver(IReadOnlyList<DeviceInfo> devices)
    {
        if (_testDevicesEnabled && _testDevicesOnly)
        {
            return;
        }

        bool hasDriver = HasAmdOrNvidiaGpu(devices);
        if (_lastGpuDriverDetected == hasDriver)
        {
            return;
        }

        _lastGpuDriverDetected = hasDriver;
        if (hasDriver)
        {
            _pendingGpuDriverWarning = false;
            return;
        }

        if (!Visible || !IsHandleCreated)
        {
            _pendingGpuDriverWarning = true;
            return;
        }

        ShowMissingGpuDriverWarning();
    }

    private static bool HasAmdOrNvidiaGpu(IReadOnlyList<DeviceInfo> devices)
    {
        foreach (DeviceInfo device in devices)
        {
            if (device.IsTestDevice || device.Kind != DeviceKind.GPU)
            {
                continue;
            }

            string id = device.InstanceId ?? string.Empty;
            if (id.Contains("VEN_10DE", StringComparison.OrdinalIgnoreCase)
                || id.Contains("VEN_1002", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    private List<string> BuildIrqLookupKeys(DeviceBlock block)
    {
        List<string> keys = [];
        HashSet<string> seen = new(StringComparer.OrdinalIgnoreCase);

        void AddKey(string key)
        {
            if (!string.IsNullOrWhiteSpace(key) && seen.Add(key))
            {
                keys.Add(key);
            }
        }

        AddKey(GetIrqPnpKey(block.Device.InstanceId));
        AddKey(GetIrqPnpKey(block.Device.RegBase));
        AddKey(GetIrqPnpKey(GetDisplayRegPath(block.Device.InstanceId)));

        return keys;
    }

    private bool TryFindIrqInfo(
        DeviceBlock block,
        IReadOnlyDictionary<string, DeviceIrqInfo> irqCounts,
        out DeviceIrqInfo? info,
        out string matchedKey)
    {
        foreach (string key in BuildIrqLookupKeys(block))
        {
            if (irqCounts.TryGetValue(key, out info))
            {
                matchedKey = key;
                return true;
            }
        }

        info = null;
        matchedKey = string.Empty;
        return false;
    }

    private static bool IsKnownMsiStatus(string status)
    {
        return status.Equals("Enabled", StringComparison.OrdinalIgnoreCase)
            || status.Equals("Disabled", StringComparison.OrdinalIgnoreCase);
    }

    private async void CalculateIrqCounts()
    {
        foreach (DeviceBlock b in _blocks)
        {
            b.IrqLabel.Text = "IRQ Count: reading...";
        }

        Dictionary<string, DeviceIrqInfo> irqCounts;
        try
        {
            irqCounts = await Task.Run(GetDeviceIrqCounts);
        }
        catch (Exception ex)
        {
            WriteLog($"IRQ.MAP: failed to read active IRQ data: {ex.Message}");
            ShowThemedInfo($"IRQ count read failed.\n{ex.Message}");
            return;
        }

        if (IsDisposed)
        {
            return;
        }

        foreach (DeviceBlock b in _blocks)
        {
            string shortPnp = GetShortPnpId(b.Device.InstanceId);
            if (TryFindIrqInfo(b, irqCounts, out DeviceIrqInfo? info, out string matchedKey) && info is not null)
            {
                string regMsiStatus = ReadMsiStatusFromRegistry(b);
                b.IrqCount = info.Count;
                b.IrqLabel.Text = $"IRQ Count: {info.Count} (MSI: {info.MsiStatus})";
                WriteLog($"IRQ.MAP: {b.Device.InstanceId} ({shortPnp}) -> count={info.Count} activeMsi={info.MsiStatus} registryMsi={regMsiStatus} irqs=[{FormatIrqNumbers(info.IrqNumbers)}] key={matchedKey} source={info.Source}");
                if (IsKnownMsiStatus(info.MsiStatus)
                    && IsKnownMsiStatus(regMsiStatus)
                    && !info.MsiStatus.Equals(regMsiStatus, StringComparison.OrdinalIgnoreCase))
                {
                    WriteLog($"IRQ.MSI.MISMATCH: {b.Device.InstanceId} ({shortPnp}) activeMsi={info.MsiStatus} registryMsi={regMsiStatus} activeIrqs=[{FormatIrqNumbers(info.IrqNumbers)}] key={matchedKey}");
                }
            }
            else
            {
                string regMsiStatus = ReadMsiStatusFromRegistry(b);
                bool isNetwork = b.Kind is DeviceKind.NET_NDIS or DeviceKind.NET_CX;
                b.IrqCount = isNetwork ? null : 0;
                b.IrqLabel.Text = isNetwork
                    ? $"IRQ Count: N/A (MSI: {regMsiStatus})"
                    : $"IRQ Count: 0 (MSI: {regMsiStatus})";
                string triedKeys = string.Join(", ", BuildIrqLookupKeys(b));
                WriteLog($"IRQ.MAP: {b.Device.InstanceId} ({shortPnp}) -> {(isNetwork ? "N/A" : "0")} activeMsi=Unknown registryMsi={regMsiStatus} irqs=[none] source=registry-fallback keys=[{triedKeys}]");
            }
        }

        LogGuiSnapshot("irq-counts");
    }
}
