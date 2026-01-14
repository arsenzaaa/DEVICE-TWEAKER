namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private const string StorageAffinityNoteText = "(affinity masks not supported on SSD/HDD)";

    private int GetDevicesViewportWidth()
    {
        int width = _devicesHost.ClientSize.Width;
        if (width <= 0)
        {
            width = _devicesPanel.ClientSize.Width;
        }

        return width;
    }

    private static void FixRssPolicyLabelOverlap(DeviceBlock block)
    {
        int desiredLeft = block.PolicyLabel.Right + 8;
        if (desiredLeft <= block.PolicyCombo.Left)
        {
            return;
        }

        int rightPadding = 24;
        int parentLeft = block.PolicyCombo.Parent?.Left ?? 0;
        int maxWidth = block.Group.ClientSize.Width - rightPadding - parentLeft - desiredLeft;
        int newWidth = block.PolicyCombo.Width;
        if (maxWidth < newWidth)
        {
            newWidth = maxWidth;
        }

        if (newWidth < 90)
        {
            newWidth = 90;
        }

        block.PolicyCombo.Location = new Point(desiredLeft, block.PolicyCombo.Top);
        block.PolicyCombo.Width = newWidth;
    }

    private string BuildDeviceBlockTitle(DeviceInfo device)
    {
        string title = device.Name;
        if (device.Kind == DeviceKind.USB && !string.IsNullOrWhiteSpace(device.UsbRoles))
        {
            title = $"{device.Name} [{device.UsbRoles}]";
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

    private void NewDeviceBlock(DeviceInfo device, int index)
    {
        Panel grp = new();

        string title = BuildDeviceBlockTitle(device);
        string logTitle = device.Kind == DeviceKind.STOR ? $"{title} {StorageAffinityNoteText}" : title;

        WriteLog(
            $"UI.BLOCK: idx={index} title=\"{logTitle}\" kind={device.Kind} id={device.InstanceId} roles=\"{device.UsbRoles}\" audio=\"{device.AudioEndpoints}\" storage=\"{device.StorageTag}\"");

        grp.Width = GetDevicesViewportWidth() - 40;
        grp.Height = _grpHeight;
        grp.BackColor = _bgGroup;
        grp.ForeColor = _fgMain;
        grp.Font = _blockFont;
        grp.Margin = new Padding(0);
        grp.Padding = new Padding(12, 16, 12, 16);

        grp.Paint += (_, e) =>
        {
            Rectangle rect = grp.ClientRectangle;
            rect.Width -= 1;
            rect.Height -= 1;
            using Pen pen = new(_border);
            e.Graphics.DrawRectangle(pen, rect);
        };

        FlowLayoutPanel headerPanel = new()
        {
            AutoSize = false,
            Size = new Size(grp.Width - 40, 24),
            Location = new Point(18, 8),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            WrapContents = false,
            Margin = Padding.Empty,
            Padding = Padding.Empty,
            BackColor = Color.Transparent,
        };

        Label headerLabel = new()
        {
            Text = title,
            Font = _blockTitleFont,
            ForeColor = _fgMain,
            AutoSize = true,
            Margin = Padding.Empty,
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
                Margin = new Padding(6, 2, 0, 0),
            };
            headerPanel.Controls.Add(headerNote);
        }

        Panel divider = new()
        {
            BackColor = _border,
            Size = new Size(grp.Width - 32, 1),
            Location = new Point(16, 36),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
        };

        int contentTop = 48;
        Label cpuLabel = new()
        {
            Text = "CPU Affinity",
            AutoSize = true,
            ForeColor = _fgMain,
            Location = new Point(18, contentTop),
        };
        if (device.Kind == DeviceKind.STOR)
        {
            cpuLabel.ForeColor = _mutedText;
        }

        int cpuPanelTop = cpuLabel.Bottom + 6;
        int cpuPanelHeight = 150;
        Panel cpuPanel = new()
        {
            Location = new Point(16, cpuPanelTop),
            Size = new Size(308, cpuPanelHeight),
            BackColor = _bgForm,
            Padding = new Padding(8, 6, 8, 6),
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
        int checkSpacing = 22;

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
            List<(int Lp, CheckBox Control, int Ccd, int Eff)> pItems = items.Where(m => m.Eff == 0).ToList();
            List<(int Lp, CheckBox Control, int Ccd, int Eff)> eItems = items.Where(m => m.Eff > 0).ToList();
            List<(int Lp, CheckBox Control, int Ccd, int Eff)> other = items.Except(pItems).Except(eItems).ToList();
            List<(int Lp, CheckBox Control, int Ccd, int Eff)> ordered = [.. pItems, .. eItems, .. other];
            columns.Add(ordered);
        }

        int columnGap = 16;
        int startX = 10;
        int minColumnWidth = 120;

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
                int w = ordered.Max(o => o.Control.PreferredSize.Width);
                if (w > 0)
                {
                    maxWidth = Math.Max(minColumnWidth, w + 10);
                }
            }

            columnWidths.Add(maxWidth);
        }

        for (int i = 0; i < columns.Count; i++)
        {
            List<(int Lp, CheckBox Control, int Ccd, int Eff)> ordered = columns[i];
            int maxWidth = columnWidths[i];
            int y = 4;
            foreach ((int _, CheckBox control, int _, int _) in ordered)
            {
                control.Location = new Point(runningX, y);
                y += checkSpacing;
            }

            runningX += maxWidth + columnGap;
        }

        int requiredWidth = columns.Count == 0 ? cpuPanel.Width : runningX - columnGap;
        if (requiredWidth > cpuPanel.ClientSize.Width)
        {
            cpuPanel.AutoScroll = true;
            cpuPanel.AutoScrollMinSize = new Size(requiredWidth + startX, cpuPanel.Height);
        }
        else
        {
            cpuPanel.AutoScroll = false;
            cpuPanel.AutoScrollMinSize = Size.Empty;
        }

        int desiredHeight = Math.Max((maxColumnCount * checkSpacing) + 18, 150);
        if (cpuPanel.Height != desiredHeight)
        {
            cpuPanel.Height = desiredHeight;
        }

        int maskY = cpuPanel.Bottom + 10;
        Label lblMask = new()
        {
            Text = "Affinity Mask: 0x0",
            AutoSize = true,
            ForeColor = _accent,
            Location = new Point(18, maskY),
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
            Location = new Point(18, maskY + 20),
        };

        int settingsX = cpuPanel.Right + 40;
        int valueX = 132;
        int rowGap = 12;
        int rowTop = 0;
        int labelOffset = 4;

        Panel settingsPanel = new()
        {
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            BackColor = Color.Transparent,
        };

        Label lblMsi = new()
        {
            Text = "MSI Mode:",
            AutoSize = true,
            Location = new Point(0, rowTop + labelOffset),
        };

        ComboBox cmbMsi = new()
        {
            Location = new Point(valueX, rowTop),
            Size = new Size(150, 26),
            DropDownStyle = ComboBoxStyle.DropDownList,
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(18, 18, 22),
            ForeColor = _fgMain,
        };
        cmbMsi.Items.AddRange(new object[] { "Disabled", "Enabled" });

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
            Size = new Size(100, 24),
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
            Location = new Point(txtLimit.Right + 8, txtLimit.Top + 4),
        };

        settingsPanel.Controls.AddRange([lblLimit, txtLimit, lblLimitHint]);
        rowTop = txtLimit.Bottom + rowGap;

        Label lblPrio = new()
        {
            Text = "IRQ Priority:",
            AutoSize = true,
            Location = new Point(0, rowTop + labelOffset),
        };

        ComboBox cmbPrio = new()
        {
            Location = new Point(valueX, rowTop),
            Size = new Size(150, 26),
            DropDownStyle = ComboBoxStyle.DropDownList,
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(18, 18, 22),
            ForeColor = _fgMain,
        };
        cmbPrio.Items.AddRange(new object[] { "Low", "Normal", "High" });

        settingsPanel.Controls.AddRange([lblPrio, cmbPrio]);
        rowTop = cmbPrio.Bottom + rowGap;

        Label lblPolicy = new()
        {
            Text = "Policy:",
            AutoSize = true,
            Location = new Point(0, rowTop + labelOffset),
        };

        ComboBox cmbPolicy = new()
        {
            Location = new Point(valueX, rowTop),
            Size = new Size(170, 26),
            DropDownStyle = ComboBoxStyle.DropDownList,
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(18, 18, 22),
            ForeColor = _fgMain,
        };
        if (device.Kind == DeviceKind.STOR)
        {
            lblPolicy.ForeColor = _mutedText;
            cmbPolicy.Enabled = false;
        }

        settingsPanel.Controls.AddRange([lblPolicy, cmbPolicy]);
        rowTop = cmbPolicy.Bottom + rowGap;

        int imodCheckSize = 14;
        int imodCheckGap = 4;

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
            Text = "IMOD Interval:",
            AutoSize = true,
            ForeColor = _fgMain,
        };

        TextBox txtImod = new()
        {
            Size = new Size(100, 24),
            BackColor = Color.FromArgb(18, 18, 22),
            BorderStyle = BorderStyle.FixedSingle,
            ForeColor = _fgMain,
            TextAlign = HorizontalAlignment.Center,
            Text = "0x0",
        };

        Label lblImodHint = new()
        {
            Text = "(hex/dec)",
            AutoSize = true,
            ForeColor = _mutedText,
        };

        Label lblImodDefault = new()
        {
            Text = "default: 0x0",
            AutoSize = true,
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

        Label lblImodExperimental = new()
        {
            Text = "(experimental)",
            AutoSize = true,
            ForeColor = _mutedText,
        };

        Button btnImodDelete = new()
        {
            Text = "DELETE IMOD",
            Size = new Size(120, 24),
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
            ResetImodIntervalsToDefault();
            ShowThemedInfo("IMOD reset to defaults.\nStartup script and winio.sys removed.");
        };

        bool showImod = ShouldShowImod(device);
        if (showImod)
        {
            Size imodLabelSize = TextRenderer.MeasureText(lblImod.Text, _baseFont);
            int checkY = rowTop + labelOffset + Math.Max(0, (imodLabelSize.Height - imodCheckSize) / 2);
            chkImod.Location = new Point(0, checkY);
            lblImod.Location = new Point(imodCheckSize + imodCheckGap, rowTop + labelOffset);
            txtImod.Location = new Point(valueX, rowTop);
            lblImodHint.Location = new Point(txtImod.Right + 8, txtImod.Top + 4);
            btnImodDelete.Location = new Point(lblImodHint.Right + 12, txtImod.Top - 1);
            lblImodExperimental.Location = new Point(lblImod.Left, txtImod.Bottom + 4);
            lblImodDefault.Location = new Point(txtImod.Left, txtImod.Bottom + 4);
            settingsPanel.Controls.AddRange([chkImod, lblImod, txtImod, lblImodHint, lblImodExperimental, lblImodDefault, btnImodDelete]);
            rowTop = Math.Max(lblImodDefault.Bottom, lblImodExperimental.Bottom) + rowGap;
        }
        else
        {
            chkImod.Visible = false;
            lblImod.Visible = false;
            txtImod.Visible = false;
            lblImodHint.Visible = false;
            lblImodExperimental.Visible = false;
            lblImodDefault.Visible = false;
            btnImodDelete.Visible = false;
        }

        settingsPanel.PerformLayout();
        Size settingsSize = settingsPanel.PreferredSize;
        if (settingsSize.Height == 0)
        {
            settingsSize = settingsPanel.Size;
        }

        int settingsTop = cpuPanel.Top + Math.Max(0, (cpuPanel.Height - settingsSize.Height) / 2);
        settingsPanel.Location = new Point(settingsX, settingsTop);

        int infoY = Math.Max(lblIrq.Bottom + 14, cpuPanel.Bottom + 18);
        infoY = Math.Max(infoY, settingsPanel.Bottom + 18);
        Label lblInfo = new()
        {
            Text = "PNP ID: -",
            AutoEllipsis = true,
            Location = new Point(18, infoY),
            Size = new Size(grp.Width - 40, 70),
            Cursor = Cursors.Hand,
            ForeColor = _fgMain,
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
        };
        lblInfo.MouseEnter += (_, _) => lblInfo.ForeColor = _accent;
        lblInfo.MouseLeave += (_, _) => lblInfo.ForeColor = _fgMain;
        lblInfo.Click += (_, _) =>
        {
            if (lblInfo.Tag is string txt && !string.IsNullOrWhiteSpace(txt))
            {
                Clipboard.SetText(txt);
                ShowCopiedToolTip(lblInfo);
            }
        };

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

        grp.Height = Math.Max(cpuPanel.Bottom + 110, lblInfo.Bottom + 20);

        DeviceBlock block = new()
        {
            Device = device,
            Kind = device.Kind,
            Group = grp,
            CpuBoxes = cpuBoxes,
            AffinityLabel = lblMask,
            IrqLabel = lblIrq,
            MsiCombo = cmbMsi,
            LimitBox = txtLimit,
            PrioCombo = cmbPrio,
            PolicyCombo = cmbPolicy,
            PolicyLabel = lblPolicy,
            ImodAutoCheck = chkImod,
            ImodBox = txtImod,
            ImodDefaultLabel = lblImodDefault,
            InfoLabel = lblInfo,
            AffinityMask = 0,
            IrqCount = null,
        };

        foreach (CheckBox cb in cpuBoxes)
        {
            cb.Tag = block;
            cb.CheckedChanged += (_, _) =>
            {
                if (cb.Tag is DeviceBlock b && b.SuppressCpuEvents == 0)
                {
                    RecalcAffinityMask(b);
                }
            };
        }

        LoadBlockSettings(block);
        _devicesPanel.Controls.Add(grp);
        _blocks.Add(block);
    }

    private void LayoutBlocks()
    {
        int paddingX = 24;
        int gapY = 18;
        int y = 12;
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
            int width = GetDevicesViewportWidth() - (paddingX * 2);
            if (width < 360)
            {
                width = 360;
            }

            b.Group.Width = width;
            int currentHeight = b.InfoLabel.Height > 0 ? b.InfoLabel.Height : 60;
            int infoWidth = Math.Max(140, b.Group.Width - b.InfoLabel.Left - 24);
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
            int width = GetDevicesViewportWidth() - (paddingX * 2);
            if (width < 360)
            {
                width = 360;
            }

            UpdateReservedCpuSetsPanelLayout(reserved, width);
            PlaceControl(reserved);
        }

        int bottomPadding = 32;
        int contentHeight = y + bottomPadding;
        if (_devicesPanel.Height != contentHeight)
        {
            _devicesPanel.Height = contentHeight;
        }

        SyncDevicesScrollBar();
    }

    private void RefreshBlocks()
    {
        InvalidateImodCache();
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
                NewDeviceBlock(d, index);
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

    private void CalculateIrqCounts()
    {
        Dictionary<string, int> irqCounts = GetDeviceIrqCounts();
        foreach (DeviceBlock b in _blocks)
        {
            string shortPnp = GetShortPnpId(b.Device.InstanceId);
            if (!string.IsNullOrWhiteSpace(shortPnp) && irqCounts.TryGetValue(shortPnp, out int val))
            {
                b.IrqCount = val;
                b.IrqLabel.Text = $"IRQ Count: {val}";
                WriteLog($"IRQ.MAP: {b.Device.InstanceId} ({shortPnp}) -> {val}");
            }
            else
            {
                b.IrqCount = 0;
                b.IrqLabel.Text = "IRQ Count: 0";
                WriteLog($"IRQ.MAP: {b.Device.InstanceId} ({shortPnp}) -> 0");
            }
        }

        LogGuiSnapshot("irq-counts");
    }
}
