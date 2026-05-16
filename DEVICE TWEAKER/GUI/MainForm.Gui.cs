using System.Diagnostics;

namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private void InitializeGui()
    {
        UpdateUiScale();
        Text = "DEVICE TWEAKER";
        Size formSize = UiScale(1120, 875);
        Size = formSize;
        StartPosition = FormStartPosition.CenterScreen;
        AutoScaleMode = AutoScaleMode.None;
        BackColor = _bgForm;
        ForeColor = _fgMain;
        Font = _baseFont;
        KeyPreview = true;

        FormBorderStyle = FormBorderStyle.Sizable;
        MaximizeBox = true;
        MinimumSize = UiScale(1120, 840);
        MaximumSize = Size.Empty;
        SizeGripStyle = SizeGripStyle.Show;

        Panel brandPanel = new()
        {
            Dock = DockStyle.Top,
            Height = UiScale(112),
            BackColor = _bgPanel,
            Padding = new Padding(UiScale(28), UiScale(16), UiScale(28), UiScale(8)),
        };

        Label logoLabel = new()
        {
            Text = "DEVICE TWEAKER",
            AutoSize = true,
            Font = _brandFont,
            ForeColor = _accent,
            Margin = new Padding(0, UiScale(2), 0, 0),
        };

        const string developerHandle = "@arsenza";
        const string developerUrl = "https://t.me/arsenzaa";
        string subtitleText = $"alpha version 0.0.4 - developed by {developerHandle}";

        LinkLabel logoSubtitle = new()
        {
            Text = subtitleText,
            AutoSize = true,
            Font = new Font("Consolas", 10, FontStyle.Regular),
            LinkBehavior = LinkBehavior.HoverUnderline,
            LinkColor = _mutedText,
            ActiveLinkColor = _accent,
            VisitedLinkColor = _mutedText,
            DisabledLinkColor = _mutedText,
            ForeColor = _mutedText,
            MaximumSize = new Size(UiScale(940), 0),
            TextAlign = ContentAlignment.MiddleCenter,
            Margin = new Padding(0, UiScale(4), 0, 0),
        };
        int linkStart = subtitleText.IndexOf(developerHandle, StringComparison.Ordinal);
        if (linkStart >= 0)
        {
            logoSubtitle.LinkArea = new LinkArea(linkStart, developerHandle.Length);
            logoSubtitle.LinkClicked += (_, _) => OpenUrl(developerUrl);
            logoSubtitle.Cursor = Cursors.Hand;
        }

        TableLayoutPanel brandLayout = new()
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 4,
            BackColor = _bgPanel,
            Margin = Padding.Empty,
            Padding = Padding.Empty,
        };
        brandLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        brandLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
        brandLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        brandLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        brandLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));

        logoLabel.Anchor = AnchorStyles.None;
        logoSubtitle.Anchor = AnchorStyles.None;

        brandLayout.Controls.Add(logoLabel, 0, 1);
        brandLayout.Controls.Add(logoSubtitle, 0, 2);
        brandLayout.Layout += (_, _) =>
        {
            int w = Math.Max(0, brandLayout.ClientSize.Width);
            Size newMax = new(w, 0);
            if (logoSubtitle.MaximumSize != newMax)
            {
                logoSubtitle.MaximumSize = newMax;
            }
        };

        brandPanel.Controls.Add(brandLayout);

        Panel statusPanel = new()
        {
            Dock = DockStyle.Top,
            Height = UiScale(74),
            BackColor = _bgPanel,
            Padding = new Padding(UiScale(28), UiScale(2), UiScale(28), UiScale(2)),
        };

        Label NewCpuFlagPrefix(string text) => new()
        {
            Text = text,
            AutoSize = true,
            Font = _htFont,
            ForeColor = _fgMain,
            Margin = new Padding(0),
        };

        Label NewCpuFlagStatus() => new()
        {
            AutoSize = true,
            Font = _htFont,
            ForeColor = _statusInactive,
            Margin = new Padding(UiScale(4), 0, 0, 0),
        };

        void AddCpuFlag(FlowLayoutPanel panel, Label prefix, Label status)
        {
            panel.Controls.Add(prefix);
            panel.Controls.Add(status);
        }

        void AddCpuFlagSeparator(FlowLayoutPanel panel)
        {
            panel.Controls.Add(new Label
            {
                Text = "|",
                AutoSize = true,
                Font = _htFont,
                ForeColor = _statusSeparator,
                Margin = new Padding(UiScale(14), 0, UiScale(14), 0),
            });
        }

        _htPrefixLabel = NewCpuFlagPrefix("Hyper-Threading");
        _htStatusLabel = NewCpuFlagStatus();
        _hybridCpuPrefixLabel = NewCpuFlagPrefix("Hybrid CPU");
        _hybridCpuStatusLabel = NewCpuFlagStatus();
        _cppcPrefixLabel = NewCpuFlagPrefix("CPPC");
        _cppcStatusLabel = NewCpuFlagStatus();
        _dualCcdPrefixLabel = NewCpuFlagPrefix("Dual-CCD");
        _dualCcdStatusLabel = NewCpuFlagStatus();

        _cpuHeaderLabel = new Label
        {
            Text = _cpuHeaderText,
            AutoSize = true,
            ForeColor = _mutedText,
            Font = _headerFont,
            TextAlign = ContentAlignment.MiddleCenter,
            Margin = new Padding(0, 0, 0, UiScale(1)),
        };

        FlowLayoutPanel cpuFlagsPanel = new()
        {
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            WrapContents = false,
            Padding = new Padding(UiScale(10), UiScale(3), UiScale(10), UiScale(3)),
            Margin = new Padding(0),
            BackColor = _bgPanel,
        };
        _cpuFlagsPanel = cpuFlagsPanel;
        AddCpuFlag(cpuFlagsPanel, _htPrefixLabel, _htStatusLabel);
        AddCpuFlagSeparator(cpuFlagsPanel);
        AddCpuFlag(cpuFlagsPanel, _hybridCpuPrefixLabel, _hybridCpuStatusLabel);
        AddCpuFlagSeparator(cpuFlagsPanel);
        AddCpuFlag(cpuFlagsPanel, _cppcPrefixLabel, _cppcStatusLabel);
        AddCpuFlagSeparator(cpuFlagsPanel);
        AddCpuFlag(cpuFlagsPanel, _dualCcdPrefixLabel, _dualCcdStatusLabel);

        TableLayoutPanel statusLayout = new()
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 4,
            BackColor = _bgPanel,
            Margin = Padding.Empty,
            Padding = Padding.Empty,
        };
        statusLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        statusLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
        statusLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        statusLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        statusLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
        statusLayout.Layout += (_, _) =>
        {
            int w = Math.Max(0, statusLayout.ClientSize.Width);
            Size newMax = new(w, 0);
            if (_cpuHeaderLabel.MaximumSize != newMax)
            {
                _cpuHeaderLabel.MaximumSize = newMax;
            }
        };
        _cpuHeaderLabel.Anchor = AnchorStyles.None;
        cpuFlagsPanel.Anchor = AnchorStyles.None;
        _cpuHeaderLabel.Margin = new Padding(0, 0, 0, UiScale(6));
        cpuFlagsPanel.Margin = new Padding(0);
        statusLayout.Controls.Add(_cpuHeaderLabel, 0, 1);
        statusLayout.Controls.Add(cpuFlagsPanel, 0, 2);
        statusPanel.Controls.Add(statusLayout);

        UpdateCpuHeaderUi();

        Panel buttonPanel = new()
        {
            Dock = DockStyle.Top,
            Height = UiScale(108),
            BackColor = _bgPanel,
            Padding = new Padding(UiScale(24), UiScale(4), UiScale(24), UiScale(16)),
            Margin = Padding.Empty,
        };

        Button btnScan = NewTopButton("REFRESH");
        Button btnApply = NewTopButton("APPLY");
        Button btnAuto = NewTopButton("AUTO-OPTIMIZATION");
        Button btnIrq = NewTopButton("CALCULATE IRQ COUNTS");
        Button btnReset = NewTopButton("RESET ALL");
        Button btnRestore = NewTopButton("RESTORE");

        foreach (Button b in new[] { btnScan, btnApply, btnAuto, btnIrq, btnReset, btnRestore })
        {
            SetTopButtonBaseStyle(b);
            b.MouseEnter += (_, _) => SetTopButtonHoverStyle(b);
            b.MouseLeave += (_, _) => SetTopButtonBaseStyle(b);
        }

        _btnLog = NewTopButton("ENABLE LOGGING");
        UpdateLoggingButtonUi();
        _btnLog.MouseEnter += (_, _) => SetTopButtonHoverStyle(_btnLog);
        _btnLog.MouseLeave += (_, _) => UpdateLoggingButtonUi();
        _btnLog.Click += (_, _) =>
        {
            if (_detailedLogEnabled)
            {
                DisableDetailedLog();
            }
            else
            {
                EnableDetailedLog();
                WriteLog("UI: LOG button turned ON -> triggering REFRESH");
                RefreshBlocks();
            }

            UpdateLoggingButtonUi();
        };

        TableLayoutPanel buttonsHost = new()
        {
            Dock = DockStyle.Fill,
            ColumnCount = 3,
            RowCount = 1,
            BackColor = _bgPanel,
            Margin = Padding.Empty,
            Padding = Padding.Empty,
        };
        buttonsHost.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
        buttonsHost.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        buttonsHost.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
        buttonsHost.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

        TableLayoutPanel buttonsGrid = new()
        {
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            ColumnCount = 3,
            RowCount = 2,
            BackColor = _bgPanel,
            Anchor = AnchorStyles.None,
            Margin = Padding.Empty,
            Padding = Padding.Empty,
        };
        buttonsGrid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        buttonsGrid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        buttonsGrid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        buttonsGrid.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        buttonsGrid.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        buttonsGrid.Controls.Add(btnApply, 0, 0);
        buttonsGrid.Controls.Add(btnAuto, 1, 0);
        buttonsGrid.Controls.Add(btnScan, 2, 0);
        buttonsGrid.Controls.Add(btnReset, 0, 1);
        buttonsGrid.Controls.Add(btnIrq, 1, 1);
        buttonsGrid.Controls.Add(btnRestore, 2, 1);

        buttonsHost.Controls.Add(buttonsGrid, 1, 0);
        buttonPanel.Controls.Add(buttonsHost);

        Panel accentStrip = new()
        {
            Dock = DockStyle.Top,
            Height = UiScale(1),
            BackColor = _accent,
        };

        _devicesHost = new BufferedPanel
        {
            Dock = DockStyle.Fill,
            BackColor = _bgForm,
            Padding = Padding.Empty,
            Margin = Padding.Empty,
        };

        int scrollWidth = UiScale(12);
        _devicesPanel = new BufferedPanel
        {
            Dock = DockStyle.None,
            BackColor = _bgForm,
            AutoScroll = false,
            Padding = new Padding(UiScale(24), UiScale(12), UiScale(24), UiScale(32)),
        };
        _devicesPanel.Location = new Point(0, 0);
        _devicesPanel.SizeChanged += (_, _) => SyncDevicesScrollBar();

        _devicesScroll = new ThemedScrollBar
        {
            Width = scrollWidth,
            BackColor = _bgForm,
            TrackColor = _bgForm,
            RailColor = _bgForm,
            ThumbColor = _accent,
            ThumbWidth = UiScale(9),
            RailWidth = 0,
            ThumbCornerRadius = UiScale(7),
            Visible = false,
            Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Right,
        };
        _devicesScroll.ValueChanged += (_, _) =>
        {
            if (_syncingScroll)
            {
                return;
            }

            SetDevicesScrollOffset(_devicesScroll.Value);
            SyncDevicesScrollBar();
        };

        _devicesHost.Controls.Add(_devicesPanel);
        _devicesHost.Controls.Add(_devicesScroll);
        _devicesHost.SizeChanged += (_, _) =>
        {
            UpdateDevicesHostLayout();
            UpdateDevicesScrollLayout();
            SyncDevicesScrollBar();
        };
        _devicesHost.MouseEnter += (_, _) => _devicesHost.Focus();
        _devicesHost.MouseWheel += (_, e) => HandleDevicesMouseWheel(e);
        _devicesHost.TabStop = true;

        Controls.Add(_devicesHost);
        Controls.Add(accentStrip);
        Controls.Add(buttonPanel);
        Controls.Add(statusPanel);
        Controls.Add(brandPanel);

        ApplyDarkScrollBarTheme(_devicesPanel);
        SyncDevicesScrollBar();
        UpdateDevicesScrollLayout();
        UpdateDevicesHostLayout();

        _copyToolTip = new ToolTip
        {
            UseFading = true,
            UseAnimation = true,
            IsBalloon = false,
            ShowAlways = true,
        };
        _layoutRefreshTimer = new System.Windows.Forms.Timer
        {
            Interval = 220,
        };
        _layoutRefreshTimer.Tick += (_, _) =>
        {
            _layoutRefreshTimer.Stop();
            RebuildDeviceBlocksForLayout();
        };

        btnScan.Click += (_, _) =>
        {
            WriteLog("UI: REFRESH button clicked");
            RefreshBlocks();
        };
        btnApply.Click += (_, _) =>
        {
            WriteLog("UI: APPLY button clicked");
            CreateDeviceTweakerBackup("pre-apply", showDialog: false);
            foreach (DeviceBlock b in _blocks)
            {
                if (b.Device.Wifi)
                {
                    WriteLog($"APPLY.SKIP: {b.Device.InstanceId} Kind={b.Kind} reason=wifi");
                    continue;
                }

                SaveBlockSettings(b);
            }

            _ = ApplyImodSettings(out string? imodNote);
            if (!string.IsNullOrWhiteSpace(imodNote))
            {
                WriteLog($"IMOD.NOTE: {imodNote}");
            }
            RefreshImodCurrentValues(reason: "apply");
            string message = "All changes have been applied and saved.";
            message += "\nPlease reboot your PC to finish applying them.";

            LogGuiSnapshot("apply");
            ShowThemedInfo(message);
        };
        btnAuto.Click += (_, _) =>
        {
            WriteLog("UI: AUTO-OPTIMIZATION button clicked");
            bool hasUsbImodTarget = _blocks.Any(b => IsUsbImodTarget(b.Device));
            bool optimizeUsbImod = false;
            if (hasUsbImodTarget)
            {
                optimizeUsbImod = ShowThemedConfirm(
                    "USB IMOD tuning is available for detected XHCI controller(s).\n\nApply it during auto-optimization?",
                    "USB IMOD TUNING",
                    "APPLY",
                    "SKIP");
                WriteLog($"AUTO.IMOD.PROMPT: {(optimizeUsbImod ? "accepted" : "declined")}");
            }
            else
            {
                WriteLog("AUTO.IMOD.PROMPT: skipped (no eligible XHCI controllers)");
            }

            if (!_testAutoDryRun)
            {
                BackupLocation backupLocation = PromptBackupLocationForAuto();
                CreateDeviceTweakerBackup("pre-auto", showDialog: false, backupLocation);
            }

            InvokeAutoOptimization(optimizeUsbImod);
            bool applyImod = optimizeUsbImod && hasUsbImodTarget;

            if (_testAutoDryRun)
            {
                WriteLog("AUTO.DRYRUN: enabled -> skipping registry writes");
                if (applyImod)
                {
                    WriteLog("AUTO.DRYRUN: IMOD apply skipped");
                }

                ShowThemedInfo("Auto-optimization preview completed.\nDry-run mode is ON (no registry changes).");
                return;
            }

            foreach (DeviceBlock b in _blocks)
            {
                if (b.Device.Wifi)
                {
                    WriteLog($"AUTO.APPLY.SKIP: {b.Device.InstanceId} Kind={b.Kind} reason=wifi");
                    continue;
                }

                SaveBlockSettings(b, msiOnlyForIntegratedGpu: true);
            }

            WriteLog("UI: AUTO-OPTIMIZATION applied and saved");
            if (applyImod)
            {
                _ = ApplyImodSettings(out string? imodNote);
                if (!string.IsNullOrWhiteSpace(imodNote))
                {
                    WriteLog($"IMOD.NOTE: {imodNote}");
                }
            }
            else
            {
                WriteLog("IMOD skipped (AUTO-OPTIMIZATION): no eligible XHCI controllers");
            }
            string autoMessage = "Auto-optimization completed and saved.";
            autoMessage += "\nPlease reboot your PC to finish applying the changes.";
            ShowThemedInfo(autoMessage);

            WriteLog("UI: AUTO-OPTIMIZATION done -> triggering REFRESH");
            RefreshBlocks();
        };
        btnIrq.Click += (_, _) =>
        {
            WriteLog("UI: CALCULATE IRQ COUNTS button clicked");
            CalculateIrqCounts();
        };
        btnReset.Click += (_, _) =>
        {
            WriteLog("UI: RESET ALL button clicked");
            CreateDeviceTweakerBackup("pre-reset", showDialog: false);
            ResetAllTweaks();
        };
        btnRestore.Click += (_, _) =>
        {
            WriteLog("UI: RESTORE button clicked");
            RestoreLatestDeviceTweakerBackup();
        };

        _lastLayoutViewportWidth = _devicesHost.ClientSize.Width;
        _lastLayoutDpi = GetCurrentWindowDpi();
        Resize += (_, _) =>
        {
            LayoutBlocks();
        };
        ResizeEnd += (_, _) => LayoutBlocks();
        DpiChanged += (_, _) =>
        {
            UpdateUiScale();
            _initialDeviceViewportHeightAdjusted = false;
            QueueDeviceLayoutRebuild(force: true);
        };
        MouseWheel += (_, e) => HandleDevicesMouseWheel(e);
        KeyDown += OnMainFormKeyDown;
    }

    private int GetCurrentWindowDpi()
    {
        try
        {
            if (IsHandleCreated)
            {
                return NativeUser32.GetDpiForWindow(Handle);
            }

            return NativeUser32.GetDpiForSystem();
        }
        catch
        {
            return 96;
        }
    }

    private void QueueDeviceLayoutRebuild(bool force)
    {
        if (_layoutRefreshTimer is null || _devicesHost is null || IsDisposed)
        {
            return;
        }

        int viewportWidth = _devicesHost.ClientSize.Width;
        int dpi = GetCurrentWindowDpi();
        if (!force
            && Math.Abs(viewportWidth - _lastLayoutViewportWidth) < UiScale(24)
            && dpi == _lastLayoutDpi)
        {
            return;
        }

        _lastLayoutViewportWidth = viewportWidth;
        _lastLayoutDpi = dpi;
        _layoutRefreshTimer.Stop();
        _layoutRefreshTimer.Start();
    }

    private void RebuildDeviceBlocksForLayout()
    {
        if (IsDisposed)
        {
            return;
        }

        UpdateUiScale();
        if (_blocks.Count == 0)
        {
            LayoutBlocks();
            return;
        }

        RefreshBlocks(includeImodReadback: false);
    }

    private bool TryExpandMainWindowForViewportWidth(int desiredViewportWidth)
    {
        if (IsDisposed || !IsHandleCreated || WindowState != FormWindowState.Normal || _expandingMainWindowForLayout)
        {
            return false;
        }

        int currentViewportWidth = GetDevicesViewportWidth();
        int delta = desiredViewportWidth - currentViewportWidth;
        if (delta <= UiScale(2))
        {
            return true;
        }

        Rectangle workingArea = Screen.FromControl(this).WorkingArea;
        int maxWidthToRight = Math.Max(Width, workingArea.Right - Left);
        int nextWidth = Math.Min(maxWidthToRight, Width + delta);
        if (nextWidth <= Width + UiScale(2))
        {
            return false;
        }

        _expandingMainWindowForLayout = true;
        try
        {
            Width = nextWidth;
            UpdateDevicesScrollLayout();
            UpdateDevicesHostLayout();
        }
        finally
        {
            _expandingMainWindowForLayout = false;
        }

        return GetDevicesViewportWidth() >= desiredViewportWidth - UiScale(2);
    }

    private void AdjustInitialDeviceViewportHeight()
    {
        if (_initialDeviceViewportHeightAdjusted
            || IsDisposed
            || !IsHandleCreated
            || WindowState != FormWindowState.Normal
            || _devicesHost is null
            || _blocks.Count == 0)
        {
            return;
        }

        Control firstBlock = _blocks[0].Group;
        int desiredViewportHeight = firstBlock.Bottom + UiScale(10);
        if (_blocks.Count > 1)
        {
            desiredViewportHeight = Math.Min(desiredViewportHeight, _blocks[1].Group.Top - UiScale(2));
        }

        desiredViewportHeight = Math.Max(firstBlock.Bottom + UiScale(2), desiredViewportHeight);

        int delta = desiredViewportHeight - _devicesHost.ClientSize.Height;
        if (Math.Abs(delta) <= UiScale(3))
        {
            _initialDeviceViewportHeightAdjusted = true;
            return;
        }

        Rectangle workingArea = Screen.FromControl(this).WorkingArea;
        int maxHeight = Math.Max(Height, workingArea.Bottom - Top);
        int nextHeight = Math.Max(MinimumSize.Height, Math.Min(Height + delta, maxHeight));
        if (Math.Abs(nextHeight - Height) <= UiScale(2))
        {
            _initialDeviceViewportHeightAdjusted = true;
            return;
        }

        _initialDeviceViewportHeightAdjusted = true;
        Height = nextHeight;
        UpdateDevicesScrollLayout();
        UpdateDevicesHostLayout();
        LayoutBlocks();
    }

    private void ApplyDarkScrollBarTheme(Control control)
    {
        if (!OperatingSystem.IsWindowsVersionAtLeast(10, 0, 17763))
        {
            return;
        }

        void ApplyTheme()
        {
            try
            {
                if (!_darkModeInitialized)
                {
                    _ = NativeUxTheme.SetPreferredAppMode(NativeUxTheme.PreferredAppMode.ForceDark);
                    NativeUxTheme.RefreshImmersiveColorPolicyState();
                    _darkModeInitialized = true;
                }

                _ = NativeUxTheme.AllowDarkModeForWindow(control.Handle, true);
                _ = NativeUxTheme.SetWindowTheme(control.Handle, "DarkMode_Explorer", null);
                HideNativeScrollBars(control);
            }
            catch
            {
            }
        }

        if (control.IsHandleCreated)
        {
            ApplyTheme();
        }
        else
        {
            control.HandleCreated += (_, _) => ApplyTheme();
        }
    }

    private void HideNativeScrollBars(Control control)
    {
        if (!control.IsHandleCreated)
        {
            control.HandleCreated += (_, _) => HideNativeScrollBars(control);
            return;
        }

        _ = NativeUser32.ShowScrollBar(control.Handle, NativeUser32.SbVert, false);
        _ = NativeUser32.ShowScrollBar(control.Handle, NativeUser32.SbHorz, false);
    }

    private void SyncDevicesScrollBar()
    {
        if (_devicesScroll is null || _devicesPanel is null)
        {
            return;
        }

        UpdateDevicesScrollLayout();
        UpdateDevicesHostLayout();

        if (_devicesPanel.IsHandleCreated)
        {
            HideNativeScrollBars(_devicesPanel);
        }

        int contentHeight = _devicesPanel.Height;
        int viewportHeight = _devicesHost.ClientSize.Height;
        int offset = Math.Max(0, -_devicesPanel.Top);
        bool needsScroll = contentHeight > viewportHeight + 1;

        _devicesScroll.Visible = needsScroll;

        _syncingScroll = true;
        _devicesScroll.Maximum = Math.Max(contentHeight, 1);
        _devicesScroll.ViewportSize = Math.Max(viewportHeight, 1);
        _devicesScroll.Value = needsScroll ? offset : 0;
        _syncingScroll = false;

        if (!needsScroll)
        {
            _devicesPanel.Location = new Point(0, 0);
        }
    }

    private void UpdateDevicesScrollLayout()
    {
        if (_devicesScroll is null)
        {
            return;
        }

        Control? host = _devicesScroll.Parent;
        if (host is null)
        {
            return;
        }

        int width = _devicesScroll.Width;
        _devicesScroll.Location = new Point(Math.Max(0, host.ClientSize.Width - width), 0);
        _devicesScroll.Height = host.ClientSize.Height;
        _devicesScroll.BringToFront();
    }

    private void UpdateDevicesHostLayout()
    {
        if (_devicesHost is null || _devicesPanel is null)
        {
            return;
        }

        int contentWidth = GetDevicesViewportWidth();
        if (_devicesPanel.Width != contentWidth)
        {
            _devicesPanel.Width = contentWidth;
        }

        if (_devicesPanel.Left != 0)
        {
            _devicesPanel.Left = 0;
        }

        int maxOffset = Math.Max(0, _devicesPanel.Height - _devicesHost.ClientSize.Height);
        int offset = Math.Max(0, -_devicesPanel.Top);
        if (offset > maxOffset)
        {
            _devicesPanel.Top = -maxOffset;
        }
    }

    private void SetDevicesScrollOffset(int offset)
    {
        if (_devicesHost is null || _devicesPanel is null)
        {
            return;
        }

        int maxOffset = Math.Max(0, _devicesPanel.Height - _devicesHost.ClientSize.Height);
        int next = Math.Max(0, Math.Min(maxOffset, offset));
        _devicesPanel.Location = new Point(0, -next);
    }

    private void HandleDevicesMouseWheel(MouseEventArgs e)
    {
        if (_devicesScroll is null || !_devicesScroll.Visible)
        {
            return;
        }

        if (!IsCursorOverDevicesHost())
        {
            return;
        }

        int delta = e.Delta > 0 ? -_devicesScroll.SmallChange : _devicesScroll.SmallChange;
        _devicesScroll.Value += delta;
    }

    private void ForwardDevicesMouseWheel(object? sender, MouseEventArgs e)
    {
        int before = _devicesScroll is null ? 0 : _devicesScroll.Value;
        HandleDevicesMouseWheel(e);

        if (_devicesScroll is not null
            && _devicesScroll.Value != before
            && e is HandledMouseEventArgs handled)
        {
            handled.Handled = true;
        }
    }

    private void WireDevicesMouseWheelForwarding(Control root)
    {
        root.MouseWheel += ForwardDevicesMouseWheel;
        foreach (Control child in root.Controls)
        {
            WireDevicesMouseWheelForwarding(child);
        }
    }

    private bool IsCursorOverDevicesHost()
    {
        if (_devicesHost is null)
        {
            return false;
        }

        Point p = _devicesHost.PointToClient(Cursor.Position);
        return p.X >= 0 && p.Y >= 0 && p.X < _devicesHost.ClientSize.Width && p.Y < _devicesHost.ClientSize.Height;
    }

    private Button NewTopButton(string text)
    {
        return new Button
        {
            Text = text,
            Size = UiScale(186, 36),
            Margin = new Padding(UiScale(8), UiScale(4), UiScale(8), UiScale(4)),
            FlatStyle = FlatStyle.Flat,
            Font = _buttonFont,
            UseVisualStyleBackColor = false,
            Cursor = Cursors.Hand,
        };
    }

    private void SetTopButtonBaseStyle(Button btn)
    {
        btn.FlatAppearance.BorderSize = 1;
        btn.BackColor = _bgForm;
        btn.ForeColor = _fgMain;
        btn.FlatAppearance.BorderColor = _accent;
    }

    private void SetTopButtonHoverStyle(Button btn)
    {
        btn.BackColor = _accent;
        btn.ForeColor = Color.FromArgb(15, 15, 15);
    }

    private void UpdateLoggingButtonUi()
    {
        if (_btnLog is null)
        {
            return;
        }

        _btnLog.Text = _detailedLogEnabled ? "DISABLE LOGGING" : "ENABLE LOGGING";
        _btnLog.FlatAppearance.BorderSize = 1;
        _btnLog.BackColor = _detailedLogEnabled ? _bgPanel : _bgForm;
        _btnLog.ForeColor = _detailedLogEnabled ? _accent : _fgMain;
        _btnLog.FlatAppearance.BorderColor = _detailedLogEnabled ? _accentDark : _accent;
    }

    private void ShowCopiedToolTip(Control target)
    {
        try
        {
            _copyToolTip.Hide(target);
            Point screenPos = Cursor.Position;
            Point clientPos = target.PointToClient(screenPos);
            Point point = new(clientPos.X, clientPos.Y + 20);
            _copyToolTip.Show("Copied", target, point, 1200);
        }
        catch
        {
        }
    }

    private static void OpenUrl(string url)
    {
        try
        {
            ProcessStartInfo startInfo = new()
            {
                FileName = url,
                UseShellExecute = true,
            };
            Process.Start(startInfo);
        }
        catch
        {
        }
    }
}
