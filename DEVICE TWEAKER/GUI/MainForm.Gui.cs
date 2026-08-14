using System.Diagnostics;
using System.Reflection;


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
        string informationalVersion = Assembly.GetEntryAssembly()
            ?.GetCustomAttribute<AssemblyInformationalVersionAttribute>()
            ?.InformationalVersion
            ?? "0.0.4-alpha.2";
        string subtitleText = $"alpha version {informationalVersion} - developed by {developerHandle}";

        LinkLabel logoSubtitle = new()
        {
            Text = subtitleText,
            AutoSize = true,
            Font = _subtitleFont,
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
            Height = UiScale(86),
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
        _sandboxPrefixLabel = NewCpuFlagPrefix("Sandbox");
        _sandboxStatusLabel = NewCpuFlagStatus();

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
        AddCpuFlagSeparator(cpuFlagsPanel);
        AddCpuFlag(cpuFlagsPanel, _sandboxPrefixLabel, _sandboxStatusLabel);

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
        Button btnReset = NewTopButton("RESET ALL");
        Button btnRestore = NewTopButton("RESTORE");
        int buttonGap = UiScale(8);
        int buttonRowGap = UiScale(8);
        btnApply.Margin = Padding.Empty;
        btnAuto.Margin = new Padding(buttonGap, 0, 0, 0);
        btnScan.Margin = new Padding(buttonGap, 0, 0, 0);
        btnReset.Margin = Padding.Empty;
        btnRestore.Margin = new Padding(buttonGap, 0, 0, 0);

        foreach (Button b in new[] { btnScan, btnApply, btnAuto, btnReset, btnRestore })
        {
            SetTopButtonBaseStyle(b);
            b.MouseEnter += (_, _) => SetTopButtonHoverStyle(b);
            b.MouseLeave += (_, _) => SetTopButtonBaseStyle(b);
        }

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

        FlowLayoutPanel bottomButtons = new()
        {
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            BackColor = _bgPanel,
            FlowDirection = FlowDirection.LeftToRight,
            Margin = new Padding(0, buttonRowGap, 0, 0),
            Padding = Padding.Empty,
            WrapContents = false,
            Anchor = AnchorStyles.None,
        };
        bottomButtons.Controls.Add(btnReset);
        bottomButtons.Controls.Add(btnRestore);
        buttonsGrid.Controls.Add(bottomButtons, 0, 1);
        buttonsGrid.SetColumnSpan(bottomButtons, 3);

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
        EnsureDevicesBusyOverlay();
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

        _copyToolTip = new ThemedToolTip(showAlways: true, font: _technicalFont)
        {
            AutoPopDelay = 20000,
            InitialDelay = 400,
            ReshowDelay = 200,
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
            OperationReport report = new();
            bool sandboxDryRun = IsSandboxDryRunActive();
            if (!sandboxDryRun)
            {
                if (!CreateDeviceTweakerBackup("pre-apply", showDialog: false))
                {
                    report.AddError("Automatic backup", "backup could not be created; changes were not applied");
                    ShowOperationResult(
                        report,
                        successMessage: string.Empty,
                        partialMessage: "APPLY was cancelled because the automatic backup failed.");
                    return;
                }
            }
            else
            {
                WriteLog("APPLY: dry-run -> skipped pre-apply backup");
            }

            int applyTotal = Math.Max(1, _blocks.Count) + 3;
            BeginDevicesBusyWork(sandboxDryRun ? "Previewing APPLY..." : "Applying changes...", applyTotal);
            try
            {
                int saved = 0;
                int saveTotal = Math.Max(1, _blocks.Count);
                foreach (DeviceBlock b in _blocks)
                {
                    saved++;
                    if (b.Device.Wifi)
                    {
                        WriteLog($"APPLY.SKIP: {b.Device.InstanceId} Kind={b.Kind} reason=wifi");
                        TickDevicesBusy($"Applying device settings ({saved}/{saveTotal})", 1);
                        continue;
                    }

                    if (sandboxDryRun && !b.Device.IsTestDevice)
                    {
                        WriteLog($"APPLY.DRYRUN: {b.Device.InstanceId} Kind={b.Kind} skipped (real device)");
                        TickDevicesBusy($"Previewing device settings ({saved}/{saveTotal})", 1);
                        continue;
                    }

                    TickDevicesBusy(
                        sandboxDryRun
                            ? $"Previewing device settings ({saved}/{saveTotal})"
                            : $"Applying device settings ({saved}/{saveTotal})",
                        1);
                    SaveBlockSettings(b, report: report);
                }

                TickDevicesBusy(sandboxDryRun ? "Skipping USB selective suspend..." : "Applying USB selective suspend...", 1);
                if (sandboxDryRun)
                {
                    WriteLog("APPLY.DRYRUN: USB selective suspend skipped");
                }
                else
                {
                    ApplyUsbSelectiveSuspendPowerPlan(forceDisable: false, report);
                }

                TickDevicesBusy(sandboxDryRun ? "Skipping USB IMOD..." : "Applying USB IMOD...", 1);
                if (sandboxDryRun)
                {
                    WriteLog("APPLY.DRYRUN: IMOD apply skipped");
                }
                else
                {
                    ImodApplyOutcome imodOutcome = ApplyImodSettings(out string? imodNote);
                    if (!string.IsNullOrWhiteSpace(imodNote))
                    {
                        WriteLog($"IMOD.NOTE: {imodNote}");
                    }

                    if (imodOutcome == ImodApplyOutcome.Failed)
                    {
                        report.AddError("IMOD", imodNote ?? "apply failed");
                    }
                    else if (!string.IsNullOrWhiteSpace(imodNote)
                        && (imodNote.Contains("failure", StringComparison.OrdinalIgnoreCase)
                            || imodNote.Contains("failed", StringComparison.OrdinalIgnoreCase)))
                    {
                        report.AddError("IMOD", imodNote);
                    }
                }

                TickDevicesBusy("Updating IMOD / IRQ display...", 1);
                WaitForBackgroundUiTasks(
                    RefreshImodCurrentValuesAsync(showReadingStatus: true, reason: sandboxDryRun ? "apply-dry-run" : "apply"),
                    CalculateIrqCountsAsync(sandboxDryRun ? "apply-dry-run" : "apply"));
                LogGuiSnapshot(sandboxDryRun ? "apply-dry-run" : "apply");
                _devicesBusyDone = _devicesBusyTotal;
                UpdateDevicesBusy("Ready", 100);
                WriteLog(
                    sandboxDryRun
                        ? $"UI: APPLY dry-run completed blocks={_blocks.Count} errors={report.Errors.Count}"
                        : $"UI: APPLY completed blocks={_blocks.Count} errors={report.Errors.Count}");
            }
            finally
            {
                EndDevicesBusy();
            }

            if (sandboxDryRun)
            {
                ShowThemedInfo("APPLY preview completed.\nSandbox dry-run is ON (no registry changes).");
            }
            else
            {
                ShowOperationResult(
                    report,
                    successMessage: "All changes have been applied and saved.\nPlease reboot your PC to finish applying them.",
                    partialMessage: "APPLY finished with errors. Some changes may be incomplete.");
            }
        };
        btnAuto.Click += (_, _) =>
        {
            WriteLog("UI: AUTO-OPTIMIZATION button clicked");
            bool hasUsbImodTarget = _blocks.Any(b => IsUsbImodTarget(b.Device));
            bool optimizeUsbImod = false;
            if (hasUsbImodTarget)
            {
                optimizeUsbImod = ShowThemedConfirm(
                    "USB IMOD tuning is available for detected XHCI controller(s).\n\nDTIMOD driver access can be blocked by Windows Defender or anti-cheats.\n\nApply it during AUTO-OPTIMIZATION?",
                    "USB IMOD TUNING",
                    "APPLY",
                    "SKIP");
                WriteLog($"AUTO.IMOD.PROMPT: {(optimizeUsbImod ? "accepted" : "declined")}");
            }
            else
            {
                WriteLog("AUTO.IMOD.PROMPT: skipped (no eligible XHCI controllers)");
            }

            OperationReport report = new();
            if (!_testAutoDryRun)
            {
                AutoBackupChoice backupChoice = PromptBackupLocationForAuto();
                if (backupChoice == AutoBackupChoice.Cancel)
                {
                    WriteLog("BACKUP.PROMPT.AUTO: cancelled by user");
                    return;
                }

                if (backupChoice == AutoBackupChoice.Local || backupChoice == AutoBackupChoice.Roaming)
                {
                    BackupLocation backupLocation = backupChoice == AutoBackupChoice.Local ? BackupLocation.Local : BackupLocation.Roaming;
                    if (!CreateDeviceTweakerBackup("pre-auto", showDialog: false, backupLocation))
                    {
                        report.AddError("Automatic backup", "backup could not be created; changes were not applied");
                        ShowOperationResult(
                            report,
                            successMessage: string.Empty,
                            partialMessage: "AUTO-OPTIMIZATION was cancelled because the automatic backup failed.");
                        return;
                    }
                }
                else
                {
                    WriteLog("BACKUP.PROMPT.AUTO: skipped by user");
                }
            }

            int saveTotal = Math.Max(1, _blocks.Count);
            // Work units: plan + each block apply + USB SS + optional IMOD + refresh.
            int autoUnits = 1 + saveTotal + 1 + 1;
            string? dryRunInfo = null;
            bool showAutoResult = true;
            BeginDevicesBusyWork("Running AUTO-OPTIMIZATION...", autoUnits);
            try
            {
                TickDevicesBusy("Planning AUTO-OPTIMIZATION...", 1);
                bool planBuilt = InvokeAutoOptimization(optimizeUsbImod, report);
                bool applyImod = optimizeUsbImod && hasUsbImodTarget;
                if (applyImod)
                {
                    SetDevicesBusyWork(_devicesBusyTotal + 1, _devicesBusyDone);
                }

                if (!planBuilt)
                {
                    WriteLog("AUTO: plan was not built -> skipping apply/save");
                    _devicesBusyDone = _devicesBusyTotal;
                    UpdateDevicesBusy("Ready", 100);
                    showAutoResult = true;
                    return;
                }

                if (_testAutoDryRun)
                {
                    WriteLog("AUTO.DRYRUN: enabled -> skipping registry writes");
                    if (applyImod)
                    {
                        WriteLog("AUTO.DRYRUN: IMOD apply skipped");
                    }

                    _devicesBusyDone = _devicesBusyTotal;
                    UpdateDevicesBusy("Ready", 100);
                    dryRunInfo = "AUTO-OPTIMIZATION preview completed.\nSandbox dry-run is ON (no registry changes).";
                    return;
                }

                int saved = 0;
                foreach (DeviceBlock b in _blocks)
                {
                    saved++;
                    if (b.Device.Wifi)
                    {
                        WriteLog($"AUTO.APPLY.SKIP: {b.Device.InstanceId} Kind={b.Kind} reason=wifi");
                        TickDevicesBusy($"Applying device settings ({saved}/{saveTotal})", 1);
                        continue;
                    }

                    TickDevicesBusy($"Applying device settings ({saved}/{saveTotal})", 1);
                    SaveBlockSettings(b, msiOnlyForIntegratedGpu: true, report: report);
                }

                TickDevicesBusy("Applying USB selective suspend...", 1);
                ApplyUsbSelectiveSuspendPowerPlan(forceDisable: true, report);

                WriteLog("UI: AUTO-OPTIMIZATION applied and saved");
                if (applyImod)
                {
                    TickDevicesBusy("Applying USB IMOD...", 1);
                    ImodApplyOutcome imodOutcome = ApplyImodSettings(out string? imodNote);
                    if (!string.IsNullOrWhiteSpace(imodNote))
                    {
                        WriteLog($"IMOD.NOTE: {imodNote}");
                    }

                    if (imodOutcome == ImodApplyOutcome.Failed)
                    {
                        report.AddError("IMOD", imodNote ?? "apply failed");
                    }
                    else if (!string.IsNullOrWhiteSpace(imodNote)
                        && (imodNote.Contains("failure", StringComparison.OrdinalIgnoreCase)
                            || imodNote.Contains("failed", StringComparison.OrdinalIgnoreCase)))
                    {
                        report.AddError("IMOD", imodNote);
                    }
                }
                else
                {
                    WriteLog("IMOD skipped (AUTO-OPTIMIZATION): no eligible XHCI controllers");
                }

                WriteLog($"UI: AUTO-OPTIMIZATION done errors={report.Errors.Count} -> triggering REFRESH");
                SetDevicesBusyStage("Refreshing devices...");
                RefreshBlocks();
                _devicesBusyDone = _devicesBusyTotal;
                UpdateDevicesBusy("Ready", 100);
            }
            finally
            {
                EndDevicesBusy();
            }

            // Dialog only after progress overlay is closed — never over a mid-stage %.
            if (!showAutoResult)
            {
                return;
            }

            if (dryRunInfo is not null)
            {
                ShowThemedInfo(dryRunInfo);
            }
            else
            {
                ShowOperationResult(
                    report,
                    successMessage: "AUTO-OPTIMIZATION completed and saved.\nPlease reboot your PC to finish applying the changes.",
                    partialMessage: "AUTO-OPTIMIZATION finished with errors. Some changes may be incomplete.");
            }
        };
        btnReset.Click += (_, _) =>
        {
            WriteLog("UI: RESET ALL button clicked");
            OperationReport report = new();
            if (!_testAutoDryRun)
            {
                if (!CreateDeviceTweakerBackup("pre-reset", showDialog: false))
                {
                    report.AddError("Automatic backup", "backup could not be created; changes were not applied");
                    ShowOperationResult(
                        report,
                        successMessage: string.Empty,
                        partialMessage: "RESET ALL was cancelled because the automatic backup failed.");
                    return;
                }
            }
            else
            {
                WriteLog("RESET: dry-run -> skipped pre-reset backup");
            }

            BeginDevicesBusyWork("Running RESET ALL...", Math.Max(4, _blocks.Count + 4));
            try
            {
                ResetAllTweaks(report);
            }
            finally
            {
                EndDevicesBusy();
            }
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
            Size = UiScale(178, 36),
            Margin = new Padding(UiScale(8), UiScale(4), UiScale(8), UiScale(4)),
            FlatStyle = FlatStyle.Flat,
            Font = _buttonFont,
            UseVisualStyleBackColor = false,
            Cursor = Cursors.Hand,
            TabStop = false,
        };
    }

    private void SetTopButtonBaseStyle(Button btn)
    {
        bool isPrimary = string.Equals(btn.Text, "AUTO-OPTIMIZATION", StringComparison.OrdinalIgnoreCase);
        btn.FlatAppearance.BorderSize = 1;
        btn.BackColor = _bgForm;
        btn.ForeColor = _fgMain;
        btn.FlatAppearance.BorderColor = isPrimary ? _accent : Color.FromArgb(150, 150, 158);
    }

    /// <summary>Configure padding on a ThemedTextBox host panel.</summary>
    private void StyleDarkTextBox(ThemedTextBox box, int leftMargin = 6, int rightMargin = 4)
    {
        box.ContentLeftPadding = leftMargin;
        box.ContentRightPadding = rightMargin;
        box.ApplyContentLayout();
    }

    /// <summary>Legacy no-op kept for any remaining plain TextBox call sites.</summary>
    private void StyleDarkTextBox(TextBox box, int leftMargin = 6, int rightMargin = 4)
    {
        _ = box;
        _ = leftMargin;
        _ = rightMargin;
    }

    private void SetTopButtonHoverStyle(Button btn)
    {
        btn.BackColor = _accent;
        btn.ForeColor = Color.FromArgb(15, 15, 15);
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

    private void OpenUrl(string url)
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
        catch (Exception ex)
        {
            WriteLog($"UI.URL.ERROR: url=\"{url}\" error=\"{FlattenLogText(ex.ToString())}\"");
        }
    }
}
