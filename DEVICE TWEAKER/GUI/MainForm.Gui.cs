using System.Diagnostics;
using System.Globalization;
using System.Reflection;


namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private void InitializeGui()
    {
        UpdateUiScale();
        Text = "DEVICE TWEAKER";
        // Allocate the complete two-column USB/IMOD layout before the first
        // visible frame. Previously the form started at 1120 and grew after
        // enumeration, which made the header jump and briefly clipped settings.
        Size formSize = UiScale(1172, 875);
        Size = formSize;
        StartPosition = FormStartPosition.CenterScreen;
        AutoScaleMode = AutoScaleMode.None;
        BackColor = _bgForm;
        ForeColor = _fgMain;
        Font = _baseFont;
        KeyPreview = true;

        FormBorderStyle = FormBorderStyle.Sizable;
        MaximizeBox = true;
        MinimumSize = UiScale(1172, 840);
        MaximumSize = Size.Empty;
        SizeGripStyle = SizeGripStyle.Show;
        ApplyQaInitialWindowBounds();

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
            Padding = new Padding(UiScale(2), 0, UiScale(2), 0),
            Margin = new Padding(0, UiScale(2), 0, 0),
        };

        const string developerHandle = "@arsenza";
        const string developerUrl = "https://t.me/arsenzaa";
        string informationalVersion = Assembly.GetEntryAssembly()
            ?.GetCustomAttribute<AssemblyInformationalVersionAttribute>()
            ?.InformationalVersion
            ?? "0.0.4-alpha.2";
        if (informationalVersion.Contains('+'))
        {
            informationalVersion = informationalVersion.Split('+')[0];
        }
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
            MaximumSize = new Size(UiScale(940), UiScale(24)),
            MinimumSize = new Size(1, UiScale(20)),
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
        brandPanel.Controls.Add(brandLayout);
        AddLanguageSelector(brandPanel);

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
            Height = UiScale(68),
            BackColor = _bgPanel,
            Padding = new Padding(UiScale(24), UiScale(4), UiScale(24), UiScale(16)),
            Margin = Padding.Empty,
        };

        Button btnScan = NewTopButton("REFRESH");
        Button btnApply = NewTopButton("APPLY");
        Button btnAuto = NewTopButton("AUTO-OPTIMIZATION");
        Button btnRestore = NewTopButton("RESTORE");
        _btnScanRef = btnScan;
        _btnApplyRef = btnApply;
        _btnAutoRef = btnAuto;
        _btnRestoreRef = btnRestore;
        _operationButtons = [btnScan, btnApply, btnAuto, btnRestore];
        int buttonGap = UiScale(8);
        btnApply.Margin = Padding.Empty;
        btnAuto.Margin = new Padding(buttonGap, 0, 0, 0);
        btnScan.Margin = new Padding(buttonGap, 0, 0, 0);
        btnRestore.Margin = new Padding(buttonGap, 0, 0, 0);

        foreach (Button b in new[] { btnScan, btnApply, btnAuto, btnRestore })
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
            ColumnCount = 4,
            RowCount = 1,
            BackColor = _bgPanel,
            Anchor = AnchorStyles.None,
            Margin = Padding.Empty,
            Padding = Padding.Empty,
        };
        buttonsGrid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        buttonsGrid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        buttonsGrid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        buttonsGrid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        buttonsGrid.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        buttonsGrid.Controls.Add(btnApply, 0, 0);
        buttonsGrid.Controls.Add(btnAuto, 1, 0);
        buttonsGrid.Controls.Add(btnScan, 2, 0);
        buttonsGrid.Controls.Add(btnRestore, 3, 0);

        buttonsHost.Controls.Add(buttonsGrid, 1, 0);
        buttonPanel.Controls.Add(buttonsHost);

        Panel accentStrip = new()
        {
            Dock = DockStyle.Top,
            Height = UiScale(1),
            BackColor = _accent,
        };

        Panel filterDivider = new()
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

        int scrollWidth = UiScale(14);
        _devicesPanel = new BufferedPanel
        {
            Dock = DockStyle.None,
            BackColor = _bgForm,
            AutoScroll = false,
            Padding = new Padding(UiScale(24), UiScale(18), UiScale(24), UiScale(24)),
        };
        _devicesPanel.Location = new Point(0, 0);
        _devicesPanel.SizeChanged += (_, _) => SyncDevicesScrollBar();

        _devicesScroll = new ThemedScrollBar
        {
            Width = scrollWidth,
            BackColor = _bgForm,
            TrackColor = _bgForm,
            RailColor = _bgForm,
            RailWidth = 0,
            ThumbColor = _accent,
            ThumbHoverColor = Color.White,
            ThumbDragColor = Color.FromArgb(210, 210, 210),
            ThumbWidth = UiScale(10),
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

        _noMatchesLabel = new Label
        {
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleCenter,
            Font = _dialogFont,
            ForeColor = _statusInactive,
            BackColor = _bgForm,
            Text = $"{UiLanguage.Text("NO MATCHING DEVICES")}\n{UiLanguage.Text("No devices match the current filter criteria.")}",
            Visible = false,
        };
        _devicesHost.Controls.Add(_noMatchesLabel);

        InitializeFilterToolbar();

        Controls.Add(_devicesHost);
        Controls.Add(filterDivider);
        Controls.Add(_filterPanel);
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
        UpdateFilterToolbarLocalization();
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
            if (_devicesBusyDepth > 0)
            {
                WriteLog("UI: REFRESH ignored because another operation is active");
                return;
            }

            WriteLog("UI: REFRESH button clicked");
            RefreshBlocks();
        };
        btnApply.Click += (_, _) =>
        {
            if (_devicesBusyDepth > 0)
            {
                WriteLog("UI: APPLY ignored because another operation is active");
                return;
            }

            WriteLog("UI: APPLY button clicked");
            OperationReport report = new();
            bool sandboxDryRun = IsSandboxDryRunActive();
            if (!sandboxDryRun)
            {
                if (!CreateDeviceTweakerBackup("pre-apply", showDialog: false))
                {
                    report.MarkNoChangesMade();
                    report.AddError("Automatic backup", "backup could not be created; changes were not applied");
                    ShowOperationResult(
                        report,
                        successMessage: string.Empty,
                        partialMessage: "APPLY was cancelled because the automatic backup failed.",
                        operationName: "APPLY");
                    return;
                }

                report.SetBackupPath(_lastBackupPath);
            }
            else
            {
                WriteLog("APPLY: dry-run -> skipped pre-apply backup");
            }

            int applyTotal = Math.Max(1, _blocks.Count) + 3;
            BeginDevicesBusyWork(sandboxDryRun ? "Previewing APPLY..." : "Applying changes...", applyTotal);
            try
            {
                int errorsBeforeDevices = report.Errors.Count;
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
                if (report.Errors.Count == errorsBeforeDevices)
                {
                    report.AddSuccess("DEVICE SETTINGS", $"{_blocks.Count} processed");
                }

                int errorsBeforePower = report.Errors.Count;
                TickDevicesBusy(sandboxDryRun ? "Skipping USB selective suspend..." : "Applying USB selective suspend...", 1);
                if (sandboxDryRun)
                {
                    WriteLog("APPLY.DRYRUN: USB selective suspend skipped");
                }
                else
                {
                    ApplyUsbSelectiveSuspendPowerPlan(forceDisable: false, report);
                }
                if (!sandboxDryRun && report.Errors.Count == errorsBeforePower)
                {
                    report.AddSuccess("USB POWER", "Applied");
                }

                TickDevicesBusy(sandboxDryRun ? "Skipping USB IMOD..." : "Applying USB IMOD...", 1);
                if (sandboxDryRun)
                {
                    WriteLog("APPLY.DRYRUN: IMOD apply skipped");
                }
                else
                {
                    ImodApplyOutcome imodOutcome = ApplyImodSettings(out string? imodNote, out string? imodTechnicalDetails);
                    if (!string.IsNullOrWhiteSpace(imodNote))
                    {
                        WriteLog($"IMOD.NOTE: {imodNote}");
                    }

                    AddImodResultToReport(report, imodOutcome, imodNote, imodTechnicalDetails);
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

            if (!sandboxDryRun)
            {
                UpdateAllBlocksInitialState();
                UpdateApplyButtonDirtyCount();
            }

            if (sandboxDryRun)
            {
                ShowThemedInfo("APPLY preview completed.\nSandbox dry-run is ON (no registry changes).");
            }
            else
            {
                ShowOperationResult(
                    report,
                    successMessage: "Please reboot your PC to finish applying the changes.",
                    partialMessage: "Applied steps were saved. One or more steps were not completed.",
                    operationName: "APPLY");
                MaybeOfferVulnerableDriverBlocklistDisable(report);
            }
        };
        btnAuto.Click += (_, _) =>
        {
            if (_devicesBusyDepth > 0)
            {
                WriteLog("UI: AUTO-OPTIMIZATION ignored because another operation is active");
                return;
            }

            WriteLog("UI: AUTO-OPTIMIZATION button clicked");
            bool hasUsbImodTarget = _blocks.Any(b => IsUsbImodTarget(b.Device));
            bool optimizeUsbImod = false;
            if (hasUsbImodTarget)
            {
                optimizeUsbImod = ShowThemedConfirm(
                    "USB IMOD tuning is available for detected XHCI controller(s).\n\nDTIMOD can be blocked by Vulnerable Driver Blocklist, Windows driver signature protection, antivirus, or anti-cheats.\n\nApply it during AUTO-OPTIMIZATION?",
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
                        report.MarkNoChangesMade();
                        report.AddError("Automatic backup", "backup could not be created; changes were not applied");
                        ShowOperationResult(
                            report,
                            successMessage: string.Empty,
                            partialMessage: "AUTO-OPTIMIZATION was cancelled because the automatic backup failed.",
                            operationName: "AUTO-OPTIMIZATION");
                        return;
                    }

                    report.SetBackupPath(_lastBackupPath);
                }
                else
                {
                    if (!EnsureOriginalDeviceTweakerBackup(BackupLocation.Local))
                    {
                        report.MarkNoChangesMade();
                        report.AddError("Original-state backup", "the initial recovery snapshot could not be created; changes were not applied");
                        ShowOperationResult(
                            report,
                            successMessage: string.Empty,
                            partialMessage: "AUTO-OPTIMIZATION was cancelled because the original-state backup failed.",
                            operationName: "AUTO-OPTIMIZATION");
                        return;
                    }

                    WriteLog("BACKUP.PROMPT.AUTO: operation backup skipped; original-state snapshot retained");
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
                bool planBuilt = InvokeAutoOptimization(optimizeUsbImod, hasUsbImodTarget, report);
                bool applyImod = optimizeUsbImod && hasUsbImodTarget;
                if (applyImod)
                {
                    SetDevicesBusyWork(_devicesBusyTotal + 1, _devicesBusyDone);
                }

                if (!planBuilt)
                {
                    report.MarkNoChangesMade();
                    WriteLog("AUTO: plan was not built -> skipping apply/save");
                    _devicesBusyDone = _devicesBusyTotal;
                    UpdateDevicesBusy("Ready", 100);
                    showAutoResult = true;
                    goto AutoCompleted;
                }
                report.AddSuccess("AUTO PLAN", "Built");
                if (!applyImod)
                {
                    string skipReason = GetAutoImodSkipReason(hasUsbImodTarget);
                    WriteLog($"AUTO.IMOD.RESULT: status=skipped reason={skipReason}");
                }

                if (_testAutoDryRun)
                {
                    WriteLog("AUTO.DRYRUN: enabled -> skipping registry writes");
                    if (applyImod)
                    {
                        WriteLog("AUTO.IMOD.RESULT: status=preview-only reason=dry-run");
                    }

                    _devicesBusyDone = _devicesBusyTotal;
                    UpdateDevicesBusy("Ready", 100);
                    dryRunInfo = "AUTO-OPTIMIZATION preview completed.\nSandbox dry-run is ON (no registry changes).";
                    goto AutoCompleted;
                }

                int errorsBeforeDevices = report.Errors.Count;
                int saved = 0;
                int deviceWriteAttempts = 0;
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
                    deviceWriteAttempts++;
                    SaveBlockSettings(b, autoMsiOnly: IsAutoMsiOnlyDevice(b), report: report);
                }
                if (report.Errors.Count == errorsBeforeDevices)
                {
                    report.AddSuccess("DEVICE SETTINGS", $"{deviceWriteAttempts} processed");
                }

                int errorsBeforePower = report.Errors.Count;
                TickDevicesBusy("Applying USB selective suspend...", 1);
                ApplyUsbSelectiveSuspendPowerPlan(forceDisable: true, report);
                if (report.Errors.Count == errorsBeforePower)
                {
                    report.AddSuccess("USB POWER", "Applied");
                }

                WriteLog("UI: AUTO-OPTIMIZATION applied and saved");
                if (applyImod)
                {
                    TickDevicesBusy("Applying USB IMOD...", 1);
                    ImodApplyOutcome imodOutcome = ApplyImodSettings(out string? imodNote, out string? imodTechnicalDetails);
                    if (!string.IsNullOrWhiteSpace(imodNote))
                    {
                        WriteLog($"IMOD.NOTE: {imodNote}");
                    }

                    AddImodResultToReport(report, imodOutcome, imodNote, imodTechnicalDetails);
                    WriteLog($"AUTO.IMOD.RESULT: status={imodOutcome}");
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

        AutoCompleted:
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
                    successMessage: "Please reboot your PC to finish applying the changes.",
                    partialMessage: "Applied steps were saved. One or more steps were not completed.",
                    operationName: "AUTO-OPTIMIZATION");
                MaybeOfferVulnerableDriverBlocklistDisable(report);
            }
        };
        btnRestore.Click += (_, _) =>
        {
            if (_devicesBusyDepth > 0)
            {
                WriteLog("UI: RESTORE ignored because another operation is active");
                return;
            }

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

    private void ApplyQaInitialWindowBounds()
    {
        string? value = Environment.GetEnvironmentVariable("DEVICE_TWEAKER_QA_WINDOW_BOUNDS");
        if (string.IsNullOrWhiteSpace(value))
        {
            return;
        }

        string[] parts = value.Split(',');
        if (parts.Length != 4
            || !int.TryParse(parts[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out int x)
            || !int.TryParse(parts[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out int y)
            || !int.TryParse(parts[2], NumberStyles.Integer, CultureInfo.InvariantCulture, out int width)
            || !int.TryParse(parts[3], NumberStyles.Integer, CultureInfo.InvariantCulture, out int height)
            || width < MinimumSize.Width
            || height < MinimumSize.Height)
        {
            return;
        }

        Rectangle requested = new(x, y, width, height);
        bool fitsKnownScreen = Screen.AllScreens.Any(screen => screen.WorkingArea.Contains(requested));
        if (!fitsKnownScreen)
        {
            return;
        }

        StartPosition = FormStartPosition.Manual;
        Bounds = requested;
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

    private void AdjustInitialDeviceViewportHeight()
    {
        if (_initialDeviceViewportHeightAdjusted
            || IsDisposed
            || !IsHandleCreated
            || _devicesHost is null)
        {
            return;
        }

        // Device cards use the internal scrollbar. Resizing the top-level form
        // after it is already visible makes the header and language selector
        // jump during startup/REFRESH, so the user-selected window bounds stay
        // authoritative for the complete session.
        _initialDeviceViewportHeightAdjusted = true;
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

    private void ResetDevicesScroll()
    {
        if (_devicesHost is null || _devicesPanel is null)
        {
            return;
        }

        _devicesPanel.Location = new Point(0, 0);
        if (_devicesScroll is not null)
        {
            _syncingScroll = true;
            _devicesScroll.Value = 0;
            _syncingScroll = false;
        }

        SyncDevicesScrollBar();
        _devicesPanel.Invalidate();
        _devicesHost.Invalidate();
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
        return new ThemedButton
        {
            Name = text,
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
        bool isPrimary = string.Equals(btn.Name, "AUTO-OPTIMIZATION", StringComparison.OrdinalIgnoreCase);
        btn.FlatAppearance.BorderSize = 1;
        btn.BackColor = _bgForm;
        btn.ForeColor = _fgMain;
        if (btn == _btnApplyRef && _blocks.Any(b => b.IsDirty))
        {
            btn.FlatAppearance.BorderColor = _statusWarn;
        }
        else
        {
            btn.FlatAppearance.BorderColor = isPrimary ? _accent : Color.FromArgb(150, 150, 158);
        }
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

    private void InitializeFilterToolbar()
    {
        _filterPanel = new Panel
        {
            Dock = DockStyle.Top,
            Height = UiScale(54),
            BackColor = _bgPanel,
            Padding = new Padding(UiScale(24), UiScale(12), UiScale(24), UiScale(12)),
            Margin = Padding.Empty,
        };

        FlowLayoutPanel searchHost = new()
        {
            Dock = DockStyle.Left,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            BackColor = _bgPanel,
            Margin = Padding.Empty,
            Padding = Padding.Empty,
        };

        _searchFilterBox = new ThemedTextBox
        {
            Width = UiScale(240),
            Height = UiScale(30),
            PlaceholderText = UiLanguage.Text("Filter devices... (Ctrl+F)"),
            Font = _baseFont,
            Margin = Padding.Empty,
            TabStop = false,
        };
        _searchFilterBox.Inner.TabStop = false;
        _searchFilterBox.TextChanged += (_, _) =>
        {
            string newQuery = _searchFilterBox.Text.Trim();
            if (string.Equals(_searchFilterText, newQuery, StringComparison.Ordinal))
            {
                return;
            }

            _searchFilterText = newQuery;
            _btnFilterClear.Visible = !string.IsNullOrEmpty(_searchFilterText);
            ResetDevicesScroll();
            ApplyDeviceFilter();
            int visibleCount = _blocks.Count(b => b.Group.Visible);
            WriteLog($"UI: Filter search text=\"{_searchFilterText}\" visibleDevices={visibleCount}/{_blocks.Count}");
        };

        _btnFilterClear = new ThemedButton
        {
            Text = "✕",
            Size = new Size(UiScale(24), UiScale(30)),
            Visible = false,
            FlatStyle = FlatStyle.Flat,
            Font = _baseFont,
            ForeColor = _statusInactive,
            BackColor = _bgPanel,
            Cursor = Cursors.Hand,
            TabStop = false,
            Margin = new Padding(UiScale(2), 0, 0, 0),
        };
        _btnFilterClear.FlatAppearance.BorderSize = 0;
        _btnFilterClear.Click += (_, _) => ClearFilterBox();

        searchHost.Controls.Add(_searchFilterBox);
        searchHost.Controls.Add(_btnFilterClear);

        _filterCategoriesHost = new FlowLayoutPanel
        {
            Dock = DockStyle.Right,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            BackColor = _bgPanel,
            Margin = Padding.Empty,
            Padding = Padding.Empty,
        };

        _filterPanel.Controls.Add(_filterCategoriesHost);
        _filterPanel.Controls.Add(searchHost);

        RebuildFilterCategoryButtons();
    }

    private void UpdateCategoryButtonSize(Button btn)
    {
        int textWidth = TextRenderer.MeasureText(btn.Text, btn.Font).Width;
        btn.Size = new Size(textWidth + UiScale(16), UiScale(30));
    }

    private static Color GetCategoryAccentColor(string category)
    {
        return category?.ToUpperInvariant() switch
        {
            "ALL" => Color.FromArgb(0, 150, 255),          // Modern Azure
            "MOUSE" => Color.FromArgb(64, 215, 240),        // Precision Cyan
            "KEYBOARD" => Color.FromArgb(245, 185, 75),     // Mechanical Gold
            "GAMEPAD" => Color.FromArgb(185, 135, 255),     // Controller Violet
            "USB" => Color.FromArgb(110, 195, 255),         // USB Sky Blue
            "GPU" => Color.FromArgb(120, 225, 100),         // GPU Neon Green
            "NETWORK" or "NET" => Color.FromArgb(70, 220, 145), // Network Emerald
            "AUDIO" => Color.FromArgb(255, 140, 105),       // Audio Coral
            "STORAGE" or "STOR" => Color.FromArgb(145, 205, 235), // Storage Ice Silver
            _ => Color.FromArgb(180, 185, 195),
        };
    }

    private void RebuildFilterCategoryButtons()
    {
        if (_filterCategoriesHost == null)
        {
            return;
        }

        _filterCategoriesHost.SuspendLayout();
        _filterCategoriesHost.Controls.Clear();
        _filterCategoryButtons.Clear();

        List<string> availableCategories = ["ALL"];
        string[] candidates = ["MOUSE", "KEYBOARD", "GAMEPAD", "USB", "GPU", "NETWORK", "AUDIO", "STORAGE"];
        foreach (string candidate in candidates)
        {
            if (_blocks.Any(b => MatchesCategory(b, candidate)))
            {
                availableCategories.Add(candidate);
            }
        }

        string? envFilter = Environment.GetEnvironmentVariable("DEVICE_TWEAKER_CATEGORY_FILTER");
        if (!string.IsNullOrWhiteSpace(envFilter) && availableCategories.Contains(envFilter, StringComparer.OrdinalIgnoreCase))
        {
            _activeCategoryFilter = envFilter.ToUpperInvariant();
        }
        else if (!availableCategories.Contains(_activeCategoryFilter, StringComparer.OrdinalIgnoreCase))
        {
            _activeCategoryFilter = "ALL";
        }

        foreach (string cat in availableCategories)
        {
            Button btnCat = new ThemedButton
            {
                Name = $"FILTER_{cat}",
                Text = UiLanguage.Text($"[ {cat} ]"),
                Tag = cat,
                AutoSize = false,
                Height = UiScale(30),
                FlatStyle = FlatStyle.Flat,
                Font = _headerFont,
                Cursor = Cursors.Hand,
                TabStop = false,
                Margin = new Padding(UiScale(4), 0, 0, 0),
            };
            btnCat.MouseEnter += (_, _) =>
            {
                if (!string.Equals(btnCat.Tag as string, _activeCategoryFilter, StringComparison.OrdinalIgnoreCase))
                {
                    btnCat.BackColor = Color.FromArgb(28, 30, 38);
                    btnCat.ForeColor = Color.FromArgb(220, 225, 235);
                    btnCat.FlatAppearance.BorderColor = Color.FromArgb(75, 80, 95);
                }
            };
            btnCat.MouseLeave += (_, _) =>
            {
                if (!string.Equals(btnCat.Tag as string, _activeCategoryFilter, StringComparison.OrdinalIgnoreCase))
                {
                    btnCat.BackColor = _bgPanel;
                    btnCat.ForeColor = Color.FromArgb(140, 145, 155);
                    btnCat.FlatAppearance.BorderColor = Color.FromArgb(50, 52, 62);
                }
            };
            UpdateCategoryButtonSize(btnCat);
            btnCat.Click += (_, _) => SetCategoryFilter(cat);
            _filterCategoryButtons.Add(btnCat);
            _filterCategoriesHost.Controls.Add(btnCat);
        }

        _filterCategoriesHost.ResumeLayout(true);
        UpdateCategoryButtonsStyle();
        WriteLog($"UI: Filter categories available=[{string.Join(", ", availableCategories)}] active=\"{_activeCategoryFilter}\"");
    }

    private void SetCategoryFilter(string category)
    {
        if (string.Equals(_activeCategoryFilter, category, StringComparison.OrdinalIgnoreCase))
        {
            ResetDevicesScroll();
            return;
        }

        string prev = _activeCategoryFilter;
        _activeCategoryFilter = category;
        UpdateCategoryButtonsStyle();
        ResetDevicesScroll();
        ApplyDeviceFilter();
        int visibleCount = _blocks.Count(b => b.Group.Visible);
        WriteLog($"UI: Filter category selected: {category} (previous={prev}) visibleDevices={visibleCount}/{_blocks.Count}");
    }

    private void UpdateCategoryButtonsStyle()
    {
        foreach (Button btn in _filterCategoryButtons)
        {
            string cat = btn.Tag as string ?? string.Empty;
            bool isActive = string.Equals(cat, _activeCategoryFilter, StringComparison.OrdinalIgnoreCase);
            Color catAccent = GetCategoryAccentColor(cat);

            if (isActive)
            {
                btn.BackColor = Color.FromArgb(32, 35, 46);
                btn.ForeColor = catAccent;
                btn.FlatAppearance.BorderColor = catAccent;
                btn.FlatAppearance.BorderSize = 1;
            }
            else
            {
                btn.BackColor = _bgPanel;
                btn.ForeColor = Color.FromArgb(140, 145, 155);
                btn.FlatAppearance.BorderColor = Color.FromArgb(50, 52, 62);
                btn.FlatAppearance.BorderSize = 1;
            }
        }
    }

    private void ApplyDeviceFilter()
    {
        ResetDevicesScroll();

        bool anyVisible = false;
        foreach (DeviceBlock b in _blocks)
        {
            bool catMatch = MatchesCategory(b, _activeCategoryFilter);
            bool searchMatch = MatchesSearchText(b, _searchFilterText);
            bool visible = catMatch && searchMatch;
            b.Group.Visible = visible;
            if (visible)
            {
                anyVisible = true;
            }
        }

        if (_noMatchesLabel != null)
        {
            _noMatchesLabel.Visible = !anyVisible && _blocks.Count > 0;
            if (_noMatchesLabel.Visible)
            {
                _noMatchesLabel.BringToFront();
            }
        }

        LayoutBlocks();
    }

    private static bool MatchesCategory(DeviceBlock b, string cat)
    {
        if (string.Equals(cat, "ALL", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        if (cat.Contains(',') || cat.Contains(';'))
        {
            string[] parts = cat.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            return parts.Any(p => MatchesCategory(b, p));
        }

        DeviceInfo info = b.Device;
        if (string.Equals(cat, "MOUSE", StringComparison.OrdinalIgnoreCase))
        {
            return info.Kind == DeviceKind.USB && info.UsbRoles.Contains("Mouse", StringComparison.OrdinalIgnoreCase);
        }
        if (string.Equals(cat, "KEYBOARD", StringComparison.OrdinalIgnoreCase))
        {
            return info.Kind == DeviceKind.USB && info.UsbRoles.Contains("Keyboard", StringComparison.OrdinalIgnoreCase);
        }
        if (string.Equals(cat, "GAMEPAD", StringComparison.OrdinalIgnoreCase))
        {
            return info.Kind == DeviceKind.USB && (
                info.UsbRoles.Contains("Gamepad", StringComparison.OrdinalIgnoreCase) ||
                info.UsbRoles.Contains("Controller", StringComparison.OrdinalIgnoreCase) ||
                info.UsbRoles.Contains("Joystick", StringComparison.OrdinalIgnoreCase));
        }
        if (string.Equals(cat, "USB", StringComparison.OrdinalIgnoreCase))
        {
            return info.Kind == DeviceKind.USB;
        }
        if (string.Equals(cat, "GPU", StringComparison.OrdinalIgnoreCase))
        {
            return info.Kind == DeviceKind.GPU;
        }
        if (string.Equals(cat, "NETWORK", StringComparison.OrdinalIgnoreCase) || string.Equals(cat, "NET", StringComparison.OrdinalIgnoreCase))
        {
            return info.Kind is DeviceKind.NET_NDIS or DeviceKind.NET_CX;
        }
        if (string.Equals(cat, "AUDIO", StringComparison.OrdinalIgnoreCase))
        {
            return info.Kind == DeviceKind.AUDIO;
        }
        if (string.Equals(cat, "STORAGE", StringComparison.OrdinalIgnoreCase) || string.Equals(cat, "STOR", StringComparison.OrdinalIgnoreCase))
        {
            return info.Kind == DeviceKind.STOR;
        }

        return true;
    }

    private static bool MatchesSearchText(DeviceBlock b, string search)
    {
        if (string.IsNullOrWhiteSpace(search))
        {
            return true;
        }

        DeviceInfo info = b.Device;
        return (info.Name?.Contains(search, StringComparison.OrdinalIgnoreCase) == true)
            || (info.InstanceId?.Contains(search, StringComparison.OrdinalIgnoreCase) == true)
            || (info.Class?.Contains(search, StringComparison.OrdinalIgnoreCase) == true)
            || (info.RegBase?.Contains(search, StringComparison.OrdinalIgnoreCase) == true)
            || (info.UsbRoles?.Contains(search, StringComparison.OrdinalIgnoreCase) == true)
            || (info.AudioEndpoints?.Contains(search, StringComparison.OrdinalIgnoreCase) == true)
            || (info.StorageTag?.Contains(search, StringComparison.OrdinalIgnoreCase) == true);
    }

    private void FocusFilterBox()
    {
        _searchFilterBox?.Inner.Focus();
        _searchFilterBox?.Inner.SelectAll();
        WriteLog("UI: Filter search box focused");
    }

    private void ClearFilterBox()
    {
        if (_searchFilterBox != null)
        {
            _searchFilterBox.Text = string.Empty;
        }
        _searchFilterText = string.Empty;
        _btnFilterClear.Visible = false;
        ResetDevicesScroll();
        ApplyDeviceFilter();
        int visibleCount = _blocks.Count(b => b.Group.Visible);
        WriteLog($"UI: Filter search cleared visibleDevices={visibleCount}/{_blocks.Count}");
    }

    private void UpdateApplyButtonDirtyCount()
    {
        if (_btnApplyRef == null)
        {
            return;
        }

        int dirtyCount = _blocks.Count(b => b.IsDirty);
        _btnApplyRef.Text = UiLanguage.Text("APPLY");
        SetTopButtonBaseStyle(_btnApplyRef);

        if (dirtyCount > 0)
        {
            _copyToolTip.SetToolTip(_btnApplyRef, $"Apply changes (Ctrl+S) — modified: {dirtyCount}");
        }
        else
        {
            _copyToolTip.SetToolTip(_btnApplyRef, "Apply changes (Ctrl+S)");
        }
    }

    private void UpdateAllBlocksInitialState()
    {
        foreach (DeviceBlock b in _blocks)
        {
            b.CaptureInitialState();
        }
    }

    private void OnBlockSettingChanged(DeviceBlock block)
    {
        block.CheckIsDirty();
        UpdateApplyButtonDirtyCount();
    }

    private void UpdateFilterToolbarLocalization()
    {
        if (_searchFilterBox != null)
        {
            _searchFilterBox.PlaceholderText = UiLanguage.Text("Filter devices... (Ctrl+F)");
        }
        foreach (Button btn in _filterCategoryButtons)
        {
            if (btn.Tag is string cat)
            {
                btn.Text = UiLanguage.Text($"[ {cat} ]");
                UpdateCategoryButtonSize(btn);
            }
        }
        if (_noMatchesLabel != null)
        {
            _noMatchesLabel.Text = $"{UiLanguage.Text("NO MATCHING DEVICES")}\n{UiLanguage.Text("No devices match the current filter criteria.")}";
        }
        if (_btnScanRef != null)
        {
            _copyToolTip.SetToolTip(_btnScanRef, "Refresh devices (F5 / Ctrl+R)");
        }
        if (_btnApplyRef != null)
        {
            UpdateApplyButtonDirtyCount();
        }
        if (_btnAutoRef != null)
        {
            _copyToolTip.SetToolTip(_btnAutoRef, "Auto-optimization (Ctrl+O)");
        }
        if (_btnRestoreRef != null)
        {
            _copyToolTip.SetToolTip(_btnRestoreRef, "Restore settings (Ctrl+Z)");
        }
    }
}
