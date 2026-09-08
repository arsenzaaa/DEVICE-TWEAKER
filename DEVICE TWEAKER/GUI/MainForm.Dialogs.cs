
namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private void StyleThemedDialogSurface(Form dialog)
    {
        dialog.BackColor = _bgForm;
        dialog.ForeColor = _fgMain;
    }

    private DialogResult ShowDialogDimmed(Form dialog)
    {
        LocalizeControlTree(dialog);
        Form? dimmer = null;
        bool ownDimmer = false;
        try
        {
            if (_dialogDimDepth == 0
                && IsHandleCreated
                && Visible
                && WindowState != FormWindowState.Minimized)
            {
                dimmer = new Form
                {
                    FormBorderStyle = FormBorderStyle.None,
                    ShowInTaskbar = false,
                    StartPosition = FormStartPosition.Manual,
                    BackColor = Color.Black,
                    Opacity = 0.52,
                    Bounds = Bounds,
                    Owner = this,
                };
                dimmer.Show(this);
                ownDimmer = true;
            }

            _dialogDimDepth++;
            return dialog.ShowDialog(this);
        }
        finally
        {
            _dialogDimDepth = Math.Max(0, _dialogDimDepth - 1);
            if (ownDimmer && dimmer is not null)
            {
                dimmer.Close();
                dimmer.Dispose();
            }
        }
    }

    private void ShowThemedInfo(string message, string title)
    {
        using Form dialog = new ThemedDialogForm();
        dialog.Name = "INFO_DIALOG";
        dialog.Text = title;
        dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
        dialog.StartPosition = FormStartPosition.CenterParent;
        dialog.MaximizeBox = false;
        dialog.MinimizeBox = false;
        dialog.ShowInTaskbar = false;
        dialog.AutoScaleMode = AutoScaleMode.None;
        StyleThemedDialogSurface(dialog);
        dialog.Font = _dialogFont;
        dialog.Icon = Icon;

        string normalized = NormalizeDialogMessage(UiLanguage.Text(message));

        int padding = UiScale(20);
        int maxTextWidth = UiScale(520);
        int minWidth = UiScale(360);
        int buttonWidth = UiScale(92);
        int buttonHeight = UiScale(32);
        int buttonGap = UiScale(16);

        Label messageLabel = new()
        {
            AutoSize = true,
            MaximumSize = new Size(maxTextWidth, 0),
            Text = normalized,
            ForeColor = _fgMain,
            BackColor = _bgForm,
            UseMnemonic = false,
            TextAlign = ContentAlignment.TopCenter,
            Font = _dialogFont,
            UseCompatibleTextRendering = false,
        };

        Size textSize = messageLabel.GetPreferredSize(new Size(maxTextWidth, 0));
        int clientWidth = Math.Max(minWidth, textSize.Width + (padding * 2));
        int labelWidth = clientWidth - (padding * 2);
        messageLabel.MaximumSize = new Size(labelWidth, 0);
        textSize = messageLabel.GetPreferredSize(new Size(labelWidth, 0));

        int clientHeight = padding + textSize.Height + buttonGap + buttonHeight + padding;
        dialog.ClientSize = new Size(clientWidth, clientHeight);

        messageLabel.AutoSize = false;
        messageLabel.Size = new Size(labelWidth, textSize.Height);
        messageLabel.Location = new Point(padding, padding);

        Button okButton = new()
        {
            Text = "OK",
            DialogResult = DialogResult.OK,
            Size = new Size(buttonWidth, buttonHeight),
            BackColor = _bgForm,
            ForeColor = _accent,
            FlatStyle = FlatStyle.Flat,
            Font = _buttonFont,
            Location = new Point((clientWidth - buttonWidth) / 2, padding + textSize.Height + buttonGap),
        };
        okButton.FlatAppearance.BorderColor = _accent;
        okButton.FlatAppearance.BorderSize = 1;

        dialog.Controls.Add(messageLabel);
        dialog.Controls.Add(okButton);

        dialog.AcceptButton = okButton;
        WireThemedTitleBar(dialog);

        ShowDialogDimmed(dialog);
    }

    private void ShowThemedInfo(string message)
    {
        ShowThemedInfo(message, "DEVICE TWEAKER");
    }

    private void ShowOperationResult(
        OperationReport report,
        string successMessage,
        string partialMessage,
        string? operationName = null)
    {
        string operation = string.IsNullOrWhiteSpace(operationName)
            ? InferOperationName(report.Succeeded ? successMessage : partialMessage)
            : operationName.Trim();
        bool failedBeforeChanges = !report.Succeeded && report.NoChangesMade;
        string state = report.Succeeded
            ? (report.Warnings.Count > 0 ? "APPLIED WITH WARNINGS" : "COMPLETED")
            : failedBeforeChanges ? "NOT APPLIED" : "PARTIALLY APPLIED";

        using Form dialog = new ThemedDialogForm
        {
            Name = "OPERATION_RESULT_DIALOG",
            Text = operation,
            FormBorderStyle = FormBorderStyle.FixedDialog,
            StartPosition = FormStartPosition.CenterParent,
            MaximizeBox = false,
            MinimizeBox = false,
            ShowInTaskbar = false,
            AutoScaleMode = AutoScaleMode.None,
            Font = _dialogFont,
            Icon = Icon,
        };
        StyleThemedDialogSurface(dialog);
        dialog.BackColor = _bgPanel;

        List<OperationIssue> actionableIssues = report.Issues
            .Where(issue => issue.Severity != OperationIssueSeverity.Success)
            .ToList();
        IReadOnlyList<OperationIssue> visibleIssues = actionableIssues.Take(3).ToList();
        string technicalText = BuildOperationTechnicalDetails(report);
        bool hasTechnical = technicalText.Length > 0;
        string backupFolder = ResolveResultBackupFolder(report.BackupPath);
        string logsFolder = AppDiagnostics.LogDirectory;
        // Always offer LOGS/BACKUPS: folders are created on open if missing.
        // BACKUPS opens this operation's backup folder when known; otherwise EXE\Backups
        // (Local). AppData is used only when that is where the last backup was written
        // or when Local does not exist yet but Roaming already does.

        int padding = UiScale(24);
        int buttonGap = UiScale(12);
        int sectionGap = UiScale(14);
        int buttonHeight = UiScale(34);
        int detailsHeight = UiScale(220);
        int maxTextWidth = UiScale(actionableIssues.Count > 0 ? 540 : 440);
        int minWidth = UiScale(400);
        Rectangle workingArea = Screen.FromControl(this).WorkingArea;

        Label statusLabel = new()
        {
            AutoSize = true,
            Text = UiLanguage.Text(state),
            Font = _titleFont,
            ForeColor = state switch
            {
                "COMPLETED" => _statusSuccess,
                "APPLIED WITH WARNINGS" => _statusWarn,
                _ => _statusDanger,
            },
            BackColor = _bgPanel,
            TextAlign = ContentAlignment.MiddleCenter,
            UseMnemonic = false,
        };
        Label summaryLabel = new()
        {
            AutoSize = true,
            Text = NormalizeDialogMessage(UiLanguage.Text(report.Succeeded ? successMessage : partialMessage)),
            Font = _dialogFont,
            ForeColor = _mutedText,
            BackColor = _bgPanel,
            TextAlign = ContentAlignment.TopCenter,
            UseMnemonic = false,
            MaximumSize = new Size(maxTextWidth, 0),
        };

        List<Control> issueControls = [];
        foreach (OperationIssue issue in visibleIssues)
        {
            string issueState = issue.Severity switch
            {
                OperationIssueSeverity.Warning => "WARNING",
                _ => "NOT APPLIED",
            };
            issueControls.Add(new Label
            {
                AutoSize = true,
                Text = $"{UiLanguage.Text(issue.Component).ToUpperInvariant()}  —  {UiLanguage.Text(issueState)}",
                Font = _blockTitleFont,
                ForeColor = issue.Severity == OperationIssueSeverity.Error ? _statusDanger : _statusWarn,
                BackColor = _bgPanel,
                TextAlign = ContentAlignment.MiddleCenter,
                UseMnemonic = false,
                MaximumSize = new Size(maxTextWidth, 0),
            });
            if (!string.IsNullOrWhiteSpace(issue.UserMessage))
            {
                issueControls.Add(new Label
                {
                    AutoSize = true,
                    Text = NormalizeDialogMessage(UiLanguage.Text(issue.UserMessage)),
                    Font = _dialogFont,
                    ForeColor = _fgMain,
                    BackColor = _bgPanel,
                    TextAlign = ContentAlignment.TopCenter,
                    UseMnemonic = false,
                    MaximumSize = new Size(maxTextWidth, 0),
                });
            }
        }
        if (actionableIssues.Count > visibleIssues.Count)
        {
            issueControls.Add(new Label
            {
                AutoSize = true,
                Text = UiLanguage.IsRussian
                    ? $"+ ЕЩЁ: {actionableIssues.Count - visibleIssues.Count} — ОТКРОЙТЕ ПОДРОБНОСТИ"
                    : $"+ {actionableIssues.Count - visibleIssues.Count} MORE — OPEN DETAILS",
                Font = _technicalFont,
                ForeColor = _statusInactive,
                BackColor = _bgPanel,
                TextAlign = ContentAlignment.MiddleCenter,
                UseMnemonic = false,
                MaximumSize = new Size(maxTextWidth, 0),
            });
        }

        List<Button> visibleButtons = [];
        // Primary dismiss first (left): OK should not sit on the far right of a long row.
        Button okButton = NewOperationResultButton(UiLanguage.Text("OK"), buttonHeight, UiScale(100));
        okButton.Name = "OK";
        okButton.DialogResult = DialogResult.OK;
        visibleButtons.Add(okButton);

        Button? detailsButton = null;
        if (hasTechnical)
        {
            detailsButton = NewOperationResultButton(UiLanguage.Text("DETAILS"), buttonHeight, UiScale(132));
            detailsButton.Name = "DETAILS";
            visibleButtons.Add(detailsButton);
        }

        Button logsButton = NewOperationResultButton(UiLanguage.Text("LOGS"), buttonHeight, UiScale(110));
        logsButton.Name = "LOGS";
        logsButton.AccessibleDescription = logsFolder;
        visibleButtons.Add(logsButton);

        Button backupsButton = NewOperationResultButton(UiLanguage.Text("BACKUPS"), buttonHeight, UiScale(120));
        backupsButton.Name = "BACKUPS";
        backupsButton.AccessibleDescription = backupFolder;
        visibleButtons.Add(backupsButton);

        int buttonRowWidth = visibleButtons.Sum(button => button.Width)
            + (Math.Max(0, visibleButtons.Count - 1) * buttonGap);

        Size Measure(Label label)
        {
            label.MaximumSize = new Size(maxTextWidth, 0);
            return label.GetPreferredSize(new Size(maxTextWidth, 0));
        }

        Size statusSize = Measure(statusLabel);
        Size summarySize = Measure(summaryLabel);
        List<Size> issueSizes = issueControls.OfType<Label>().Select(Measure).ToList();

        int contentWidth = Math.Max(statusSize.Width, summarySize.Width);
        foreach (Size size in issueSizes)
        {
            contentWidth = Math.Max(contentWidth, size.Width);
        }

        int clientWidth = Math.Max(minWidth, Math.Max(contentWidth + (padding * 2), buttonRowWidth + (padding * 2)));
        clientWidth = Math.Min(clientWidth, Math.Max(minWidth, workingArea.Width - UiScale(60)));
        int labelWidth = clientWidth - (padding * 2);

        statusLabel.MaximumSize = new Size(labelWidth, 0);
        summaryLabel.MaximumSize = new Size(labelWidth, 0);
        statusSize = statusLabel.GetPreferredSize(new Size(labelWidth, 0));
        summarySize = summaryLabel.GetPreferredSize(new Size(labelWidth, 0));
        for (int i = 0; i < issueControls.Count; i++)
        {
            if (issueControls[i] is Label issueLabel)
            {
                issueLabel.MaximumSize = new Size(labelWidth, 0);
                issueSizes[i] = issueLabel.GetPreferredSize(new Size(labelWidth, 0));
            }
        }

        int y = padding;
        statusLabel.AutoSize = false;
        statusLabel.Size = new Size(labelWidth, statusSize.Height);
        statusLabel.Location = new Point(padding, y);
        y += statusSize.Height + UiScale(12);

        summaryLabel.AutoSize = false;
        summaryLabel.Size = new Size(labelWidth, summarySize.Height);
        summaryLabel.Location = new Point(padding, y);
        y += summarySize.Height + sectionGap;

        for (int i = 0; i < issueControls.Count; i++)
        {
            if (issueControls[i] is not Label issueLabel)
            {
                continue;
            }

            issueLabel.AutoSize = false;
            issueLabel.Size = new Size(labelWidth, issueSizes[i].Height);
            issueLabel.Location = new Point(padding, y);
            bool isTitle = issueLabel.Font == _blockTitleFont;
            y += issueSizes[i].Height + (isTitle ? UiScale(4) : UiScale(10));
        }
        if (issueControls.Count > 0)
        {
            y += UiScale(2);
        }

        Panel detailsHost = new()
        {
            BackColor = _bgGroup,
            BorderStyle = BorderStyle.FixedSingle,
            Visible = false,
            Location = new Point(padding, y),
            Size = new Size(labelWidth, detailsHeight),
        };
        Panel detailsContent = new()
        {
            BackColor = _bgGroup,
            Location = Point.Empty,
        };
        Label detailsLabel = new()
        {
            AutoSize = true,
            Text = technicalText,
            Font = _technicalFont,
            ForeColor = _mutedText,
            BackColor = _bgGroup,
            UseMnemonic = false,
            Location = new Point(UiScale(10), UiScale(8)),
        };
        detailsContent.Controls.Add(detailsLabel);
        ThemedScrollBar detailsScroll = new()
        {
            Width = UiScale(13),
            Dock = DockStyle.Right,
            BackColor = _bgGroup,
            TrackColor = _bgGroup,
            RailColor = _bgGroup,
            ThumbColor = _accent,
            ThumbWidth = UiScale(8),
            RailWidth = 0,
            ThumbCornerRadius = UiScale(6),
            Visible = false,
        };
        detailsHost.Controls.Add(detailsContent);
        detailsHost.Controls.Add(detailsScroll);

        bool syncingDetailsScroll = false;
        void SyncDetailsLayout()
        {
            int contentW = Math.Max(1, detailsHost.ClientSize.Width - detailsScroll.Width - UiScale(22));
            detailsLabel.MaximumSize = new Size(contentW, 0);
            Size preferred = detailsLabel.GetPreferredSize(new Size(contentW, 0));
            detailsLabel.Size = preferred;
            detailsContent.Width = detailsHost.ClientSize.Width - detailsScroll.Width;
            detailsContent.Height = preferred.Height + UiScale(16);

            int maxOffset = Math.Max(0, detailsContent.Height - detailsHost.ClientSize.Height);
            int offset = Math.Max(0, Math.Min(maxOffset, -detailsContent.Top));
            detailsContent.Location = new Point(0, -offset);
            detailsScroll.Visible = maxOffset > 0;
            syncingDetailsScroll = true;
            detailsScroll.Maximum = Math.Max(detailsContent.Height, 1);
            detailsScroll.ViewportSize = Math.Max(detailsHost.ClientSize.Height, 1);
            detailsScroll.Value = offset;
            syncingDetailsScroll = false;
        }
        detailsScroll.ValueChanged += (_, _) =>
        {
            if (!syncingDetailsScroll)
            {
                detailsContent.Top = -detailsScroll.Value;
            }
        };
        detailsHost.SizeChanged += (_, _) => SyncDetailsLayout();
        detailsHost.MouseWheel += (_, e) =>
        {
            if (detailsScroll.Visible)
            {
                detailsScroll.Value += e.Delta > 0 ? -detailsScroll.SmallChange : detailsScroll.SmallChange;
            }
        };

        int detailsTop = y;
        int dividerTop = y + UiScale(2);
        Panel divider = new()
        {
            BackColor = Color.FromArgb(48, 48, 54),
            Location = new Point(padding, dividerTop),
            Size = new Size(labelWidth, 1),
        };
        int buttonsTop = dividerTop + UiScale(14);
        int rowLeft = (clientWidth - buttonRowWidth) / 2;
        int buttonX = rowLeft;
        foreach (Button button in visibleButtons)
        {
            button.Location = new Point(buttonX, buttonsTop);
            button.BackColor = _bgPanel;
            buttonX += button.Width + buttonGap;
        }
        okButton.FlatAppearance.BorderColor = _accent;

        int compactHeight = buttonsTop + buttonHeight + padding;
        dialog.ClientSize = new Size(clientWidth, Math.Min(compactHeight, workingArea.Height - UiScale(60)));

        dialog.Controls.Add(statusLabel);
        dialog.Controls.Add(summaryLabel);
        foreach (Control control in issueControls)
        {
            dialog.Controls.Add(control);
        }
        dialog.Controls.Add(detailsHost);
        dialog.Controls.Add(divider);
        foreach (Button button in visibleButtons)
        {
            dialog.Controls.Add(button);
        }

        void RelayoutExpanded(bool expanded)
        {
            int nextY = detailsTop;
            detailsHost.Visible = expanded;
            if (expanded)
            {
                detailsHost.Location = new Point(padding, nextY);
                detailsHost.Size = new Size(labelWidth, detailsHeight);
                nextY += detailsHeight + sectionGap;
            }

            divider.Location = new Point(padding, nextY + UiScale(2));
            int nextButtonsTop = nextY + UiScale(14);
            int nextRowLeft = (clientWidth - buttonRowWidth) / 2;
            int nextX = nextRowLeft;
            foreach (Button button in visibleButtons)
            {
                button.Location = new Point(nextX, nextButtonsTop);
                nextX += button.Width + buttonGap;
            }

            int requested = nextButtonsTop + buttonHeight + padding;
            dialog.ClientSize = new Size(clientWidth, Math.Min(requested, workingArea.Height - UiScale(60)));
            if (expanded)
            {
                SyncDetailsLayout();
            }
        }

        logsButton.Click += (_, _) => TryOpenResultFolder(logsFolder, "UI.RESULT.LOGS");
        backupsButton.Click += (_, _) => TryOpenResultFolder(backupFolder, "UI.RESULT.BACKUPS");
        if (detailsButton is not null)
        {
            detailsButton.Click += (_, _) =>
            {
                bool expand = !detailsHost.Visible;
                detailsButton.Text = expand ? UiLanguage.Text("HIDE DETAILS") : UiLanguage.Text("DETAILS");
                detailsButton.Width = expand ? UiScale(160) : UiScale(132);
                buttonRowWidth = visibleButtons.Sum(button => button.Width)
                    + (Math.Max(0, visibleButtons.Count - 1) * buttonGap);
                RelayoutExpanded(expand);
                if (expand)
                {
                    detailsHost.Focus();
                }
            };
        }

        dialog.AcceptButton = okButton;
        WireThemedTitleBar(dialog);
        int successCount = report.Issues.Count(issue => issue.Severity == OperationIssueSeverity.Success);
        WriteLog(
            $"UI.RESULT.SUMMARY: operation=\"{SanitizeLogValue(operation)}\" state={state} " +
            $"results={report.Issues.Count} successes={successCount} warnings={report.Warnings.Count} " +
            $"errors={report.Errors.Count} noChanges={report.NoChangesMade} " +
            $"backup=\"{SanitizeLogValue(report.BackupPath)}\" backupsFolder=\"{SanitizeLogValue(backupFolder)}\"");
        foreach (OperationIssue issue in report.Issues)
        {
            WriteLog(
                $"UI.RESULT.ITEM: severity={issue.Severity} component=\"{SanitizeLogValue(issue.Component)}\" " +
                $"message=\"{SanitizeLogValue(issue.UserMessage)}\" " +
                $"technicalDetails=\"{SanitizeLogValue(issue.TechnicalDetails)}\"");
        }
        ShowDialogDimmed(dialog);
    }

    private string ResolveResultBackupFolder(string? backupPath)
    {
        if (!string.IsNullOrWhiteSpace(backupPath))
        {
            try
            {
                string? directory = Path.GetDirectoryName(Path.GetFullPath(backupPath));
                if (!string.IsNullOrWhiteSpace(directory))
                {
                    return directory;
                }
            }
            catch
            {
                // Fall through to managed locations.
            }
        }

        string local = GetBackupDirectory(BackupLocation.Local);
        string roaming = GetBackupDirectory(BackupLocation.Roaming);
        if (Directory.Exists(local))
        {
            return local;
        }

        if (Directory.Exists(roaming))
        {
            return roaming;
        }

        // Default target for new installs / first open: next to the EXE.
        return local;
    }

    private void TryOpenResultFolder(string folderPath, string logPrefix)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(folderPath))
            {
                return;
            }

            Directory.CreateDirectory(folderPath);
            WriteLog($"{logPrefix}: open=\"{SanitizeLogValue(folderPath)}\"");
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = $"\"{folderPath}\"",
                UseShellExecute = true,
            });
        }
        catch (Exception ex)
        {
            WriteLog($"{logPrefix}: failed to open folder: {FlattenLogText(ex.ToString())}");
        }
    }

    private Button NewOperationResultButton(string text, int height, int? width = null)
    {
        Button button = new()
        {
            Name = text,
            Text = text,
            Size = new Size(width ?? UiScale(100), height),
            FlatStyle = FlatStyle.Flat,
            Font = _buttonFont,
            Margin = Padding.Empty,
            UseVisualStyleBackColor = false,
            Cursor = Cursors.Hand,
        };
        SetTopButtonBaseStyle(button);
        button.MouseEnter += (_, _) => SetTopButtonHoverStyle(button);
        button.MouseLeave += (_, _) => SetTopButtonBaseStyle(button);
        return button;
    }

    private static string InferOperationName(string text)
    {
        string normalized = NormalizeDialogMessage(text).Trim();
        if (normalized.Length == 0)
        {
            return "DEVICE TWEAKER";
        }

        int end = normalized.IndexOfAny([' ', '\r', '\n']);
        return (end > 0 ? normalized[..end] : normalized).Trim().ToUpperInvariant();
    }

    private static string BuildOperationTechnicalDetails(OperationReport report)
    {
        List<string> sections = [];
        int index = 0;
        foreach (OperationIssue issue in report.Issues)
        {
            if (issue.Severity == OperationIssueSeverity.Success
                && string.IsNullOrWhiteSpace(issue.TechnicalDetails))
            {
                continue;
            }

            index++;
            string severity = UiLanguage.Text(issue.Severity.ToString().ToUpperInvariant());
            string title = $"{index}. [{severity}] {UiLanguage.Text(issue.Component)}";
            string body = issue.TechnicalDetails
                ?? UiLanguage.Text(issue.UserMessage)
                ?? UiLanguage.Text("No additional details.");
            sections.Add($"{title}{Environment.NewLine}{body}");
        }

        return string.Join(Environment.NewLine + Environment.NewLine, sections);
    }

    private bool ShowThemedConfirm(string message, string title, string yesText = "YES", string noText = "NO")
    {
        using Form dialog = new ThemedDialogForm();
        dialog.Name = "CONFIRM_DIALOG";
        dialog.Text = title;
        dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
        dialog.StartPosition = FormStartPosition.CenterParent;
        dialog.MaximizeBox = false;
        dialog.MinimizeBox = false;
        dialog.ShowInTaskbar = false;
        dialog.AutoScaleMode = AutoScaleMode.None;
        StyleThemedDialogSurface(dialog);
        dialog.Font = _dialogFont;
        dialog.Icon = Icon;

        string normalized = NormalizeDialogMessage(UiLanguage.Text(message));

        int padding = UiScale(20);
        int maxTextWidth = UiScale(520);
        int minWidth = UiScale(360);
        int buttonWidth = UiScale(120);
        int buttonHeight = UiScale(32);
        int buttonGap = UiScale(16);

        Label messageLabel = new()
        {
            AutoSize = true,
            MaximumSize = new Size(maxTextWidth, 0),
            Text = normalized,
            ForeColor = _fgMain,
            BackColor = _bgForm,
            UseMnemonic = false,
            TextAlign = ContentAlignment.MiddleCenter,
            Font = _dialogFont,
            UseCompatibleTextRendering = false,
        };

        Size textSize = messageLabel.GetPreferredSize(new Size(maxTextWidth, 0));
        int buttonRowWidth = (buttonWidth * 2) + buttonGap;
        int clientWidth = Math.Max(minWidth, Math.Max(textSize.Width + (padding * 2), buttonRowWidth + (padding * 2)));
        int labelWidth = clientWidth - (padding * 2);
        messageLabel.MaximumSize = new Size(labelWidth, 0);
        textSize = messageLabel.GetPreferredSize(new Size(labelWidth, 0));
        int clientHeight = padding + textSize.Height + buttonGap + buttonHeight + padding;
        dialog.ClientSize = new Size(clientWidth, clientHeight);

        messageLabel.AutoSize = false;
        messageLabel.Size = new Size(labelWidth, textSize.Height);
        messageLabel.Location = new Point(padding, padding);

        int buttonsTop = padding + textSize.Height + buttonGap;
        int rowLeft = (clientWidth - buttonRowWidth) / 2;

        Button yesButton = new()
        {
            Name = yesText,
            Text = yesText,
            DialogResult = DialogResult.Yes,
            Size = new Size(buttonWidth, buttonHeight),
            Location = new Point(rowLeft, buttonsTop),
            FlatStyle = FlatStyle.Flat,
            Font = _buttonFont,
            UseVisualStyleBackColor = false,
            Cursor = Cursors.Hand,
        };
        SetTopButtonBaseStyle(yesButton);
        yesButton.MouseEnter += (_, _) => SetTopButtonHoverStyle(yesButton);
        yesButton.MouseLeave += (_, _) => SetTopButtonBaseStyle(yesButton);

        Button noButton = new()
        {
            Name = noText,
            Text = noText,
            DialogResult = DialogResult.No,
            Size = new Size(buttonWidth, buttonHeight),
            Location = new Point(rowLeft + buttonWidth + buttonGap, buttonsTop),
            FlatStyle = FlatStyle.Flat,
            Font = _buttonFont,
            UseVisualStyleBackColor = false,
            Cursor = Cursors.Hand,
        };
        SetTopButtonBaseStyle(noButton);
        noButton.MouseEnter += (_, _) => SetTopButtonHoverStyle(noButton);
        noButton.MouseLeave += (_, _) => SetTopButtonBaseStyle(noButton);

        dialog.Controls.Add(messageLabel);
        dialog.Controls.Add(yesButton);
        dialog.Controls.Add(noButton);

        dialog.AcceptButton = yesButton;
        dialog.CancelButton = noButton;
        WireThemedTitleBar(dialog);

        return ShowDialogDimmed(dialog) == DialogResult.Yes;
    }

    private bool ShowThemedConfirm(string message)
    {
        return ShowThemedConfirm(message, "DEVICE TWEAKER");
    }

    private RestoreChoice ShowRestoreChoiceDialog(IReadOnlyList<BackupSnapshotInfo> backups, out string? selectedBackupPath)
    {
        selectedBackupPath = null;
        string? selectedPath = null;
        using Form dialog = new ThemedDialogForm();
        dialog.Name = "RESTORE_DIALOG";
        dialog.Text = "RESTORE";
        dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
        dialog.StartPosition = FormStartPosition.CenterParent;
        dialog.MaximizeBox = false;
        dialog.MinimizeBox = false;
        dialog.ShowInTaskbar = false;
        dialog.AutoScaleMode = AutoScaleMode.None;
        StyleThemedDialogSurface(dialog);
        dialog.Font = _dialogFont;
        dialog.Icon = Icon;

        bool hasBackup = backups.Count > 0;
        Button? deleteButton = null;
        string message = hasBackup
            ? "Choose how to restore DEVICE TWEAKER settings.\n\n"
                + "RESTORE LAST uses the newest snapshot. RESTORE SELECTED uses the highlighted snapshot.\n"
                + "RESET WINDOWS DEFAULT restores supported settings to Windows defaults.\n"
                + "ORIGINAL STATE is protected and never pruned."
            : "No backup snapshot was found.\n\n"
                + "RESET WINDOWS DEFAULT restores supported settings to Windows defaults.";
        string normalized = NormalizeDialogMessage(UiLanguage.Text(message));

        int padding = UiScale(20);
        int maxTextWidth = UiScale(hasBackup ? 700 : 500);
        int minWidth = UiScale(hasBackup ? 1020 : 540);
        int standardButtonWidth = UiScale(170);
        int selectedButtonWidth = UiScale(186);
        int resetButtonWidth = UiScale(220);
        int cancelButtonWidth = UiScale(150);
        int buttonHeight = UiScale(32);
        int buttonGap = UiScale(12);
        int backupListHeight = hasBackup ? UiScale(112) : 0;

        Label messageLabel = new()
        {
            AutoSize = true,
            MaximumSize = new Size(maxTextWidth, 0),
            Text = normalized,
            ForeColor = _fgMain,
            BackColor = _bgForm,
            UseMnemonic = false,
            TextAlign = ContentAlignment.MiddleCenter,
            Font = _dialogFont,
            UseCompatibleTextRendering = false,
        };

        Size textSize = messageLabel.GetPreferredSize(new Size(maxTextWidth, 0));
        int buttonRowWidth = hasBackup
            ? standardButtonWidth + selectedButtonWidth + resetButtonWidth + standardButtonWidth + cancelButtonWidth + (buttonGap * 4)
            : resetButtonWidth + cancelButtonWidth + buttonGap;
        int clientWidth = Math.Max(minWidth, Math.Max(textSize.Width + (padding * 2), buttonRowWidth + (padding * 2)));
        int labelWidth = clientWidth - (padding * 2);
        messageLabel.MaximumSize = new Size(labelWidth, 0);
        textSize = messageLabel.GetPreferredSize(new Size(labelWidth, 0));
        int clientHeight = padding + textSize.Height + (hasBackup ? buttonGap + backupListHeight : 0) + buttonGap + buttonHeight + padding;
        dialog.ClientSize = new Size(clientWidth, clientHeight);

        messageLabel.AutoSize = false;
        messageLabel.Size = new Size(labelWidth, textSize.Height);
        messageLabel.Location = new Point(padding, padding);

        ListBox? backupList = null;
        if (hasBackup)
        {
            backupList = new ListBox
            {
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.FromArgb(18, 18, 22),
                ForeColor = _fgMain,
                Font = _dialogFont,
                IntegralHeight = false,
                HorizontalScrollbar = true,
                Size = new Size(labelWidth, backupListHeight),
                Location = new Point(padding, messageLabel.Bottom + buttonGap),
            };

            foreach (BackupSnapshotInfo backup in backups)
            {
                backupList.Items.Add(backup);
            }

            backupList.SelectedIndex = 0;
            selectedPath = backups[0].Path;
            backupList.SelectedIndexChanged += (_, _) =>
            {
                int index = backupList.SelectedIndex;
                selectedPath = index >= 0 && index < backups.Count ? backups[index].Path : null;
                UpdateDeleteButtonState();
            };
        }

        RestoreChoice choice = RestoreChoice.Cancel;
        int buttonsTop = (backupList?.Bottom ?? messageLabel.Bottom) + buttonGap;
        int rowLeft = (clientWidth - buttonRowWidth) / 2;

        Button MakeButton(string text, RestoreChoice value, int left, int width, bool enabled = true)
        {
            Button button = new()
            {
                Name = text,
                Text = text,
                Size = new Size(width, buttonHeight),
                Location = new Point(left, buttonsTop),
                FlatStyle = FlatStyle.Flat,
                Font = _buttonFont,
                UseVisualStyleBackColor = false,
                Cursor = enabled ? Cursors.Hand : Cursors.Default,
                Enabled = enabled,
            };

            SetTopButtonBaseStyle(button);
            if (!enabled)
            {
                button.ForeColor = Color.FromArgb(120, 120, 125);
                button.FlatAppearance.BorderColor = Color.FromArgb(80, 80, 86);
            }
            else
            {
                button.MouseEnter += (_, _) => SetTopButtonHoverStyle(button);
                button.MouseLeave += (_, _) => SetTopButtonBaseStyle(button);
                button.Click += (_, _) =>
                {
                    choice = value;
                    dialog.DialogResult = DialogResult.OK;
                    dialog.Close();
                };
            }

            return button;
        }

        void UpdateDeleteButtonState()
        {
            if (deleteButton is null || backupList is null)
            {
                return;
            }

            int index = backupList.SelectedIndex;
            bool canDelete = index >= 0 && index < backups.Count && !backups[index].IsOriginal;
            deleteButton.Enabled = canDelete;
            deleteButton.Cursor = canDelete ? Cursors.Hand : Cursors.Default;
            if (canDelete)
            {
                SetTopButtonBaseStyle(deleteButton);
            }
            else
            {
                deleteButton.ForeColor = Color.FromArgb(120, 120, 125);
                deleteButton.FlatAppearance.BorderColor = Color.FromArgb(80, 80, 86);
            }
        }

        Button? latestButton = hasBackup
            ? MakeButton("RESTORE LAST", RestoreChoice.RestoreLatest, rowLeft, standardButtonWidth)
            : null;
        int selectedLeft = rowLeft + standardButtonWidth + buttonGap;
        Button? backupButton = hasBackup
            ? MakeButton("RESTORE SELECTED", RestoreChoice.RestoreBackup, selectedLeft, selectedButtonWidth)
            : null;
        int safeResetLeft = hasBackup
            ? selectedLeft + selectedButtonWidth + buttonGap
            : rowLeft;
        Button resetButton = MakeButton("RESET WINDOWS DEFAULT", RestoreChoice.SafeReset, safeResetLeft, resetButtonWidth);
        int deleteLeft = safeResetLeft + resetButtonWidth + buttonGap;
        deleteButton = hasBackup
            ? MakeButton("DELETE SELECTED", RestoreChoice.DeleteBackup, deleteLeft, standardButtonWidth)
            : null;
        UpdateDeleteButtonState();
        int cancelLeft = hasBackup
            ? deleteLeft + standardButtonWidth + buttonGap
            : safeResetLeft + resetButtonWidth + buttonGap;
        Button cancelButton = MakeButton("CANCEL", RestoreChoice.Cancel, cancelLeft, cancelButtonWidth);

        dialog.Controls.Add(messageLabel);
        if (backupList is not null)
        {
            dialog.Controls.Add(backupList);
        }
        if (latestButton is not null)
        {
            dialog.Controls.Add(latestButton);
        }
        if (backupButton is not null)
        {
            dialog.Controls.Add(backupButton);
        }
        dialog.Controls.Add(resetButton);
        if (deleteButton is not null)
        {
            dialog.Controls.Add(deleteButton);
        }
        dialog.Controls.Add(cancelButton);

        dialog.AcceptButton = latestButton ?? resetButton;
        dialog.CancelButton = cancelButton;
        WireThemedTitleBar(dialog);

        RestoreChoice result = ShowDialogDimmed(dialog) == DialogResult.OK ? choice : RestoreChoice.Cancel;
        selectedBackupPath = selectedPath;
        return result;
    }

    private AutoBackupChoice ShowAutoBackupChoiceDialog()
    {
        using Form dialog = new ThemedDialogForm();
        dialog.Name = "AUTO_BACKUP_DIALOG";
        dialog.Text = "AUTO BACKUP";
        dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
        dialog.StartPosition = FormStartPosition.CenterParent;
        dialog.MaximizeBox = false;
        dialog.MinimizeBox = false;
        dialog.ShowInTaskbar = false;
        dialog.AutoScaleMode = AutoScaleMode.None;
        StyleThemedDialogSurface(dialog);
        dialog.Font = _dialogFont;
        dialog.Icon = Icon;

        string message = NormalizeDialogMessage(UiLanguage.Text(
            "Where should DEVICE TWEAKER save the pre-auto backup?\n\n"
            + "EXE FOLDER = portable backup next to the app.\n"
            + "APPDATA = user profile backup that survives app folder cleanup.\n"
            + "SKIP = do not create an additional pre-auto backup.\n"
            + "The protected ORIGINAL STATE snapshot is always retained.\n"
            + "Close (X) = cancel AUTO-OPTIMIZATION."));

        int padding = UiScale(20);
        int maxTextWidth = UiScale(620);
        int buttonWidth = UiScale(150);
        int buttonHeight = UiScale(32);
        int buttonGap = UiScale(12);

        Label messageLabel = new()
        {
            AutoSize = true,
            MaximumSize = new Size(maxTextWidth, 0),
            Text = message,
            ForeColor = _fgMain,
            BackColor = _bgForm,
            UseMnemonic = false,
            TextAlign = ContentAlignment.MiddleCenter,
            Font = _dialogFont,
            UseCompatibleTextRendering = false,
        };

        Size textSize = messageLabel.GetPreferredSize(new Size(maxTextWidth, 0));
        int buttonRowWidth = (buttonWidth * 3) + (buttonGap * 2);
        int clientWidth = Math.Max(textSize.Width + (padding * 2), buttonRowWidth + (padding * 2));
        int labelWidth = clientWidth - (padding * 2);
        messageLabel.AutoSize = false;
        messageLabel.Size = new Size(labelWidth, textSize.Height);
        messageLabel.Location = new Point(padding, padding);

        int buttonsTop = messageLabel.Bottom + UiScale(18);
        int rowLeft = (clientWidth - buttonRowWidth) / 2;
        int clientHeight = buttonsTop + buttonHeight + padding;
        dialog.ClientSize = new Size(clientWidth, clientHeight);

        AutoBackupChoice choice = AutoBackupChoice.Skip;
        Button MakeButton(string text, AutoBackupChoice value, int left)
        {
            Button button = new()
            {
                Name = text,
                Text = text,
                Size = new Size(buttonWidth, buttonHeight),
                Location = new Point(left, buttonsTop),
                FlatStyle = FlatStyle.Flat,
                Font = _buttonFont,
                UseVisualStyleBackColor = false,
                Cursor = Cursors.Hand,
            };
            SetTopButtonBaseStyle(button);
            button.MouseEnter += (_, _) => SetTopButtonHoverStyle(button);
            button.MouseLeave += (_, _) => SetTopButtonBaseStyle(button);
            button.Click += (_, _) =>
            {
                choice = value;
                dialog.DialogResult = DialogResult.OK;
                dialog.Close();
            };
            return button;
        }

        Button exeButton = MakeButton("EXE FOLDER", AutoBackupChoice.Local, rowLeft);
        Button appDataButton = MakeButton("APPDATA", AutoBackupChoice.Roaming, rowLeft + buttonWidth + buttonGap);
        Button skipButton = MakeButton("SKIP", AutoBackupChoice.Skip, rowLeft + ((buttonWidth + buttonGap) * 2));

        dialog.Controls.Add(messageLabel);
        dialog.Controls.Add(exeButton);
        dialog.Controls.Add(appDataButton);
        dialog.Controls.Add(skipButton);
        dialog.AcceptButton = exeButton;
        // Do not bind CancelButton to SKIP — Esc/X must cancel AUTO, not skip backup.
        dialog.CancelButton = null;
        WireThemedTitleBar(dialog);

        DialogResult result = ShowDialogDimmed(dialog);
        return result == DialogResult.OK ? choice : AutoBackupChoice.Cancel;
    }

    private static string NormalizeDialogMessage(string message)
    {
        if (string.IsNullOrEmpty(message))
        {
            return string.Empty;
        }

        string normalized = message.Replace("\r\n", "\n").Replace("\r", "\n");
        return normalized.Replace("\n", Environment.NewLine);
    }

    private void WireThemedTitleBar(Form form)
    {
        form.HandleCreated += (_, _) => ApplyTitleBarTheme(form);
        form.Shown += (_, _) => ApplyTitleBarTheme(form);
    }

    private void ApplyTitleBarTheme(Form form)
    {
        try
        {
            if (!OperatingSystem.IsWindowsVersionAtLeast(10, 0, 17763))
            {
                return;
            }

            int bgResult = 0;
            int fgResult = 0;
            if (OperatingSystem.IsWindowsVersionAtLeast(10, 0, 22000))
            {
                int bg = ColorToColorRef(_bgForm);
                bgResult = DwmSetWindowAttribute(form.Handle, DwmwaCaptionColor, ref bg, sizeof(int));

                int fg = ColorToColorRef(_fgMain);
                fgResult = DwmSetWindowAttribute(form.Handle, DwmwaTextColor, ref fg, sizeof(int));
            }

            int darkMode = 1;
            int darkResult = DwmSetWindowAttribute(form.Handle, DwmwaUseImmersiveDarkMode, ref darkMode, sizeof(int));
            if (bgResult != 0 || fgResult != 0 || darkResult != 0)
            {
                WriteLog(
                    $"UI.TITLEBAR.WARN: form=\"{form.Text}\" captionResult=0x{bgResult:X8} " +
                    $"textResult=0x{fgResult:X8} darkResult=0x{darkResult:X8}");
            }
        }
        catch (Exception ex)
        {
            WriteLog($"UI.TITLEBAR.ERROR: form=\"{form.Text}\" error=\"{FlattenLogText(ex.ToString())}\"");
        }
    }
}
