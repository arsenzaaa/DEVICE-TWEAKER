
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

    private Panel AttachThemedDialogFooter(Form dialog, IReadOnlyList<Button> buttons, int buttonGap = 10)
    {
        int footerHeight = UiScale(50);
        Panel footerPanel = new()
        {
            Dock = DockStyle.Bottom,
            Height = footerHeight,
            BackColor = _bgForm,
        };

        void LayoutButtons()
        {
            int totalWidth = buttons.Sum(b => b.Width) + Math.Max(0, buttons.Count - 1) * buttonGap;
            int startX = Math.Max(UiScale(14), (footerPanel.ClientSize.Width - totalWidth) / 2);
            int currentX = startX;
            int btnY = (footerHeight - (buttons.Count > 0 ? buttons[0].Height : UiScale(32))) / 2;

            foreach (Button btn in buttons)
            {
                btn.Location = new Point(currentX, btnY);
                currentX += btn.Width + buttonGap;
            }
        }

        foreach (Button btn in buttons)
        {
            footerPanel.Controls.Add(btn);
        }

        LayoutButtons();
        footerPanel.Resize += (_, _) => LayoutButtons();
        dialog.Controls.Add(footerPanel);

        return footerPanel;
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
        dialog.BackColor = _bgForm;
        WireThemedTitleBar(dialog);

        string normalized = NormalizeDialogMessage(UiLanguage.Text(message));

        int padding = UiScale(20);
        int maxTextWidth = UiScale(580);
        int minWidth = UiScale(460);
        int buttonWidth = UiScale(96);
        int buttonHeight = UiScale(32);
        int footerHeight = UiScale(50);

        bool hasMultipleLines = normalized.Contains('\n');
        Label messageLabel = new()
        {
            AutoSize = true,
            MaximumSize = new Size(maxTextWidth, 0),
            Text = normalized,
            ForeColor = _fgMain,
            BackColor = _bgForm,
            UseMnemonic = false,
            TextAlign = hasMultipleLines ? ContentAlignment.TopLeft : ContentAlignment.MiddleCenter,
            Font = _dialogFont,
        };

        Size textSize = messageLabel.GetPreferredSize(new Size(maxTextWidth, 0));
        int clientWidth = Math.Max(minWidth, textSize.Width + (padding * 2));
        int innerWidth = clientWidth - (padding * 2);

        messageLabel.MaximumSize = new Size(innerWidth, 0);
        textSize = messageLabel.GetPreferredSize(new Size(innerWidth, 0));
        messageLabel.AutoSize = false;
        messageLabel.Size = new Size(innerWidth, textSize.Height);
        messageLabel.Location = new Point(padding, padding);
        dialog.Controls.Add(messageLabel);

        Button okButton = NewDialogButton(UiLanguage.Text("OK"), buttonHeight, buttonWidth, isPrimary: true);
        okButton.Name = "OK";
        okButton.DialogResult = DialogResult.OK;

        AttachThemedDialogFooter(dialog, [okButton]);

        int clientHeight = padding + textSize.Height + padding + footerHeight;
        dialog.ClientSize = new Size(clientWidth, clientHeight);
        dialog.AcceptButton = okButton;
        dialog.CancelButton = okButton;

        ShowDialogDimmed(dialog);
    }

    private void ShowThemedInfo(string message)
    {
        ShowThemedInfo(message, "DEVICE TWEAKER");
    }

    private sealed class IssueItemToken
    {
        public required string Text { get; init; }
        public required Font Font { get; init; }
        public required Color Color { get; init; }
        public Rectangle Bounds { get; set; }
    }

    private sealed class IssueItemView : Control
    {
        private readonly List<IssueItemToken> _tokens = [];

        public IssueItemView(
            string component,
            string detail,
            string stateTag,
            Color tagColor,
            Font textFont,
            Font boldFont,
            Color backColor,
            int maxWidth,
            int indent)
        {
            SetStyle(
                ControlStyles.UserPaint
                | ControlStyles.AllPaintingInWmPaint
                | ControlStyles.OptimizedDoubleBuffer
                | ControlStyles.ResizeRedraw,
                true);
            BackColor = backColor;
            TabStop = false;

            // 1. Dash marker
            _tokens.Add(new IssueItemToken
            {
                Text = "— ",
                Font = boldFont,
                Color = Color.FromArgb(150, 155, 170),
            });

            // 2. Component name
            string compText = string.IsNullOrWhiteSpace(detail)
                ? component + " "
                : component + ": ";
            _tokens.Add(new IssueItemToken
            {
                Text = compText,
                Font = boldFont,
                Color = Color.FromArgb(130, 195, 245),
            });

            // 3. Detail words (if any)
            if (!string.IsNullOrWhiteSpace(detail))
            {
                string[] words = detail.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                for (int w = 0; w < words.Length; w++)
                {
                    _tokens.Add(new IssueItemToken
                    {
                        Text = words[w] + " ",
                        Font = textFont,
                        Color = Color.FromArgb(225, 228, 235),
                    });
                }
            }

            // 4. State tag with non-breaking space inside parentheses
            string safeTag = $"({stateTag.Replace(' ', '\u00A0')})";
            _tokens.Add(new IssueItemToken
            {
                Text = safeTag,
                Font = boldFont,
                Color = tagColor,
            });

            // Measure and layout with hanging indent
            int currentX = 0;
            int currentY = 0;
            int lineHeight = Math.Max(boldFont.Height, textFont.Height) + 2;

            for (int i = 0; i < _tokens.Count; i++)
            {
                IssueItemToken token = _tokens[i];
                Size sz = TextRenderer.MeasureText(token.Text, token.Font, Size.Empty, TextFormatFlags.NoPadding);

                if (currentX > indent && currentX + sz.Width > maxWidth)
                {
                    currentX = indent;
                    currentY += lineHeight;
                }

                token.Bounds = new Rectangle(currentX, currentY, sz.Width, sz.Height);
                currentX += sz.Width;
            }

            Size = new Size(maxWidth, currentY + lineHeight);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
            using SolidBrush bg = new(BackColor);
            e.Graphics.FillRectangle(bg, ClientRectangle);

            foreach (IssueItemToken token in _tokens)
            {
                TextRenderer.DrawText(
                    e.Graphics,
                    token.Text,
                    token.Font,
                    token.Bounds.Location,
                    token.Color,
                    TextFormatFlags.NoPadding);
            }
        }
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
        bool isSuccess = report.Succeeded && report.Warnings.Count == 0;
        bool isWarnings = report.Succeeded && report.Warnings.Count > 0;
        bool isPartial = !report.Succeeded && !failedBeforeChanges;

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
            BackColor = _bgForm,
            ForeColor = _fgMain,
        };
        StyleThemedDialogSurface(dialog);
        WireThemedTitleBar(dialog);

        List<OperationIssue> actionableIssues = report.Issues
            .Where(issue => issue.Severity != OperationIssueSeverity.Success)
            .ToList();
        IReadOnlyList<OperationIssue> visibleIssues = actionableIssues.Take(3).ToList();
        string technicalText = BuildOperationTechnicalDetails(report);
        bool hasTechnical = technicalText.Length > 0;
        string backupFolder = ResolveResultBackupFolder(report.BackupPath);
        string logsFolder = AppDiagnostics.LogDirectory;

        int padding = UiScale(20);
        int sectionGap = UiScale(12);
        int buttonHeight = UiScale(32);
        int footerHeight = UiScale(50);
        int detailsHeight = UiScale(220);
        Rectangle workingArea = Screen.FromControl(this).WorkingArea;

        List<Button> visibleButtons = [];
        Button okButton = NewDialogButton(UiLanguage.Text("OK"), buttonHeight, UiScale(96), isPrimary: true);
        okButton.Name = "OK";
        okButton.DialogResult = DialogResult.OK;
        visibleButtons.Add(okButton);

        Button? detailsButton = null;
        if (hasTechnical)
        {
            detailsButton = NewDialogButton(UiLanguage.Text("DETAILS"), buttonHeight, UiScale(116), isPrimary: false);
            detailsButton.Name = "DETAILS";
            visibleButtons.Add(detailsButton);
        }

        Button logsButton = NewDialogButton(UiLanguage.Text("LOGS"), buttonHeight, UiScale(96), isPrimary: false);
        logsButton.Name = "LOGS";
        logsButton.AccessibleDescription = logsFolder;
        visibleButtons.Add(logsButton);

        Button backupsButton = NewDialogButton(UiLanguage.Text("BACKUPS"), buttonHeight, UiScale(104), isPrimary: false);
        backupsButton.Name = "BACKUPS";
        backupsButton.AccessibleDescription = backupFolder;
        visibleButtons.Add(backupsButton);

        int clientWidth = isSuccess ? UiScale(480) : UiScale(680);
        int innerWidth = clientWidth - (padding * 2);
        int currentY = padding;

        if (isSuccess)
        {
            string msg = NormalizeDialogMessage(
                string.IsNullOrWhiteSpace(successMessage)
                    ? (UiLanguage.IsRussian
                        ? "Оптимизация успешно завершена и сохранена.\nПожалуйста, перезагрузите ПК для применения изменений."
                        : "Auto-optimization completed and saved.\nPlease reboot your PC to finish applying the changes.")
                    : UiLanguage.Text(successMessage));

            Label msgLabel = new()
            {
                AutoSize = true,
                MaximumSize = new Size(innerWidth, 0),
                Text = msg,
                Font = _dialogFont,
                ForeColor = _fgMain,
                BackColor = _bgForm,
                TextAlign = ContentAlignment.MiddleCenter,
                UseMnemonic = false,
                Location = new Point(padding, currentY),
            };
            Size msgSz = msgLabel.GetPreferredSize(new Size(innerWidth, 0));
            msgLabel.AutoSize = false;
            msgLabel.Size = new Size(innerWidth, msgSz.Height);
            dialog.Controls.Add(msgLabel);
            currentY = msgLabel.Bottom + UiScale(12);
        }
        else
        {
            string statusHeader = isWarnings
                ? (UiLanguage.IsRussian ? "Применено с замечаниями" : "Applied with warnings")
                : isPartial
                    ? (UiLanguage.IsRussian ? "Применено частично" : "Partially applied")
                    : (UiLanguage.IsRussian ? "Операция не выполнена" : "Operation not applied");

            Color headerColor = isWarnings || isPartial
                ? Color.FromArgb(235, 180, 75)
                : (report.NoChangesMade ? Color.FromArgb(245, 95, 80) : Color.FromArgb(95, 205, 135));

            Label statusLabel = new()
            {
                AutoSize = true,
                MaximumSize = new Size(innerWidth, 0),
                Text = statusHeader,
                Font = _blockTitleFont,
                ForeColor = headerColor,
                BackColor = _bgForm,
                TextAlign = ContentAlignment.TopLeft,
                UseMnemonic = false,
                Location = new Point(padding, currentY),
            };
            Size sSz = statusLabel.GetPreferredSize(new Size(innerWidth, 0));
            statusLabel.AutoSize = false;
            statusLabel.Size = new Size(innerWidth, sSz.Height);
            dialog.Controls.Add(statusLabel);
            currentY = statusLabel.Bottom + UiScale(4);

            string leadText;
            if (report.NoChangesMade)
            {
                leadText = !string.IsNullOrWhiteSpace(partialMessage)
                    ? NormalizeDialogMessage(UiLanguage.Text(partialMessage))
                    : (UiLanguage.IsRussian ? "Действия были остановлены до внесения изменений." : "The operation stopped before any changes were made.");
            }
            else
            {
                leadText = NormalizeDialogMessage(UiLanguage.Text(report.Succeeded ? successMessage : partialMessage));
                if (string.IsNullOrWhiteSpace(leadText) || leadText.Equals("DEVICE TWEAKER", StringComparison.OrdinalIgnoreCase))
                {
                    leadText = UiLanguage.IsRussian ? "Основные параметры были сохранены." : "Core settings were successfully configured.";
                }
            }

            Label leadLabel = new()
            {
                AutoSize = true,
                MaximumSize = new Size(innerWidth, 0),
                Text = leadText,
                Font = _dialogFont,
                ForeColor = Color.FromArgb(170, 175, 188),
                BackColor = _bgForm,
                TextAlign = ContentAlignment.TopLeft,
                UseMnemonic = false,
                Location = new Point(padding, currentY),
            };
            Size leadSize = leadLabel.GetPreferredSize(new Size(innerWidth, 0));
            leadLabel.AutoSize = false;
            leadLabel.Size = new Size(innerWidth, leadSize.Height);
            dialog.Controls.Add(leadLabel);
            currentY = leadLabel.Bottom + UiScale(10);

            for (int i = 0; i < visibleIssues.Count; i++)
            {
                OperationIssue issue = visibleIssues[i];
                bool isOptionalSkip = !report.NoChangesMade && (
                    issue.Severity == OperationIssueSeverity.Warning ||
                    issue.Component.Contains("IMOD", StringComparison.OrdinalIgnoreCase) ||
                    (issue.UserMessage?.Contains("недоступ", StringComparison.OrdinalIgnoreCase) ?? false) ||
                    (issue.UserMessage?.Contains("unavail", StringComparison.OrdinalIgnoreCase) ?? false) ||
                    (issue.TechnicalDetails?.Contains("blocklist", StringComparison.OrdinalIgnoreCase) ?? false)
                );

                string stateTag = isOptionalSkip
                    ? (UiLanguage.IsRussian ? "пропущено" : "Skipped")
                    : issue.Severity switch
                    {
                        OperationIssueSeverity.Warning => (UiLanguage.IsRussian ? "замечание" : "Warning"),
                        _ => (UiLanguage.IsRussian ? "ошибка" : "Not applied"),
                    };

                string issueDetail = !string.IsNullOrWhiteSpace(issue.UserMessage)
                    ? NormalizeDialogMessage(UiLanguage.Text(issue.UserMessage))
                    : string.Empty;

                Color tagColor = isOptionalSkip
                    ? Color.FromArgb(160, 165, 178)
                    : issue.Severity switch
                    {
                        OperationIssueSeverity.Warning => Color.FromArgb(240, 205, 125),
                        _ => Color.FromArgb(250, 95, 80),
                    };

                IssueItemView itemView = new(
                    UiLanguage.Text(issue.Component),
                    issueDetail,
                    stateTag,
                    tagColor,
                    _dialogFont,
                    _blockTitleFont,
                    _bgForm,
                    innerWidth - UiScale(8),
                    UiScale(18));

                itemView.Location = new Point(padding + UiScale(4), currentY);
                dialog.Controls.Add(itemView);
                currentY = itemView.Bottom + UiScale(8);
            }

            if (actionableIssues.Count > visibleIssues.Count)
            {
                Label moreLabel = new()
                {
                    AutoSize = true,
                    MaximumSize = new Size(innerWidth, 0),
                    Text = UiLanguage.IsRussian
                        ? $"+ Ещё: {actionableIssues.Count - visibleIssues.Count} — нажмите «Подробности» для просмотра"
                        : $"+ {actionableIssues.Count - visibleIssues.Count} more — open details to view all",
                    Font = _baseFont,
                    ForeColor = _statusInactive,
                    BackColor = _bgForm,
                    TextAlign = ContentAlignment.TopLeft,
                    UseMnemonic = false,
                    Location = new Point(padding + UiScale(4), currentY + UiScale(2)),
                };
                Size moreSz = moreLabel.GetPreferredSize(new Size(innerWidth, 0));
                moreLabel.AutoSize = false;
                moreLabel.Size = new Size(innerWidth, moreSz.Height);
                dialog.Controls.Add(moreLabel);
                currentY = moreLabel.Bottom + UiScale(4);
            }
        }

        // Details host
        DeviceCardPanel detailsHost = new()
        {
            BorderColor = Color.FromArgb(38, 38, 48),
            BackColor = Color.FromArgb(7, 7, 10),
            Visible = false,
            Location = new Point(padding, currentY + sectionGap),
            Size = new Size(innerWidth, detailsHeight),
        };

        Panel detailsHeaderBar = new()
        {
            Location = new Point(1, 1),
            Size = new Size(innerWidth - 2, UiScale(26)),
            BackColor = Color.FromArgb(14, 14, 20),
        };
        detailsHeaderBar.Paint += (_, pe) =>
        {
            using Pen p = new(Color.FromArgb(32, 32, 42), 1);
            pe.Graphics.DrawLine(p, 0, detailsHeaderBar.Height - 1, detailsHeaderBar.Width, detailsHeaderBar.Height - 1);
        };
        Label detailsHeader = new()
        {
            AutoSize = false,
            Text = UiLanguage.Text("DIAGNOSTICS & SYSTEM LOG"),
            Font = _blockTitleFont,
            ForeColor = Color.FromArgb(140, 145, 160),
            BackColor = Color.Transparent,
            Location = new Point(UiScale(10), 0),
            Size = new Size(detailsHeaderBar.Width - UiScale(20), detailsHeaderBar.Height),
            TextAlign = ContentAlignment.MiddleLeft,
            UseMnemonic = false,
        };
        detailsHeaderBar.Controls.Add(detailsHeader);
        detailsHost.Controls.Add(detailsHeaderBar);

        int detailsHeaderH = detailsHeaderBar.Height;
        Panel detailsContent = new()
        {
            BackColor = Color.FromArgb(7, 7, 10),
            Location = new Point(1, detailsHeaderH + 1),
        };
        Label detailsLabel = new()
        {
            AutoSize = true,
            Text = technicalText,
            Font = _technicalFont,
            ForeColor = Color.FromArgb(180, 185, 200),
            BackColor = Color.FromArgb(7, 7, 10),
            UseMnemonic = false,
            Location = new Point(UiScale(10), UiScale(6)),
        };
        detailsContent.Controls.Add(detailsLabel);
        ThemedScrollBar detailsScroll = new()
        {
            Width = UiScale(14),
            Dock = DockStyle.Right,
            BackColor = Color.FromArgb(7, 7, 10),
            TrackColor = Color.FromArgb(7, 7, 10),
            RailColor = Color.FromArgb(7, 7, 10),
            ThumbColor = Color.FromArgb(48, 48, 56),
            ThumbHoverColor = Color.FromArgb(80, 80, 92),
            ThumbDragColor = Color.FromArgb(120, 120, 135),
            ThumbWidth = UiScale(8),
            RailWidth = 0,
            ThumbCornerRadius = UiScale(6),
            Visible = false,
        };
        detailsHost.Controls.Add(detailsContent);
        detailsHost.Controls.Add(detailsScroll);
        dialog.Controls.Add(detailsHost);

        bool syncingDetailsScroll = false;
        void SyncDetailsLayout()
        {
            int contentW = Math.Max(1, detailsHost.ClientSize.Width - detailsScroll.Width - UiScale(24));
            detailsLabel.MaximumSize = new Size(contentW, 0);
            Size preferred = detailsLabel.GetPreferredSize(new Size(contentW, 0));
            detailsLabel.Size = preferred;
            detailsContent.Width = detailsHost.ClientSize.Width - detailsScroll.Width - 2;
            detailsContent.Height = preferred.Height + UiScale(16);

            int viewH = Math.Max(1, detailsHost.ClientSize.Height - detailsHeaderH - 2);
            int maxOffset = Math.Max(0, detailsContent.Height - viewH);
            int offset = Math.Max(0, Math.Min(maxOffset, -(detailsContent.Top - detailsHeaderH - 1)));
            detailsContent.Location = new Point(1, detailsHeaderH + 1 - offset);
            detailsScroll.Visible = maxOffset > 0;
            syncingDetailsScroll = true;
            detailsScroll.Maximum = Math.Max(detailsContent.Height, 1);
            detailsScroll.ViewportSize = viewH;
            detailsScroll.Value = offset;
            syncingDetailsScroll = false;
        }
        detailsScroll.ValueChanged += (_, _) =>
        {
            if (!syncingDetailsScroll)
            {
                detailsContent.Top = detailsHeaderH + 1 - detailsScroll.Value;
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

        AttachThemedDialogFooter(dialog, visibleButtons, UiScale(10));

        int compactHeight = currentY + UiScale(12) + footerHeight;
        dialog.ClientSize = new Size(clientWidth, Math.Min(compactHeight, workingArea.Height - UiScale(60)));
        int baseClientHeight = dialog.ClientSize.Height;

        void RelayoutExpanded(bool expanded)
        {
            detailsHost.Visible = expanded;
            int requested = expanded ? baseClientHeight + detailsHeight + sectionGap : baseClientHeight;
            dialog.ClientSize = new Size(clientWidth, Math.Min(requested, workingArea.Height - UiScale(60)));
            if (expanded)
            {
                SyncDetailsLayout();
                detailsHost.Focus();
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
                detailsButton.Width = expand ? UiScale(150) : UiScale(116);
                RelayoutExpanded(expand);
            };
        }

        dialog.AcceptButton = okButton;
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

    private Button NewDialogButton(string text, int height, int? width = null, bool isPrimary = false)
    {
        Button button = new ThemedButton
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
        StyleDialogButton(button, isPrimary);
        return button;
    }

    private Button NewOperationResultButton(string text, int height, int? width = null)
        => NewDialogButton(text, height, width, isPrimary: false);

    private void StyleDialogButton(Button button, bool isPrimary = false)
    {
        button.FlatStyle = FlatStyle.Flat;
        button.FlatAppearance.BorderSize = 1;
        button.Font = _buttonFont;
        button.Cursor = Cursors.Hand;
        button.UseVisualStyleBackColor = false;

        Color baseBg = isPrimary ? Color.FromArgb(18, 18, 22) : Color.FromArgb(12, 12, 16);
        Color baseFg = isPrimary ? Color.FromArgb(245, 245, 245) : Color.FromArgb(180, 180, 190);
        Color baseBorder = isPrimary ? Color.FromArgb(215, 215, 222) : Color.FromArgb(75, 75, 84);

        Color hoverBg = isPrimary ? Color.FromArgb(34, 34, 40) : Color.FromArgb(24, 24, 30);
        Color hoverFg = Color.White;
        Color hoverBorder = isPrimary ? Color.FromArgb(255, 255, 255) : Color.FromArgb(135, 135, 145);

        Color disabledBg = Color.FromArgb(10, 10, 13);
        Color disabledFg = Color.FromArgb(85, 85, 92);
        Color disabledBorder = Color.FromArgb(36, 36, 42);

        void ApplyState(bool isHover = false)
        {
            if (!button.Enabled)
            {
                button.BackColor = disabledBg;
                button.ForeColor = disabledFg;
                button.FlatAppearance.BorderColor = disabledBorder;
                button.Cursor = Cursors.Default;
                return;
            }

            button.Cursor = Cursors.Hand;
            if (isHover)
            {
                button.BackColor = hoverBg;
                button.ForeColor = hoverFg;
                button.FlatAppearance.BorderColor = hoverBorder;
            }
            else
            {
                button.BackColor = baseBg;
                button.ForeColor = baseFg;
                button.FlatAppearance.BorderColor = baseBorder;
            }
        }

        ApplyState();
        button.MouseEnter += (_, _) => ApplyState(isHover: true);
        button.MouseLeave += (_, _) => ApplyState(isHover: false);
        button.EnabledChanged += (_, _) => ApplyState(isHover: false);
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
            string? userMsg = !string.IsNullOrWhiteSpace(issue.UserMessage)
                ? UiLanguage.Text(issue.UserMessage)
                : null;
            string? techDetails = !string.IsNullOrWhiteSpace(issue.TechnicalDetails)
                ? issue.TechnicalDetails
                : null;

            string body;
            if (userMsg != null && techDetails != null)
            {
                body = techDetails.Contains(userMsg, StringComparison.OrdinalIgnoreCase)
                    ? techDetails
                    : $"{userMsg}{Environment.NewLine}{Environment.NewLine}{techDetails}";
            }
            else
            {
                body = techDetails ?? userMsg ?? UiLanguage.Text("No additional details.");
            }
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
        dialog.BackColor = _bgForm;
        WireThemedTitleBar(dialog);

        string normalized = NormalizeDialogMessage(UiLanguage.Text(message));

        int padding = UiScale(20);
        int clientWidth = UiScale(640);
        int innerWidth = clientWidth - (padding * 2);
        int buttonWidth = UiScale(110);
        int buttonHeight = UiScale(32);
        int footerHeight = UiScale(50);
        int currentY = padding;

        string[] paragraphs = normalized
            .Replace("\r\n", "\n")
            .Split(["\n\n"], StringSplitOptions.RemoveEmptyEntries)
            .Select(p => p.Trim())
            .Where(p => p.Length > 0)
            .ToArray();

        if (paragraphs.Length <= 1)
        {
            string singleText = paragraphs.Length == 1 ? paragraphs[0] : normalized;
            bool hasMultipleLines = singleText.Contains('\n');
            Label msgLabel = new()
            {
                AutoSize = true,
                MaximumSize = new Size(innerWidth, 0),
                Text = singleText,
                ForeColor = _fgMain,
                BackColor = _bgForm,
                UseMnemonic = false,
                TextAlign = hasMultipleLines ? ContentAlignment.TopLeft : ContentAlignment.MiddleCenter,
                Font = _dialogFont,
            };

            Size textSize = msgLabel.GetPreferredSize(new Size(innerWidth, 0));
            int actualWidth = Math.Max(UiScale(520), textSize.Width + (padding * 2));
            int actualInner = actualWidth - (padding * 2);

            msgLabel.MaximumSize = new Size(actualInner, 0);
            textSize = msgLabel.GetPreferredSize(new Size(actualInner, 0));
            msgLabel.AutoSize = false;
            msgLabel.Size = new Size(actualInner, textSize.Height);
            msgLabel.Location = new Point(padding, currentY);
            dialog.Controls.Add(msgLabel);
            currentY = msgLabel.Bottom + UiScale(14);
            clientWidth = actualWidth;
        }
        else if (paragraphs.Length == 2 && (paragraphs[0].EndsWith("?") || paragraphs[0].EndsWith("?\n")))
        {
            // Case: Question on top, details below (e.g. Reset Adapter Settings)
            Label qLabel = new()
            {
                AutoSize = true,
                MaximumSize = new Size(innerWidth, 0),
                Text = paragraphs[0],
                Font = _blockTitleFont,
                ForeColor = _fgMain,
                BackColor = _bgForm,
                TextAlign = ContentAlignment.TopLeft,
                UseMnemonic = false,
                Location = new Point(padding, currentY),
            };
            Size qSz = qLabel.GetPreferredSize(new Size(innerWidth, 0));
            qLabel.AutoSize = false;
            qLabel.Size = new Size(innerWidth, qSz.Height);
            dialog.Controls.Add(qLabel);
            currentY = qLabel.Bottom + UiScale(10);

            Label detailLabel = new()
            {
                AutoSize = true,
                MaximumSize = new Size(innerWidth, 0),
                Text = paragraphs[1],
                Font = _dialogFont,
                ForeColor = Color.FromArgb(170, 175, 188),
                BackColor = _bgForm,
                TextAlign = ContentAlignment.TopLeft,
                UseMnemonic = false,
                Location = new Point(padding, currentY),
            };
            Size dSz = detailLabel.GetPreferredSize(new Size(innerWidth, 0));
            detailLabel.AutoSize = false;
            detailLabel.Size = new Size(innerWidth, dSz.Height);
            dialog.Controls.Add(detailLabel);
            currentY = detailLabel.Bottom + UiScale(16);
        }
        else
        {
            // Multi-part structured message (Lead Context, Advisory/Notice Card, Action Prompt)
            // 1. Lead paragraph
            string leadText = paragraphs[0];
            Label leadLabel = new()
            {
                AutoSize = true,
                MaximumSize = new Size(innerWidth, 0),
                Text = leadText,
                Font = _dialogFont,
                ForeColor = _fgMain,
                BackColor = _bgForm,
                TextAlign = ContentAlignment.TopLeft,
                UseMnemonic = false,
                Location = new Point(padding, currentY),
            };
            Size leadSz = leadLabel.GetPreferredSize(new Size(innerWidth, 0));
            leadLabel.AutoSize = false;
            leadLabel.Size = new Size(innerWidth, leadSz.Height);
            dialog.Controls.Add(leadLabel);
            currentY = leadLabel.Bottom + UiScale(12);

            // 2. Middle paragraph(s): Advisory / Warning card
            for (int i = 1; i < paragraphs.Length - 1; i++)
            {
                string noticeText = paragraphs[i];
                int cardPad = UiScale(12);
                int cardInnerW = innerWidth - (cardPad * 2);

                DeviceCardPanel card = new()
                {
                    Location = new Point(padding, currentY),
                    BackColor = Color.FromArgb(16, 18, 25),
                    BorderColor = Color.FromArgb(44, 48, 62),
                };

                int cardY = cardPad;
                Label noticeTag = new()
                {
                    AutoSize = true,
                    Text = UiLanguage.IsRussian ? "ВНИМАНИЕ" : "NOTICE",
                    Font = _technicalFont,
                    ForeColor = Color.FromArgb(235, 180, 75),
                    BackColor = Color.FromArgb(30, 26, 18),
                    Location = new Point(cardPad, cardY),
                    Padding = new Padding(UiScale(6), UiScale(2), UiScale(6), UiScale(2)),
                    UseMnemonic = false,
                };
                card.Controls.Add(noticeTag);
                cardY = noticeTag.Bottom + UiScale(6);

                Label noticeLabel = new()
                {
                    AutoSize = true,
                    MaximumSize = new Size(cardInnerW, 0),
                    Text = noticeText,
                    Font = _dialogFont,
                    ForeColor = Color.FromArgb(215, 218, 228),
                    BackColor = Color.Transparent,
                    TextAlign = ContentAlignment.TopLeft,
                    UseMnemonic = false,
                    Location = new Point(cardPad, cardY),
                };
                Size nSz = noticeLabel.GetPreferredSize(new Size(cardInnerW, 0));
                noticeLabel.AutoSize = false;
                noticeLabel.Size = new Size(cardInnerW, nSz.Height);
                card.Controls.Add(noticeLabel);

                card.Size = new Size(innerWidth, noticeLabel.Bottom + cardPad);
                dialog.Controls.Add(card);
                currentY = card.Bottom + UiScale(12);
            }

            // 3. Action prompt (last paragraph)
            string promptText = paragraphs[^1];
            Label promptLabel = new()
            {
                AutoSize = true,
                MaximumSize = new Size(innerWidth, 0),
                Text = promptText,
                Font = _blockTitleFont,
                ForeColor = Color.FromArgb(245, 248, 255),
                BackColor = _bgForm,
                TextAlign = ContentAlignment.TopLeft,
                UseMnemonic = false,
                Location = new Point(padding, currentY),
            };
            Size pSz = promptLabel.GetPreferredSize(new Size(innerWidth, 0));
            promptLabel.AutoSize = false;
            promptLabel.Size = new Size(innerWidth, pSz.Height);
            dialog.Controls.Add(promptLabel);
            currentY = promptLabel.Bottom + UiScale(14);
        }

        Button yesButton = NewDialogButton(UiLanguage.Text(yesText), buttonHeight, buttonWidth, isPrimary: true);
        yesButton.Name = yesText;
        yesButton.DialogResult = DialogResult.Yes;

        Button noButton = NewDialogButton(UiLanguage.Text(noText), buttonHeight, buttonWidth, isPrimary: false);
        noButton.Name = noText;
        noButton.DialogResult = DialogResult.No;

        AttachThemedDialogFooter(dialog, [yesButton, noButton], UiScale(12));

        dialog.ClientSize = new Size(clientWidth, currentY + footerHeight);
        dialog.AcceptButton = yesButton;
        dialog.CancelButton = noButton;

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
        dialog.Text = UiLanguage.Text("RESTORE");
        dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
        dialog.StartPosition = FormStartPosition.CenterParent;
        dialog.MaximizeBox = false;
        dialog.MinimizeBox = false;
        dialog.ShowInTaskbar = false;
        dialog.AutoScaleMode = AutoScaleMode.None;
        StyleThemedDialogSurface(dialog);
        dialog.Font = _dialogFont;
        dialog.Icon = Icon;
        dialog.BackColor = _bgForm;
        WireThemedTitleBar(dialog);

        bool hasBackup = backups.Count > 0;
        Button? deleteButton = null;

        int padding = UiScale(20);
        int clientWidth = UiScale(hasBackup ? 940 : 620);
        int standardButtonWidth = UiScale(160);
        int selectedButtonWidth = UiScale(180);
        int resetButtonWidth = UiScale(210);
        int cancelButtonWidth = UiScale(110);
        int buttonHeight = UiScale(32);
        int footerHeight = UiScale(50);
        int innerWidth = clientWidth - (padding * 2);

        RestoreChoice choice = RestoreChoice.Cancel;
        int currentY = padding;

        // Prompt Header
        string titleText = hasBackup
            ? (UiLanguage.IsRussian
                ? "Выберите точку восстановления для отката настроек:"
                : "Select a snapshot to restore system configuration:")
            : (UiLanguage.IsRussian
                ? "Точки восстановления конфигурации не найдены."
                : "No backup snapshots were found.");

        Label promptTitle = new()
        {
            AutoSize = true,
            MaximumSize = new Size(innerWidth, 0),
            Text = titleText,
            Font = _dialogFont,
            ForeColor = _fgMain,
            BackColor = _bgForm,
            TextAlign = ContentAlignment.MiddleLeft,
            UseMnemonic = false,
            Location = new Point(padding, currentY),
        };
        dialog.Controls.Add(promptTitle);
        currentY = promptTitle.Bottom + UiScale(4);

        string subText = hasBackup
            ? (UiLanguage.IsRussian
                ? "Откатывает службы, реестр и системные параметры к зафиксированному моменту времени."
                : "Reverts services, registry values, and device parameters back to captured state.")
            : (UiLanguage.IsRussian
                ? "Вы можете выполнить безопасный сброс настроек к стандартным значениям Windows."
                : "RESET WINDOWS DEFAULT restores supported settings back to clean Windows defaults.");

        Label promptSub = new()
        {
            AutoSize = true,
            MaximumSize = new Size(innerWidth, 0),
            Text = subText,
            Font = _technicalFont,
            ForeColor = Color.FromArgb(145, 150, 165),
            BackColor = _bgForm,
            TextAlign = ContentAlignment.MiddleLeft,
            UseMnemonic = false,
            Location = new Point(padding, currentY),
        };
        dialog.Controls.Add(promptSub);
        currentY = promptSub.Bottom + UiScale(12);

        // Snapshots List
        ListBox? backupList = null;
        if (hasBackup)
        {
            int itemHeight = UiScale(34);
            int visibleCount = Math.Min(backups.Count, 5);
            int listHeight = Math.Max(UiScale(68), (visibleCount * itemHeight) + UiScale(4));
            bool needsScroll = backups.Count > visibleCount;
            int scrollWidth = UiScale(14);

            Panel listPanel = new()
            {
                BackColor = Color.FromArgb(10, 10, 14),
                Location = new Point(padding, currentY),
                Size = new Size(innerWidth, listHeight),
                Padding = Padding.Empty,
            };
            listPanel.Paint += (_, e) =>
            {
                using Pen pen = new(Color.FromArgb(38, 38, 50), 1);
                e.Graphics.DrawRectangle(pen, 0, 0, listPanel.Width - 1, listPanel.Height - 1);
            };

            ThemedScrollBar? listScroll = null;
            if (needsScroll)
            {
                listScroll = new ThemedScrollBar
                {
                    Width = scrollWidth,
                    Dock = DockStyle.Right,
                    BackColor = Color.FromArgb(10, 10, 14),
                    TrackColor = Color.FromArgb(10, 10, 14),
                    RailColor = Color.FromArgb(10, 10, 14),
                    ThumbColor = Color.FromArgb(48, 48, 56),
                    ThumbHoverColor = Color.FromArgb(80, 80, 92),
                    ThumbDragColor = Color.FromArgb(120, 120, 135),
                    ThumbWidth = UiScale(8),
                    RailWidth = 0,
                    ThumbCornerRadius = UiScale(6),
                    Maximum = backups.Count,
                    ViewportSize = visibleCount,
                    SmallChange = 1,
                    LargeChange = visibleCount,
                    Value = 0,
                };
                listPanel.Controls.Add(listScroll);
            }

            int contentAreaWidth = needsScroll ? (innerWidth - scrollWidth - UiScale(2)) : (innerWidth - 2);

            Panel clipPanel = new()
            {
                Location = new Point(1, 1),
                Size = new Size(contentAreaWidth, listHeight - 2),
                BackColor = Color.FromArgb(10, 10, 14),
            };

            backupList = new ListBox
            {
                BorderStyle = BorderStyle.None,
                BackColor = Color.FromArgb(10, 10, 14),
                ForeColor = _fgMain,
                Font = _dialogFont,
                IntegralHeight = false,
                HorizontalScrollbar = false,
                DrawMode = DrawMode.OwnerDrawFixed,
                ItemHeight = itemHeight,
                Location = new Point(0, 0),
                Height = clipPanel.Height,
                Width = needsScroll
                    ? (clipPanel.Width + SystemInformation.VerticalScrollBarWidth + UiScale(16))
                    : clipPanel.Width,
            };

            backupList.DrawItem += (_, e) =>
            {
                if (e.Index < 0 || e.Index >= backupList.Items.Count)
                {
                    return;
                }

                if (backupList.Items[e.Index] is not BackupSnapshotInfo info)
                {
                    return;
                }

                e.Graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
                bool selected = (e.State & DrawItemState.Selected) != 0;
                Color bg = selected ? Color.FromArgb(20, 30, 46) : Color.FromArgb(12, 13, 18);
                using SolidBrush bgBrush = new(bg);
                e.Graphics.FillRectangle(bgBrush, e.Bounds);

                using Pen rowSepPen = new(Color.FromArgb(26, 28, 38), 1);
                e.Graphics.DrawLine(rowSepPen, e.Bounds.Left, e.Bounds.Bottom - 1, e.Bounds.Right, e.Bounds.Bottom - 1);

                if (selected)
                {
                    using SolidBrush barBrush = new(Color.FromArgb(90, 165, 255));
                    e.Graphics.FillRectangle(barBrush, e.Bounds.X, e.Bounds.Y, UiScale(3), e.Bounds.Height);

                    using Pen borderPen = new(Color.FromArgb(56, 84, 122), 1);
                    e.Graphics.DrawRectangle(borderPen, e.Bounds.X, e.Bounds.Y, e.Bounds.Width - 1, e.Bounds.Height - 1);
                }

                int textY = e.Bounds.Top + (itemHeight - UiScale(18)) / 2;
                int badgeH = UiScale(20);
                int badgeY = e.Bounds.Top + (itemHeight - badgeH) / 2;
                int x = e.Bounds.Left + UiScale(14);

                // Column 1: Date & Time
                DateTime stamp = info.CreatedAt ?? info.LastWriteUtc.ToLocalTime();
                string dateStr = stamp.ToString("yyyy-MM-dd  HH:mm:ss");
                Color dateColor = selected ? Color.FromArgb(248, 250, 255) : Color.FromArgb(175, 182, 196);
                TextRenderer.DrawText(e.Graphics, dateStr, _dialogFont, new Point(x, textY), dateColor, TextFormatFlags.NoPadding);
                x += UiScale(205);

                // Column 2: Location Badge
                int locBadgeW = UiScale(68);
                Rectangle locRect = new(x, badgeY, locBadgeW, badgeH);
                bool isLocalExe = info.Location.Contains("EXE", StringComparison.OrdinalIgnoreCase);
                Color locBgColor = isLocalExe ? Color.FromArgb(18, 38, 62) : Color.FromArgb(24, 32, 46);
                Color locBorderColor = isLocalExe ? Color.FromArgb(50, 96, 150) : Color.FromArgb(65, 85, 115);
                Color locTextColor = isLocalExe ? Color.FromArgb(140, 205, 255) : Color.FromArgb(190, 210, 235);
                using SolidBrush locBg = new(locBgColor);
                using Pen locBorder = new(locBorderColor, 1);
                e.Graphics.FillRectangle(locBg, locRect);
                e.Graphics.DrawRectangle(locBorder, locRect);
                TextRenderer.DrawText(e.Graphics, info.Location, _blockTitleFont, locRect, locTextColor, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);
                x += locBadgeW + UiScale(14);

                // Column 3: Type & Description (clean semantic contrast, no emoji shields)
                if (info.IsOriginal)
                {
                    string origText = "BASELINE";
                    Size origSz = TextRenderer.MeasureText(e.Graphics, origText, _blockTitleFont, Size.Empty, TextFormatFlags.NoPadding);
                    int origW = origSz.Width + UiScale(14);
                    Rectangle origRect = new(x, badgeY, origW, badgeH);
                    using SolidBrush origBg = new(Color.FromArgb(42, 32, 14));
                    using Pen origBorder = new(Color.FromArgb(135, 95, 30), 1);
                    e.Graphics.FillRectangle(origBg, origRect);
                    e.Graphics.DrawRectangle(origBorder, origRect);
                    TextRenderer.DrawText(e.Graphics, origText, _blockTitleFont, origRect, Color.FromArgb(255, 195, 70), TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);
                    x += origW + UiScale(12);

                    string desc = UiLanguage.IsRussian ? "Исходная конфигурация Windows (защищена)" : "Original Windows Baseline (Protected)";
                    TextRenderer.DrawText(e.Graphics, desc, _technicalFont, new Point(x, textY + UiScale(1)), Color.FromArgb(215, 220, 230), TextFormatFlags.NoPadding);
                }
                else
                {
                    string snapText = "SNAPSHOT";
                    Size snapSz = TextRenderer.MeasureText(e.Graphics, snapText, _blockTitleFont, Size.Empty, TextFormatFlags.NoPadding);
                    int snapW = snapSz.Width + UiScale(14);
                    Rectangle snapRect = new(x, badgeY, snapW, badgeH);
                    using SolidBrush snapBg = new(Color.FromArgb(26, 28, 36));
                    using Pen snapBorder = new(Color.FromArgb(65, 70, 85), 1);
                    e.Graphics.FillRectangle(snapBg, snapRect);
                    e.Graphics.DrawRectangle(snapBorder, snapRect);
                    TextRenderer.DrawText(e.Graphics, snapText, _blockTitleFont, snapRect, Color.FromArgb(180, 185, 200), TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);
                    x += snapW + UiScale(12);

                    TextRenderer.DrawText(e.Graphics, info.Reason, _dialogFont, new Point(x, textY), Color.FromArgb(240, 244, 252), TextFormatFlags.NoPadding);
                }
            };

            foreach (BackupSnapshotInfo backup in backups)
            {
                backupList.Items.Add(backup);
            }

            bool syncingScroll = false;
            if (needsScroll && listScroll is not null)
            {
                listScroll.ValueChanged += (_, _) =>
                {
                    if (!syncingScroll && backupList.Items.Count > 0)
                    {
                        syncingScroll = true;
                        backupList.TopIndex = Math.Clamp(listScroll.Value, 0, backupList.Items.Count - 1);
                        syncingScroll = false;
                    }
                };

                backupList.MouseWheel += (_, e) =>
                {
                    int delta = e.Delta > 0 ? -1 : 1;
                    listScroll.Value = Math.Clamp(listScroll.Value + delta, 0, Math.Max(0, backups.Count - visibleCount));
                };
            }

            backupList.SelectedIndex = 0;
            selectedPath = backups[0].Path;
            backupList.SelectedIndexChanged += (_, _) =>
            {
                int index = backupList.SelectedIndex;
                selectedPath = index >= 0 && index < backups.Count ? backups[index].Path : null;
                UpdateDeleteButtonState();
                if (!syncingScroll && needsScroll && listScroll is not null)
                {
                    syncingScroll = true;
                    listScroll.Value = Math.Clamp(backupList.TopIndex, 0, Math.Max(0, backups.Count - visibleCount));
                    syncingScroll = false;
                }
            };

            clipPanel.Controls.Add(backupList);
            listPanel.Controls.Add(clipPanel);
            dialog.Controls.Add(listPanel);
            currentY = listPanel.Bottom + UiScale(10);
        }

        if (hasBackup)
        {
            Label noteLabel = new()
            {
                AutoSize = false,
                Text = UiLanguage.IsRussian
                    ? "Исходный снимок (BASELINE) защищен от удаления."
                    : "Original baseline snapshot is permanently protected.",
                Font = _technicalFont,
                ForeColor = _statusInactive,
                BackColor = _bgForm,
                TextAlign = ContentAlignment.MiddleLeft,
                UseMnemonic = false,
                Location = new Point(padding, currentY),
                Size = new Size(innerWidth, UiScale(20)),
            };
            dialog.Controls.Add(noteLabel);
            currentY = noteLabel.Bottom + UiScale(8);
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
        }

        List<Button> footerButtons = [];

        Button? latestButton = null;
        if (hasBackup)
        {
            latestButton = NewDialogButton(UiLanguage.Text("RESTORE LAST"), buttonHeight, standardButtonWidth, isPrimary: true);
            latestButton.Name = "RESTORE LAST";
            latestButton.Click += (_, _) =>
            {
                choice = RestoreChoice.RestoreLatest;
                dialog.DialogResult = DialogResult.OK;
                dialog.Close();
            };
            footerButtons.Add(latestButton);
        }

        Button? backupButton = null;
        if (hasBackup)
        {
            backupButton = NewDialogButton(UiLanguage.Text("RESTORE SELECTED"), buttonHeight, selectedButtonWidth, isPrimary: false);
            backupButton.Name = "RESTORE SELECTED";
            backupButton.Click += (_, _) =>
            {
                choice = RestoreChoice.RestoreBackup;
                dialog.DialogResult = DialogResult.OK;
                dialog.Close();
            };
            footerButtons.Add(backupButton);
        }

        Button resetButton = NewDialogButton(UiLanguage.Text("RESET WINDOWS DEFAULT"), buttonHeight, resetButtonWidth, isPrimary: !hasBackup);
        resetButton.Name = "RESET WINDOWS DEFAULT";
        resetButton.Click += (_, _) =>
        {
            choice = RestoreChoice.SafeReset;
            dialog.DialogResult = DialogResult.OK;
            dialog.Close();
        };
        footerButtons.Add(resetButton);

        if (hasBackup)
        {
            deleteButton = NewDialogButton(UiLanguage.Text("DELETE SELECTED"), buttonHeight, standardButtonWidth, isPrimary: false);
            deleteButton.Name = "DELETE SELECTED";
            deleteButton.Click += (_, _) =>
            {
                choice = RestoreChoice.DeleteBackup;
                dialog.DialogResult = DialogResult.OK;
                dialog.Close();
            };
            UpdateDeleteButtonState();
            footerButtons.Add(deleteButton);
        }

        Button cancelButton = NewDialogButton(UiLanguage.Text("CANCEL"), buttonHeight, cancelButtonWidth, isPrimary: false);
        cancelButton.Name = "CANCEL";
        cancelButton.Click += (_, _) =>
        {
            choice = RestoreChoice.Cancel;
            dialog.DialogResult = DialogResult.Cancel;
            dialog.Close();
        };
        footerButtons.Add(cancelButton);

        AttachThemedDialogFooter(dialog, footerButtons, UiScale(10));

        dialog.ClientSize = new Size(clientWidth, currentY + footerHeight);
        dialog.AcceptButton = latestButton ?? resetButton;
        dialog.CancelButton = cancelButton;

        RestoreChoice result = ShowDialogDimmed(dialog) == DialogResult.OK ? choice : RestoreChoice.Cancel;
        selectedBackupPath = selectedPath;
        return result;
    }

    private AutoBackupChoice ShowAutoBackupChoiceDialog()
    {
        using Form dialog = new ThemedDialogForm();
        dialog.Name = "AUTO_BACKUP_DIALOG";
        dialog.Text = UiLanguage.Text("AUTO BACKUP");
        dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
        dialog.StartPosition = FormStartPosition.CenterParent;
        dialog.MaximizeBox = false;
        dialog.MinimizeBox = false;
        dialog.ShowInTaskbar = false;
        dialog.AutoScaleMode = AutoScaleMode.None;
        StyleThemedDialogSurface(dialog);
        dialog.Font = _dialogFont;
        dialog.Icon = Icon;
        dialog.BackColor = _bgForm;
        WireThemedTitleBar(dialog);

        int padding = UiScale(20);
        int buttonWidth = UiScale(140);
        int buttonHeight = UiScale(32);
        int footerHeight = UiScale(50);
        int clientWidth = UiScale(680);
        int innerWidth = clientWidth - (padding * 2);

        AutoBackupChoice choice = AutoBackupChoice.Cancel;
        int currentY = padding;

        Label promptLabel = new()
        {
            AutoSize = false,
            Text = UiLanguage.Text("Where should DEVICE TWEAKER save the pre-auto backup?"),
            Font = _blockTitleFont,
            ForeColor = _fgMain,
            BackColor = _bgForm,
            TextAlign = ContentAlignment.MiddleLeft,
            UseMnemonic = false,
            Location = new Point(padding, currentY),
            Size = new Size(innerWidth, UiScale(24)),
        };
        dialog.Controls.Add(promptLabel);
        currentY = promptLabel.Bottom + UiScale(4);

        Label subLabel = new()
        {
            AutoSize = false,
            Text = UiLanguage.IsRussian
                ? "Выберите расположение для создания точки отката:"
                : "Select backup snapshot storage before applying optimizations:",
            Font = _dialogFont,
            ForeColor = Color.FromArgb(155, 160, 175),
            BackColor = _bgForm,
            TextAlign = ContentAlignment.MiddleLeft,
            UseMnemonic = false,
            Location = new Point(padding, currentY),
            Size = new Size(innerWidth, UiScale(20)),
        };
        dialog.Controls.Add(subLabel);
        currentY = subLabel.Bottom + UiScale(12);

        (string badge, string badgeText, string desc, AutoBackupChoice optChoice, Color badgeFg, Color badgeBg, Color badgeBorder)[] options =
        [
            (
                "EXE FOLDER",
                UiLanguage.Text("EXE FOLDER"),
                UiLanguage.IsRussian ? "Портативный бэкап рядом с исполняемым файлом" : "Portable backup next to application executable",
                AutoBackupChoice.Local,
                Color.FromArgb(130, 195, 245),
                Color.FromArgb(20, 28, 40),
                Color.FromArgb(45, 65, 95)
            ),
            (
                "APPDATA",
                UiLanguage.Text("APPDATA"),
                UiLanguage.IsRussian ? "Хранилище профиля пользователя (сохранится при переносе)" : "User profile storage surviving folder moves/cleanups",
                AutoBackupChoice.Roaming,
                Color.FromArgb(180, 195, 225),
                Color.FromArgb(24, 26, 36),
                Color.FromArgb(50, 56, 76)
            ),
            (
                "SKIP",
                UiLanguage.Text("SKIP"),
                UiLanguage.IsRussian ? "Продолжить без создания новой точки отката" : "Continue without creating an additional snapshot",
                AutoBackupChoice.Skip,
                Color.FromArgb(160, 165, 178),
                Color.FromArgb(22, 24, 30),
                Color.FromArgb(45, 48, 60)
            ),
        ];

        int cardHeight = UiScale(42);
        foreach (var opt in options)
        {
            DeviceCardPanel card = new()
            {
                Location = new Point(padding, currentY),
                Size = new Size(innerWidth, cardHeight),
                BackColor = Color.FromArgb(14, 16, 22),
                BorderColor = Color.FromArgb(34, 38, 50),
                Cursor = Cursors.Hand,
            };

            int pillPaddingH = UiScale(10);
            int pillH = UiScale(24);
            int pillY = (cardHeight - pillH) / 2;
            int pillX = UiScale(10);
            Size badgeSz = TextRenderer.MeasureText(opt.badgeText, _blockTitleFont, Size.Empty, TextFormatFlags.NoPadding);
            int pillW = badgeSz.Width + (pillPaddingH * 2);

            DeviceCardPanel pillPanel = new()
            {
                Location = new Point(pillX, pillY),
                Size = new Size(pillW, pillH),
                BackColor = opt.badgeBg,
                BorderColor = opt.badgeBorder,
                Cursor = Cursors.Hand,
            };
            Label pillLabel = new()
            {
                AutoSize = false,
                Dock = DockStyle.Fill,
                Text = opt.badgeText,
                Font = _blockTitleFont,
                ForeColor = opt.badgeFg,
                BackColor = Color.Transparent,
                TextAlign = ContentAlignment.MiddleCenter,
                UseMnemonic = false,
                Cursor = Cursors.Hand,
            };
            pillPanel.Controls.Add(pillLabel);
            card.Controls.Add(pillPanel);

            int dividerX = pillX + pillW + UiScale(10);
            Panel divider = new()
            {
                Location = new Point(dividerX, pillY + UiScale(2)),
                Size = new Size(1, pillH - UiScale(4)),
                BackColor = Color.FromArgb(42, 46, 60),
            };
            card.Controls.Add(divider);

            int descX = dividerX + UiScale(10);
            Label descLabel = new()
            {
                AutoSize = false,
                Location = new Point(descX, 0),
                Size = new Size(innerWidth - descX - UiScale(8), cardHeight),
                Text = opt.desc,
                Font = _dialogFont,
                ForeColor = Color.FromArgb(220, 224, 235),
                BackColor = Color.Transparent,
                TextAlign = ContentAlignment.MiddleLeft,
                UseMnemonic = false,
                Cursor = Cursors.Hand,
            };
            card.Controls.Add(descLabel);

            void SelectChoice()
            {
                choice = opt.optChoice;
                dialog.DialogResult = DialogResult.OK;
                dialog.Close();
            }

            void OnEnter()
            {
                card.BorderColor = Color.FromArgb(65, 75, 100);
                card.BackColor = Color.FromArgb(18, 22, 30);
                card.Invalidate();
            }

            void OnLeave()
            {
                card.BorderColor = Color.FromArgb(34, 38, 50);
                card.BackColor = Color.FromArgb(14, 16, 22);
                card.Invalidate();
            }

            card.MouseEnter += (_, _) => OnEnter();
            card.MouseLeave += (_, _) => OnLeave();
            pillPanel.MouseEnter += (_, _) => OnEnter();
            pillPanel.MouseLeave += (_, _) => OnLeave();
            pillLabel.MouseEnter += (_, _) => OnEnter();
            pillLabel.MouseLeave += (_, _) => OnLeave();
            descLabel.MouseEnter += (_, _) => OnEnter();
            descLabel.MouseLeave += (_, _) => OnLeave();

            card.Click += (_, _) => SelectChoice();
            pillPanel.Click += (_, _) => SelectChoice();
            pillLabel.Click += (_, _) => SelectChoice();
            descLabel.Click += (_, _) => SelectChoice();

            dialog.Controls.Add(card);
            currentY = card.Bottom + UiScale(6);
        }

        currentY += UiScale(8);

        Panel notePanel = new()
        {
            Location = new Point(padding, currentY),
            Size = new Size(innerWidth, UiScale(22)),
            BackColor = _bgForm,
        };

        Label noteTag = new()
        {
            AutoSize = true,
            Text = UiLanguage.IsRussian ? "ИНФО:" : "NOTE:",
            Font = _technicalFont,
            ForeColor = Color.FromArgb(100, 150, 210),
            BackColor = _bgForm,
            Location = new Point(0, 0),
            UseMnemonic = false,
        };
        notePanel.Controls.Add(noteTag);

        Label noteDesc = new()
        {
            AutoSize = true,
            Text = UiLanguage.IsRussian
                ? "Снимок исходного состояния Windows сохраняется независимо от выбора."
                : "Original Windows baseline snapshot is always retained independently.",
            Font = _technicalFont,
            ForeColor = Color.FromArgb(145, 150, 165),
            BackColor = _bgForm,
            Location = new Point(noteTag.PreferredWidth + UiScale(8), 0),
            UseMnemonic = false,
        };
        notePanel.Controls.Add(noteDesc);
        dialog.Controls.Add(notePanel);
        currentY = notePanel.Bottom + UiScale(14);

        Button exeButton = NewDialogButton(UiLanguage.Text("EXE FOLDER"), buttonHeight, buttonWidth, isPrimary: true);
        exeButton.Name = "EXE FOLDER";
        exeButton.Click += (_, _) =>
        {
            choice = AutoBackupChoice.Local;
            dialog.DialogResult = DialogResult.OK;
            dialog.Close();
        };

        Button appDataButton = NewDialogButton(UiLanguage.Text("APPDATA"), buttonHeight, buttonWidth, isPrimary: false);
        appDataButton.Name = "APPDATA";
        appDataButton.Click += (_, _) =>
        {
            choice = AutoBackupChoice.Roaming;
            dialog.DialogResult = DialogResult.OK;
            dialog.Close();
        };

        Button skipButton = NewDialogButton(UiLanguage.Text("SKIP"), buttonHeight, buttonWidth, isPrimary: false);
        skipButton.Name = "SKIP";
        skipButton.Click += (_, _) =>
        {
            choice = AutoBackupChoice.Skip;
            dialog.DialogResult = DialogResult.OK;
            dialog.Close();
        };

        AttachThemedDialogFooter(dialog, [exeButton, appDataButton, skipButton], UiScale(12));

        dialog.ClientSize = new Size(clientWidth, currentY + footerHeight);
        dialog.AcceptButton = exeButton;
        dialog.CancelButton = skipButton;

        DialogResult result = ShowDialogDimmed(dialog);
        return result == DialogResult.OK ? choice : AutoBackupChoice.Cancel;
    }

    private static string PreventAwkwardHyphenBreaks(string text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return string.Empty;
        }

        // 1. Non-breaking hyphen \u2011 prevents ugly line-breaks splitting technical compound words across lines
        string result = text
            .Replace("anti-cheats", "anti\u2011cheats")
            .Replace("anti-cheat", "anti\u2011cheat")
            .Replace("анти-читами", "анти\u2011читами")
            .Replace("анти-читы", "анти\u2011читы")
            .Replace("анти-чит", "анти\u2011чит")
            .Replace("XHCI-", "XHCI\u2011")
            .Replace("USB-", "USB\u2011")
            .Replace("P-Core", "P\u2011Core")
            .Replace("E-Core", "E\u2011Core")
            .Replace("P-ядра", "P\u2011ядра")
            .Replace("E-ядра", "E\u2011ядра")
            .Replace("MSI-X", "MSI\u2011X")
            .Replace("pre-auto", "pre\u2011auto")
            .Replace("SHA-256", "SHA\u2011256");

        // 2. Non-breaking space \u00A0 prevents technical acronyms, units, and short Russian prepositions from dangling at line ends
        result = result
            .Replace("USB IMOD", "USB\u00A0IMOD")
            .Replace("8000 Гц", "8000\u00A0Гц")
            .Replace("1000 Гц", "1000\u00A0Гц")
            .Replace("0 мкс", "0\u00A0мкс")
            .Replace("50 мкс", "50\u00A0мкс")
            .Replace("250 нс", "250\u00A0нс")
            .Replace("перезагрузите ПК", "перезагрузите\u00A0ПК")
            .Replace("сброса Windows", "сброса\u00A0Windows")
            .Replace("сброс Windows", "сброс\u00A0Windows");

        // Bind short Russian prepositions (в, во, на, с, со, к, по, для, не, или, и, от, до, за, из, о, об) to the next word
        result = System.Text.RegularExpressions.Regex.Replace(
            result,
            @"(?<=(?:^|[\s(«""]))([вВ][оО]?|[нН][аА]|[сС][оО]?|[кК]|[пП][оО]|[дД][лЛ][яЯ]|[нН][еЕ]|[иИ][лЛ][иИ]|[иИ]|[оО][тТ]|[дД][оО]|[зЗ][аА]|[иИ][зЗ]|[оО][бБ]?)\s+",
            "$1\u00A0");

        return result;
    }

    private static string NormalizeDialogMessage(string message)
    {
        if (string.IsNullOrEmpty(message))
        {
            return string.Empty;
        }

        string normalized = PreventAwkwardHyphenBreaks(message)
            .Replace("\r\n", "\n")
            .Replace("\r", "\n");
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
