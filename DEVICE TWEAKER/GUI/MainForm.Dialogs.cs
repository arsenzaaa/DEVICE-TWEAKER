
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

    private void ShowThemedInfo(string message, string title, string? filePath = null)
    {
        using Form dialog = new ThemedDialogForm
        {
            Name = "INFO_DIALOG", Text = UiLanguage.Text(title),
            FormBorderStyle = FormBorderStyle.FixedDialog, StartPosition = FormStartPosition.CenterParent,
            MaximizeBox = false, MinimizeBox = false, ShowInTaskbar = false,
            AutoScaleMode = AutoScaleMode.None, Font = _dialogFont, Icon = Icon,
            BackColor = _bgForm, ForeColor = _fgMain,
        };
        WireThemedTitleBar(dialog);
        int padding = UiScale(24), width = UiScale(620), inner = width - padding * 2;
        int y = padding;
        foreach (string paragraph in NormalizeDialogMessage(UiLanguage.Text(message)).Replace("\r\n", "\n").Split("\n\n", StringSplitOptions.RemoveEmptyEntries))
            AddCenteredDialogText(dialog, paragraph.Trim(), inner, padding, ref y, _fgMain);
        if (filePath is not null)
        {
            Label file = new()
            {
                Name = "BACKUP_FILE", Text = Path.GetFileName(filePath), AccessibleDescription = filePath,
                AutoEllipsis = true, UseMnemonic = false, Font = _technicalFont, ForeColor = _mutedText,
                BackColor = _bgForm, TextAlign = ContentAlignment.MiddleCenter,
                Location = new Point(padding, y), Size = new Size(inner, UiScale(26)),
            };
            dialog.Controls.Add(file);
            y = file.Bottom + UiScale(16);
        }
        Button ok = NewDialogButton(UiLanguage.Text("OK"), UiScale(32), UiScale(96), isPrimary: true);
        ok.Name = "OK"; ok.DialogResult = DialogResult.OK;
        List<Button> buttons = [ok];
        if (filePath is not null)
        {
            Button details = NewDialogButton(UiLanguage.Text("DETAILS"), UiScale(32));
            details.Name = "DETAILS";
            details.Click += (_, _) => ShowFileLocationDialog(dialog, filePath);
            buttons.Add(details);
            Button folder = NewDialogButton(UiLanguage.Text("BACKUPS"), UiScale(32));
            folder.Name = "BACKUPS";
            folder.Click += (_, _) => TryOpenResultFolder(ResolveResultBackupFolder(filePath), "UI.BACKUP.FOLDER");
            buttons.Add(folder);
        }
        AttachThemedDialogFooter(dialog, buttons);
        dialog.ClientSize = new Size(width, y + UiScale(8) + UiScale(50));
        dialog.AcceptButton = ok; dialog.CancelButton = ok;
        ShowDialogDimmed(dialog);
    }

    private void ShowBackupFileResult(string path, bool restored)
        => ShowThemedInfo(restored
            ? "Backup restored.\n\nRestart Windows to finish restoring settings."
            : "Backup saved.", "DEVICE TWEAKER", path);

    private void ShowFileLocationDialog(Form owner, string path)
    {
        using Form dialog = new ThemedDialogForm
        {
            Name = "FILE_LOCATION_DIALOG", Text = UiLanguage.Text("BACKUP FILE"),
            FormBorderStyle = FormBorderStyle.FixedDialog, StartPosition = FormStartPosition.CenterParent,
            MaximizeBox = false, MinimizeBox = false, ShowInTaskbar = false,
            AutoScaleMode = AutoScaleMode.None, Font = _dialogFont, Icon = Icon,
            BackColor = _bgForm, ForeColor = _fgMain,
        };
        WireThemedTitleBar(dialog);
        Label hint = new() { Text = UiLanguage.Text("Full path (select to copy):"), AutoSize = true,
            Font = _technicalFont, ForeColor = _mutedText, Location = new Point(UiScale(24), UiScale(20)) };
        TextBox pathBox = new()
        {
            Name = "BACKUP_FULL_PATH", Text = path, ReadOnly = true, Multiline = false,
            Font = _dialogFont, BackColor = _bgForm, ForeColor = _fgMain, BorderStyle = BorderStyle.FixedSingle,
            Location = new Point(UiScale(24), hint.Bottom + UiScale(12)), Width = UiScale(632),
        };
        dialog.Controls.Add(hint); dialog.Controls.Add(pathBox);
        Button ok = NewDialogButton(UiLanguage.Text("OK"), UiScale(32), UiScale(96), isPrimary: true);
        ok.Name = "OK"; ok.DialogResult = DialogResult.OK;
        AttachThemedDialogFooter(dialog, [ok]);
        dialog.ClientSize = new Size(UiScale(680), pathBox.Bottom + UiScale(24) + UiScale(50));
        dialog.AcceptButton = ok; dialog.CancelButton = ok;
        LocalizeControlTree(dialog);
        dialog.ShowDialog(owner);
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
            AccessibleRole = AccessibleRole.StaticText;
            AccessibleName = $"{component}: {detail} ({stateTag})";

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
        bool hasTechnical = technicalText.Length > 0 && (!isSuccess
            || report.Issues.Any(issue => !string.IsNullOrWhiteSpace(issue.TechnicalDetails)));
        string backupFolder = ResolveResultBackupFolder(report.BackupPath);
        string logsFolder = AppDiagnostics.LogDirectory;

        int padding = UiScale(20);
        int buttonHeight = UiScale(32);
        int footerHeight = UiScale(50);
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

        int clientWidth = isSuccess ? UiScale(480) : UiScale(600);
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

            foreach (string line in msg.Replace("\r\n", "\n").Split('\n', StringSplitOptions.RemoveEmptyEntries))
                AddCenteredDialogText(dialog, line, innerWidth, padding, ref currentY, _fgMain);
        }
        else
        {
            string statusHeader = isWarnings
                ? (UiLanguage.IsRussian ? "Применено с замечаниями" : "Applied with warnings")
                : isPartial
                    ? (UiLanguage.IsRussian ? "Применено частично" : "Partially applied")
                    : (UiLanguage.IsRussian ? "Операция не выполнена" : "Operation not applied");
            AddCenteredDialogText(dialog, statusHeader, innerWidth, padding, ref currentY, _fgMain);
            string lead = NormalizeDialogMessage(UiLanguage.Text(report.Succeeded ? successMessage : partialMessage));
            if (report.NoChangesMade && string.IsNullOrWhiteSpace(lead))
                lead = UiLanguage.Text("The operation stopped before any changes were made.");
            string comparableLead = lead.Replace('\u00a0', ' ');
            if (isWarnings && comparableLead.Equals(UiLanguage.Text("Applied with warnings. Core settings were successfully configured.").Replace('\u00a0', ' '), StringComparison.OrdinalIgnoreCase))
                lead = string.Empty;
            else if (comparableLead.StartsWith(statusHeader + ".", StringComparison.OrdinalIgnoreCase))
                lead = lead[(statusHeader.Length + 1)..].Trim();
            else if (comparableLead.StartsWith(statusHeader + " (", StringComparison.OrdinalIgnoreCase))
                lead = lead[(statusHeader.Length + 2)..].TrimEnd('.', ')');
            if (lead.Length > 0 && !lead.Equals(statusHeader, StringComparison.OrdinalIgnoreCase))
                AddCenteredDialogText(dialog, lead, innerWidth, padding, ref currentY, _mutedText);
            foreach (OperationIssue issue in visibleIssues)
            {
                string detail = !string.IsNullOrWhiteSpace(issue.UserMessage)
                    ? UiLanguage.Text(issue.UserMessage)
                    : UiLanguage.Text(issue.IsSkipped ? "Skipped" : issue.Severity == OperationIssueSeverity.Error ? "Not applied" : "Warning");
                if (report.NoChangesMade)
                {
                    foreach (string suffix in new[] { " No changes were made.", " Настройки не изменены." })
                        if (detail.Length > suffix.Length && detail.Replace('\u00a0', ' ').EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
                            detail = detail[..^suffix.Length];
                }
                string skip = issue.IsSkipped ? " (" + UiLanguage.Text("Skipped") + ")" : string.Empty;
                AddCenteredDialogText(dialog, NormalizeDialogMessage(UiLanguage.Text(issue.Component) + skip + ": " + detail), innerWidth, padding, ref currentY, _mutedText);
            }
            if (hasTechnical)
                AddCenteredDialogText(dialog, UiLanguage.Text(isPartial
                    ? "See DETAILS for applied / skipped settings and their reasons."
                    : isWarnings
                        ? "Full results / warnings are in DETAILS."
                        : "See DETAILS and LOGS for the error."), innerWidth, padding, ref currentY, _mutedText);
        }

        AttachThemedDialogFooter(dialog, visibleButtons, UiScale(10));
        int compactHeight = currentY + UiScale(12) + footerHeight;
        dialog.ClientSize = new Size(clientWidth, Math.Min(compactHeight, workingArea.Height - UiScale(60)));
        logsButton.Click += (_, _) => TryOpenResultFolder(logsFolder, "UI.RESULT.LOGS");
        backupsButton.Click += (_, _) => TryOpenResultFolder(backupFolder, "UI.RESULT.BACKUPS");
        if (detailsButton is not null)
            detailsButton.Click += (_, _) => ShowOperationDetailsDialog(dialog, technicalText);

        dialog.AcceptButton = okButton;
        int successCount = report.Issues.Count(issue => issue.Severity == OperationIssueSeverity.Success);
        string previewScope = report.IsDeveloperPreview ? " scope=qa-dialog-preview" : string.Empty;
        WriteLog(
            $"UI.RESULT.SUMMARY: operation=\"{SanitizeLogValue(operation)}\" state={state} " +
            $"results={report.Issues.Count} successes={successCount} warnings={report.Warnings.Count} " +
            $"errors={report.Errors.Count} noChanges={report.NoChangesMade} " +
            $"backup=\"{SanitizeLogValue(report.BackupPath)}\" backupsFolder=\"{SanitizeLogValue(backupFolder)}\"{previewScope}");
        foreach (OperationIssue issue in report.Issues)
        {
            WriteLog(
                $"UI.RESULT.ITEM: severity={issue.Severity} component=\"{SanitizeLogValue(issue.Component)}\" " +
                $"message=\"{SanitizeLogValue(issue.UserMessage)}\" " +
                $"technicalDetails=\"{SanitizeLogValue(issue.TechnicalDetails)}\"{previewScope}");
        }
        ShowDialogDimmed(dialog);
    }

    private void AddCenteredDialogText(Form dialog, string text, int width, int padding, ref int y, Color color)
    {
        Label label = new()
        {
            Text = text, Font = _dialogFont, ForeColor = color, BackColor = _bgForm,
            Location = new Point(padding, y),
            TextAlign = ContentAlignment.MiddleCenter, UseMnemonic = false,
        };
        // Keep the same line count while narrowing the paragraph. This balances
        // centered lines instead of leaving a few words alone on the final line.
        int height = label.GetPreferredSize(new Size(width, 0)).Height;
        int lower = width * 3 / 5, upper = width;
        if (!text.Contains('\n') && height > label.GetPreferredSize(new Size(int.MaxValue, 0)).Height)
        {
            while (lower < upper)
            {
                int candidate = lower + (upper - lower) / 2;
                if (label.GetPreferredSize(new Size(candidate, 0)).Height <= height)
                    upper = candidate;
                else
                    lower = candidate + 1;
            }
        }
        int paragraphWidth = Math.Min(width, upper + UiScale(4));
        label.MaximumSize = new Size(paragraphWidth, 0);
        label.Location = new Point(padding + (width - paragraphWidth) / 2, y);
        label.Size = new Size(paragraphWidth, label.GetPreferredSize(new Size(paragraphWidth, 0)).Height);
        dialog.Controls.Add(label);
        y = label.Bottom + UiScale(14);
    }

    private void ShowOperationDetailsDialog(Form owner, string text)
    {
        using Form dialog = new ThemedDialogForm
        {
            Name = "OPERATION_DETAILS_DIALOG", Text = UiLanguage.Text("DETAILS"),
            FormBorderStyle = FormBorderStyle.FixedDialog, StartPosition = FormStartPosition.CenterParent,
            MaximizeBox = false, MinimizeBox = false, ShowInTaskbar = false,
            AutoScaleMode = AutoScaleMode.None, Font = _dialogFont, Icon = Icon,
            BackColor = _bgForm, ForeColor = _fgMain,
        };
        WireThemedTitleBar(dialog);
        int padding = UiScale(24), width = UiScale(680);
        int maxHeight = Math.Min(UiScale(360), Screen.FromControl(owner).WorkingArea.Height - UiScale(180));
        Panel viewport = new() { BackColor = _bgForm, Location = new Point(padding, padding), Size = new Size(width - padding * 2, maxHeight) };
        Panel content = new() { BackColor = _bgForm };
        Label body = new() { Text = NormalizeDialogMessage(text), Font = _technicalFont, ForeColor = _mutedText, BackColor = _bgForm, UseMnemonic = false };
        int bodyWidth = viewport.Width - UiScale(22);
        body.MaximumSize = new Size(bodyWidth, 0);
        body.Size = new Size(bodyWidth, body.GetPreferredSize(new Size(bodyWidth, 0)).Height);
        int height = Math.Min(maxHeight, Math.Max(UiScale(64), body.Height));
        viewport.Height = height;
        content.Size = body.Size;
        content.Controls.Add(body);
        ThemedScrollBar scroll = new()
        {
            Name = "DETAILS_SCROLL", Width = UiScale(14), Dock = DockStyle.Right, BackColor = _bgForm, TrackColor = _bgForm, RailColor = _bgForm,
            RailWidth = 0, ThumbWidth = UiScale(8), ThumbCornerRadius = UiScale(6),
            ThumbColor = Color.FromArgb(48, 48, 56), ThumbHoverColor = Color.FromArgb(80, 80, 92), ThumbDragColor = Color.FromArgb(120, 120, 135),
            Maximum = Math.Max(1, content.Height), ViewportSize = height, SmallChange = UiScale(40), Visible = content.Height > height,
        };
        scroll.ValueChanged += (_, _) => content.Top = -scroll.Value;
        dialog.KeyPreview = true;
        dialog.KeyDown += (_, e) =>
        {
            int? next = e.KeyCode switch
            {
                Keys.Home => 0, Keys.End => Math.Max(0, content.Height - height),
                Keys.PageUp => scroll.Value - height, Keys.PageDown => scroll.Value + height,
                Keys.Up => scroll.Value - scroll.SmallChange, Keys.Down => scroll.Value + scroll.SmallChange,
                _ => null,
            };
            if (next is null) return;
            scroll.Value = next.Value;
            e.Handled = true; e.SuppressKeyPress = true;
        };
        body.MouseWheel += (_, e) => scroll.Value += e.Delta > 0 ? -scroll.SmallChange : scroll.SmallChange;
        viewport.Controls.Add(content); viewport.Controls.Add(scroll); dialog.Controls.Add(viewport);
        Button ok = NewDialogButton(UiLanguage.Text("OK"), UiScale(32), UiScale(96), isPrimary: true);
        ok.Name = "OK"; ok.DialogResult = DialogResult.OK;
        AttachThemedDialogFooter(dialog, [ok]);
        dialog.ClientSize = new Size(width, viewport.Bottom + padding + UiScale(50));
        dialog.AcceptButton = ok; dialog.CancelButton = ok;
        LocalizeControlTree(dialog);
        dialog.ShowDialog(owner);
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
        if (QaRun.TryBlock("EXTERNAL FOLDER")) return;
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

    private int DialogButtonTextWidth(string text)
        => TextRenderer.MeasureText(text, _buttonFont, Size.Empty, TextFormatFlags.NoPadding | TextFormatFlags.SingleLine).Width + UiScale(24);

    private Button NewDialogButton(string text, int height, int? width = null, bool isPrimary = false)
    {
        Button button = new ThemedButton
        {
            Name = text,
            Text = text,
            Size = new Size(Math.Max(width ?? UiScale(100), DialogButtonTextWidth(text)), height),
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
        button.FlatAppearance.MouseOverBackColor = hoverBg;
        button.FlatAppearance.MouseDownBackColor = Color.FromArgb(38, 38, 44);

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
            index++;
            string severity = UiLanguage.Text(issue.IsSkipped ? "SKIPPED" : issue.Severity.ToString().ToUpperInvariant());
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
                body = techDetails ?? userMsg ?? UiLanguage.Text(issue.Severity == OperationIssueSeverity.Success
                    ? "Applied" : "No additional details.");
            }
            sections.Add($"{title}{Environment.NewLine}{body}");
        }

        return string.Join(Environment.NewLine + Environment.NewLine, sections);
    }

    private bool ShowThemedConfirm(string message, string title, string yesText = "YES", string noText = "NO", DriverEnvironmentSnapshot? driverEnvironment = null)
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
        int clientWidth = UiScale(driverEnvironment is null ? 640 : 600);
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

        foreach (string paragraph in paragraphs)
            AddCenteredDialogText(dialog, paragraph, innerWidth, padding, ref currentY,
                currentY == padding ? _fgMain : _mutedText);

        Button yesButton = NewDialogButton(UiLanguage.Text(yesText), buttonHeight, buttonWidth, isPrimary: true);
        yesButton.Name = yesText;
        yesButton.DialogResult = DialogResult.Yes;

        Button noButton = NewDialogButton(UiLanguage.Text(noText), buttonHeight, buttonWidth, isPrimary: false);
        noButton.Name = noText;
        noButton.DialogResult = DialogResult.No;

        List<Button> confirmButtons = [yesButton, noButton];
        if (driverEnvironment is not null)
        {
            Button checkButton = NewDialogButton(UiLanguage.Text("SYSTEM CHECK"), buttonHeight, UiScale(116));
            checkButton.Name = "SYSTEM_CHECK_TOGGLE";
            checkButton.FlatAppearance.MouseOverBackColor = Color.FromArgb(24, 24, 30);
            checkButton.FlatAppearance.MouseDownBackColor = Color.FromArgb(34, 34, 40);
            checkButton.Click += (_, _) => ShowDriverEnvironmentDetails(dialog, driverEnvironment);
            confirmButtons.Add(checkButton);
        }
        AttachThemedDialogFooter(dialog, confirmButtons, UiScale(12));

        dialog.ClientSize = new Size(clientWidth, currentY + footerHeight);
        dialog.AcceptButton = yesButton;
        dialog.CancelButton = noButton;

        return ShowDialogDimmed(dialog) == DialogResult.Yes;
    }

    private void ShowDriverEnvironmentDetails(Form owner, DriverEnvironmentSnapshot snapshot)
    {
        using Form dialog = new ThemedDialogForm
        {
            Name = "DRIVER_ENVIRONMENT_DIALOG", Text = UiLanguage.Text("SYSTEM CHECK"),
            FormBorderStyle = FormBorderStyle.FixedDialog, StartPosition = FormStartPosition.CenterParent,
            MaximizeBox = false, MinimizeBox = false, ShowInTaskbar = false,
            AutoScaleMode = AutoScaleMode.None, Font = _dialogFont, Icon = Icon,
            BackColor = _bgForm, ForeColor = _fgMain,
        };
        WireThemedTitleBar(dialog);
        int padding = UiScale(24), width = UiScale(740), inner = width - padding * 2;
        int y = padding;
        string RegistryState(DriverRegistryState state) => state.Value is uint value
            ? UiLanguage.Text(value == 0 ? "Off" : "On") + $" ({value})"
            : UiLanguage.Text(state.Evidence == "missing" ? "No registry setting" : "Could not read");
        string ServiceState(DriverServiceState service) => UiLanguage.Text(service.Exists switch
        {
            false => "Service not installed", null => "Could not read",
            _ => service.State switch
            {
                "running" => "Service running", "stopped" => "Service stopped", "starting" => "Starting", "stopping" => "Stopping",
                "continuing" => "Continuing", "pausing" => "Pausing", "paused" => "Paused", _ => "Could not read",
            },
        });
        (string Label, string Value)[] rows =
        [
            ("Vulnerable driver blocklist", RegistryState(snapshot.Blocklist)),
            ("Memory integrity (HVCI)", UiLanguage.Text(snapshot.HvciRunning switch { true => "On", false => "Off", _ => "Could not read" })),
            ("HVCI registry setting", RegistryState(snapshot.HvciConfigured)),
            ("Microsoft Defender", ServiceState(snapshot.Defender)),
            ("FACEIT", ServiceState(snapshot.Faceit)),
            ("Vanguard (vgk)", ServiceState(snapshot.VanguardDriver)),
            ("Vanguard (vgc)", ServiceState(snapshot.VanguardService)),
        ];
        for (int i = 0; i < rows.Length; i++)
        {
            if (i is 0 or 3)
            {
                if (i == 3) y += UiScale(12);
                Label section = new() { Text = UiLanguage.Text(i == 0 ? "Windows settings" : "Services"),
                    Font = _blockTitleFont, ForeColor = _mutedText, BackColor = _bgForm,
                    AutoSize = true, Location = new Point(padding, y) };
                dialog.Controls.Add(section);
                y += section.Height + UiScale(12);
            }
            Label name = new() { Name = $"SYSTEM_CHECK_NAME_{i}", Text = UiLanguage.Text(rows[i].Label), Font = _dialogFont,
                ForeColor = _fgMain, BackColor = _bgForm, UseMnemonic = false, Location = new Point(padding, y), Size = new Size(UiScale(390), UiScale(26)) };
            Label value = new() { Name = $"SYSTEM_CHECK_VALUE_{i}", Text = rows[i].Value, Font = _dialogFont,
                ForeColor = _mutedText, BackColor = _bgForm, UseMnemonic = false, Location = new Point(padding + UiScale(390), y), Size = new Size(inner - UiScale(390), UiScale(26)) };
            dialog.Controls.Add(name); dialog.Controls.Add(value);
            y += UiScale(28);
        }
        y += UiScale(12);
        foreach (string text in new[]
        {
            "This check reads settings and service status. It does not test DTIMOD.sys loading.",
        })
        {
            Label note = new()
            {
                Name = "SYSTEM_CHECK_NOTE_" + y, Text = NormalizeDialogMessage(UiLanguage.Text(text)),
                Font = _technicalFont, ForeColor = _mutedText, BackColor = _bgForm, UseMnemonic = false,
                Location = new Point(padding, y), MaximumSize = new Size(inner, 0),
            };
            note.Size = new Size(inner, note.GetPreferredSize(new Size(inner, 0)).Height);
            dialog.Controls.Add(note);
            y = note.Bottom + UiScale(12);
        }
        Button closeButton = NewDialogButton(UiLanguage.Text("OK"), UiScale(32), UiScale(96), isPrimary: true);
        closeButton.Name = "OK"; closeButton.DialogResult = DialogResult.OK;
        AttachThemedDialogFooter(dialog, [closeButton]);
        dialog.ClientSize = new Size(width, y + UiScale(6) + UiScale(50));
        dialog.AcceptButton = closeButton; dialog.CancelButton = closeButton;
        LocalizeControlTree(dialog);
        WriteLog("UI.IMOD.PREFLIGHT.DETAILS: opened separate system check");
        dialog.ShowDialog(owner);
    }

    private bool ShowThemedConfirm(string message)
    {
        return ShowThemedConfirm(message, "DEVICE TWEAKER");
    }

    private RestoreChoice ShowRestoreChoiceDialog(IReadOnlyList<BackupSnapshotInfo> sourceBackups, out string? selectedBackupPath, Func<BackupSnapshotInfo, int>? deleteBackup = null)
    {
        List<BackupSnapshotInfo> backups = sourceBackups.ToList();
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
        Button? latestButton = null;
        Button? backupButton = null;
        Label? menuStatus = null;
        ThemedScrollBar? listScroll = null;
        int visibleCount = Math.Min(backups.Count, 5);

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
                ? "Выберите бекап для восстановления настроек:"
                : "Select a backup to restore settings:")
            : (UiLanguage.IsRussian
                ? "Бекапы не найдены."
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
                ? "Настройки устройств / сохранённый профиль IMOD и NIC ITR."
                : "Device settings / saved IMOD and NIC ITR profile.")
            : (UiLanguage.IsRussian
                ? "Сброс возвращает поддерживаемые настройки к значениям Windows."
                : "Reset restores supported settings to Windows defaults.");

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
                Name = "BACKUP_LIST",
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
                    string origText = UiLanguage.Text("ORIGINAL");
                    Size origSz = TextRenderer.MeasureText(e.Graphics, origText, _blockTitleFont, Size.Empty, TextFormatFlags.NoPadding);
                    int origW = origSz.Width + UiScale(14);
                    Rectangle origRect = new(x, badgeY, origW, badgeH);
                    using SolidBrush origBg = new(Color.FromArgb(42, 32, 14));
                    using Pen origBorder = new(Color.FromArgb(135, 95, 30), 1);
                    e.Graphics.FillRectangle(origBg, origRect);
                    e.Graphics.DrawRectangle(origBorder, origRect);
                    TextRenderer.DrawText(e.Graphics, origText, _blockTitleFont, origRect, Color.FromArgb(255, 195, 70), TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);
                    x += origW + UiScale(12);

                    string desc = UiLanguage.IsRussian ? "Исходные настройки устройств (защищены)" : "Original device settings (protected)";
                    TextRenderer.DrawText(e.Graphics, desc, _technicalFont, new Point(x, textY + UiScale(1)), Color.FromArgb(215, 220, 230), TextFormatFlags.NoPadding);
                }
                else
                {
                    string snapText = UiLanguage.Text("BACKUP");
                    Size snapSz = TextRenderer.MeasureText(e.Graphics, snapText, _blockTitleFont, Size.Empty, TextFormatFlags.NoPadding);
                    int snapW = snapSz.Width + UiScale(14);
                    Rectangle snapRect = new(x, badgeY, snapW, badgeH);
                    using SolidBrush snapBg = new(Color.FromArgb(26, 28, 36));
                    using Pen snapBorder = new(Color.FromArgb(65, 70, 85), 1);
                    e.Graphics.FillRectangle(snapBg, snapRect);
                    e.Graphics.DrawRectangle(snapBorder, snapRect);
                    TextRenderer.DrawText(e.Graphics, snapText, _blockTitleFont, snapRect, Color.FromArgb(180, 185, 200), TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);
                    x += snapW + UiScale(12);

                    TextRenderer.DrawText(e.Graphics, UiLanguage.Text(info.Reason), _dialogFont, new Point(x, textY), Color.FromArgb(240, 244, 252), TextFormatFlags.NoPadding);
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
                    ? "Бекап исходных настроек защищён от удаления."
                    : "The original-settings backup is protected from deletion.",
                Font = _technicalFont,
                ForeColor = _statusInactive,
                BackColor = _bgForm,
                TextAlign = ContentAlignment.MiddleLeft,
                UseMnemonic = false,
                Location = new Point(padding, currentY),
                Size = new Size(innerWidth, UiScale(40)),
            };
            menuStatus = noteLabel;
            noteLabel.Name = "BACKUP_MENU_STATUS";
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
            bool selected = index >= 0 && index < backups.Count;
            deleteButton.Enabled = selected && !backups[index].IsOriginal;
            if (latestButton is not null) latestButton.Enabled = backups.Count > 0;
            if (backupButton is not null) backupButton.Enabled = selected;
        }

        List<Button> footerButtons = [];

        if (hasBackup)
        {
            latestButton = NewDialogButton(UiLanguage.Text("RESTORE LAST"), buttonHeight, standardButtonWidth, isPrimary: true);
            latestButton.Name = "RESTORE LAST";
            latestButton.Click += (_, _) =>
            {
                selectedPath = backups.FirstOrDefault()?.Path;
                choice = RestoreChoice.RestoreBackup;
                dialog.DialogResult = DialogResult.OK;
                dialog.Close();
            };
            footerButtons.Add(latestButton);
        }

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
                if (deleteBackup is null)
                {
                    choice = RestoreChoice.DeleteBackup;
                    dialog.DialogResult = DialogResult.OK;
                    dialog.Close();
                    return;
                }
                int index = backupList!.SelectedIndex;
                if (index < 0 || index >= backups.Count || backups[index].IsOriginal) return;
                BackupSnapshotInfo selected = backups[index];
                try
                {
                    int deleted = deleteBackup(selected);
                    if (deleted == 0)
                    {
                        menuStatus!.Text = UiLanguage.Text("The backup could not be deleted. See LOGS for details.");
                        return;
                    }
                    backupList.BeginUpdate();
                    try
                    {
                        backups.RemoveAt(index);
                        backupList.Items.RemoveAt(index);
                        backupList.SelectedIndex = backups.Count == 0 ? -1 : Math.Min(index, backups.Count - 1);
                        selectedPath = backupList.SelectedItem is BackupSnapshotInfo remaining ? remaining.Path : null;
                    }
                    finally { backupList.EndUpdate(); }
                    menuStatus!.Text = UiLanguage.Text(backups.Count == 0
                        ? "Backup deleted. No backups remain."
                        : "Backup deleted. Select another backup or close this window.");
                    if (listScroll is not null)
                    {
                        listScroll.Maximum = Math.Max(1, backups.Count);
                        listScroll.Value = Math.Min(listScroll.Value, Math.Max(0, backups.Count - visibleCount));
                        listScroll.Visible = backups.Count > visibleCount;
                    }
                    UpdateDeleteButtonState();
                }
                catch (Exception ex)
                {
                    WriteLog($"BACKUP.DELETE: failed path={selected.Path} error={FlattenLogText(ex.ToString())}");
                    menuStatus!.Text = UiLanguage.Text("The backup could not be deleted. See LOGS for details.");
                }
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
        int clientWidth = UiScale(600);
        int innerWidth = clientWidth - (padding * 2);

        AutoBackupChoice choice = AutoBackupChoice.Cancel;
        int currentY = padding;

        AddCenteredDialogText(dialog, UiLanguage.Text("Where should the additional backup be saved?"), innerWidth, padding, ref currentY, _fgMain);
        AddCenteredDialogText(dialog, UiLanguage.Text("The EXE folder keeps the backup next to the program.\nAppData keeps it in your user profile."), innerWidth, padding, ref currentY, _mutedText);
        AddCenteredDialogText(dialog, UiLanguage.Text("Skipping this copy keeps the original-settings backup."), innerWidth, padding, ref currentY, _mutedText);

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
            @"(?<=(?:^|[\s(\u00AB""]))([вВ][оО]?|[нН][аА]|[сС][оО]?|[кК]|[пП][оО]|[дД][лЛ][яЯ]|[нН][еЕ]|[иИ][лЛ][иИ]|[иИ]|[оО][тТ]|[дД][оО]|[зЗ][аА]|[иИ][зЗ]|[оО][бБ]?)[ \t]+",
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
