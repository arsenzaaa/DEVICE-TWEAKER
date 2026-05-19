namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private void ShowThemedInfo(string message, string title)
    {
        using Form dialog = new();
        dialog.Text = title;
        dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
        dialog.StartPosition = FormStartPosition.CenterParent;
        dialog.MaximizeBox = false;
        dialog.MinimizeBox = false;
        dialog.ShowInTaskbar = false;
        dialog.AutoScaleMode = AutoScaleMode.None;
        dialog.BackColor = _bgForm;
        dialog.ForeColor = _fgMain;
        dialog.Font = _dialogFont;
        dialog.Icon = Icon;

        string normalized = NormalizeDialogMessage(message);

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
            TextAlign = ContentAlignment.MiddleCenter,
            Font = _dialogFont,
            UseCompatibleTextRendering = true,
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
        dialog.Shown += (_, _) => ApplyTitleBarTheme(dialog);

        dialog.ShowDialog(this);
    }

    private void ShowThemedInfo(string message)
    {
        ShowThemedInfo(message, "DEVICE TWEAKER");
    }

    private bool ShowThemedConfirm(string message, string title, string yesText = "YES", string noText = "NO")
    {
        using Form dialog = new();
        dialog.Text = title;
        dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
        dialog.StartPosition = FormStartPosition.CenterParent;
        dialog.MaximizeBox = false;
        dialog.MinimizeBox = false;
        dialog.ShowInTaskbar = false;
        dialog.AutoScaleMode = AutoScaleMode.None;
        dialog.BackColor = _bgForm;
        dialog.ForeColor = _fgMain;
        dialog.Font = _dialogFont;
        dialog.Icon = Icon;

        string normalized = NormalizeDialogMessage(message);

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
            UseCompatibleTextRendering = true,
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
        dialog.Shown += (_, _) => ApplyTitleBarTheme(dialog);

        return dialog.ShowDialog(this) == DialogResult.Yes;
    }

    private bool ShowThemedConfirm(string message)
    {
        return ShowThemedConfirm(message, "DEVICE TWEAKER");
    }

    private RestoreChoice ShowRestoreChoiceDialog(IReadOnlyList<BackupSnapshotInfo> backups, out string? selectedBackupPath)
    {
        selectedBackupPath = null;
        string? selectedPath = null;
        using Form dialog = new();
        dialog.Text = "RESTORE";
        dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
        dialog.StartPosition = FormStartPosition.CenterParent;
        dialog.MaximizeBox = false;
        dialog.MinimizeBox = false;
        dialog.ShowInTaskbar = false;
        dialog.AutoScaleMode = AutoScaleMode.None;
        dialog.BackColor = _bgForm;
        dialog.ForeColor = _fgMain;
        dialog.Font = _dialogFont;
        dialog.Icon = Icon;

        bool hasBackup = backups.Count > 0;
        string message = hasBackup
            ? "Choose restore mode.\n\n"
                + "RESET TO DEFAULT clears DEVICE TWEAKER changes and ignores backup files.\n"
                + "RESTORE BACKUP restores the selected snapshot; it may already contain applied tweaks."
            : "Choose restore mode.\n\n"
                + "RESET TO DEFAULT clears DEVICE TWEAKER changes.\n"
                + "No backup snapshot was found.";
        string normalized = NormalizeDialogMessage(message);

        int padding = UiScale(20);
        int maxTextWidth = UiScale(hasBackup ? 700 : 500);
        int minWidth = UiScale(hasBackup ? 860 : 500);
        int buttonWidth = UiScale(hasBackup ? 150 : 160);
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
            UseCompatibleTextRendering = true,
        };

        Size textSize = messageLabel.GetPreferredSize(new Size(maxTextWidth, 0));
        int buttonCount = hasBackup ? 5 : 2;
        int buttonRowWidth = (buttonWidth * buttonCount) + (buttonGap * (buttonCount - 1));
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
            };
        }

        RestoreChoice choice = RestoreChoice.Cancel;
        int buttonsTop = (backupList?.Bottom ?? messageLabel.Bottom) + buttonGap;
        int rowLeft = (clientWidth - buttonRowWidth) / 2;

        Button MakeButton(string text, RestoreChoice value, int left, bool enabled = true)
        {
            Button button = new()
            {
                Text = text,
                Size = new Size(buttonWidth, buttonHeight),
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

        Button resetButton = MakeButton("RESET DEFAULT", RestoreChoice.ResetDefault, rowLeft);
        Button? backupButton = hasBackup
            ? MakeButton("RESTORE BACKUP", RestoreChoice.RestoreBackup, rowLeft + buttonWidth + buttonGap)
            : null;
        Button? deleteButton = hasBackup
            ? MakeButton("DELETE SELECTED", RestoreChoice.DeleteBackup, rowLeft + ((buttonWidth + buttonGap) * 2))
            : null;
        Button? deleteAllButton = hasBackup
            ? MakeButton("DELETE ALL", RestoreChoice.DeleteAllBackups, rowLeft + ((buttonWidth + buttonGap) * 3))
            : null;
        int cancelLeft = hasBackup
            ? rowLeft + ((buttonWidth + buttonGap) * 4)
            : rowLeft + buttonWidth + buttonGap;
        Button cancelButton = MakeButton("CANCEL", RestoreChoice.Cancel, cancelLeft);

        dialog.Controls.Add(messageLabel);
        if (backupList is not null)
        {
            dialog.Controls.Add(backupList);
        }
        dialog.Controls.Add(resetButton);
        if (backupButton is not null)
        {
            dialog.Controls.Add(backupButton);
        }
        if (deleteButton is not null)
        {
            dialog.Controls.Add(deleteButton);
        }
        if (deleteAllButton is not null)
        {
            dialog.Controls.Add(deleteAllButton);
        }
        dialog.Controls.Add(cancelButton);

        dialog.AcceptButton = resetButton;
        dialog.CancelButton = cancelButton;
        dialog.Shown += (_, _) => ApplyTitleBarTheme(dialog);

        RestoreChoice result = dialog.ShowDialog(this) == DialogResult.OK ? choice : RestoreChoice.Cancel;
        selectedBackupPath = selectedPath;
        return result;
    }

    private AutoBackupChoice ShowAutoBackupChoiceDialog()
    {
        using Form dialog = new();
        dialog.Text = "AUTO BACKUP";
        dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
        dialog.StartPosition = FormStartPosition.CenterParent;
        dialog.MaximizeBox = false;
        dialog.MinimizeBox = false;
        dialog.ShowInTaskbar = false;
        dialog.AutoScaleMode = AutoScaleMode.None;
        dialog.BackColor = _bgForm;
        dialog.ForeColor = _fgMain;
        dialog.Font = _dialogFont;
        dialog.Icon = Icon;

        string message = NormalizeDialogMessage(
            "Where should DEVICE TWEAKER save the pre-auto backup?\n\n"
            + "EXE FOLDER = portable backup next to the app.\n"
            + "APPDATA = user profile backup that survives app folder cleanup.\n"
            + "SKIP = run auto-optimization without creating a backup.");

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
            UseCompatibleTextRendering = true,
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
        dialog.CancelButton = skipButton;
        dialog.Shown += (_, _) => ApplyTitleBarTheme(dialog);

        return dialog.ShowDialog(this) == DialogResult.OK ? choice : AutoBackupChoice.Skip;
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

    private void ApplyTitleBarTheme(Form form)
    {
        try
        {
            if (!OperatingSystem.IsWindowsVersionAtLeast(10, 0, 17763))
            {
                return;
            }

            int bg = ColorToColorRef(_bgForm);
            _ = DwmSetWindowAttribute(form.Handle, DwmwaCaptionColor, ref bg, sizeof(int));

            int fg = ColorToColorRef(_fgMain);
            _ = DwmSetWindowAttribute(form.Handle, DwmwaTextColor, ref fg, sizeof(int));
        }
        catch
        {
        }
    }
}
