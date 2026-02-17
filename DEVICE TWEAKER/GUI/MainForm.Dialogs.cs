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

    private bool ShowThemedConfirm(string message, string title)
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
            Text = "YES",
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
            Text = "NO",
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
