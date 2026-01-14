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

        int padding = UiScale(20);
        int maxTextWidth = UiScale(520);
        int minWidth = UiScale(360);
        int buttonWidth = UiScale(92);
        int buttonHeight = UiScale(32);
        int buttonGap = UiScale(16);

        Size textSize = TextRenderer.MeasureText(
            message,
            _dialogFont,
            new Size(maxTextWidth, int.MaxValue),
            TextFormatFlags.WordBreak | TextFormatFlags.TextBoxControl);

        int clientWidth = Math.Max(minWidth, textSize.Width + (padding * 2));
        int clientHeight = padding + textSize.Height + buttonGap + buttonHeight + padding;

        dialog.ClientSize = new Size(clientWidth, clientHeight);

        Label messageLabel = new()
        {
            AutoSize = false,
            Size = new Size(clientWidth - (padding * 2), textSize.Height),
            Location = new Point(padding, padding),
            Text = message,
            ForeColor = _fgMain,
            BackColor = _bgForm,
            UseMnemonic = false,
            TextAlign = ContentAlignment.MiddleCenter,
            Font = _dialogFont,
        };

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

        int padding = UiScale(20);
        int maxTextWidth = UiScale(520);
        int minWidth = UiScale(360);
        int buttonWidth = UiScale(120);
        int buttonHeight = UiScale(32);
        int buttonGap = UiScale(16);

        Size textSize = TextRenderer.MeasureText(
            message,
            _dialogFont,
            new Size(maxTextWidth, int.MaxValue),
            TextFormatFlags.WordBreak | TextFormatFlags.TextBoxControl);

        int buttonRowWidth = (buttonWidth * 2) + buttonGap;
        int clientWidth = Math.Max(minWidth, Math.Max(textSize.Width + (padding * 2), buttonRowWidth + (padding * 2)));
        int clientHeight = padding + textSize.Height + buttonGap + buttonHeight + padding;
        dialog.ClientSize = new Size(clientWidth, clientHeight);

        Label messageLabel = new()
        {
            AutoSize = false,
            Size = new Size(clientWidth - (padding * 2), textSize.Height),
            Location = new Point(padding, padding),
            Text = message,
            ForeColor = _fgMain,
            BackColor = _bgForm,
            UseMnemonic = false,
            TextAlign = ContentAlignment.MiddleCenter,
            Font = _dialogFont,
        };

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
