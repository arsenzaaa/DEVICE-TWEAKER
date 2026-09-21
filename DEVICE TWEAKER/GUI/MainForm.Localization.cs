using System.Runtime.CompilerServices;

namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private sealed class LanguageSelectorButton : Button
    {
        private Color _borderColor;

        public LanguageSelectorButton()
        {
            SetStyle(
                ControlStyles.UserPaint
                | ControlStyles.AllPaintingInWmPaint
                | ControlStyles.OptimizedDoubleBuffer
                | ControlStyles.ResizeRedraw,
                true);
            FlatStyle = FlatStyle.Flat;
            FlatAppearance.BorderSize = 0;
            UseVisualStyleBackColor = false;
        }

        public Color BorderColor
        {
            get => _borderColor;
            set
            {
                if (_borderColor == value)
                {
                    return;
                }

                _borderColor = value;
                Invalidate();
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.Clear(BackColor);

            Rectangle border = ClientRectangle;
            border.Width = Math.Max(0, border.Width - 1);
            border.Height = Math.Max(0, border.Height - 1);
            if (border.Width > 0 && border.Height > 0)
            {
                using Pen pen = new(BorderColor);
                e.Graphics.DrawRectangle(pen, border);
            }

            TextRenderer.DrawText(
                e.Graphics,
                Text,
                Font,
                ClientRectangle,
                ForeColor,
                TextFormatFlags.HorizontalCenter
                | TextFormatFlags.VerticalCenter
                | TextFormatFlags.SingleLine
                | TextFormatFlags.NoPrefix);

            if (Focused && ShowFocusCues)
            {
                Rectangle focus = Rectangle.Inflate(ClientRectangle, -3, -3);
                ControlPaint.DrawFocusRectangle(e.Graphics, focus, ForeColor, BackColor);
            }
        }
    }

    private sealed class LocalizedControlState(string sourceText)
    {
        public string SourceText { get; set; } = sourceText;
        public bool EventsAttached { get; set; }
    }

    private static readonly object LocalizationIgnoreTag = new();
    private readonly ConditionalWeakTable<Control, LocalizedControlState> _localizedControls = new();
    private bool _applyingLocalization;
    private Button? _englishLanguageButton;
    private Button? _russianLanguageButton;

    private void AddLanguageSelector(Panel brandPanel)
    {
        Size languagePanelSize = UiScale(91, 28);
        Panel languagePanel = new()
        {
            AutoSize = false,
            Size = languagePanelSize,
            MinimumSize = languagePanelSize,
            MaximumSize = languagePanelSize,
            BackColor = _bgPanel,
            Tag = LocalizationIgnoreTag,
            TabStop = false,
            Anchor = AnchorStyles.Top | AnchorStyles.Right,
        };

        Size languageButtonSize = UiScale(43, 28);
        Button NewLanguageButton(string name, string text, int left)
        {
            LanguageSelectorButton button = new()
            {
                Name = name,
                Text = text,
                AutoSize = false,
                Location = new Point(left, 0),
                Size = languageButtonSize,
                MinimumSize = languageButtonSize,
                MaximumSize = languageButtonSize,
                BackColor = _bgPanel,
                ForeColor = _mutedText,
                Font = _baseFont,
                Cursor = Cursors.Hand,
                TabStop = true,
                Tag = LocalizationIgnoreTag,
                UseMnemonic = false,
            };
            return button;
        }

        _englishLanguageButton = NewLanguageButton("LANGUAGE_EN", "EN", 0);
        _russianLanguageButton = NewLanguageButton("LANGUAGE_RU", "RU", UiScale(48));
        _englishLanguageButton.Click += (_, _) => SetUiLanguage(UiLanguageCode.English);
        _russianLanguageButton.Click += (_, _) => SetUiLanguage(UiLanguageCode.Russian);
        languagePanel.Controls.Add(_englishLanguageButton);
        languagePanel.Controls.Add(_russianLanguageButton);

        void PositionLanguagePanel()
        {
            languagePanel.Location = new Point(
                Math.Max(UiScale(8), brandPanel.ClientSize.Width - languagePanel.Width - UiScale(18)),
                UiScale(14));
        }

        brandPanel.Controls.Add(languagePanel);
        languagePanel.BringToFront();
        PositionLanguagePanel();
    }

    private void InitializeLocalization()
    {
        UiLanguage.Initialize();
        UiLanguage.Changed += HandleUiLanguageChanged;
        LocalizeControlTree(this);
        UpdateLanguageSelectorStyle();
    }

    private void SetUiLanguage(UiLanguageCode language)
    {
        if (UiLanguage.Current == language)
        {
            return;
        }

        WriteLog($"UI.LANGUAGE: {UiLanguage.Current} -> {language}");
        UiLanguage.Set(language);
    }

    private void HandleUiLanguageChanged(object? sender, EventArgs e)
    {
        if (IsDisposed)
        {
            return;
        }

        if (InvokeRequired)
        {
            BeginInvoke(new Action(() => HandleUiLanguageChanged(sender, e)));
            return;
        }

        LocalizeControlTree(this);
        UpdateLanguageSelectorStyle();
        UpdateCpuHeaderUi();
        UpdateFilterToolbarLocalization();
        UpdateApplyButtonDirtyCount();
        foreach (DeviceBlock block in _blocks)
        {
            if (block.ModifiedBadge is not null)
            {
                block.ModifiedBadge.Text = UiLanguage.Text("[ MODIFIED ]");
            }
            if (block.Kind == DeviceKind.STOR)
            {
                _copyToolTip.SetToolTip(block.AffinityLabel, "NVMe/SATA storage uses Windows Multi-Queue steering. Pinned core affinity is intentionally disabled to ensure maximum SSD speed and low latency.");
            }
            else if (block.Kind == DeviceKind.AUDIO && (IsDisplayHdmiaudio(block.Device.InstanceId, block.Device.Name) || IsDisplayAudioEndpointsText(block.Device.AudioEndpoints)))
            {
                _copyToolTip.SetToolTip(block.AffinityLabel, "Display/HDMI audio shares PCIe bus with the GPU and uses Windows default interrupt steering (0x0).");
            }
            else
            {
                _copyToolTip.SetToolTip(block.AffinityLabel, "Interrupt Affinity Mask (AssignmentSetOverride). Strictly routes hardware interrupt service routines (ISRs) and deferred procedure calls (DPCs) to the selected CPU logical cores.");
            }
            if (block.InfoLabel is not null)
            {
                _copyToolTip.SetToolTip(block.InfoLabel, "Click to copy full registry path to clipboard");
            }
            block.RelayoutAction?.Invoke();
            ResetHorizontalView(block.LimitBox);
            ResetHorizontalView(block.ImodBox);
            ResetHorizontalView(block.NicItrBox);
            block.MsiCombo.Invalidate();
            block.PrioCombo.Invalidate();
            block.PolicyCombo.Invalidate();
            block.NdisModeCombo?.Invalidate();
            block.ImodModeCombo?.Invalidate();
            block.RawMouseThrottleCombo?.Invalidate();
        }

        PerformLayout();
        Invalidate(true);
        LogGuiSnapshot("language-change");
    }

    private static void ResetHorizontalView(TextBox? editor)
    {
        if (editor is null || editor.IsDisposed || editor.Focused || editor.TextLength == 0)
        {
            return;
        }

        editor.Select(0, 0);
        editor.ScrollToCaret();
    }

    private void UpdateLanguageSelectorStyle()
    {
        StyleLanguageButton(_englishLanguageButton, UiLanguage.Current == UiLanguageCode.English, "English");
        StyleLanguageButton(_russianLanguageButton, UiLanguage.Current == UiLanguageCode.Russian, "Русский");
    }

    private void StyleLanguageButton(Button? button, bool active, string accessibleName)
    {
        if (button is null)
        {
            return;
        }

        button.ForeColor = active ? _accent : _mutedText;
        Color borderColor = active ? _accent : _border;
        if (button is LanguageSelectorButton languageButton)
        {
            languageButton.BorderColor = borderColor;
        }
        else
        {
            button.FlatAppearance.BorderColor = borderColor;
            button.FlatAppearance.BorderSize = 1;
        }
        button.AccessibleName = accessibleName;
        button.AccessibleDescription = active
            ? (UiLanguage.IsRussian ? "Выбранный язык интерфейса" : "Selected interface language")
            : (UiLanguage.IsRussian ? "Переключить язык интерфейса" : "Switch interface language");
        button.Invalidate();
    }

    private void LocalizeControlTree(Control root)
    {
        if (root.IsDisposed)
        {
            return;
        }

        LocalizeControl(root);
        foreach (Control child in root.Controls)
        {
            LocalizeControlTree(child);
        }
    }

    private void LocalizeControl(Control control)
    {
        if (ReferenceEquals(control.Tag, LocalizationIgnoreTag))
        {
            return;
        }

        LocalizedControlState state = _localizedControls.GetValue(
            control,
            static item => new LocalizedControlState(item.Text ?? string.Empty));
        if (!state.EventsAttached)
        {
            state.EventsAttached = true;
            control.ControlAdded += HandleLocalizedControlAdded;
            if (ShouldTranslateControlText(control))
            {
                control.TextChanged += HandleLocalizedControlTextChanged;
            }
        }

        if (ShouldTranslateControlText(control))
        {
            SetLocalizedControlText(control, state.SourceText);
        }

        if (control is ComboBox or ThemedDropDownPicker)
        {
            control.Invalidate();
        }

        if (control is ImodMapTextBox imodMap)
        {
            imodMap.RefreshLocalizedAccessibility();
        }
        else if (control is NicItrTableLabel nicItrTable)
        {
            nicItrTable.RefreshLocalizedAccessibility();
        }
    }

    private static bool ShouldTranslateControlText(Control control)
        => control is Form or Label or ButtonBase or GroupBox or InfoTextBox;

    private string GetSourceControlText(Control control)
        => _localizedControls.TryGetValue(control, out LocalizedControlState? state)
            ? state.SourceText
            : control.Text ?? string.Empty;

    private void HandleLocalizedControlAdded(object? sender, ControlEventArgs e)
    {
        if (e.Control is Control control)
        {
            LocalizeControlTree(control);
        }
    }

    private void HandleLocalizedControlTextChanged(object? sender, EventArgs e)
    {
        if (_applyingLocalization || sender is not Control control || control.IsDisposed)
        {
            return;
        }

        LocalizedControlState state = _localizedControls.GetValue(
            control,
            static item => new LocalizedControlState(item.Text ?? string.Empty));
        state.SourceText = control.Text ?? string.Empty;
        SetLocalizedControlText(control, state.SourceText);
    }

    private void SetLocalizedControlText(Control control, string sourceText)
    {
        string localized = UiLanguage.Text(sourceText);
        if (control is Button && sourceText is ("SET" or "SAVE" or "CHECK" or "DELETE"))
        {
            int minimumWidth = sourceText switch
            {
                "SET" => UiScale(54),
                "SAVE" => UiScale(58),
                _ => UiScale(76),
            };
            int measuredWidth = TextRenderer.MeasureText(
                localized,
                control.Font,
                Size.Empty,
                TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix | TextFormatFlags.SingleLine).Width;
            // Leave real breathing room for Cyrillic glyph overhang and the
            // one-pixel focus/border inset used by the custom dark buttons.
            control.Width = Math.Max(minimumWidth, measuredWidth + UiScale(26));
        }
        else if (control is Button && string.Equals(control.FindForm()?.Name, "TEST_ADMIN_DIALOG", StringComparison.Ordinal))
        {
            int measuredWidth = TextRenderer.MeasureText(
                localized,
                control.Font,
                Size.Empty,
                TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix | TextFormatFlags.SingleLine).Width;
            control.Width = Math.Max(control.Width, measuredWidth + UiScale(26));
        }

        if (string.Equals(control.Text, localized, StringComparison.Ordinal))
        {
            return;
        }

        _applyingLocalization = true;
        try
        {
            control.Text = localized;
        }
        finally
        {
            _applyingLocalization = false;
        }
    }
}
