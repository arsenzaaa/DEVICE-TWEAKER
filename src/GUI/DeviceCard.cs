namespace DeviceTweakerCS;

internal sealed record DeviceCardStyle(float Scale, Font BodyFont, Font TitleFont,
    Color Background, Color Foreground, Color Muted, Color Border);

internal sealed record CardColumns(bool Stacked, int AffinityWidth, int SettingsLeft, int SettingsWidth)
{
    internal static CardColumns Measure(int width, float scale, int requiredAffinityWidth, int desiredSettingsWidth)
    {
        int S(int value) => (int)Math.Round(value * scale);
        int minimum = S(320);
        int fullMaximum = Math.Max(minimum, width - S(16) - S(24));
        int settingsMinimum = Math.Min(desiredSettingsWidth, Math.Max(minimum, width - S(48)));
        int sideMaximum = Math.Max(minimum, width - S(16) - S(20) - settingsMinimum - S(24));
        int needed = Math.Max(minimum, requiredAffinityWidth + S(12));
        bool stacked = needed > sideMaximum + S(4);
        int affinityWidth = Math.Min(stacked ? fullMaximum : sideMaximum, needed);
        int settingsLeft = stacked ? S(16) : S(16) + affinityWidth + S(20);
        int settingsWidth = Math.Max(S(260), stacked ? fullMaximum : width - settingsLeft - S(24));
        return new(stacked, affinityWidth, settingsLeft, settingsWidth);
    }
}

// Owns card chrome and the common section layout. The form supplies editors and
// binds operations; this control has no registry, discovery or driver dependency.
internal sealed class DeviceCard : UserControl
{
    private readonly DeviceCardStyle _style;
    private readonly Label _kind;
    internal FlowLayoutPanel Header { get; }
    internal Label TitleLabel { get; }
    internal Label ModifiedBadge { get; }
    internal Panel Divider { get; }

    internal DeviceCard(DeviceCardStyle style, string title, string kind, Color kindColor)
    {
        _style = style;
        AutoScaleMode = AutoScaleMode.None;
        AutoScroll = false;
        TabStop = false;
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer
            | ControlStyles.ResizeRedraw | ControlStyles.UserPaint, true);
        BackColor = style.Background;
        ForeColor = style.Foreground;
        Font = style.BodyFont;
        Margin = Padding.Empty;
        Padding = new Padding(S(12), S(16), S(12), S(16));
        Header = new()
        {
            Size = new Size(Math.Max(1, Width - S(40)), S(24)), Location = new Point(S(18), S(8)),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            WrapContents = false, Margin = Padding.Empty, Padding = Padding.Empty, BackColor = Color.Transparent,
        };
        _kind = new Label
        {
            Text = kind, Font = style.TitleFont, ForeColor = kindColor, AutoSize = true,
            Margin = new Padding(0, 0, S(6), 0),
        };
        TitleLabel = new HighlightLabel
        {
            Text = title, Font = style.TitleFont, ForeColor = style.Foreground,
            AutoSize = false, AutoEllipsis = true, UseMnemonic = false, Margin = Padding.Empty,
            HighlightText = "Mouse scanning", HighlightColor = Color.FromArgb(120, 120, 120),
        };
        ModifiedBadge = new Label
        {
            Text = UiLanguage.Text("Changed"), Font = style.BodyFont, ForeColor = style.Muted,
            AutoSize = true, Visible = false, Margin = new Padding(S(12), S(1), 0, 0),
        };
        Header.Controls.AddRange([_kind, TitleLabel, ModifiedBadge]);
        Divider = new()
        {
            BackColor = style.Border, Size = new Size(Math.Max(1, Width - S(32)), S(1)),
            Location = new Point(S(16), S(36)), Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
        };
        Controls.AddRange([Header, Divider]);
        TitleLabel.TextChanged += (_, _) => ArrangeHeader();
        TitleLabel.FontChanged += (_, _) => ArrangeHeader();
        ModifiedBadge.TextChanged += (_, _) => ArrangeHeader();
        ModifiedBadge.VisibleChanged += (_, _) => ArrangeHeader();
        Header.SizeChanged += (_, _) => ArrangeHeader();
        ArrangeHeader();
    }

    internal float LayoutScale => _style.Scale;

    private int S(int value) => (int)Math.Round(value * _style.Scale);

    private void ArrangeHeader()
    {
        Size desired = TitleLabel.GetPreferredSize(Size.Empty);
        int badgeWidth = ModifiedBadge.GetPreferredSize(Size.Empty).Width + ModifiedBadge.Margin.Horizontal;
        int kindWidth = _kind.GetPreferredSize(Size.Empty).Width + _kind.Margin.Horizontal;
        int available = Math.Max(1, Header.ClientSize.Width - kindWidth - badgeWidth);
        TitleLabel.Size = new Size(Math.Min(desired.Width, available), desired.Height);
    }

    internal void ComposeBody(IEnumerable<Control> sections) => Controls.AddRange(sections.ToArray());

    internal CardColumns ArrangeAffinity(BufferedPanel affinity, int requiredWidth,
        int contentRight, int contentHeight, int desiredSettingsWidth)
    {
        CardColumns columns = CardColumns.Measure(Width, _style.Scale, requiredWidth, desiredSettingsWidth);
        affinity.Size = new Size(columns.AffinityWidth, contentHeight);
        bool scroll = contentRight > affinity.ClientSize.Width + S(2);
        affinity.AutoScroll = scroll;
        if (scroll)
            affinity.AutoScrollMinSize = new Size(requiredWidth + S(12), Math.Max(contentHeight, affinity.ClientSize.Height));
        else
        {
            affinity.AutoScrollMinSize = Size.Empty;
            affinity.AutoScrollPosition = Point.Empty;
            affinity.HorizontalScroll.Visible = false;
            affinity.VerticalScroll.Visible = false;
        }
        return columns;
    }

    internal void ArrangeDiagnostics(Control info, Control divider, int lowerContentBottom)
    {
        int width = Math.Max(S(140), Width - info.Left - S(24));
        int preferred = string.IsNullOrWhiteSpace(info.Text) ? info.PreferredSize.Height
            : TextRenderer.MeasureText(info.Text, info.Font, new Size(Math.Max(1, width), int.MaxValue), TextFormatFlags.WordBreak).Height;
        info.Bounds = new Rectangle(info.Left, lowerContentBottom + S(18), width, Math.Max(S(70), preferred + S(20)));
        divider.Bounds = new Rectangle(info.Left, info.Top - S(9), width, divider.Height);
        Height = info.Bottom + S(20);
    }

    protected override void SetBoundsCore(int x, int y, int width, int height, BoundsSpecified specified)
    {
        Control? parent = Parent;
        Rectangle old = Bounds;
        base.SetBoundsCore(x, y, width, height, specified);
        if (parent is not null && !old.IsEmpty) parent.Invalidate(Rectangle.Union(old, Bounds), false);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        base.OnPaint(e);
        Rectangle bounds = ClientRectangle;
        if (bounds.Width <= 1 || bounds.Height <= 1) return;
        bounds.Width--; bounds.Height--;
        using Pen border = new(_style.Border);
        e.Graphics.DrawRectangle(border, bounds);
    }
}
