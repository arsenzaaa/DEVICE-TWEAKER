using System.Drawing.Drawing2D;

namespace DeviceTweakerCS;

internal sealed class ThemedCheckBox : CheckBox
{
    private bool _hovered;
    private bool _pressed;
    private bool _syncCheckedStateText;

    public Color BorderColor { get; set; } = Color.FromArgb(150, 150, 150);
    public Color BoxBackColor { get; set; } = Color.FromArgb(18, 18, 22);
    public Color CheckedBackColor { get; set; } = Color.FromArgb(28, 28, 34);
    public Color HoverBackColor { get; set; } = Color.FromArgb(36, 36, 42);
    public Color PressedBackColor { get; set; } = Color.FromArgb(48, 48, 56);
    public Color CheckColor { get; set; } = Color.FromArgb(240, 240, 240);

    /// <summary>
    /// When true, <see cref="Control.Text"/> follows Checked: Enabled / Disabled.
    /// </summary>
    public bool SyncCheckedStateText
    {
        get => _syncCheckedStateText;
        set
        {
            _syncCheckedStateText = value;
            if (value)
            {
                SyncTextFromChecked();
            }
        }
    }

    public ThemedCheckBox()
    {
        SetStyle(
            ControlStyles.UserPaint
            | ControlStyles.AllPaintingInWmPaint
            | ControlStyles.OptimizedDoubleBuffer
            | ControlStyles.ResizeRedraw,
            true);
        UseVisualStyleBackColor = false;
        FlatStyle = FlatStyle.Flat;
    }

    private void SyncTextFromChecked()
    {
        if (!_syncCheckedStateText)
        {
            return;
        }

        string next = Checked ? "Enabled" : "Disabled";
        if (!string.Equals(Text, next, StringComparison.Ordinal))
        {
            Text = next;
        }
    }

    protected override void OnMouseEnter(EventArgs e)
    {
        _hovered = true;
        Invalidate();
        base.OnMouseEnter(e);
    }

    protected override void OnMouseLeave(EventArgs e)
    {
        _hovered = false;
        _pressed = false;
        Invalidate();
        base.OnMouseLeave(e);
    }

    protected override void OnMouseDown(MouseEventArgs e)
    {
        _pressed = true;
        Invalidate();
        base.OnMouseDown(e);
    }

    protected override void OnMouseUp(MouseEventArgs e)
    {
        _pressed = false;
        Invalidate();
        base.OnMouseUp(e);
    }

    protected override void OnCheckedChanged(EventArgs e)
    {
        SyncTextFromChecked();
        Invalidate();
        base.OnCheckedChanged(e);
    }

    protected override void OnEnabledChanged(EventArgs e)
    {
        Invalidate();
        base.OnEnabledChanged(e);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        Graphics g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.PixelOffsetMode = PixelOffsetMode.HighQuality;
        g.CompositingQuality = CompositingQuality.HighQuality;
        g.Clear(BackColor);

        // Keep 1px inset so anti-aliased stroke/fringe is not clipped by the control edge
        // (clipping is what makes the pill/knob look stair-stepped).
        int trackWidth = Math.Max(28, Math.Min(34, Height + 12));
        int trackHeight = Math.Max(14, Math.Min(18, Height - 2));
        float trackTop = Math.Max(1f, (Height - trackHeight) / 2f);
        RectangleF trackRect = new(1f, trackTop, trackWidth - 2f, trackHeight);

        Color trackFill = Checked ? CheckedBackColor : BoxBackColor;
        if (_pressed)
        {
            trackFill = PressedBackColor;
        }
        else if (_hovered)
        {
            trackFill = HoverBackColor;
        }

        float radius = trackRect.Height / 2f;
        using GraphicsPath trackPath = CreateRoundRect(trackRect, radius);
        using SolidBrush trackBrush = new(trackFill);
        g.FillPath(trackBrush, trackPath);

        using Pen borderPen = new(Enabled ? BorderColor : Color.FromArgb(85, 85, 90), 1.25f);
        borderPen.Alignment = PenAlignment.Center;
        g.DrawPath(borderPen, trackPath);

        float knobSize = Math.Max(8f, trackRect.Height - 6f);
        float knobPad = Math.Max(2f, (trackRect.Height - knobSize) / 2f);
        float knobLeft = Checked
            ? trackRect.Right - knobSize - knobPad
            : trackRect.Left + knobPad;
        RectangleF knobRect = new(knobLeft, trackRect.Top + knobPad, knobSize, knobSize);
        using SolidBrush knobBrush = new(Enabled ? CheckColor : Color.FromArgb(120, 120, 125));
        g.FillEllipse(knobBrush, knobRect);

        int textLeft = (int)Math.Ceiling(trackRect.Right) + 7;
        Rectangle textRect = new(
            textLeft,
            0,
            Math.Max(0, Width - textLeft),
            Height);
        TextRenderer.DrawText(
            g,
            Text,
            Font,
            textRect,
            Enabled ? ForeColor : Color.FromArgb(120, 120, 125),
            BackColor,
            TextFormatFlags.Left
            | TextFormatFlags.VerticalCenter
            | TextFormatFlags.NoPrefix
            | TextFormatFlags.NoPadding
            | TextFormatFlags.EndEllipsis);
    }

    private static GraphicsPath CreateRoundRect(RectangleF rect, float radius)
    {
        float diameter = Math.Max(1f, radius * 2f);
        if (diameter > rect.Width)
        {
            diameter = rect.Width;
        }

        if (diameter > rect.Height)
        {
            diameter = rect.Height;
        }

        GraphicsPath path = new();
        path.AddArc(rect.Left, rect.Top, diameter, diameter, 180, 90);
        path.AddArc(rect.Right - diameter, rect.Top, diameter, diameter, 270, 90);
        path.AddArc(rect.Right - diameter, rect.Bottom - diameter, diameter, diameter, 0, 90);
        path.AddArc(rect.Left, rect.Bottom - diameter, diameter, diameter, 90, 90);
        path.CloseFigure();
        return path;
    }
}
