using System.Drawing.Drawing2D;

namespace DeviceTweakerCS;

internal sealed class ThemedCheckBox : CheckBox
{
    private bool _hovered;
    private bool _pressed;

    public Color BorderColor { get; set; } = Color.FromArgb(150, 150, 150);
    public Color BoxBackColor { get; set; } = Color.FromArgb(8, 8, 10);
    public Color CheckedBackColor { get; set; } = Color.FromArgb(14, 14, 17);
    public Color HoverBackColor { get; set; } = Color.FromArgb(22, 22, 26);
    public Color PressedBackColor { get; set; } = Color.FromArgb(32, 32, 38);
    public Color CheckColor { get; set; } = Color.FromArgb(240, 240, 240);

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
        g.Clear(BackColor);

        int trackWidth = Math.Max(26, Math.Min(32, Height + 10));
        int trackHeight = Math.Max(14, Math.Min(16, Height - 4));
        Rectangle trackRect = new(0, Math.Max(0, (Height - trackHeight) / 2), trackWidth, trackHeight);

        Color trackFill = Checked ? CheckedBackColor : BoxBackColor;
        if (_pressed)
        {
            trackFill = PressedBackColor;
        }
        else if (_hovered)
        {
            trackFill = HoverBackColor;
        }

        using GraphicsPath trackPath = CreateRoundRect(trackRect, trackHeight / 2);
        using SolidBrush trackBrush = new(trackFill);
        g.FillPath(trackBrush, trackPath);

        using Pen borderPen = new(Enabled ? BorderColor : Color.FromArgb(85, 85, 90));
        g.DrawPath(borderPen, trackPath);

        int knobSize = trackHeight - 6;
        int knobLeft = Checked
            ? trackRect.Right - knobSize - 3
            : trackRect.Left + 3;
        Rectangle knobRect = new(knobLeft, trackRect.Top + 3, knobSize, knobSize);
        using SolidBrush knobBrush = new(Enabled ? CheckColor : Color.FromArgb(120, 120, 125));
        g.FillEllipse(knobBrush, knobRect);

        Rectangle textRect = new(
            trackRect.Right + 7,
            0,
            Math.Max(0, Width - trackRect.Right - 7),
            Height);
        TextRenderer.DrawText(
            g,
            Text,
            Font,
            textRect,
            Enabled ? ForeColor : Color.FromArgb(120, 120, 125),
            BackColor,
            TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix | TextFormatFlags.EndEllipsis);
    }

    private static GraphicsPath CreateRoundRect(Rectangle rect, int radius)
    {
        int diameter = Math.Max(1, radius * 2);
        GraphicsPath path = new();
        path.AddArc(rect.Left, rect.Top, diameter, diameter, 180, 90);
        path.AddArc(rect.Right - diameter, rect.Top, diameter, diameter, 270, 90);
        path.AddArc(rect.Right - diameter, rect.Bottom - diameter, diameter, diameter, 0, 90);
        path.AddArc(rect.Left, rect.Bottom - diameter, diameter, diameter, 90, 90);
        path.CloseFigure();
        return path;
    }
}
