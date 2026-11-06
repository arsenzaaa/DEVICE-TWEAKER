using System.ComponentModel;
namespace DeviceTweakerCS;

// Shared flat-button presentation. Logical clicks never wait for visual feedback.
internal sealed class ThemedButton : Button
{
    private ControlAnimation? _appearance;
    private bool _painted, _hovered, _pressed;
    private Color _fromBack, _toBack, _fromFore, _toFore, _fromBorder, _toBorder;
    internal bool IsAnimating => _appearance?.IsRunning == true;

    [DefaultValue(typeof(Color), "125, 125, 132")]
    public Color DisabledForeColor { get; set; } = Color.FromArgb(125, 125, 132);

    protected override void OnMouseEnter(EventArgs e)
    { _hovered = true; base.OnMouseEnter(e); Invalidate(); }
    protected override void OnMouseLeave(EventArgs e)
    { _hovered = false; _pressed = false; base.OnMouseLeave(e); Invalidate(); }
    protected override void OnMouseDown(MouseEventArgs e)
    { if (e.Button == MouseButtons.Left) _pressed = true; base.OnMouseDown(e); Invalidate(); }
    protected override void OnMouseUp(MouseEventArgs e)
    { _pressed = false; base.OnMouseUp(e); Invalidate(); }
    protected override void OnMouseCaptureChanged(EventArgs e)
    { if (!Capture) _pressed = false; base.OnMouseCaptureChanged(e); Invalidate(); }
    protected override void OnKeyDown(KeyEventArgs e)
    { if (e.KeyCode == Keys.Space) _pressed = true; base.OnKeyDown(e); Invalidate(); }
    protected override void OnKeyUp(KeyEventArgs e)
    { _pressed = false; base.OnKeyUp(e); Invalidate(); }
    protected override void OnEnabledChanged(EventArgs e)
    {
        _appearance?.Complete(); _painted = false; _pressed = false;
        base.OnEnabledChanged(e); Invalidate();
    }
    protected override void OnVisibleChanged(EventArgs e)
    {
        if (!Visible) { _hovered = false; _pressed = false; _painted = false; }
        base.OnVisibleChanged(e);
    }
    protected override void Dispose(bool disposing)
    { if (disposing) _appearance?.Dispose(); base.Dispose(disposing); }

    protected override void OnPaint(PaintEventArgs pevent)
    {
        // Keep native support for image-bearing/non-flat buttons.
        if (Image is not null || BackgroundImage is not null || FlatStyle != FlatStyle.Flat)
        { base.OnPaint(pevent); return; }

        Color back = BackColor, fore = Enabled ? ForeColor : DisabledForeColor;
        Color borderColor = Enabled ? FlatAppearance.BorderColor : Color.FromArgb(82, 82, 90);
        if (borderColor.IsEmpty) borderColor = ForeColor;
        if (Enabled && (_hovered || _pressed))
        {
            Color configured = _pressed ? FlatAppearance.MouseDownBackColor : FlatAppearance.MouseOverBackColor;
            back = configured.IsEmpty ? UiMotion.Mix(back, fore, _pressed ? .13f : .06f) : configured;
        }
        float progress = _appearance?.Position ?? 1;
        Color shownBack = UiMotion.Mix(_fromBack, _toBack, progress);
        Color shownFore = UiMotion.Mix(_fromFore, _toFore, progress);
        Color shownBorder = UiMotion.Mix(_fromBorder, _toBorder, progress);
        if (!_painted || !Enabled)
        {
            _fromBack = _toBack = back; _fromFore = _toFore = fore;
            _fromBorder = _toBorder = borderColor; _painted = true;
        }
        else if (back != _toBack || fore != _toFore || borderColor != _toBorder)
        {
            _fromBack = shownBack; _fromFore = shownFore; _fromBorder = shownBorder;
            _toBack = back; _toFore = fore; _toBorder = borderColor;
            _appearance ??= new(this, 100, _ => Invalidate());
            _appearance.MoveTo(0, animate: false);
            _appearance.MoveTo(1);
        }
        progress = _appearance?.Position ?? 1;
        Graphics graphics = pevent.Graphics;
        graphics.Clear(UiMotion.Mix(_fromBack, _toBack, progress));
        Rectangle border = ClientRectangle;
        border.Width = Math.Max(0, border.Width - 1); border.Height = Math.Max(0, border.Height - 1);
        if (FlatAppearance.BorderSize > 0)
        {
            using Pen pen = new(UiMotion.Mix(_fromBorder, _toBorder, progress), FlatAppearance.BorderSize);
            graphics.DrawRectangle(pen, border);
        }
        Rectangle textBounds = Rectangle.FromLTRB(Padding.Left + 4, Padding.Top + 2,
            Width - Padding.Right - 4, Height - Padding.Bottom - 2);
        TextFormatFlags flags = TextFormatFlags.SingleLine | TextFormatFlags.EndEllipsis;
        flags |= TextAlign is ContentAlignment.TopLeft or ContentAlignment.MiddleLeft or ContentAlignment.BottomLeft
            ? TextFormatFlags.Left : TextAlign is ContentAlignment.TopRight or ContentAlignment.MiddleRight or ContentAlignment.BottomRight
                ? TextFormatFlags.Right : TextFormatFlags.HorizontalCenter;
        flags |= TextAlign is ContentAlignment.TopLeft or ContentAlignment.TopCenter or ContentAlignment.TopRight
            ? TextFormatFlags.Top : TextAlign is ContentAlignment.BottomLeft or ContentAlignment.BottomCenter or ContentAlignment.BottomRight
                ? TextFormatFlags.Bottom : TextFormatFlags.VerticalCenter;
        if (!UseMnemonic) flags |= TextFormatFlags.NoPrefix;
        else if (!ShowKeyboardCues) flags |= TextFormatFlags.HidePrefix;
        TextRenderer.DrawText(graphics, Text, Font, textBounds, UiMotion.Mix(_fromFore, _toFore, progress), flags);
        if (Focused && ShowFocusCues && Enabled)
            ControlPaint.DrawFocusRectangle(graphics, Rectangle.Inflate(ClientRectangle, -3, -3), ForeColor, BackColor);
    }
}
