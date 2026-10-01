namespace DeviceTweakerCS;

/// <summary>
/// Keeps disabled flat buttons readable in the dark theme. The stock WinForms
/// renderer substitutes system disabled colors and can produce black-on-black text.
/// </summary>
internal sealed class ThemedButton : Button
{
    public Color DisabledForeColor { get; set; } = Color.FromArgb(125, 125, 132);

    protected override void OnEnabledChanged(EventArgs e)
    {
        base.OnEnabledChanged(e);
        Invalidate();
    }

    protected override void OnPaint(PaintEventArgs pevent)
    {
        if (Enabled)
        {
            base.OnPaint(pevent);
            return;
        }

        Graphics graphics = pevent.Graphics;
        graphics.Clear(BackColor);

        Rectangle border = ClientRectangle;
        border.Width = Math.Max(0, border.Width - 1);
        border.Height = Math.Max(0, border.Height - 1);
        using Pen pen = new(Color.FromArgb(82, 82, 90));
        graphics.DrawRectangle(pen, border);

        Rectangle textBounds = Rectangle.Inflate(ClientRectangle, -4, -2);
        TextRenderer.DrawText(
            graphics,
            Text,
            Font,
            textBounds,
            DisabledForeColor,
            TextFormatFlags.HorizontalCenter
            | TextFormatFlags.VerticalCenter
            | TextFormatFlags.SingleLine
            | TextFormatFlags.EndEllipsis
            | TextFormatFlags.NoPrefix);
    }
}
