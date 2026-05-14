namespace DeviceTweakerCS;

internal sealed class HighlightLabel : Label
{
    public string HighlightText { get; set; } = string.Empty;
    public Color HighlightColor { get; set; } = Color.Gray;

    protected override void OnPaint(PaintEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(HighlightText))
        {
            base.OnPaint(e);
            return;
        }

        int index = Text.IndexOf(HighlightText, StringComparison.OrdinalIgnoreCase);
        if (index < 0)
        {
            base.OnPaint(e);
            return;
        }

        OnPaintBackground(e);

        TextFormatFlags flags = TextFormatFlags.NoPadding
            | TextFormatFlags.NoPrefix
            | TextFormatFlags.SingleLine
            | TextFormatFlags.VerticalCenter;

        string before = Text[..index];
        string highlight = Text.Substring(index, HighlightText.Length);
        string after = Text[(index + HighlightText.Length)..];

        int x = 0;
        DrawTextPart(e.Graphics, before, ForeColor, ref x, flags);
        DrawTextPart(e.Graphics, highlight, HighlightColor, ref x, flags);
        DrawTextPart(e.Graphics, after, ForeColor, ref x, flags);
    }

    private void DrawTextPart(Graphics graphics, string text, Color color, ref int x, TextFormatFlags flags)
    {
        if (string.IsNullOrEmpty(text))
        {
            return;
        }

        Rectangle bounds = new(x, 0, Width - x, Height);
        TextRenderer.DrawText(graphics, text, Font, bounds, color, BackColor, flags);
        Size size = TextRenderer.MeasureText(graphics, text, Font, Size.Empty, flags);
        x += size.Width;
    }
}
