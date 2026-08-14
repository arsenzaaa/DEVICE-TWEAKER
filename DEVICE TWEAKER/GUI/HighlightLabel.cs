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

        TextFormatFlags flags = TextFormatFlags.NoPrefix
            | TextFormatFlags.NoPadding
            | TextFormatFlags.SingleLine
            | TextFormatFlags.VerticalCenter
            | TextFormatFlags.PreserveGraphicsClipping
            | TextFormatFlags.EndEllipsis;

        // One draw per region — never overpaint the same glyphs twice.
        // Double TextRenderer passes with ClearType cause cyan/red fringing
        // on dark OLED backgrounds (visible on "Mouse scanning" headers).
        string prefix = Text[..index];
        int highlightEnd = index + HighlightText.Length;
        int highlightLeft = string.IsNullOrEmpty(prefix)
            ? 0
            : TextRenderer.MeasureText(e.Graphics, prefix, Font, Size.Empty, flags & ~TextFormatFlags.EndEllipsis).Width;
        int highlightWidth = TextRenderer.MeasureText(
            e.Graphics,
            Text.Substring(index, HighlightText.Length),
            Font,
            Size.Empty,
            flags & ~TextFormatFlags.EndEllipsis).Width;
        if (highlightWidth <= 0)
        {
            TextRenderer.DrawText(e.Graphics, Text, Font, ClientRectangle, ForeColor, flags);
            return;
        }

        Rectangle bounds = ClientRectangle;
        System.Drawing.Drawing2D.GraphicsState state = e.Graphics.Save();
        try
        {
            if (highlightLeft > 0)
            {
                e.Graphics.SetClip(new Rectangle(0, 0, Math.Min(highlightLeft, Width), Height));
                TextRenderer.DrawText(e.Graphics, Text, Font, bounds, ForeColor, flags);
            }

            int clipHighlightWidth = Math.Max(0, Math.Min(highlightWidth, Width - highlightLeft));
            if (clipHighlightWidth > 0)
            {
                e.Graphics.SetClip(new Rectangle(highlightLeft, 0, clipHighlightWidth, Height));
                TextRenderer.DrawText(e.Graphics, Text, Font, bounds, HighlightColor, flags);
            }

            int right = highlightLeft + highlightWidth;
            if (right < Width && highlightEnd < Text.Length)
            {
                e.Graphics.SetClip(new Rectangle(right, 0, Width - right, Height));
                TextRenderer.DrawText(e.Graphics, Text, Font, bounds, ForeColor, flags);
            }
        }
        finally
        {
            e.Graphics.Restore(state);
        }
    }
}
