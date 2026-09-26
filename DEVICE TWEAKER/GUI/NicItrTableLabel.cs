using System.Text.RegularExpressions;

namespace DeviceTweakerCS;

/// <summary>Renders NIC queue readback with pixel-positioned columns.</summary>
internal sealed partial class NicItrTableLabel : Label
{
    private sealed record QueueRow(string Queue, string Value, string Rx, string Tx);
    private readonly List<QueueRow> _rows = [];

    public Color HeaderColor { get; set; } = Color.FromArgb(172, 180, 190);
    public Color ValueColor { get; set; } = Color.FromArgb(240, 240, 240);

    public NicItrTableLabel()
    {
        AutoSize = false;
        DoubleBuffered = true;
        UseMnemonic = false;
    }

    protected override void OnTextChanged(EventArgs e)
    {
        base.OnTextChanged(e);
        ParseRows();
        RefreshLocalizedAccessibility();
        Invalidate();
    }

    internal void RefreshLocalizedAccessibility()
    {
        if (_rows.Count == 0)
        {
            AccessibleName = UiLanguage.Text(Text ?? string.Empty);
            return;
        }

        IEnumerable<string> rows = _rows.Select(row =>
            $"{row.Queue}; {row.Value}; RX {UiLanguage.Text(row.Rx)}; TX {UiLanguage.Text(row.Tx)}");
        AccessibleName = $"{UiLanguage.Text("QUEUE")}; {UiLanguage.Text("VALUE")}; RX; TX"
            + Environment.NewLine
            + string.Join(Environment.NewLine, rows);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        e.Graphics.Clear(BackColor);
        if (_rows.Count == 0)
        {
            TextRenderer.DrawText(e.Graphics, Text, Font, ClientRectangle, ForeColor, BackColor,
                TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix | TextFormatFlags.WordBreak);
            return;
        }

        int lineHeight = Math.Max(Font.Height, 15);
        int width = Math.Max(1, ClientSize.Width);
        (int queueX, int valueX, int rxX, int txX) = Columns(width);
        TextFormatFlags flags = TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix | TextFormatFlags.SingleLine | TextFormatFlags.EndEllipsis;
        int y = 0;
        DrawCell(e.Graphics, UiLanguage.Text("QUEUE"), queueX, y, valueX - queueX - 8, HeaderColor, flags, lineHeight);
        DrawCell(e.Graphics, UiLanguage.Text("VALUE"), valueX, y, rxX - valueX - 8, HeaderColor, flags, lineHeight);
        DrawCell(e.Graphics, "RX", rxX, y, txX - rxX - 8, HeaderColor, flags, lineHeight);
        DrawCell(e.Graphics, "TX", txX, y, width - txX, HeaderColor, flags, lineHeight);
        y += lineHeight;

        foreach (QueueRow row in _rows)
        {
            DrawCell(e.Graphics, row.Queue, queueX, y, valueX - queueX - 8, ValueColor, flags, lineHeight);
            DrawCell(e.Graphics, row.Value, valueX, y, rxX - valueX - 8, ValueColor, flags, lineHeight);
            DrawCell(e.Graphics, row.Rx, rxX, y, txX - rxX - 8, ForeColor, flags, lineHeight);
            DrawCell(e.Graphics, row.Tx, txX, y, width - txX, ForeColor, flags, lineHeight);
            y += lineHeight;
        }
    }

    internal bool ValidateColumnLayout(int width)
    {
        (int queueX, int valueX, int rxX, int txX) = Columns(width);
        return queueX < valueX && valueX < rxX && rxX < txX && txX < width;
    }

    private void ParseRows()
    {
        _rows.Clear();
        foreach (string line in (Text ?? string.Empty).Split(["\r\n", "\n"], StringSplitOptions.RemoveEmptyEntries))
        {
            Match match = QueueLineRegex().Match(line.Trim());
            if (!match.Success)
            {
                continue;
            }

            string detail = match.Groups["detail"].Value.Trim();
            string[] directions = detail.Split([" | "], 2, StringSplitOptions.None);
            string rx;
            string tx;
            if (detail.Equals("Off", StringComparison.OrdinalIgnoreCase))
            {
                rx = "Off";
                tx = "Off";
            }
            else if (directions.Length == 2)
            {
                rx = RemoveDirectionPrefix(directions[0], "RX");
                tx = RemoveDirectionPrefix(directions[1], "TX");
            }
            else
            {
                rx = detail;
                tx = string.Empty;
            }

            _rows.Add(new QueueRow(match.Groups["queue"].Value, match.Groups["value"].Value, rx, tx));
        }
    }

    private static string RemoveDirectionPrefix(string value, string prefix)
    {
        string text = value.Trim();
        return text.StartsWith(prefix + " ", StringComparison.OrdinalIgnoreCase)
            ? text[(prefix.Length + 1)..]
            : text;
    }

    private static (int Queue, int Value, int Rx, int Tx) Columns(int width)
    {
        int value = Math.Max(72, (int)(width * 0.22));
        int rx = Math.Max(value + 96, (int)(width * 0.50));
        int tx = Math.Max(rx + 110, (int)(width * 0.76));
        return (0, value, rx, Math.Min(tx, Math.Max(rx + 1, width - 1)));
    }

    private void DrawCell(Graphics graphics, string text, int x, int y, int width, Color color, TextFormatFlags flags, int lineHeight)
    {
        if (!string.IsNullOrEmpty(text) && width > 0)
        {
            TextRenderer.DrawText(graphics, text, Font, new Rectangle(x, y, width, lineHeight), color, BackColor, flags);
        }
    }

    [GeneratedRegex(@"^(?<queue>[^:]+):\s*(?<value>\S+)\s*\|\s*(?<detail>.+)$")]
    private static partial Regex QueueLineRegex();
}
