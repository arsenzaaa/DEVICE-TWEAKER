using System.Text.RegularExpressions;

namespace DeviceTweakerCS;

/// <summary>
/// Draws IMOD readback in pixel-positioned columns. Text remains available for
/// logs and copying, but spaces in that text never control the visual layout.
/// </summary>
internal sealed partial class ImodMapTextBox : ScrollableControl
{
    private sealed record DeviceRow(string Name, string Irq, string Value, string Delay);
    private sealed record InterruptRow(string LeftIrq, string LeftValue, string LeftDelay, string RightIrq, string RightValue, string RightDelay);

    private readonly List<DeviceRow> _devices = [];
    private readonly List<InterruptRow> _interrupts = [];
    private string[] _fallbackLines = [];

    public Color PrefixColor { get; set; } = Color.FromArgb(172, 180, 190);
    public Color RoleColor { get; set; } = Color.FromArgb(208, 230, 250);
    public Color ValueColor { get; set; } = Color.FromArgb(240, 240, 240);

    public int DisplayRowCount
    {
        get
        {
            if (_devices.Count == 0 && _interrupts.Count == 0)
            {
                return Math.Max(1, _fallbackLines.Length);
            }

            return (_devices.Count > 0 ? _devices.Count + 1 : 0)
                + (_devices.Count > 0 && _interrupts.Count > 0 ? 1 : 0)
                + (_interrupts.Count > 0 ? _interrupts.Count + 1 : 0);
        }
    }

    public ImodMapTextBox()
    {
        AutoScroll = true;
        BackColor = Color.FromArgb(12, 12, 15);
        Cursor = Cursors.Hand;
        DoubleBuffered = true;
        Margin = Padding.Empty;
        TabStop = false;
    }

    protected override void OnTextChanged(EventArgs e)
    {
        base.OnTextChanged(e);
        ParseText();
        RefreshLocalizedAccessibility();
        UpdateScrollExtent();
        Invalidate();
    }

    internal void RefreshLocalizedAccessibility()
    {
        if (_devices.Count == 0 && _interrupts.Count == 0)
        {
            AccessibleName = string.Join(Environment.NewLine, _fallbackLines.Select(UiLanguage.Text));
            return;
        }

        List<string> lines = [];
        if (_devices.Count > 0)
        {
            lines.Add($"{UiLanguage.Text("DEVICE")}; IRQ; {UiLanguage.Text("VALUE")}; {UiLanguage.Text("DELAY")}");
            lines.AddRange(_devices.Select(row =>
                $"{UiLanguage.RoleText(row.Name)}; IRQ {row.Irq}; {row.Value}; {row.Delay}"));
        }

        lines.AddRange(_interrupts.SelectMany(row => new[]
        {
            $"IRQ {row.LeftIrq}; {row.LeftValue}; {row.LeftDelay}",
            string.IsNullOrWhiteSpace(row.RightIrq)
                ? string.Empty
                : $"IRQ {row.RightIrq}; {row.RightValue}; {row.RightDelay}",
        }).Where(line => line.Length > 0));
        AccessibleName = string.Join(Environment.NewLine, lines);
    }

    protected override void OnFontChanged(EventArgs e)
    {
        base.OnFontChanged(e);
        UpdateScrollExtent();
        Invalidate();
    }

    protected override void OnForeColorChanged(EventArgs e)
    {
        base.OnForeColorChanged(e);
        Invalidate();
    }

    protected override void OnResize(EventArgs e)
    {
        base.OnResize(e);
        UpdateScrollExtent();
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        base.OnPaint(e);
        e.Graphics.Clear(BackColor);

        int lineHeight = LineHeight;
        int y = AutoScrollPosition.Y;
        int width = Math.Max(1, ClientSize.Width - (VerticalScroll.Visible ? SystemInformation.VerticalScrollBarWidth : 0));
        TextFormatFlags flags = TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix | TextFormatFlags.SingleLine | TextFormatFlags.EndEllipsis;

        if (_devices.Count == 0 && _interrupts.Count == 0)
        {
            foreach (string line in _fallbackLines)
            {
                DrawCell(e.Graphics, UiLanguage.Text(line), 0, y, width, ForeColor, flags);
                y += lineHeight;
            }
            return;
        }

        (int nameX, int irqX, int valueX, int delayX) = DeviceColumns(width);
        if (_devices.Count > 0)
        {
            DrawCell(e.Graphics, UiLanguage.Text("DEVICE"), nameX, y, irqX - nameX - 8, PrefixColor, flags);
            DrawCell(e.Graphics, "IRQ", irqX, y, valueX - irqX - 8, PrefixColor, flags);
            DrawCell(e.Graphics, UiLanguage.Text("VALUE"), valueX, y, delayX - valueX - 8, PrefixColor, flags);
            DrawCell(e.Graphics, UiLanguage.Text("DELAY"), delayX, y, width - delayX, PrefixColor, flags);
            y += lineHeight;

            foreach (DeviceRow row in _devices)
            {
                DrawCell(e.Graphics, UiLanguage.RoleText(row.Name), nameX, y, irqX - nameX - 8, RoleColor, flags);
                DrawCell(e.Graphics, row.Irq, irqX, y, valueX - irqX - 8, ValueColor, flags);
                DrawCell(e.Graphics, row.Value, valueX, y, delayX - valueX - 8, ValueColor, flags);
                DrawCell(e.Graphics, row.Delay, delayX, y, width - delayX, ForeColor, flags);
                y += lineHeight;
            }
        }

        if (_devices.Count > 0 && _interrupts.Count > 0)
        {
            y += lineHeight;
        }

        if (_interrupts.Count > 0)
        {
            (int irq1X, int value1X, int delay1X, int irq2X, int value2X, int delay2X) = InterruptColumns(width);
            DrawCell(e.Graphics, "IRQ", irq1X, y, value1X - irq1X - 8, PrefixColor, flags);
            DrawCell(e.Graphics, UiLanguage.Text("VALUE"), value1X, y, delay1X - value1X - 8, PrefixColor, flags);
            DrawCell(e.Graphics, UiLanguage.Text("DELAY"), delay1X, y, irq2X - delay1X - 12, PrefixColor, flags);
            DrawCell(e.Graphics, "IRQ", irq2X, y, value2X - irq2X - 8, PrefixColor, flags);
            DrawCell(e.Graphics, UiLanguage.Text("VALUE"), value2X, y, delay2X - value2X - 8, PrefixColor, flags);
            DrawCell(e.Graphics, UiLanguage.Text("DELAY"), delay2X, y, width - delay2X, PrefixColor, flags);
            y += lineHeight;

            foreach (InterruptRow row in _interrupts)
            {
                DrawCell(e.Graphics, row.LeftIrq, irq1X, y, value1X - irq1X - 8, ValueColor, flags);
                DrawCell(e.Graphics, row.LeftValue, value1X, y, delay1X - value1X - 8, ValueColor, flags);
                DrawCell(e.Graphics, row.LeftDelay, delay1X, y, irq2X - delay1X - 12, ForeColor, flags);
                DrawCell(e.Graphics, row.RightIrq, irq2X, y, value2X - irq2X - 8, ValueColor, flags);
                DrawCell(e.Graphics, row.RightValue, value2X, y, delay2X - value2X - 8, ValueColor, flags);
                DrawCell(e.Graphics, row.RightDelay, delay2X, y, width - delay2X, ForeColor, flags);
                y += lineHeight;
            }
        }
    }

    internal bool ValidateColumnLayout(int width)
    {
        (int nameX, int irqX, int valueX, int delayX) = DeviceColumns(width);
        (int irq1X, int value1X, int delay1X, int irq2X, int value2X, int delay2X) = InterruptColumns(width);
        return nameX < irqX && irqX < valueX && valueX < delayX && delayX < width
            && irq1X < value1X && value1X < delay1X && delay1X < irq2X
            && irq2X < value2X && value2X < delay2X && delay2X < width;
    }

    private int LineHeight => Math.Max(Font?.Height ?? 16, 15);

    private void ParseText()
    {
        _devices.Clear();
        _interrupts.Clear();
        _fallbackLines = (Text ?? string.Empty).Split(["\r\n", "\n"], StringSplitOptions.None);

        foreach (string sourceLine in _fallbackLines)
        {
            string line = sourceLine.Trim();
            if (line.StartsWith("devices:", StringComparison.OrdinalIgnoreCase))
            {
                line = line["devices:".Length..].Trim();
            }

            Match device = DeviceLineRegex().Match(line);
            if (device.Success)
            {
                _devices.Add(new DeviceRow(
                    device.Groups["name"].Value.Trim(),
                    device.Groups["irq"].Value,
                    device.Groups["value"].Value,
                    device.Groups["delay"].Value.Trim()));
                continue;
            }

            if (!line.StartsWith("interrupters ", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            int colon = line.IndexOf(':');
            if (colon < 0)
            {
                continue;
            }

            string[] pair = line[(colon + 1)..].Split([" | "], 2, StringSplitOptions.None);
            if (TryParseInterrupt(pair[0], out (string Irq, string Value, string Delay) left))
            {
                (string Irq, string Value, string Delay) right = pair.Length == 2
                    && TryParseInterrupt(pair[1], out (string Irq, string Value, string Delay) parsedRight)
                        ? parsedRight
                        : (string.Empty, string.Empty, string.Empty);
                _interrupts.Add(new InterruptRow(left.Irq, left.Value, left.Delay, right.Irq, right.Value, right.Delay));
            }
        }
    }

    private static bool TryParseInterrupt(string text, out (string Irq, string Value, string Delay) row)
    {
        Match match = InterruptLineRegex().Match(text.Trim());
        if (!match.Success)
        {
            row = default;
            return false;
        }

        row = (match.Groups["irq"].Value, match.Groups["value"].Value, match.Groups["delay"].Value.Trim());
        return true;
    }

    private void UpdateScrollExtent()
    {
        AutoScrollMinSize = new Size(0, (DisplayRowCount * LineHeight) + 2);
    }

    private static (int Name, int Irq, int Value, int Delay) DeviceColumns(int width)
    {
        int irq = Math.Max(100, (int)(width * 0.30));
        int value = Math.Max(irq + 36, (int)(width * 0.44));
        int delay = Math.Min(width - 96, Math.Max(value + 72, (int)(width * 0.66)));
        return (0, irq, value, Math.Min(delay, Math.Max(value + 1, width - 1)));
    }

    private static (int Irq1, int Value1, int Delay1, int Irq2, int Value2, int Delay2) InterruptColumns(int width)
    {
        int colWidth = width / 2;
        int valueOffset = Math.Max(36, (int)(colWidth * 0.16));
        int delayOffset = Math.Max(valueOffset + 76, (int)(colWidth * 0.50));

        int irq1 = 0;
        int value1 = valueOffset;
        int delay1 = delayOffset;

        int irq2 = colWidth;
        int value2 = colWidth + valueOffset;
        int delay2 = colWidth + delayOffset;

        return (irq1, value1, delay1, irq2, value2, delay2);
    }

    private void DrawCell(Graphics graphics, string text, int x, int y, int width, Color color, TextFormatFlags flags)
    {
        if (!string.IsNullOrEmpty(text) && width > 0)
        {
            TextRenderer.DrawText(graphics, text, Font, new Rectangle(x, y, width, LineHeight), color, BackColor, flags);
        }
    }

    [GeneratedRegex(@"^(?<name>.+?)\s*->\s*intr(?<irq>\d+|\?)\s+(?<value>\S+)\s*(?<delay>.*)$", RegexOptions.IgnoreCase)]
    private static partial Regex DeviceLineRegex();

    [GeneratedRegex(@"^intr(?<irq>\d+|\?)\s+(?<value>\S+)\s+(?<delay>.+)$", RegexOptions.IgnoreCase)]
    private static partial Regex InterruptLineRegex();
}
