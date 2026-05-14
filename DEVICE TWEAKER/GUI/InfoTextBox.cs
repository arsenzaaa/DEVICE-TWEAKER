namespace DeviceTweakerCS;

internal sealed class InfoTextBox : RichTextBox
{
    private static readonly string[] KnownPrefixes =
    [
        "TEST DEVICE",
        "PNP ID:",
        "Class:",
        "Registry:",
        "HID:",
        "Polling:",
        "Net type:",
        "RSS:",
        "NIC ITR:",
        "Type:",
        "Audio endpoints:",
        "Raw input throttle:",
    ];

    private bool _formatting;

    public Color PrefixColor { get; set; } = Color.FromArgb(172, 180, 190);
    public Color ValueColor { get; set; } = Color.FromArgb(240, 240, 240);
    public Color SeparatorColor { get; set; } = Color.FromArgb(70, 70, 78);

    public InfoTextBox()
    {
        BorderStyle = BorderStyle.None;
        DetectUrls = false;
        HideSelection = true;
        Multiline = true;
        ReadOnly = true;
        ScrollBars = RichTextBoxScrollBars.None;
        ShortcutsEnabled = true;
        WordWrap = true;
        ZoomFactor = 1.0f;
    }

    protected override void OnTextChanged(EventArgs e)
    {
        base.OnTextChanged(e);
        ApplySyntaxColors();
    }

    protected override void OnForeColorChanged(EventArgs e)
    {
        base.OnForeColorChanged(e);
        ApplySyntaxColors();
    }

    public void ApplySyntaxColors()
    {
        if (_formatting || IsDisposed || TextLength == 0)
        {
            return;
        }

        _formatting = true;
        int selectionStart = SelectionStart;
        int selectionLength = SelectionLength;

        try
        {
            string text = Text;
            SelectAll();
            SelectionColor = ForeColor;

            int lineStart = 0;
            while (lineStart < text.Length)
            {
                int lineEnd = text.IndexOf('\n', lineStart);
                if (lineEnd < 0)
                {
                    lineEnd = text.Length;
                }

                int lineLength = lineEnd - lineStart;
                if (lineLength > 0 && text[lineStart + lineLength - 1] == '\r')
                {
                    lineLength--;
                }

                ColorLine(text, lineStart, lineLength);

                if (lineEnd >= text.Length)
                {
                    break;
                }

                lineStart = lineEnd + 1;
            }
        }
        finally
        {
            int safeStart = Math.Min(selectionStart, TextLength);
            int safeLength = Math.Min(selectionLength, Math.Max(0, TextLength - safeStart));
            Select(safeStart, safeLength);
            _formatting = false;
        }
    }

    private void ColorLine(string text, int lineStart, int lineLength)
    {
        if (lineLength <= 0)
        {
            return;
        }

        int lineEnd = lineStart + lineLength;
        int first = lineStart;
        while (first < lineEnd && char.IsWhiteSpace(text[first]))
        {
            first++;
        }

        foreach (string prefix in KnownPrefixes)
        {
            if (!StartsWithAt(text, first, lineEnd, prefix))
            {
                continue;
            }

            ApplyColor(first, prefix.Length, PrefixColor);

            int valueStart = first + prefix.Length;
            while (valueStart < lineEnd && char.IsWhiteSpace(text[valueStart]))
            {
                valueStart++;
            }

            if (valueStart < lineEnd)
            {
                ApplyColor(valueStart, lineEnd - valueStart, ValueColor);
                ColorSeparators(text, valueStart, lineEnd);
            }

            break;
        }
    }

    private void ColorSeparators(string text, int start, int end)
    {
        int pos = start;
        while (pos < end)
        {
            int separator = text.IndexOf('|', pos, end - pos);
            if (separator < 0)
            {
                break;
            }

            ApplyColor(separator, 1, SeparatorColor);
            pos = separator + 1;
        }
    }

    private void ApplyColor(int start, int length, Color color)
    {
        if (length <= 0 || start < 0 || start >= TextLength)
        {
            return;
        }

        Select(start, Math.Min(length, TextLength - start));
        SelectionColor = color;
    }

    private static bool StartsWithAt(string text, int start, int end, string value)
    {
        return start >= 0
            && start + value.Length <= end
            && string.Compare(text, start, value, 0, value.Length, StringComparison.OrdinalIgnoreCase) == 0;
    }
}
