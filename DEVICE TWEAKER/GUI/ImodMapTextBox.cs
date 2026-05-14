namespace DeviceTweakerCS;

internal sealed class ImodMapTextBox : RichTextBox
{
    private bool _formatting;

    public Color PrefixColor { get; set; } = Color.FromArgb(172, 180, 190);
    public Color RoleColor { get; set; } = Color.FromArgb(208, 230, 250);
    public Color ValueColor { get; set; } = Color.FromArgb(240, 240, 240);

    public ImodMapTextBox()
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

        if (StartsWithAt(text, first, lineEnd, "devices:"))
        {
            ApplyColor(first, "devices:".Length, PrefixColor);
        }
        else if (StartsWithAt(text, first, lineEnd, "interrupters"))
        {
            int colon = text.IndexOf(':', first, lineEnd - first);
            if (colon >= first)
            {
                ApplyColor(first, colon - first + 1, PrefixColor);
            }
        }

        ColorRoleLabels(text, lineStart, lineEnd);
        ColorIntrValues(text, lineStart, lineEnd);
    }

    private void ColorRoleLabels(string text, int lineStart, int lineEnd)
    {
        int segmentStart = lineStart;
        while (segmentStart < lineEnd)
        {
            int segmentEnd = text.IndexOf('|', segmentStart, lineEnd - segmentStart);
            if (segmentEnd < 0)
            {
                segmentEnd = lineEnd;
            }

            int arrow = IndexOf(text, "->", segmentStart, segmentEnd);
            if (arrow > segmentStart)
            {
                int labelStart = segmentStart;
                while (labelStart < arrow && char.IsWhiteSpace(text[labelStart]))
                {
                    labelStart++;
                }

                if (StartsWithAt(text, labelStart, arrow, "devices:"))
                {
                    labelStart += "devices:".Length;
                    while (labelStart < arrow && char.IsWhiteSpace(text[labelStart]))
                    {
                        labelStart++;
                    }
                }

                int labelEnd = arrow;
                while (labelEnd > labelStart && char.IsWhiteSpace(text[labelEnd - 1]))
                {
                    labelEnd--;
                }

                if (labelEnd > labelStart)
                {
                    ApplyColor(labelStart, labelEnd - labelStart, RoleColor);
                }
            }

            segmentStart = segmentEnd + 1;
        }
    }

    private void ColorIntrValues(string text, int lineStart, int lineEnd)
    {
        int pos = lineStart;
        while (pos < lineEnd)
        {
            int intr = IndexOf(text, "intr", pos, lineEnd);
            if (intr < 0)
            {
                break;
            }
            if (!IsIntrValueToken(text, intr, lineEnd))
            {
                pos = intr + 4;
                continue;
            }

            int end = intr;
            while (end < lineEnd && text[end] != '|')
            {
                end++;
            }
            while (end > intr && char.IsWhiteSpace(text[end - 1]))
            {
                end--;
            }

            ApplyColor(intr, end - intr, ValueColor);
            pos = end;
        }
    }

    private static bool IsIntrValueToken(string text, int intrStart, int lineEnd)
    {
        int markerEnd = intrStart + 4;
        return markerEnd < lineEnd && (char.IsDigit(text[markerEnd]) || text[markerEnd] == '?');
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

    private static int IndexOf(string text, string value, int start, int end)
    {
        int length = Math.Max(0, end - start);
        int index = text.IndexOf(value, start, length, StringComparison.OrdinalIgnoreCase);
        return index >= end ? -1 : index;
    }
}
