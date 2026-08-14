using System.Diagnostics.CodeAnalysis;

namespace DeviceTweakerCS;

/// <summary>
/// Dark single-line field: outer panel draws the border; an inner borderless TextBox
/// is padded and vertically centered. Avoids EM_SETRECT / WM_NCCALCSIZE hacks.
/// </summary>
internal sealed class ThemedTextBox : Panel
{
    private readonly TextBox _edit;
    private Color _borderColor = Color.FromArgb(120, 120, 128);

    public int ContentLeftPadding { get; set; } = 6;
    public int ContentRightPadding { get; set; } = 4;

    public TextBox Inner => _edit;

    public ThemedTextBox()
    {
        SetStyle(
            ControlStyles.AllPaintingInWmPaint
            | ControlStyles.OptimizedDoubleBuffer
            | ControlStyles.UserPaint
            | ControlStyles.ResizeRedraw,
            true);
        DoubleBuffered = true;
        TabStop = false;
        Padding = Padding.Empty;
        Margin = Padding.Empty;

        // Create child BEFORE setting colors - setters raise On*Changed which touch _edit.
        _edit = new TextBox
        {
            BorderStyle = BorderStyle.None,
            Margin = Padding.Empty,
            TabStop = true,
            AutoSize = false,
        };
        Controls.Add(_edit);

        BackColor = Color.FromArgb(18, 18, 22);
        ForeColor = Color.FromArgb(240, 240, 240);

        Click += (_, _) => _edit.Focus();
        _edit.KeyDown += (_, e) =>
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
            }
        };
    }

    [AllowNull]
    public override string Text
    {
        get => _edit.Text;
        set => _edit.Text = value ?? string.Empty;
    }

    public HorizontalAlignment TextAlign
    {
        get => _edit.TextAlign;
        set => _edit.TextAlign = value;
    }

    public new event EventHandler TextChanged
    {
        add => _edit.TextChanged += value;
        remove => _edit.TextChanged -= value;
    }

    public void ApplyContentLayout()
    {
        PerformLayout();
        Invalidate();
    }

    protected override void OnHandleCreated(EventArgs e)
    {
        base.OnHandleCreated(e);
        _edit.Font = Font;
        PerformLayout();
    }

    protected override void OnFontChanged(EventArgs e)
    {
        base.OnFontChanged(e);
        _edit.Font = Font;
        PerformLayout();
    }

    protected override void OnBackColorChanged(EventArgs e)
    {
        base.OnBackColorChanged(e);
        _edit.BackColor = BackColor;
        Invalidate();
    }

    protected override void OnForeColorChanged(EventArgs e)
    {
        base.OnForeColorChanged(e);
        _edit.ForeColor = ForeColor;
    }

    protected override void OnEnabledChanged(EventArgs e)
    {
        base.OnEnabledChanged(e);
        _edit.Enabled = Enabled;
    }

    protected override void OnVisibleChanged(EventArgs e)
    {
        base.OnVisibleChanged(e);
        _edit.Visible = Visible;
    }

    protected override void OnSizeChanged(EventArgs e)
    {
        base.OnSizeChanged(e);
        PerformLayout();
    }

    protected override void OnLayout(LayoutEventArgs levent)
    {
        base.OnLayout(levent);
        LayoutEdit();
    }

    protected override void OnGotFocus(EventArgs e)
    {
        base.OnGotFocus(e);
        if (!_edit.Focused)
        {
            _edit.Focus();
        }
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        e.Graphics.Clear(BackColor);
        Rectangle border = ClientRectangle;
        border.Width = Math.Max(0, border.Width - 1);
        border.Height = Math.Max(0, border.Height - 1);
        using Pen pen = new(Enabled ? _borderColor : Color.FromArgb(70, 70, 76));
        e.Graphics.DrawRectangle(pen, border);
    }

    private void LayoutEdit()
    {
        const int border = 1;
        int left = border + Math.Max(2, ContentLeftPadding);
        int right = border + Math.Max(2, ContentRightPadding);
        int innerWidth = Math.Max(4, Width - left - right);
        int textHeight = Math.Max(
            _edit.Font.Height,
            TextRenderer.MeasureText(
                "Ag",
                _edit.Font,
                Size.Empty,
                TextFormatFlags.NoPadding | TextFormatFlags.SingleLine).Height);

        int innerHeight = Height - (border * 2);
        int top = border + Math.Max(0, (innerHeight - textHeight) / 2);
        if (innerHeight - textHeight >= 4)
        {
            top = Math.Min(top + 1, border + innerHeight - textHeight);
        }

        _edit.SetBounds(left, top, innerWidth, Math.Min(textHeight + 1, Math.Max(textHeight, innerHeight)));
    }
}
