namespace DeviceTweakerCS;

internal sealed class ThemedDropDownPicker : Control, IMessageFilter
{
    private const int WmEraseBackground = 0x0014;
    private const int WmMouseActivate = 0x0021;
    private const int WmLButtonDown = 0x0201;
    private const int WmRButtonDown = 0x0204;
    private const int WmNcLButtonDown = 0x00A1;
    private const int WmMouseWheel = 0x020A;
    private const int MaNoActivate = 3;
    private const int WsExToolWindow = 0x00000080;
    private const int WsExNoActivate = 0x08000000;

    private DropDownForm? _popup;
    private int _selectedIndex = -1;

    public List<object> Items { get; } = [];
    public int MaxDropDownItems { get; set; } = 8;
    public int DropDownWidth { get; set; }
    public int ItemHeight { get; set; } = 18;
    public Color BorderColor { get; set; } = Color.FromArgb(150, 150, 150);
    public Color ButtonColor { get; set; } = Color.FromArgb(14, 14, 17);
    public Color SelectedBackColor { get; set; } = Color.FromArgb(34, 34, 40);
    public Color SelectedForeColor { get; set; } = Color.FromArgb(240, 240, 240);
    public Color ArrowColor { get; set; } = Color.FromArgb(240, 240, 240);

    public event EventHandler? SelectedIndexChanged;

    public int SelectedIndex
    {
        get => _selectedIndex;
        set
        {
            int normalized = value >= 0 && value < Items.Count ? value : -1;
            if (_selectedIndex == normalized)
            {
                return;
            }

            _selectedIndex = normalized;
            Invalidate();
            SelectedIndexChanged?.Invoke(this, EventArgs.Empty);
        }
    }

    public object? SelectedItem
    {
        get => _selectedIndex >= 0 && _selectedIndex < Items.Count ? Items[_selectedIndex] : null;
        set
        {
            if (value is null)
            {
                SelectedIndex = -1;
                return;
            }

            int index = Items.IndexOf(value);
            if (index >= 0)
            {
                SelectedIndex = index;
            }
        }
    }

    public ThemedDropDownPicker()
    {
        SetStyle(
            ControlStyles.UserPaint
            | ControlStyles.AllPaintingInWmPaint
            | ControlStyles.OptimizedDoubleBuffer
            | ControlStyles.ResizeRedraw
            | ControlStyles.Opaque
            | ControlStyles.Selectable,
            true);

        TabStop = true;
        Cursor = Cursors.Hand;
    }

    public bool PreFilterMessage(ref Message m)
    {
        if (_popup is not { Visible: true })
        {
            return false;
        }

        if (m.Msg is not (WmLButtonDown or WmRButtonDown or WmNcLButtonDown or WmMouseWheel))
        {
            return false;
        }

        Point mouse = MousePosition;
        Rectangle ownerBounds = RectangleToScreen(ClientRectangle);
        Rectangle popupBounds = _popup.Bounds;
        if (!ownerBounds.Contains(mouse) && !popupBounds.Contains(mouse))
        {
            ClosePopup();
        }

        return false;
    }

    protected override void OnPaintBackground(PaintEventArgs e)
    {
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        using SolidBrush backgroundBrush = new(BackColor);
        e.Graphics.FillRectangle(backgroundBrush, ClientRectangle);

        Rectangle bounds = ClientRectangle;
        int arrowWidth = Math.Max(18, Math.Min(24, Height));
        Rectangle arrowRect = new(bounds.Right - arrowWidth - 1, bounds.Top + 1, arrowWidth, bounds.Height - 2);
        using SolidBrush buttonBrush = new(ButtonColor);
        e.Graphics.FillRectangle(buttonBrush, arrowRect);

        string text = SelectedItem?.ToString() ?? string.Empty;
        Rectangle textRect = new(bounds.Left + 6, bounds.Top + 1, Math.Max(0, bounds.Width - arrowWidth - 10), bounds.Height - 2);
        Color textColor = Enabled ? ForeColor : Color.FromArgb(120, 120, 125);
        TextRenderer.DrawText(
            e.Graphics,
            text,
            Font,
            textRect,
            textColor,
            BackColor,
            TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix | TextFormatFlags.EndEllipsis);

        int centerX = arrowRect.Left + arrowRect.Width / 2;
        int centerY = arrowRect.Top + arrowRect.Height / 2 + 1;
        Point[] arrow =
        [
            new(centerX - 4, centerY - 1),
            new(centerX + 4, centerY - 1),
            new(centerX, centerY + 3),
        ];
        using SolidBrush arrowBrush = new(Enabled ? ArrowColor : Color.FromArgb(115, 115, 120));
        e.Graphics.FillPolygon(arrowBrush, arrow);

        using Pen borderPen = new(Enabled ? BorderColor : Color.FromArgb(80, 80, 86));
        Rectangle borderRect = bounds;
        borderRect.Width -= 1;
        borderRect.Height -= 1;
        e.Graphics.DrawRectangle(borderPen, borderRect);
    }

    protected override void OnMouseDown(MouseEventArgs e)
    {
        Focus();
        TogglePopup();
        base.OnMouseDown(e);
    }

    protected override void OnKeyDown(KeyEventArgs e)
    {
        if (e.KeyCode is Keys.Enter or Keys.Space or Keys.F4)
        {
            TogglePopup();
            e.Handled = true;
            return;
        }

        base.OnKeyDown(e);
    }

    protected override void OnEnabledChanged(EventArgs e)
    {
        if (!Enabled)
        {
            ClosePopup();
        }

        base.OnEnabledChanged(e);
        Invalidate();
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            ClosePopup();
        }

        base.Dispose(disposing);
    }

    protected override void WndProc(ref Message m)
    {
        if (m.Msg == WmEraseBackground)
        {
            m.Result = IntPtr.Zero;
            return;
        }

        base.WndProc(ref m);
    }

    private void TogglePopup()
    {
        if (!Enabled || Items.Count == 0)
        {
            return;
        }

        if (_popup is { Visible: true })
        {
            ClosePopup();
            return;
        }

        ShowPopup();
    }

    private void ShowPopup()
    {
        ClosePopup();

        int visibleItems = Math.Min(Math.Max(1, MaxDropDownItems), Items.Count);
        int rowHeight = Math.Max(ItemHeight, 18);
        int contentHeight = visibleItems * rowHeight;
        int popupWidth = Math.Max(Width, DropDownWidth > 0 ? DropDownWidth : Width);
        DropDownList list = new(this, popupWidth - 2, contentHeight, rowHeight, visibleItems);
        DropDownForm popup = new(this, list, BackColor, BorderColor)
        {
            Size = new Size(popupWidth, contentHeight + 2),
            Location = PointToScreen(new Point(0, Height - 1)),
        };

        _popup = popup;
        popup.FormClosed += (_, _) =>
        {
            if (ReferenceEquals(_popup, popup))
            {
                _popup = null;
            }

            Application.RemoveMessageFilter(this);
            Invalidate();
        };

        Application.AddMessageFilter(this);
        Form? ownerForm = FindForm();
        if (ownerForm is not null)
        {
            popup.Show(ownerForm);
        }
        else
        {
            popup.Show();
        }
    }

    private void ClosePopup()
    {
        DropDownForm? popup = _popup;
        _popup = null;
        Application.RemoveMessageFilter(this);

        if (popup is not null && !popup.IsDisposed)
        {
            popup.Close();
        }
    }

    private void CommitSelection(int index)
    {
        SelectedIndex = index;
        ClosePopup();
    }

    private sealed class DropDownForm : Form
    {
        private readonly ThemedDropDownPicker _owner;
        private readonly Color _background;
        private readonly Color _border;

        protected override bool ShowWithoutActivation => true;

        public DropDownForm(ThemedDropDownPicker owner, DropDownList content, Color background, Color border)
        {
            _owner = owner;
            _background = background;
            _border = border;

            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            BackColor = background;
            AutoScaleMode = AutoScaleMode.None;
            Padding = Padding.Empty;

            SetStyle(
                ControlStyles.UserPaint
                | ControlStyles.AllPaintingInWmPaint
                | ControlStyles.OptimizedDoubleBuffer
                | ControlStyles.ResizeRedraw,
                true);

            content.Location = new Point(1, 1);
            Controls.Add(content);
        }

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= WsExToolWindow | WsExNoActivate;
                return cp;
            }
        }

        protected override void OnPaintBackground(PaintEventArgs e)
        {
            using SolidBrush brush = new(_background);
            e.Graphics.FillRectangle(brush, ClientRectangle);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            using SolidBrush brush = new(_background);
            e.Graphics.FillRectangle(brush, ClientRectangle);

            Rectangle rect = ClientRectangle;
            rect.Width -= 1;
            rect.Height -= 1;
            using Pen pen = new(_border);
            e.Graphics.DrawRectangle(pen, rect);
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WmEraseBackground)
            {
                m.Result = IntPtr.Zero;
                return;
            }

            if (m.Msg == WmMouseActivate)
            {
                m.Result = new IntPtr(MaNoActivate);
                return;
            }

            base.WndProc(ref m);
        }
    }

    private sealed class DropDownList : Control
    {
        private readonly ThemedDropDownPicker _owner;
        private readonly int _itemHeight;
        private readonly int _visibleItems;
        private int _firstIndex;
        private int _hotIndex;

        public DropDownList(ThemedDropDownPicker owner, int width, int height, int itemHeight, int visibleItems)
        {
            _owner = owner;
            _itemHeight = itemHeight;
            _visibleItems = Math.Max(1, visibleItems);
            _hotIndex = owner.SelectedIndex >= 0 ? owner.SelectedIndex : 0;
            _firstIndex = Math.Clamp(_hotIndex - _visibleItems + 1, 0, Math.Max(0, owner.Items.Count - _visibleItems));

            SetStyle(
                ControlStyles.UserPaint
                | ControlStyles.AllPaintingInWmPaint
                | ControlStyles.OptimizedDoubleBuffer
                | ControlStyles.ResizeRedraw
                | ControlStyles.Opaque,
                true);

            TabStop = false;
            BackColor = owner.BackColor;
            ForeColor = owner.ForeColor;
            Font = owner.Font;
            Size = new Size(width, height);
        }

        protected override void OnPaintBackground(PaintEventArgs e)
        {
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            using SolidBrush backgroundBrush = new(_owner.BackColor);
            e.Graphics.FillRectangle(backgroundBrush, ClientRectangle);

            int rows = Math.Min(_visibleItems, _owner.Items.Count - _firstIndex);
            for (int row = 0; row < rows; row++)
            {
                int index = _firstIndex + row;
                Rectangle itemRect = new(0, row * _itemHeight, Width, _itemHeight);
                bool selected = index == _hotIndex;
                Color background = selected ? _owner.SelectedBackColor : _owner.BackColor;
                Color foreground = selected ? _owner.SelectedForeColor : _owner.ForeColor;

                using SolidBrush itemBrush = new(background);
                e.Graphics.FillRectangle(itemBrush, itemRect);

                if (selected)
                {
                    using SolidBrush markerBrush = new(_owner.BorderColor);
                    e.Graphics.FillRectangle(markerBrush, new Rectangle(itemRect.Left, itemRect.Top, 3, itemRect.Height));
                }

                string text = _owner.Items[index].ToString() ?? string.Empty;
                Rectangle textRect = new(itemRect.Left + 8, itemRect.Top, Math.Max(0, itemRect.Width - 14), itemRect.Height);
                TextRenderer.DrawText(
                    e.Graphics,
                    text,
                    _owner.Font,
                    textRect,
                    foreground,
                    background,
                    TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix);
            }
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            SetHotIndex(IndexFromPoint(e.Location));
            base.OnMouseMove(e);
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            int index = IndexFromPoint(e.Location);
            if (index >= 0)
            {
                SetHotIndex(index);
                _owner.CommitSelection(index);
            }

            base.OnMouseDown(e);
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WmEraseBackground)
            {
                m.Result = IntPtr.Zero;
                return;
            }

            base.WndProc(ref m);
        }

        private int IndexFromPoint(Point point)
        {
            if (point.X < 0 || point.X >= Width || point.Y < 0 || point.Y >= Height)
            {
                return -1;
            }

            int index = _firstIndex + (point.Y / _itemHeight);
            return index >= 0 && index < _owner.Items.Count ? index : -1;
        }

        private void SetHotIndex(int index)
        {
            if (index < 0 || index >= _owner.Items.Count || index == _hotIndex)
            {
                return;
            }

            int previousIndex = _hotIndex;
            _hotIndex = index;
            if (_hotIndex < _firstIndex)
            {
                _firstIndex = _hotIndex;
                Invalidate();
                return;
            }

            if (_hotIndex >= _firstIndex + _visibleItems)
            {
                _firstIndex = _hotIndex - _visibleItems + 1;
                Invalidate();
                return;
            }

            InvalidateItem(previousIndex);
            InvalidateItem(_hotIndex);
        }

        private void InvalidateItem(int index)
        {
            int row = index - _firstIndex;
            if (row < 0 || row >= _visibleItems)
            {
                return;
            }

            Invalidate(new Rectangle(0, row * _itemHeight, Width, _itemHeight));
        }
    }
}
