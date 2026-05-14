namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private void StyleDarkDropDownPicker(ThemedDropDownPicker picker)
    {
        picker.BackColor = Color.FromArgb(18, 18, 22);
        picker.ForeColor = _fgMain;
        picker.BorderColor = _border;
        picker.ButtonColor = Color.FromArgb(14, 14, 17);
        picker.SelectedBackColor = Color.FromArgb(48, 48, 58);
        picker.SelectedForeColor = _fgMain;
        picker.ArrowColor = _fgMain;
        picker.ItemHeight = UiScale(18);
    }

    private void StyleDarkComboBox(ComboBox comboBox)
    {
        comboBox.FlatStyle = FlatStyle.Flat;
        comboBox.BackColor = Color.FromArgb(18, 18, 22);
        comboBox.ForeColor = _fgMain;
        comboBox.DrawMode = DrawMode.OwnerDrawFixed;
        comboBox.ItemHeight = UiScale(18);
        if (comboBox is ThemedComboBox themedComboBox)
        {
            themedComboBox.BorderColor = _border;
            themedComboBox.DropButtonColor = Color.FromArgb(14, 14, 17);
            themedComboBox.SelectedBackColor = Color.FromArgb(34, 34, 40);
            themedComboBox.SelectedForeColor = _fgMain;
            themedComboBox.ArrowColor = _fgMain;
            return;
        }

        comboBox.DrawItem -= DrawDarkComboBoxItem;
        comboBox.DrawItem += DrawDarkComboBoxItem;
    }

    private void DrawDarkComboBoxItem(object? sender, DrawItemEventArgs e)
    {
        if (sender is not ComboBox comboBox)
        {
            return;
        }

        bool selected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
        bool disabled = (e.State & DrawItemState.Disabled) == DrawItemState.Disabled || !comboBox.Enabled;
        Color background = selected ? Color.FromArgb(34, 34, 40) : comboBox.BackColor;
        Color foreground = disabled ? _mutedText : comboBox.ForeColor;

        using SolidBrush backgroundBrush = new(background);
        e.Graphics.FillRectangle(backgroundBrush, e.Bounds);

        string text = (e.Index >= 0 && e.Index < comboBox.Items.Count
            ? comboBox.GetItemText(comboBox.Items[e.Index])
            : comboBox.Text) ?? string.Empty;

        Rectangle textBounds = new(
            e.Bounds.Left + UiScale(5),
            e.Bounds.Top,
            Math.Max(0, e.Bounds.Width - UiScale(10)),
            e.Bounds.Height);

        TextRenderer.DrawText(
            e.Graphics,
            text,
            comboBox.Font,
            textBounds,
            foreground,
            background,
            TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix);

    }
}

internal sealed class ThemedComboBox : ComboBox
{
    private const int WmPaint = 0x000F;
    private const int WmEraseBackground = 0x0014;
    private const int WmSetCursor = 0x0020;
    private const int WmNcPaint = 0x0085;
    private const int WmSetFocus = 0x0007;
    private const int WmKillFocus = 0x0008;
    private const int WmNcMouseMove = 0x00A0;
    private const int WmNcMouseLeave = 0x02A2;
    private const int WmMouseMove = 0x0200;
    private const int WmMouseHover = 0x02A1;
    private const int WmMouseLeave = 0x02A3;
    private const int WmLButtonDown = 0x0201;
    private const int WmLButtonDoubleClick = 0x0203;
    private const int WmKeyDown = 0x0100;
    private const int WmPrint = 0x0317;
    private const int WmPrintClient = 0x0318;
    private const int WmMouseActivate = 0x0021;
    private const int MaNoActivate = 3;
    private const int WsExToolWindow = 0x00000080;
    private const int WsExNoActivate = 0x08000000;

    private ThemedComboDropDownForm? _popup;

    public Color BorderColor { get; set; } = Color.FromArgb(150, 150, 150);
    public Color DropButtonColor { get; set; } = Color.FromArgb(18, 18, 22);
    public Color SelectedBackColor { get; set; } = Color.FromArgb(34, 34, 40);
    public Color SelectedForeColor { get; set; } = Color.FromArgb(240, 240, 240);
    public Color ArrowColor { get; set; } = Color.FromArgb(240, 240, 240);

    public ThemedComboBox()
    {
        DrawMode = DrawMode.OwnerDrawFixed;
        DropDownStyle = ComboBoxStyle.DropDownList;
        FlatStyle = FlatStyle.Flat;
        SetStyle(
            ControlStyles.AllPaintingInWmPaint
            | ControlStyles.OptimizedDoubleBuffer
            | ControlStyles.ResizeRedraw,
            true);
    }

    protected override void OnSelectedIndexChanged(EventArgs e)
    {
        base.OnSelectedIndexChanged(e);
        Invalidate();
    }

    protected override void OnTextChanged(EventArgs e)
    {
        base.OnTextChanged(e);
        Invalidate();
    }

    protected override void OnDrawItem(DrawItemEventArgs e)
    {
        if (e.Index < 0)
        {
            return;
        }

        bool selected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
        Color background = selected ? SelectedBackColor : BackColor;
        Color foreground = selected ? SelectedForeColor : ForeColor;

        using SolidBrush backgroundBrush = new(background);
        e.Graphics.FillRectangle(backgroundBrush, e.Bounds);

        string text = GetItemText(Items[e.Index]) ?? string.Empty;
        Rectangle textBounds = new(e.Bounds.Left + 6, e.Bounds.Top, Math.Max(0, e.Bounds.Width - 12), e.Bounds.Height);
        TextRenderer.DrawText(
            e.Graphics,
            text,
            Font,
            textBounds,
            foreground,
            background,
            TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix);
    }

    protected override void WndProc(ref Message m)
    {
        if (m.Msg == WmEraseBackground)
        {
            if (m.WParam != IntPtr.Zero)
            {
                using Graphics g = Graphics.FromHdc(m.WParam);
                PaintChrome(g);
            }

            m.Result = new IntPtr(1);
            return;
        }

        if (m.Msg is WmPrint or WmPrintClient && m.WParam != IntPtr.Zero)
        {
            using Graphics g = Graphics.FromHdc(m.WParam);
            PaintChrome(g);
            m.Result = IntPtr.Zero;
            return;
        }

        if (m.Msg is WmMouseMove or WmMouseHover or WmMouseLeave or WmNcMouseMove or WmNcMouseLeave)
        {
            PaintChrome();
            m.Result = IntPtr.Zero;
            return;
        }

        if (m.Msg == WmSetCursor)
        {
            Cursor.Current = Cursors.Hand;
            PaintChrome();
            m.Result = new IntPtr(1);
            return;
        }

        if (m.Msg is WmLButtonDown or WmLButtonDoubleClick)
        {
            Focus();
            ToggleThemedPopup();
            return;
        }

        if (m.Msg == WmKeyDown)
        {
            Keys key = (Keys)m.WParam.ToInt32();
            if (key is Keys.Space or Keys.Enter or Keys.F4)
            {
                ToggleThemedPopup();
                return;
            }
        }

        base.WndProc(ref m);
        if (m.Msg is WmPaint or WmNcPaint or WmSetFocus or WmKillFocus)
        {
            PaintChrome();
        }
    }

    protected override void OnEnabledChanged(EventArgs e)
    {
        base.OnEnabledChanged(e);
        Invalidate();
    }

    protected override void OnDropDown(EventArgs e)
    {
        DroppedDown = false;
        ToggleThemedPopup();
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            CloseThemedPopup();
        }

        base.Dispose(disposing);
    }

    private void ToggleThemedPopup()
    {
        if (!Enabled || Items.Count == 0)
        {
            return;
        }

        if (_popup is { Visible: true })
        {
            CloseThemedPopup();
            return;
        }

        ShowThemedPopup();
    }

    private void ShowThemedPopup()
    {
        CloseThemedPopup();

        int visibleItems = Math.Min(Math.Max(1, MaxDropDownItems), Items.Count);
        int popupItemHeight = Math.Max(ItemHeight, 18);
        int contentHeight = popupItemHeight * visibleItems;
        int popupWidth = Math.Max(Width, DropDownWidth);

        ThemedComboPopupList popupList = new(this, popupWidth - 2, contentHeight, popupItemHeight, visibleItems);
        ThemedComboDropDownForm popup = new(this, popupList, BackColor, BorderColor)
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

            Invalidate();
            Focus();
        };

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

    private void CloseThemedPopup()
    {
        ThemedComboDropDownForm? popup = _popup;
        if (popup is null)
        {
            return;
        }

        _popup = null;
        if (!popup.IsDisposed)
        {
            popup.Close();
        }
    }

    private void CommitPopupSelection(int selectedIndex)
    {
        if (selectedIndex >= 0 && selectedIndex < Items.Count)
        {
            SelectedIndex = selectedIndex;
        }

        CloseThemedPopup();
    }

    private void PaintChrome()
    {
        if (!IsHandleCreated || ClientSize.Width <= 0 || ClientSize.Height <= 0)
        {
            return;
        }

        using Graphics g = Graphics.FromHwnd(Handle);
        PaintChrome(g);
    }

    private void PaintChrome(Graphics g)
    {
        if (ClientSize.Width <= 0 || ClientSize.Height <= 0)
        {
            return;
        }

        Rectangle bounds = ClientRectangle;
        using SolidBrush backBrush = new(BackColor);
        g.FillRectangle(backBrush, bounds);

        int arrowWidth = Math.Max(18, Math.Min(24, Height));
        Rectangle arrowRect = new(bounds.Right - arrowWidth - 1, bounds.Top + 1, arrowWidth, bounds.Height - 2);
        using SolidBrush buttonBrush = new(DropButtonColor);
        g.FillRectangle(buttonBrush, arrowRect);

        string text = GetItemText(SelectedItem) ?? Text ?? string.Empty;
        Rectangle textRect = new(bounds.Left + 6, bounds.Top + 1, Math.Max(0, bounds.Width - arrowWidth - 10), bounds.Height - 2);
        Color textColor = Enabled ? ForeColor : Color.FromArgb(130, 130, 135);
        TextRenderer.DrawText(
            g,
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
        g.FillPolygon(arrowBrush, arrow);

        using Pen borderPen = new(BorderColor);
        Rectangle borderRect = bounds;
        borderRect.Width -= 1;
        borderRect.Height -= 1;
        g.DrawRectangle(borderPen, borderRect);
    }

    private sealed class ThemedComboDropDownForm : Form
    {
        private readonly ThemedComboBox _owner;
        private readonly Color _background;
        private readonly Color _border;

        protected override bool ShowWithoutActivation => true;

        public ThemedComboDropDownForm(
            ThemedComboBox owner,
            ThemedComboPopupList content,
            Color background,
            Color border)
        {
            _owner = owner;
            _background = background;
            _border = border;

            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            BackColor = background;
            Padding = Padding.Empty;
            AutoScaleMode = AutoScaleMode.None;

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

        protected override void OnDeactivate(EventArgs e)
        {
            base.OnDeactivate(e);
            _owner.CloseThemedPopup();
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

    private sealed class ThemedComboPopupList : Control
    {
        private readonly ThemedComboBox _owner;
        private readonly int _itemHeight;
        private readonly int _visibleItems;
        private int _firstIndex;
        private int _hotIndex;

        public ThemedComboPopupList(
            ThemedComboBox owner,
            int width,
            int height,
            int itemHeight,
            int visibleItems)
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

            TabStop = true;
            Font = owner.Font;
            BackColor = owner.BackColor;
            ForeColor = owner.ForeColor;
            Size = new Size(width, height);
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

                string text = _owner.GetItemText(_owner.Items[index]) ?? string.Empty;
                Rectangle textRect = new(itemRect.Left + 6, itemRect.Top, Math.Max(0, itemRect.Width - 12), itemRect.Height);
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

        protected override void OnPaintBackground(PaintEventArgs e)
        {
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
                _owner.CommitPopupSelection(index);
            }

            base.OnMouseDown(e);
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                _owner.CloseThemedPopup();
                e.Handled = true;
                return;
            }

            if (e.KeyCode == Keys.Enter)
            {
                _owner.CommitPopupSelection(_hotIndex);
                e.Handled = true;
                return;
            }

            if (e.KeyCode == Keys.Up)
            {
                SetHotIndex(Math.Max(0, _hotIndex - 1));
                e.Handled = true;
                return;
            }

            if (e.KeyCode == Keys.Down)
            {
                SetHotIndex(Math.Min(_owner.Items.Count - 1, _hotIndex + 1));
                e.Handled = true;
                return;
            }

            base.OnKeyDown(e);
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
            else if (_hotIndex >= _firstIndex + _visibleItems)
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

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WmEraseBackground)
            {
                m.Result = IntPtr.Zero;
                return;
            }

            base.WndProc(ref m);
        }
    }
}
