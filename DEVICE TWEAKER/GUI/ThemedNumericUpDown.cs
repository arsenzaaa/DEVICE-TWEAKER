using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;

namespace DeviceTweakerCS;

/// <summary>
/// Dark-themed NumericUpDown: custom-drawn spinner buttons, padded edit field.
/// Avoids setting BackColor / empty SetWindowTheme on UpDownButtons - that throws
/// "Visual Style handle creation failed" under ForceDark + EnableVisualStyles.
/// </summary>
internal sealed class ThemedNumericUpDown : NumericUpDown
{
    private const int WmPaint = 0x000F;
    private const int WmPrintClient = 0x0318;
    private const int WmMouseMove = 0x0200;
    private const int WmMouseLeave = 0x02A3;
    private const int EmSetMargins = 0x00D3;
    private const int EcLeftMargin = 0x0001;
    private const int EcRightMargin = 0x0002;

    private ButtonsPainter? _buttonsPainter;

    public Color ButtonBackColor { get; set; } = Color.FromArgb(14, 14, 17);
    public Color ButtonHoverColor { get; set; } = Color.FromArgb(40, 40, 48);
    public Color ArrowColor { get; set; } = Color.FromArgb(230, 230, 230);

    public ThemedNumericUpDown()
    {
        BorderStyle = BorderStyle.FixedSingle;
        BackColor = Color.FromArgb(18, 18, 22);
        ForeColor = Color.FromArgb(240, 240, 240);
        TextAlign = HorizontalAlignment.Center;
    }

    protected override void OnHandleCreated(EventArgs e)
    {
        base.OnHandleCreated(e);
        ApplyTheme();
    }

    protected override void OnBackColorChanged(EventArgs e)
    {
        base.OnBackColorChanged(e);
        ApplyTheme();
    }

    protected override void OnForeColorChanged(EventArgs e)
    {
        base.OnForeColorChanged(e);
        ApplyTheme();
    }

    protected override void OnSizeChanged(EventArgs e)
    {
        base.OnSizeChanged(e);
        ApplyEditMargins();
        _buttonsPainter?.Invalidate();
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _buttonsPainter?.Release();
            _buttonsPainter = null;
        }

        base.Dispose(disposing);
    }

    private void ApplyTheme()
    {
        if (!IsHandleCreated)
        {
            return;
        }

        foreach (Control child in Controls)
        {
            if (!child.IsHandleCreated)
            {
                child.HandleCreated += (_, _) => ApplyTheme();
                continue;
            }

            if (child is TextBox edit)
            {
                try
                {
                    if (edit.BackColor != BackColor)
                    {
                        edit.BackColor = BackColor;
                    }

                    if (edit.ForeColor != ForeColor)
                    {
                        edit.ForeColor = ForeColor;
                    }
                }
                catch (InvalidOperationException)
                {
                    // Visual styles can reject BackColor on some hosts - ignore.
                }

                ApplyEditMargins();
                continue;
            }

            // UpDownButtons: never set BackColor / SetWindowTheme - that causes
            // InvalidOperationException ("Visual Style handle creation failed").
            if (_buttonsPainter is null || _buttonsPainter.AssignedHandle != child.Handle)
            {
                _buttonsPainter?.Release();
                _buttonsPainter = new ButtonsPainter(this, child);
            }
        }
    }

    private void ApplyEditMargins()
    {
        foreach (Control child in Controls)
        {
            if (child is not TextBox edit || !edit.IsHandleCreated)
            {
                continue;
            }

            int left = Math.Max(4, Math.Min(8, Height / 4));
            int right = Math.Max(4, Math.Min(8, Height / 4));
            int margins = left | (right << 16);
            _ = SendMessage(edit.Handle, EmSetMargins, (IntPtr)(EcLeftMargin | EcRightMargin), (IntPtr)margins);
            try
            {
                edit.TextAlign = TextAlign;
            }
            catch
            {
            }
        }
    }

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

    private sealed class ButtonsPainter : NativeWindow
    {
        private readonly ThemedNumericUpDown _owner;
        private readonly Control _buttons;
        private int _hoverHalf; // 0 none, 1 up, 2 down
        private bool _trackingLeave;

        public IntPtr AssignedHandle { get; private set; }

        public ButtonsPainter(ThemedNumericUpDown owner, Control buttons)
        {
            _owner = owner;
            _buttons = buttons;
            AssignedHandle = buttons.Handle;
            AssignHandle(buttons.Handle);
            _buttons.Invalidated += (_, _) => Invalidate();
        }

        public void Invalidate()
        {
            if (Handle != IntPtr.Zero)
            {
                InvalidateRect(Handle, IntPtr.Zero, false);
            }
        }

        public void Release()
        {
            if (Handle != IntPtr.Zero)
            {
                ReleaseHandle();
            }

            AssignedHandle = IntPtr.Zero;
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg is WmPaint or WmPrintClient)
            {
                // Let default paint run first (keeps hit-testing / visual-style state sane),
                // then cover with our dark spinner art.
                base.WndProc(ref m);
                try
                {
                    using Graphics g = Graphics.FromHwnd(Handle);
                    PaintButtons(g);
                }
                catch
                {
                }

                return;
            }

            if (m.Msg == WmMouseMove)
            {
                int y = (short)((m.LParam.ToInt32() >> 16) & 0xFFFF);
                int half = y < Math.Max(1, _buttons.Height / 2) ? 1 : 2;
                if (half != _hoverHalf)
                {
                    _hoverHalf = half;
                    Invalidate();
                }

                _trackingLeave = false;
                TrackMouseLeave();
            }
            else if (m.Msg == WmMouseLeave)
            {
                _trackingLeave = false;
                if (_hoverHalf != 0)
                {
                    _hoverHalf = 0;
                    Invalidate();
                }
            }

            base.WndProc(ref m);

            if (m.Msg is >= 0x0201 and <= 0x020A)
            {
                Invalidate();
            }
        }

        private void TrackMouseLeave()
        {
            if (_trackingLeave)
            {
                return;
            }

            TrackMouseEventEvents tme = new()
            {
                cbSize = Marshal.SizeOf<TrackMouseEventEvents>(),
                dwFlags = 0x00000002, // TME_LEAVE
                hwndTrack = Handle,
                dwHoverTime = 0,
            };
            if (TrackMouseEvent(ref tme))
            {
                _trackingLeave = true;
            }
        }

        private void PaintButtons(Graphics g)
        {
            g.SmoothingMode = SmoothingMode.None;
            Rectangle bounds = new(0, 0, Math.Max(1, _buttons.Width), Math.Max(1, _buttons.Height));
            using (SolidBrush bg = new(_owner.BackColor))
            {
                g.FillRectangle(bg, bounds);
            }

            int mid = Math.Max(1, bounds.Height / 2);
            Rectangle up = new(0, 0, bounds.Width, mid);
            Rectangle down = new(0, mid, bounds.Width, bounds.Height - mid);

            Color upFill = _hoverHalf == 1 ? _owner.ButtonHoverColor : _owner.ButtonBackColor;
            Color downFill = _hoverHalf == 2 ? _owner.ButtonHoverColor : _owner.ButtonBackColor;
            using (SolidBrush upBrush = new(upFill))
            using (SolidBrush downBrush = new(downFill))
            {
                g.FillRectangle(upBrush, up);
                g.FillRectangle(downBrush, down);
            }

            using (Pen border = new(Color.FromArgb(90, 90, 100)))
            {
                g.DrawLine(border, 0, 0, 0, bounds.Height);
                g.DrawLine(border, 0, mid, bounds.Width, mid);
            }

            DrawArrow(g, up, true);
            DrawArrow(g, down, false);
        }

        private void DrawArrow(Graphics g, Rectangle area, bool up)
        {
            int size = Math.Max(3, Math.Min(5, Math.Min(area.Width, area.Height) / 3));
            int cx = area.Left + (area.Width / 2);
            int cy = area.Top + (area.Height / 2);
            Point[] points = up
                ?
                [
                    new Point(cx, cy - size + 1),
                    new Point(cx - size, cy + size - 2),
                    new Point(cx + size, cy + size - 2),
                ]
                :
                [
                    new Point(cx, cy + size - 1),
                    new Point(cx - size, cy - size + 2),
                    new Point(cx + size, cy - size + 2),
                ];

            using SolidBrush brush = new(_owner.Enabled ? _owner.ArrowColor : Color.FromArgb(120, 120, 120));
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.FillPolygon(brush, points);
            g.SmoothingMode = SmoothingMode.None;
        }

        [DllImport("user32.dll")]
        private static extern bool InvalidateRect(IntPtr hWnd, IntPtr lpRect, bool bErase);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool TrackMouseEvent(ref TrackMouseEventEvents lpEventTrack);

        [StructLayout(LayoutKind.Sequential)]
        private struct TrackMouseEventEvents
        {
            public int cbSize;
            public int dwFlags;
            public IntPtr hwndTrack;
            public int dwHoverTime;
        }
    }
}
