using System.Text;
using System.Text.RegularExpressions;

namespace DeviceTweakerCS;

/// <summary>
/// Dark, non-activating tooltip window without the native Win32 tooltip text limit.
/// Full sentences are preserved and stacked vertically.
/// </summary>
internal sealed class ThemedToolTip : IDisposable
{
    private const int WrapWidth = 110;
    private readonly Dictionary<Control, Registration> _registrations = [];
    private readonly System.Windows.Forms.Timer _showTimer = new();
    private readonly System.Windows.Forms.Timer _hideTimer = new();
    private readonly ToolTipWindow _window;
    private Control? _pendingTarget;
    private Control? _visibleTarget;
    private bool _disposed;

    public ThemedToolTip(bool showAlways, Font? font = null)
    {
        ShowAlways = showAlways;
        _window = new ToolTipWindow(font ?? SystemFonts.MessageBoxFont!);
        _showTimer.Tick += (_, _) =>
        {
            _showTimer.Stop();
            if (_pendingTarget is not null && _registrations.TryGetValue(_pendingTarget, out Registration? registration))
            {
                ShowPopup(registration.Text, _pendingTarget, Cursor.Position);
            }
        };
        _hideTimer.Tick += (_, _) =>
        {
            _hideTimer.Stop();
            HidePopup();
        };
    }

    public bool Active { get; set; } = true;
    public bool ShowAlways { get; set; }
    public int InitialDelay { get; set; } = 500;
    public int ReshowDelay { get; set; } = 100;
    public int AutoPopDelay { get; set; } = 5000;

    public void SetToolTip(Control control, string? caption)
    {
        ThrowIfDisposed();
        RemoveRegistration(control);

        string text = FormatText(caption);
        if (text.Length == 0)
        {
            return;
        }

        EventHandler mouseEnter = (_, _) => QueueShow(control);
        EventHandler mouseLeave = (_, _) => CancelAndHide(control);
        MouseEventHandler mouseDown = (_, _) => CancelAndHide(control);
        EventHandler disposed = (_, _) => RemoveRegistration(control);

        control.MouseEnter += mouseEnter;
        control.MouseLeave += mouseLeave;
        control.MouseDown += mouseDown;
        control.Disposed += disposed;
        _registrations[control] = new Registration(text, mouseEnter, mouseLeave, mouseDown, disposed);
    }

    public void Hide(Control control)
    {
        if (_pendingTarget == control)
        {
            _pendingTarget = null;
            _showTimer.Stop();
        }

        if (_visibleTarget == control)
        {
            HidePopup();
        }
    }

    public void Show(string? caption, Control control, Point point, int duration)
    {
        ThrowIfDisposed();
        if (!Active || control.IsDisposed)
        {
            return;
        }

        Point screenPoint = control.PointToScreen(point);
        ShowPopup(FormatText(caption), control, screenPoint);
        StartHideTimer(duration);
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;
        _showTimer.Stop();
        _hideTimer.Stop();
        foreach (Control control in _registrations.Keys.ToArray())
        {
            RemoveRegistration(control);
        }

        _showTimer.Dispose();
        _hideTimer.Dispose();
        _window.Dispose();
    }

    private void QueueShow(Control control)
    {
        if (!Active || control.IsDisposed || !_registrations.ContainsKey(control))
        {
            return;
        }

        _pendingTarget = control;
        _showTimer.Stop();
        _showTimer.Interval = Math.Max(1, _window.IsShown ? ReshowDelay : InitialDelay);
        _showTimer.Start();
    }

    private void CancelAndHide(Control control)
    {
        if (_pendingTarget == control)
        {
            _pendingTarget = null;
            _showTimer.Stop();
        }

        if (_visibleTarget == control)
        {
            HidePopup();
        }
    }

    private void ShowPopup(string text, Control target, Point anchor)
    {
        if (text.Length == 0 || target.IsDisposed || (!ShowAlways && target.FindForm() is not { ContainsFocus: true }))
        {
            return;
        }

        _pendingTarget = null;
        _visibleTarget = target;
        int dpi = target.DeviceDpi > 0 ? target.DeviceDpi : 96;
        _window.SetText(text, dpi);

        Screen screen = Screen.FromPoint(anchor);
        Rectangle workingArea = screen.WorkingArea;
        Point location = new(anchor.X + ScaleForDpi(14, dpi), anchor.Y + ScaleForDpi(20, dpi));
        if (location.X + _window.Width > workingArea.Right)
        {
            location.X = Math.Max(workingArea.Left, anchor.X - _window.Width - ScaleForDpi(10, dpi));
        }

        if (location.Y + _window.Height > workingArea.Bottom)
        {
            location.Y = Math.Max(workingArea.Top, anchor.Y - _window.Height - ScaleForDpi(10, dpi));
        }

        _window.Location = location;
        if (!_window.IsShown)
        {
            _window.ShowInactive();
        }

        StartHideTimer(AutoPopDelay);
    }

    private void StartHideTimer(int duration)
    {
        _hideTimer.Stop();
        _hideTimer.Interval = Math.Max(1, duration);
        _hideTimer.Start();
    }

    private void HidePopup()
    {
        _hideTimer.Stop();
        _visibleTarget = null;
        if (_window.IsShown)
        {
            _window.HideInactive();
        }
    }

    private void RemoveRegistration(Control control)
    {
        if (!_registrations.Remove(control, out Registration? registration))
        {
            return;
        }

        control.MouseEnter -= registration.MouseEnter;
        control.MouseLeave -= registration.MouseLeave;
        control.MouseDown -= registration.MouseDown;
        control.Disposed -= registration.Disposed;
        Hide(control);
    }

    private static string FormatText(string? caption)
    {
        if (string.IsNullOrWhiteSpace(caption))
        {
            return string.Empty;
        }

        string normalized = caption.Replace("\r\n", "\n").Replace('\r', '\n');
        string[] lines = normalized.Trim().Split('\n');
        StringBuilder result = new();
        bool previousLineWasEmpty = false;

        foreach (string sourceLine in lines)
        {
            string line = Regex.Replace(sourceLine, @"\s+", " ").Trim();
            if (line.Length == 0)
            {
                if (!previousLineWasEmpty && result.Length > 0)
                {
                    result.AppendLine();
                }

                previousLineWasEmpty = true;
                continue;
            }

            previousLineWasEmpty = false;
            string[] sentences = Regex.Split(line, @"(?<=[.!?])\s+(?=[A-Z0-9])");
            foreach (string sentence in sentences)
            {
                AppendWrappedLine(result, sentence.Trim());
            }
        }

        return result.ToString().TrimEnd();
    }

    private static void AppendWrappedLine(StringBuilder result, string text)
    {
        string remaining = text;
        while (remaining.Length > WrapWidth)
        {
            int split = remaining.LastIndexOf(' ', WrapWidth);
            if (split < 1)
            {
                split = WrapWidth;
            }

            result.AppendLine(remaining[..split].TrimEnd());
            remaining = remaining[split..].TrimStart();
        }

        result.AppendLine(remaining);
    }

    private static int ScaleForDpi(int value, int dpi)
    {
        return Math.Max(1, (int)Math.Round(value * Math.Max(96, dpi) / 96d));
    }

    private void ThrowIfDisposed()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
    }

    private sealed record Registration(
        string Text,
        EventHandler MouseEnter,
        EventHandler MouseLeave,
        MouseEventHandler MouseDown,
        EventHandler Disposed);

    private sealed class ToolTipWindow : NativeWindow, IDisposable
    {
        private const int WsPopup = unchecked((int)0x80000000);
        private const int WsExTopmost = 0x00000008;
        private const int WsExToolWindow = 0x00000080;
        private const int WsExNoActivate = 0x08000000;
        private const uint SwpNoActivate = 0x0010;
        private const uint SwpShowWindow = 0x0040;
        private const int SwHide = 0;
        private const int WmPaint = 0x000F;
        private const int WmEraseBackground = 0x0014;
        private const int WmMouseActivate = 0x0021;
        private const int WmNcHitTest = 0x0084;
        private const int MaNoActivate = 3;
        private const int HtTransparent = -1;
        private const int BorderSize = 1;
        private const int HorizontalPadding = 10;
        private const int VerticalPadding = 8;
        private const int MaximumWidth = 900;
        private static readonly IntPtr HwndTopmost = new(-1);
        private readonly Font _font;
        private int _dpi = 96;
        private string _text = string.Empty;
        private bool _disposed;

        public Point Location { get; set; }
        public int Width { get; private set; } = 1;
        public int Height { get; private set; } = 1;

        [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetWindowPos(
            IntPtr hWnd,
            IntPtr hWndInsertAfter,
            int x,
            int y,
            int cx,
            int cy,
            uint flags);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int command);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true)]
        private static extern bool InvalidateRect(IntPtr hWnd, IntPtr rect, bool erase);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr BeginPaint(IntPtr hWnd, out PaintStruct paint);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool EndPaint(IntPtr hWnd, ref PaintStruct paint);

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        private struct Rect
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        private struct PaintStruct
        {
            public IntPtr Hdc;
            [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
            public bool Erase;
            public Rect Paint;
            [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
            public bool Restore;
            [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
            public bool IncUpdate;
            [System.Runtime.InteropServices.MarshalAs(
                System.Runtime.InteropServices.UnmanagedType.ByValArray,
                SizeConst = 32)]
            public byte[] Reserved;
        }

        public ToolTipWindow(Font font)
        {
            _font = (Font)font.Clone();
            CreateParams parameters = new()
            {
                Caption = string.Empty,
                ClassName = null,
                Style = WsPopup,
                ExStyle = WsExTopmost | WsExToolWindow | WsExNoActivate,
                X = 0,
                Y = 0,
                Width = 1,
                Height = 1,
            };
            CreateHandle(parameters);
        }

        public bool IsShown => Handle != IntPtr.Zero && IsWindowVisible(Handle);

        public void SetText(string text, int dpi)
        {
            _text = text;
            _dpi = Math.Max(96, dpi);
            int border = Scale(BorderSize);
            int horizontalPadding = Scale(HorizontalPadding);
            int verticalPadding = Scale(VerticalPadding);
            int maximumWidth = Scale(MaximumWidth);
            Size measured = TextRenderer.MeasureText(
                text,
                _font,
                new Size(maximumWidth - ((horizontalPadding + border) * 2), int.MaxValue),
                TextFormatFlags.Left
                | TextFormatFlags.Top
                | TextFormatFlags.NoPrefix
                | TextFormatFlags.WordBreak);
            Width = Math.Min(
                maximumWidth,
                Math.Max(1, measured.Width + ((horizontalPadding + border) * 2)));
            Height = Math.Max(1, measured.Height + ((verticalPadding + border) * 2));
            if (Handle != IntPtr.Zero)
            {
                _ = InvalidateRect(Handle, IntPtr.Zero, erase: false);
            }
        }

        public void ShowInactive()
        {
            _ = SetWindowPos(
                Handle,
                HwndTopmost,
                Location.X,
                Location.Y,
                Width,
                Height,
                SwpNoActivate | SwpShowWindow);
            _ = InvalidateRect(Handle, IntPtr.Zero, erase: false);
        }

        public void HideInactive()
        {
            if (Handle != IntPtr.Zero)
            {
                _ = ShowWindow(Handle, SwHide);
            }
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WmEraseBackground)
            {
                m.Result = new IntPtr(1);
                return;
            }

            if (m.Msg == WmMouseActivate)
            {
                m.Result = new IntPtr(MaNoActivate);
                return;
            }

            if (m.Msg == WmNcHitTest)
            {
                m.Result = new IntPtr(HtTransparent);
                return;
            }

            if (m.Msg == WmPaint)
            {
                PaintStruct paint = new() { Reserved = new byte[32] };
                IntPtr hdc = BeginPaint(Handle, out paint);
                try
                {
                    using Graphics graphics = Graphics.FromHdc(hdc);
                    graphics.Clear(Color.FromArgb(118, 118, 126));
                    int border = Scale(BorderSize);
                    int horizontalPadding = Scale(HorizontalPadding);
                    int verticalPadding = Scale(VerticalPadding);
                    Rectangle inner = new(
                        border,
                        border,
                        Math.Max(0, Width - (border * 2)),
                        Math.Max(0, Height - (border * 2)));
                    using SolidBrush background = new(Color.FromArgb(12, 12, 15));
                    graphics.FillRectangle(background, inner);

                    Rectangle textBounds = new(
                        border + horizontalPadding,
                        border + verticalPadding,
                        Math.Max(0, Width - ((border + horizontalPadding) * 2)),
                        Math.Max(0, Height - ((border + verticalPadding) * 2)));
                    TextRenderer.DrawText(
                        graphics,
                        _text,
                        _font,
                        textBounds,
                        Color.FromArgb(235, 235, 235),
                        Color.FromArgb(12, 12, 15),
                        TextFormatFlags.Left
                        | TextFormatFlags.Top
                        | TextFormatFlags.NoPrefix
                        | TextFormatFlags.WordBreak);
                }
                finally
                {
                    _ = EndPaint(Handle, ref paint);
                }

                return;
            }

            base.WndProc(ref m);
        }

        private int Scale(int value)
        {
            return Math.Max(1, (int)Math.Round(value * _dpi / 96d));
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            HideInactive();
            if (Handle != IntPtr.Zero)
            {
                DestroyHandle();
            }
            _font.Dispose();
        }
    }
}
