using System.Drawing.Drawing2D;

namespace DeviceTweakerCS;

internal sealed class ThemedScrollBar : Control
{
    private const int MinThumbHeight = 24;

    private int _value;
    private int _maximum = 1;
    private int _viewport = 1;
    private bool _dragging;
    private int _dragOffset;

    public event EventHandler? ValueChanged;

    public int Maximum
    {
        get => _maximum;
        set
        {
            int next = Math.Max(1, value);
            if (_maximum != next)
            {
                _maximum = next;
                Invalidate();
                SetValue(_value, false);
            }
        }
    }

    public int ViewportSize
    {
        get => _viewport;
        set
        {
            int next = Math.Max(1, value);
            if (_viewport != next)
            {
                _viewport = next;
                Invalidate();
                SetValue(_value, false);
            }
        }
    }

    public int Value
    {
        get => _value;
        set => SetValue(value, true);
    }

    public int SmallChange { get; set; } = 40;
    public int LargeChange { get; set; } = 120;

    public Color ThumbColor { get; set; } = Color.White;
    public Color TrackColor { get; set; } = Color.Empty;
    public Color RailColor { get; set; } = Color.White;
    public int RailWidth { get; set; } = 2;
    public int ThumbWidth { get; set; } = 8;
    public int ThumbCornerRadius { get; set; } = 6;

    public ThemedScrollBar()
    {
        SetStyle(ControlStyles.AllPaintingInWmPaint
                 | ControlStyles.OptimizedDoubleBuffer
                 | ControlStyles.ResizeRedraw
                 | ControlStyles.UserPaint, true);
        TabStop = false;
        Cursor = Cursors.Hand;
        Width = 12;
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        base.OnPaint(e);

        e.Graphics.SmoothingMode = SmoothingMode.None;
        Color track = TrackColor.IsEmpty ? BackColor : TrackColor;
        using SolidBrush trackBrush = new(track);
        e.Graphics.FillRectangle(trackBrush, ClientRectangle);

        if (RailWidth > 0 && RailColor.A > 0)
        {
            int railWidth = Math.Min(RailWidth, ClientSize.Width);
            if (railWidth > 0)
            {
                int railX = (ClientSize.Width - railWidth) / 2;
                using SolidBrush railBrush = new(RailColor);
                e.Graphics.FillRectangle(railBrush, railX, 0, railWidth, ClientSize.Height);
            }
        }

        Rectangle thumb = GetThumbRect();
        if (thumb.Height <= 0)
        {
            return;
        }

        using SolidBrush thumbBrush = new(ThumbColor);
        int radius = Math.Min(ThumbCornerRadius, Math.Min(thumb.Width, thumb.Height) / 2);
        if (radius > 0)
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            using GraphicsPath path = CreateRoundedRect(thumb, radius);
            e.Graphics.FillPath(thumbBrush, path);
        }
        else
        {
            e.Graphics.FillRectangle(thumbBrush, thumb);
        }
    }

    protected override void OnMouseDown(MouseEventArgs e)
    {
        base.OnMouseDown(e);
        if (e.Button != MouseButtons.Left)
        {
            return;
        }

        Rectangle thumb = GetThumbRect();
        if (thumb.Contains(e.Location))
        {
            _dragging = true;
            _dragOffset = e.Y - thumb.Top;
            Capture = true;
            return;
        }

        int direction = e.Y < thumb.Top ? -1 : 1;
        SetValue(_value + (direction * LargeChange), true);
    }

    protected override void OnMouseMove(MouseEventArgs e)
    {
        base.OnMouseMove(e);
        if (!_dragging)
        {
            return;
        }

        Rectangle thumb = GetThumbRect();
        int trackHeight = ClientSize.Height;
        int range = GetRange();
        if (range <= 0 || trackHeight <= 0)
        {
            return;
        }

        int thumbHeight = thumb.Height;
        int maxTop = trackHeight - thumbHeight;
        int newTop = e.Y - _dragOffset;
        newTop = Math.Max(0, Math.Min(maxTop, newTop));

        int newValue = (int)Math.Round((float)newTop / maxTop * range);
        SetValue(newValue, true);
    }

    protected override void OnMouseUp(MouseEventArgs e)
    {
        base.OnMouseUp(e);
        if (e.Button != MouseButtons.Left)
        {
            return;
        }

        _dragging = false;
        Capture = false;
    }

    protected override void OnMouseWheel(MouseEventArgs e)
    {
        base.OnMouseWheel(e);
        int delta = e.Delta > 0 ? -SmallChange : SmallChange;
        SetValue(_value + delta, true);
    }

    private int GetRange()
    {
        return Math.Max(0, _maximum - _viewport);
    }

    private void SetValue(int value, bool raiseEvent)
    {
        int range = GetRange();
        int next = Math.Max(0, Math.Min(range, value));
        if (next == _value)
        {
            return;
        }

        _value = next;
        Invalidate();
        if (raiseEvent)
        {
            ValueChanged?.Invoke(this, EventArgs.Empty);
        }
    }

    private Rectangle GetThumbRect()
    {
        int trackHeight = ClientSize.Height;
        if (trackHeight <= 0)
        {
            return Rectangle.Empty;
        }

        int range = GetRange();
        if (range <= 0)
        {
            return new Rectangle(0, 0, ClientSize.Width, trackHeight);
        }

        float ratio = (float)_viewport / _maximum;
        int thumbHeight = Math.Max(MinThumbHeight, (int)Math.Round(trackHeight * ratio));
        if (thumbHeight > trackHeight)
        {
            thumbHeight = trackHeight;
        }

        int maxTop = trackHeight - thumbHeight;
        int top = maxTop > 0
            ? (int)Math.Round((float)_value / range * maxTop)
            : 0;

        int thumbWidth = GetThumbWidth();
        int thumbX = (ClientSize.Width - thumbWidth) / 2;
        return new Rectangle(thumbX, top, thumbWidth, thumbHeight);
    }

    private int GetThumbWidth()
    {
        int width = ThumbWidth <= 0 ? ClientSize.Width : ThumbWidth;
        width = Math.Max(2, Math.Min(width, ClientSize.Width));
        return width;
    }

    private static GraphicsPath CreateRoundedRect(Rectangle rect, int radius)
    {
        GraphicsPath path = new();
        if (radius <= 0)
        {
            path.AddRectangle(rect);
            path.CloseFigure();
            return path;
        }

        int diameter = radius * 2;
        int right = rect.Right - diameter;
        int bottom = rect.Bottom - diameter;

        path.AddArc(rect.Left, rect.Top, diameter, diameter, 180, 90);
        path.AddArc(right, rect.Top, diameter, diameter, 270, 90);
        path.AddArc(right, bottom, diameter, diameter, 0, 90);
        path.AddArc(rect.Left, bottom, diameter, diameter, 90, 90);
        path.CloseFigure();
        return path;
    }
}
