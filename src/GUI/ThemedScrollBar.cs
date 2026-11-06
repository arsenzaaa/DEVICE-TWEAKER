using System.ComponentModel;
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
    private double _wheelRemainder;
    private ControlAnimation? _scrollMotion;
    internal bool IsScrolling => _scrollMotion?.IsRunning == true;
    internal int TargetValue => IsScrolling ? (int)Math.Round(_scrollMotion!.Target) : _value;

    private ControlAnimation ScrollMotion => _scrollMotion ??= new(this, 80,
        position => SetValue((int)Math.Round(position), true));

    internal void StopScrolling()
    { _wheelRemainder = 0; _scrollMotion?.MoveTo(_value, animate: false); }

    internal void ScrollTo(int value)
    {
        int target = Math.Clamp(value, 0, GetRange());
        if (target == TargetValue && IsScrolling) return;
        if (!IsScrolling) ScrollMotion.MoveTo(_value, animate: false);
        ScrollMotion.MoveTo(target);
    }

    internal void ScrollWheel(int delta)
    {
        if (delta == 0) return;
        // Fractional deltas survive even when one event is less than one pixel.
        _wheelRemainder -= (double)delta / SystemInformation.MouseWheelScrollDelta * SmallChange;
        int change = (int)_wheelRemainder;
        _wheelRemainder -= change;
        if (change != 0) ScrollTo(TargetValue + change);
    }

    public event EventHandler? ValueChanged;

    [DefaultValue(1)]
    public int Maximum
    {
        get => _maximum;
        set
        {
            int next = Math.Max(1, value);
            if (_maximum != next)
            {
                _maximum = next;
                _scrollMotion?.MoveTo(Math.Clamp(_value, 0, GetRange()), animate: false);
                Invalidate();
                SetValue(_value, false);
            }
        }
    }

    [DefaultValue(1)]
    public int ViewportSize
    {
        get => _viewport;
        set
        {
            int next = Math.Max(1, value);
            if (_viewport != next)
            {
                _viewport = next;
                _scrollMotion?.MoveTo(Math.Clamp(_value, 0, GetRange()), animate: false);
                Invalidate();
                SetValue(_value, false);
            }
        }
    }

    [DefaultValue(0)]
    public int Value
    {
        get => _value;
        set
        {
            // A consumer may mirror Value back while handling ValueChanged.
            // Equal-value synchronization must not cancel the active gesture.
            int next = Math.Clamp(value, 0, GetRange());
            if (next == _value) return;
            _scrollMotion?.MoveTo(next, animate: false);
            SetValue(next, true);
        }
    }

    [DefaultValue(40)]
    public int SmallChange { get; set; } = 40;
    [DefaultValue(120)]
    public int LargeChange { get; set; } = 120;

    private bool _hovered;

    [DefaultValue(typeof(Color), "White")]
    public Color ThumbColor { get; set; } = Color.White;
    [DefaultValue(typeof(Color), "")]
    public Color ThumbHoverColor { get; set; } = Color.Empty;
    [DefaultValue(typeof(Color), "")]
    public Color ThumbDragColor { get; set; } = Color.Empty;
    [DefaultValue(typeof(Color), "")]
    public Color TrackColor { get; set; } = Color.Empty;
    [DefaultValue(typeof(Color), "")]
    public Color RailColor { get; set; } = Color.Empty;
    [DefaultValue(0)]
    public int RailWidth { get; set; } = 0;
    [DefaultValue(10)]
    public int ThumbWidth { get; set; } = 10;
    [DefaultValue(7)]
    public int ThumbCornerRadius { get; set; } = 7;

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

    protected override void OnMouseEnter(EventArgs e)
    {
        base.OnMouseEnter(e);
        _hovered = true;
        Invalidate();
    }

    protected override void OnMouseLeave(EventArgs e)
    {
        base.OnMouseLeave(e);
        _hovered = false;
        Invalidate();
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

        Color effectiveThumb = _dragging && !ThumbDragColor.IsEmpty
            ? ThumbDragColor
            : _hovered && !ThumbHoverColor.IsEmpty
                ? ThumbHoverColor
                : ThumbColor;

        using SolidBrush thumbBrush = new(effectiveThumb);
        int cornerDiameter = Math.Min(ThumbCornerRadius * 2, Math.Min(thumb.Width, thumb.Height));
        if (cornerDiameter > 1)
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            using GraphicsPath path = CreateRoundedRect(thumb, cornerDiameter);
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

        _scrollMotion?.MoveTo(_value, animate: false);
        Rectangle thumb = GetThumbRect();
        if (thumb.Contains(e.Location))
        {
            _dragging = true;
            _dragOffset = e.Y - thumb.Top;
            Capture = true;
            Invalidate();
            return;
        }

        int direction = e.Y < thumb.Top ? -1 : 1;
        ScrollTo(TargetValue + (direction * LargeChange));
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
        Invalidate();
    }

    protected override void OnMouseWheel(MouseEventArgs e)
    {
        base.OnMouseWheel(e);
        ScrollWheel(e.Delta);
        if (e is HandledMouseEventArgs handled) handled.Handled = true;
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing) _scrollMotion?.Dispose();
        base.Dispose(disposing);
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

    private static GraphicsPath CreateRoundedRect(Rectangle rect, int cornerDiameter)
    {
        GraphicsPath path = new();
        if (cornerDiameter <= 1)
        {
            path.AddRectangle(rect);
            path.CloseFigure();
            return path;
        }

        int diameter = Math.Min(cornerDiameter, Math.Min(rect.Width, rect.Height));
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
