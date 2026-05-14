namespace DeviceTweakerCS;

internal class BufferedPanel : Panel
{
    public BufferedPanel()
    {
        SetStyle(ControlStyles.AllPaintingInWmPaint
                 | ControlStyles.OptimizedDoubleBuffer
                 | ControlStyles.ResizeRedraw
                 | ControlStyles.UserPaint, true);
        UpdateStyles();
    }
}

internal sealed class DeviceCardPanel : BufferedPanel
{
    private Rectangle _lastBounds;

    public Color BorderColor { get; set; } = Color.White;

    protected override void OnParentChanged(EventArgs e)
    {
        base.OnParentChanged(e);
        _lastBounds = Bounds;
    }

    protected override void SetBoundsCore(int x, int y, int width, int height, BoundsSpecified specified)
    {
        Control? parent = Parent;
        Rectangle oldBounds = Bounds;

        base.SetBoundsCore(x, y, width, height, specified);

        if (parent is not null && !oldBounds.IsEmpty)
        {
            parent.Invalidate(Rectangle.Union(oldBounds, Bounds), false);
        }

        _lastBounds = Bounds;
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        base.OnPaint(e);

        Rectangle rect = ClientRectangle;
        if (rect.Width <= 1 || rect.Height <= 1)
        {
            return;
        }

        rect.Width -= 1;
        rect.Height -= 1;
        using Pen pen = new(BorderColor);
        e.Graphics.DrawRectangle(pen, rect);
    }
}
