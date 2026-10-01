namespace DeviceTweakerCS;

internal sealed class ThemedDialogForm : Form
{
    private const int WmNcActivate = 0x0086;

    public ThemedDialogForm()
    {
        SetStyle(
            ControlStyles.AllPaintingInWmPaint |
            ControlStyles.OptimizedDoubleBuffer |
            ControlStyles.UserPaint |
            ControlStyles.ResizeRedraw,
            true);
        Padding = Padding.Empty;
    }

    protected override CreateParams CreateParams
    {
        get
        {
            CreateParams cp = base.CreateParams;
            if (FormBorderStyle == FormBorderStyle.None)
            {
                cp.ClassStyle |= 0x00020000; // CS_DROPSHADOW
            }
            return cp;
        }
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        base.OnPaint(e);
        if (FormBorderStyle == FormBorderStyle.None)
        {
            using Pen borderPen = new(Color.FromArgb(68, 68, 80), 1);
            e.Graphics.DrawRectangle(borderPen, 0, 0, ClientSize.Width - 1, ClientSize.Height - 1);
        }
    }

    public static void EnableDrag(Control control, Form form)
    {
        control.MouseDown += (_, e) =>
        {
            if (e.Button == MouseButtons.Left)
            {
                NativeUser32.ReleaseCapture();
                _ = NativeUser32.SendMessage(form.Handle, NativeUser32.WmNcLButtonDown, (IntPtr)NativeUser32.HtCaption, IntPtr.Zero);
            }
        };
    }

    protected override void WndProc(ref Message m)
    {
        // Keep the native caption in the dark active visual state when a
        // non-activating popup or a nested modal window is shown. This changes
        // only non-client painting; normal focus and modality are untouched.
        if (m.Msg == WmNcActivate && m.WParam == IntPtr.Zero && WindowState != FormWindowState.Minimized)
        {
            m.WParam = new IntPtr(1);
        }

        base.WndProc(ref m);
    }
}
