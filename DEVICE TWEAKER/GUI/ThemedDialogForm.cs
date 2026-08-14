namespace DeviceTweakerCS;

internal sealed class ThemedDialogForm : Form
{
    private const int WmNcActivate = 0x0086;

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
