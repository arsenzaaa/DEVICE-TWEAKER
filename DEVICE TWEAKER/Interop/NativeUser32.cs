using System.Runtime.InteropServices;

namespace DeviceTweakerCS;

internal static class NativeUser32
{
    internal const int SbHorz = 0;
    internal const int SbVert = 1;
    internal const int SbBoth = 3;

    [DllImport("user32.dll")]
    internal static extern int GetDpiForSystem();

    [DllImport("user32.dll")]
    internal static extern int GetDpiForWindow(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError = true)]
    internal static extern bool ShowScrollBar(IntPtr hWnd, int wBar, bool bShow);
}
