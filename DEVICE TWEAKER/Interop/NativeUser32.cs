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

    internal const int EmGetRect = 0x00B2;
    internal const int EmSetRect = 0x00B3;
    internal const int EmSetMargins = 0x00D3;
    internal const int EcLeftMargin = 0x0001;
    internal const int EcRightMargin = 0x0002;

    [StructLayout(LayoutKind.Sequential)]
    internal struct Rect
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    internal static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    internal static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, ref Rect lParam);

    [DllImport("user32.dll", SetLastError = true)]
    internal static extern bool SetWindowPos(
        IntPtr hWnd,
        IntPtr hWndInsertAfter,
        int x,
        int y,
        int cx,
        int cy,
        int uFlags);
}
