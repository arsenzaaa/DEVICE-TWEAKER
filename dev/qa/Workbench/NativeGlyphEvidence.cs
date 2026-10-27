using System.Drawing.Imaging;
using System.Runtime.InteropServices;

namespace DeviceTweakerQa;

internal static class NativeGlyphEvidence
{
    [DllImport("user32.dll")]
    private static extern bool PrintWindow(IntPtr window, IntPtr dc, uint flags);

    internal static Bitmap Capture(Control control, string path)
    {
        control.Refresh();
        Bitmap image = new(control.Width, control.Height);
        using Graphics graphics = Graphics.FromImage(image);
        IntPtr dc = graphics.GetHdc();
        try { if (!PrintWindow(control.Handle, dc, 2)) throw new InvalidOperationException("Native text capture failed"); }
        finally { graphics.ReleaseHdc(dc); }
        image.Save(path, ImageFormat.Png);
        return image;
    }

    // Includes antialiasing colors, not just a bright-pixel threshold. Ignore translation only.
    internal static string Signature(Bitmap image)
    {
        int background = image.GetPixel(0, 0).ToArgb();
        List<(int X, int Y, int Color)> pixels = [];
        for (int y = 0; y < image.Height; y++)
            for (int x = 0; x < image.Width; x++)
            {
                int color = image.GetPixel(x, y).ToArgb();
                if (color != background) pixels.Add((x, y, color));
            }
        if (pixels.Count == 0) return string.Empty;
        int left = pixels.Min(p => p.X), top = pixels.Min(p => p.Y);
        return string.Join(';', pixels.Select(p => $"{p.X-left},{p.Y-top},{p.Color:X8}"));
    }
}
