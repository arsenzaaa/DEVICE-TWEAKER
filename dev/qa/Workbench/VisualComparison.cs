using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

namespace DeviceTweakerQa;

internal sealed record PixelDifference(string Status, long ChangedPixels, long ComparedPixels, double Ratio,
    int Left, int Top, int Width, int Height);

internal static class VisualComparison
{
    // Exact by default. No automatic masks, baseline replacement or tolerance widening.
    public static PixelDifference Compare(Bitmap before, Bitmap after, string? diffPath = null, int channelTolerance = 0)
    {
        if (channelTolerance is < 0 or > 255) throw new ArgumentOutOfRangeException(nameof(channelTolerance));
        if (before.Size != after.Size) return new("SIZE_CHANGED", 0, 0, 1, 0, 0, after.Width, after.Height);
        using Bitmap a = ToArgb(before), b = ToArgb(after), diff = new(a.Width, a.Height, PixelFormat.Format32bppArgb);
        Rectangle rect = new(Point.Empty, a.Size);
        BitmapData ad = a.LockBits(rect, ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
        BitmapData bd = b.LockBits(rect, ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
        BitmapData dd = diff.LockBits(rect, ImageLockMode.WriteOnly, PixelFormat.Format32bppArgb);
        long count = 0; int minX = a.Width, minY = a.Height, maxX = -1, maxY = -1;
        try
        {
            byte[] ar = new byte[a.Width * 4], br = new byte[ar.Length], dr = new byte[ar.Length];
            for (int y = 0; y < a.Height; y++)
            {
                Marshal.Copy(ad.Scan0 + y * ad.Stride, ar, 0, ar.Length);
                Marshal.Copy(bd.Scan0 + y * bd.Stride, br, 0, br.Length);
                for (int x = 0; x < a.Width; x++)
                {
                    int i = x * 4;
                    bool changed = Math.Abs(ar[i] - br[i]) > channelTolerance || Math.Abs(ar[i + 1] - br[i + 1]) > channelTolerance
                        || Math.Abs(ar[i + 2] - br[i + 2]) > channelTolerance || Math.Abs(ar[i + 3] - br[i + 3]) > channelTolerance;
                    if (changed) { count++; minX = Math.Min(minX, x); minY = Math.Min(minY, y); maxX = Math.Max(maxX, x); maxY = Math.Max(maxY, y); }
                    dr[i] = changed ? (byte)180 : (byte)(br[i] / 4);
                    dr[i + 1] = changed ? (byte)0 : (byte)(br[i + 1] / 4);
                    dr[i + 2] = changed ? (byte)255 : (byte)(br[i + 2] / 4); dr[i + 3] = 255;
                }
                Marshal.Copy(dr, 0, dd.Scan0 + y * dd.Stride, dr.Length);
            }
        }
        finally { a.UnlockBits(ad); b.UnlockBits(bd); diff.UnlockBits(dd); }
        if (diffPath is not null) { Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(diffPath))!); diff.Save(diffPath, ImageFormat.Png); }
        return new(count == 0 ? "MATCH" : "CHANGED", count, (long)a.Width * a.Height,
            count / ((double)a.Width * a.Height), count == 0 ? 0 : minX, count == 0 ? 0 : minY,
            count == 0 ? 0 : maxX - minX + 1, count == 0 ? 0 : maxY - minY + 1);
    }
    private static Bitmap ToArgb(Bitmap image)
    { Bitmap result = new(image.Width, image.Height, PixelFormat.Format32bppArgb); using Graphics g = Graphics.FromImage(result); g.DrawImageUnscaled(image, 0, 0); return result; }

    public static object CompareRuns(string beforeRoot, string afterRoot, string output)
    {
        beforeRoot = Path.GetFullPath(beforeRoot); afterRoot = Path.GetFullPath(afterRoot); output = Path.GetFullPath(output);
        if (beforeRoot.Equals(afterRoot, StringComparison.OrdinalIgnoreCase)
            || !output.Equals(Path.Combine(afterRoot, "comparison"), StringComparison.OrdinalIgnoreCase)
            || output.StartsWith(beforeRoot.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException("Comparison output must be candidate/comparison and must not modify the baseline");
        DevRun before = DevFiles.Read<DevRun>(Path.Combine(beforeRoot, "session.json")), after = DevFiles.Read<DevRun>(Path.Combine(afterRoot, "session.json"));
        DevFiles.Verify(beforeRoot, before); DevFiles.Verify(afterRoot, after);
        DevFiles.VerifyExecution(beforeRoot); DevFiles.VerifyExecution(afterRoot);
        Directory.CreateDirectory(output);
        if (before.Environment != after.Environment || before.ScenarioId != after.ScenarioId || before.ScenarioSha256 != after.ScenarioSha256)
            throw new InvalidDataException("Runs are not comparable: environment / scenario differs. Use the same scenario and display environment.");
        var a = before.Captures.Where(c => !c.Timeline).ToDictionary(c => c.Id);
        var b = after.Captures.Where(c => !c.Timeline).ToDictionary(c => c.Id);
        if (a.Count == 0 || !a.Keys.Order().SequenceEqual(b.Keys.Order())) throw new InvalidDataException("Checkpoint set differs / is empty");
        List<object> entries = []; bool changed = false;
        foreach (string id in a.Keys)
        {
            using Bitmap beforeImage = new(DevFiles.Inside(beforeRoot, a[id].Path)), afterImage = new(DevFiles.Inside(afterRoot, b[id].Path));
            string target = Path.Combine(output, id + "-diff.png");
            PixelDifference difference = Compare(beforeImage, afterImage, target);
            changed |= difference.Status != "MATCH";
            string old = Path.Combine(output, id + "-before.png"), next = Path.Combine(output, id + "-after.png");
            File.Copy(DevFiles.Inside(beforeRoot, a[id].Path), old, overwrite: true); File.Copy(DevFiles.Inside(afterRoot, b[id].Path), next, overwrite: true);
            entries.Add(new { id, difference, before = Path.GetFileName(old), after = Path.GetFileName(next), diff = File.Exists(target) ? Path.GetFileName(target) : null });
        }
        var result = new { status = changed ? "REVIEW_REQUIRED" : "MATCH", beforeRoot, afterRoot, channelTolerance = 0,
            beforeManifest = DevFiles.Hash(Path.Combine(beforeRoot, "session.json")), afterManifest = DevFiles.Hash(Path.Combine(afterRoot, "session.json")), checkpoints = entries };
        DevFiles.Write(Path.Combine(output, "comparison.json"), result);
        WorkbenchHtml.Comparison(output, result);
        DevFiles.Verify(beforeRoot, before); DevFiles.Verify(afterRoot, after);
        return result;
    }
}
