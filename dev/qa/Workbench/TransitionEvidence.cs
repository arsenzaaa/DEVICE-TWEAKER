using System.Drawing;
using System.Drawing.Imaging;

namespace DeviceTweakerQa;

internal static class TransitionEvidence
{
    public static void Write(string root, string id, DevCapture[] frames, int requestedIntervalMs)
    {
        double[] gaps = frames.Zip(frames.Skip(1), (a,b) => b.TimeMs - a.TimeMs).ToArray();
        DevFiles.Write(DevFiles.Inside(root, "timelines/" + id + ".json"), new
        {
            id, requestedIntervalMs, observedFrames = frames.Length, observedGapsMs = gaps,
            meanGapMs = gaps.Average(), maxGapMs = gaps.Max(),
            scope = "Actual native snapshot timestamps; capture/PNG overhead included. This is not display FPS or proof of smoothness.",
            frames = frames.Select(f => new { f.Id, f.Path, f.Sha256, f.TimeMs }),
        });
        for (int page = 0; page * 8 < frames.Length; page++)
        {
            using Bitmap sheet = new(1280, 540); using Graphics g = Graphics.FromImage(sheet);
            using Font font = new("Consolas", 11); g.Clear(Color.FromArgb(22,22,27));
            for (int item = 0; item < 8 && page * 8 + item < frames.Length; item++)
            {
                DevCapture frame = frames[page * 8 + item]; int x = item % 4 * 320, y = item / 4 * 270;
                using Bitmap image = new(DevFiles.Inside(root, frame.Path));
                float scale = Math.Min(310f / image.Width, 235f / image.Height);
                g.DrawString($"{frame.Id}\n{frame.TimeMs:F1} ms", font, Brushes.White, x + 4, y + 2);
                g.DrawImage(image, x + 4, y + 38, image.Width * scale, image.Height * scale);
            }
            string path = DevFiles.Inside(root, $"timelines/{id}-contact-{page}.png"); Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            sheet.Save(path, ImageFormat.Png);
        }
    }
}
