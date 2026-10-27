using System.Drawing.Imaging;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text.Json;

internal static class NativeEditorEvidence
{
    [DllImport("user32.dll")]
    private static extern bool PrintWindow(IntPtr hwnd, IntPtr dc, uint flags);

    private static Bitmap Capture(Control control, string path)
    {
        control.Refresh();
        Bitmap bitmap = new(control.Width, control.Height);
        using Graphics graphics = Graphics.FromImage(bitmap);
        IntPtr dc = graphics.GetHdc();
        try { if (!PrintWindow(control.Handle, dc, 2)) throw new InvalidOperationException("Native editor capture failed"); }
        finally { graphics.ReleaseHdc(dc); }
        bitmap.Save(path, ImageFormat.Png);
        return bitmap;
    }

    // Ignore positioning; compare the complete foreground glyph against a taller native edit.
    // Both controls have the same text, font, width, colors and single-line Win32 implementation.
    private static string Glyph(Bitmap bitmap)
    {
        List<Point> pixels = [];
        for (int y = 0; y < bitmap.Height; y++)
            for (int x = 0; x < bitmap.Width; x++)
            {
                Color color = bitmap.GetPixel(x, y);
                if (Math.Max(color.R, Math.Max(color.G, color.B)) > 100) pixels.Add(new(x, y));
            }
        if (pixels.Count == 0) throw new InvalidOperationException("Native editor text was not rendered");
        int left = pixels.Min(p => p.X), top = pixels.Min(p => p.Y);
        return string.Join(';', pixels.Select(p => $"{p.X-left},{p.Y-top}"));
    }

    internal static bool Run(Assembly assembly, Form host, string output)
    {
        string directory = Path.Combine(output, "native-editors");
        Directory.CreateDirectory(directory);
        List<object> cases = [];
        bool passed = true;
        foreach (float scale in new[] { 1f, 1.25f, 1.5f })
        foreach (string text in new[] { "0xFA0", "gypq", "Йду" })
        {
            string id = $"{(int)(scale*100)}-{cases.Count}";
            using Font font = new("Consolas", 10.5f * scale);
            using Control field = (Control)Activator.CreateInstance(assembly.GetType("DeviceTweakerCS.ThemedTextBox", true)!, nonPublic:true)!;
            field.Bounds = new(20, 120, (int)Math.Round(78*scale), (int)Math.Round(24*scale));
            field.Font = font;
            field.BackColor = Color.FromArgb(18,18,22);
            field.ForeColor = Color.FromArgb(240,240,240);
            field.Text = text;
            host.Controls.Add(field);
            TextBox inner = (TextBox)field.GetType().GetProperty("Inner")!.GetValue(field)!;
            using TextBox reference = new() { Font=font, BorderStyle=BorderStyle.None, AutoSize=false,
                BackColor=field.BackColor, ForeColor=field.ForeColor, Text=text, Bounds=new(260,120,inner.Width,inner.PreferredHeight+16), TabStop=false };
            host.Controls.Add(reference);
            Application.DoEvents();
            using Bitmap actual = Capture(inner, Path.Combine(directory,id+"-actual.png"));
            using Bitmap complete = Capture(reference, Path.Combine(directory,id+"-reference.png"));
            string expected = Glyph(complete);
            bool same = Glyph(actual) == expected;
            // A deliberately clipped real native control must fail this pixel oracle.
            reference.Height = Math.Max(3,font.Height/2);
            using Bitmap broken = Capture(reference, Path.Combine(directory,id+"-negative.png"));
            bool detectsClipping;
            try { detectsClipping = Glyph(broken) != expected; }
            catch (InvalidOperationException) { detectsClipping = true; }
            passed &= same && detectsClipping;
            cases.Add(new { id,kind="editor",scale,text,inner.Height,preferredHeight=inner.PreferredHeight,fontHeight=font.Height,
                completeGlyph=same,negativeDetected=detectsClipping });
            host.Controls.Remove(reference);
            host.Controls.Remove(field);
        }
        foreach (float scale in new[] { 1f, 1.25f, 1.5f })
        {
            string id=$"label-{(int)(scale*100)}";
            using Font font=new("Consolas",10.5f*scale);
            using Label actual=new() { AutoSize=false,Text="Аудио:",Font=font,BackColor=Color.FromArgb(18,18,22),
                ForeColor=Color.FromArgb(240,240,240),Bounds=new(20,120,(int)Math.Round(96*scale),(int)Math.Round(22*scale)) };
            using Label reference=new() { AutoSize=false,Text=actual.Text,Font=font,BackColor=actual.BackColor,
                ForeColor=actual.ForeColor,Bounds=new(260,120,actual.Width,actual.Height+16) };
            host.Controls.Add(actual);
            host.Controls.Add(reference);
            using Bitmap rendered=Capture(actual,Path.Combine(directory,id+"-actual.png"));
            using Bitmap complete=Capture(reference,Path.Combine(directory,id+"-reference.png"));
            string expected=Glyph(complete);
            bool same=Glyph(rendered)==expected;
            reference.Height=Math.Max(3,font.Height/2);
            using Bitmap broken=Capture(reference,Path.Combine(directory,id+"-negative.png"));
            bool detectsClipping;
            try { detectsClipping=Glyph(broken)!=expected; }
            catch(InvalidOperationException) { detectsClipping=true; }
            passed &= same && detectsClipping;
            cases.Add(new { id,kind="label",scale,text=actual.Text,actual.Height,fontHeight=font.Height,
                completeGlyph=same,negativeDetected=detectsClipping });
            host.Controls.Remove(reference);
            host.Controls.Remove(actual);
        }
        var images=Directory.GetFiles(directory,"*.png").Select(path=>new { path,
            sha256=Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))) }).ToArray();
        File.WriteAllText(Path.Combine(output,"native-editors.json"), JsonSerializer.Serialize(new {
            status=passed?"PASS":"FAIL",cases,images,
            scope="Native PrintWindow glyph pixels for text fields and Audio label versus taller controls with identical font/text/width; deliberately clipped controls must differ. Font/size scaling model, not real Windows DPI transitions." },new JsonSerializerOptions {WriteIndented=true}));
        return passed;
    }
}
