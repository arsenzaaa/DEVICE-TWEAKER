using System.Drawing.Imaging;

namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private bool TryHandleVisualQa(ref Message message)
    {
        if (message.Msg is not (0x8475 or 0x8476)) return false;
        if (!QaRun.IsReadOnly || QaDesktop.Expected is null) return false;
        QaDesktop.DemandAssignedDesktop();
        if (message.Msg == 0x8476)
        {
            // Use the real handler with explicit modifiers. Synthetic WM_KEYDOWN
            // can inherit a stale keyboard/focus state on an inactive desktop.
            Keys key = (Keys)message.WParam.ToInt32();
            if (!Enabled || key is not (Keys.Home or Keys.End or Keys.PageUp or Keys.PageDown))
            { message.Result = IntPtr.Zero; return true; }
            KeyEventArgs args = new(key);
            // This command exercises list navigation with focus on its header,
            // while editor navigation is verified separately by control contracts.
            ActiveControl = _btnScanRef;
            OnMainFormKeyDown(this, args);
            message.Result = args.Handled ? new IntPtr(1) : IntPtr.Zero;
            return true;
        }
        string path = Path.Combine(QaRun.OutputDirectory, "visual-snapshot.png");
        try
        {
            if (Control.FromHandle(message.WParam) is not Form target
                || (target != this && !OwnedForms.Contains(target)))
                throw new InvalidOperationException("Visual QA target is not an owned form.");
            QaRun.DemandArtifactPath(path);
            using Bitmap image = new(target.Width, target.Height);
            target.DrawToBitmap(image, new Rectangle(Point.Empty, target.Size));
            image.Save(path, ImageFormat.Png);
            message.Result = new IntPtr(1);
        }
        catch (Exception error)
        {
            File.WriteAllText(path + ".error.txt", error.ToString());
            message.Result = IntPtr.Zero;
        }
        return true;
    }
}
