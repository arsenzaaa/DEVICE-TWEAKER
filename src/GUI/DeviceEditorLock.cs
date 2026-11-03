namespace DeviceTweakerCS;

// Consume input for the device editor during APPLY/AUTO without changing the
// controls' enabled/read-only states or their native painting. Programmatic AUTO
// updates still run normally, including their availability rules.
internal sealed class DeviceEditorLock : IMessageFilter, IDisposable
{
    private readonly Control _root;
    private bool _disposed;

    internal DeviceEditorLock(Control root)
    {
        _root = root;
        foreach (ThemedDropDownPicker picker in Descendants(root).OfType<ThemedDropDownPicker>())
            picker.DismissPopup();
        Application.AddMessageFilter(this);
    }

    private static IEnumerable<Control> Descendants(Control root)
    {
        foreach (Control child in root.Controls)
        {
            yield return child;
            foreach (Control descendant in Descendants(child)) yield return descendant;
        }
    }

    public bool PreFilterMessage(ref Message message)
    {
        if (_disposed || _root.IsDisposed) return false;
        bool input = message.Msg is >= 0x100 and <= 0x109 or >= 0x200 and <= 0x20E
            or 0x300 or 0x302 or 0x303 or 0x304;
        if (!input) return false;
        Control? target = Control.FromChildHandle(message.HWnd);
        return target is not null && (target == _root || _root.Contains(target));
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        Application.RemoveMessageFilter(this);
    }
}
