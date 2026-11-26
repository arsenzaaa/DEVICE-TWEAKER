namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private sealed record BusyButtonState(string Caption, string OriginalSourceText);

    private readonly Dictionary<Button, BusyButtonState> _busyButtonCaptions = [];

    // Feedback belongs to the awaited operation, not to a timer or artificial delay.
    private IDisposable BeginButtonFeedback(Button? button, string caption, Action? refreshAfter = null)
    {
        if (button is null || button.IsDisposed) return new DeviceOperationScope(() => { });
        string original = GetIdleControlSourceText(button);
        bool wasEnabled = button.Enabled;
        _busyButtonCaptions.Add(button, new BusyButtonState(caption, original));
        button.Enabled = false;
        button.Text = caption;
        button.Refresh();
        return new DeviceOperationScope(() =>
        {
            _busyButtonCaptions.Remove(button);
            if (button.IsDisposed) return;
            button.Text = original;
            button.Enabled = wasEnabled;
            refreshAfter?.Invoke();
        });
    }

    private string GetIdleControlSourceText(Control control)
        => control is Button button && _busyButtonCaptions.TryGetValue(button, out BusyButtonState? busy)
            ? busy.OriginalSourceText
            : GetSourceControlText(control);

    private void RefreshButtonFeedbackCaptions()
    {
        foreach ((Button button, BusyButtonState busy) in _busyButtonCaptions)
            if (!button.IsDisposed) button.Text = busy.Caption;
    }
}
