namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private readonly Dictionary<Button, string> _busyButtonCaptions = [];

    // Feedback belongs to the awaited operation, not to a timer or artificial delay.
    private IDisposable BeginButtonFeedback(Button? button, string caption, Action? refreshAfter = null)
    {
        if (button is null || button.IsDisposed) return new DeviceOperationScope(() => { });
        string original = GetSourceControlText(button);
        bool wasEnabled = button.Enabled;
        _busyButtonCaptions.Add(button, caption);
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

    private void RefreshButtonFeedbackCaptions()
    {
        foreach ((Button button, string caption) in _busyButtonCaptions)
            if (!button.IsDisposed) button.Text = caption;
    }
}
