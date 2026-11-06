namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    // A posted Task continuation can run ahead of input, WM_PAINT and WM_TIMER.
    // Resume a build slice from a native UI timer after those messages can run.
    private Task YieldDeviceUiAsync()
    {
        if (_uiDisposing || IsDisposed || Disposing) return Task.CompletedTask;
        TaskCompletionSource ready = new(TaskCreationOptions.RunContinuationsAsynchronously);
        System.Windows.Forms.Timer timer = new() { Interval = 1 };
        void Complete(object? sender, EventArgs e)
        {
            timer.Stop();
            timer.Tick -= Complete;
            Disposed -= Complete;
            timer.Dispose();
            ready.TrySetResult();
        }
        timer.Tick += Complete;
        Disposed += Complete;
        timer.Start();
        return ready.Task;
    }
}
