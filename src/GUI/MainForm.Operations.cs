namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private readonly DeviceOperationCoordinator _operations = new();
    private string? _activeOperationName;

    private IDisposable? TryBeginDeviceOperation(string name)
    {
        IDisposable? lease = _operations.TryEnter();
        if (lease is null)
        {
            WriteLog($"UI.OPERATION.BLOCK: {name}");
            return null;
        }
        _activeOperationName = name;
        DeviceEditorLock? editors = name is "APPLY" or "AUTO-OPTIMIZATION"
            ? new(_devicesPanel) : null;
        SetOperationButtonsEnabled(false);
        WriteLog($"UI.OPERATION.BEGIN: {name}");
        return new DeviceOperationScope(() =>
        {
            editors?.Dispose();
            lease.Dispose();
            _activeOperationName = null;
            WriteLog($"UI.OPERATION.END: {name}");
            if (!_uiDisposing && !IsDisposed && !Disposing)
                SetOperationButtonsEnabled(_devicesBusyDepth == 0);
        });
    }

    private sealed class DeviceOperationScope(Action release) : IDisposable
    {
        private Action? _release = release;
        public void Dispose() => Interlocked.Exchange(ref _release, null)?.Invoke();
    }
}
