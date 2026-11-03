namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private readonly DeviceOperationCoordinator _operations = new();
    private string? _activeOperationName;

    private OperationReport CreateApplyReport(bool preview)
    {
#if DEVELOPER_QA
        return _devLastApplyReport = new OperationReport { IsDeveloperPreview = preview && _devExpectedApplyErrors >= 0 };
#else
        return new OperationReport();
#endif
    }

    private void ValidateApplyPreviewReport(OperationReport report)
    {
#if DEVELOPER_QA
        int expected = _devExpectedApplyErrors;
        _devExpectedApplyErrors = -1;
        if (expected >= 0 && report.Errors.Count != expected)
            throw new InvalidDataException($"APPLY model expected {expected} errors, observed {report.Errors.Count}.");
#endif
    }

    private void CompleteApplyPresentation(OperationReport report, bool acceptAsApplied)
    {
        // A partial write or an unconfirmed readback must retain the pending
        // baseline. REFRESH can reconcile values once Windows confirms them.
        if (acceptAsApplied && report.Succeeded && report.Warnings.Count == 0)
            UpdateAllBlocksInitialState();
        UpdateApplyButtonDirtyCount();
    }

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
