namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private bool _interruptModeCheckRunning;
    private Button _interruptModesReadButton = null!;
    private string? _interruptModeStatusKey;
    private object[] _interruptModeStatusArgs = [];

    private void SetInterruptModeStatus(string? key, params object[] args)
    {
        _interruptModeStatusKey = key;
        _interruptModeStatusArgs = args;
        RenderInterruptModeButton();
    }

    private void RenderInterruptModeButton()
    {
        if (_interruptModesReadButton is null || _interruptModesReadButton.IsDisposed) return;
        string caption = UiResources.Get(_interruptModeCheckRunning ? UiText.ReadingIrq : UiText.ReadIrq);
        if (_interruptModeCheckRunning && _interruptModeStatusKey == UiText.IrqPreparing) caption = "IRQ: DTIMOD...";
        else if (!_interruptModeCheckRunning && _interruptModeStatusKey == UiText.IrqCompleted
            && _interruptModeStatusArgs.Length == 2 && _interruptModeStatusArgs[0] is int total
            && _interruptModeStatusArgs[1] is int unavailable)
            caption = string.Format(System.Globalization.CultureInfo.CurrentCulture,
                UiResources.Get(UiText.IrqReadCount), total - unavailable, total);
        _interruptModesReadButton.Text = caption;
        ((ThemedButton)_interruptModesReadButton).DisabledForeColor = _interruptModeCheckRunning ? _fgMain : Color.FromArgb(125, 125, 132);
        if (_interruptModeStatusKey is not null)
            _copyToolTip.SetToolTip(_interruptModesReadButton, string.Format(System.Globalization.CultureInfo.CurrentCulture,
                UiResources.Get(_interruptModeStatusKey), _interruptModeStatusArgs));
        else _copyToolTip?.SetToolTip(_interruptModesReadButton, string.Empty);
    }

    private void UpdateInterruptModesStatus()
    {
        if (_interruptModesReadButton is null) return;
        if (!_interruptModeCheckRunning)
            _interruptModesReadButton.Enabled = _devicesBusyDepth == 0 && !_operations.IsBusy && _blocks.Any(block =>
                block.Device.InstanceId.StartsWith("PCI\\", StringComparison.OrdinalIgnoreCase));
        SetInterruptModeStatus(_interruptModeStatusKey, _interruptModeStatusArgs);
        _interruptModesReadButton.AccessibleName = UiLanguage.Text("Read interrupt modes for all PCI devices");
    }

    private void SetInterruptModeResult(DeviceBlock block, InterruptModeResult result)
    {
        block.InterruptMode = result;
        if (block.InterruptModeLabel is not { IsDisposed: false } label) return;
        label.Text = result.Text;
        label.ForeColor = _mutedText;
        // ThemedToolTip translates on display; retain English for later RU / EN switches.
        string detail = result.Detail;
        if (result.ReadAt.HasValue) detail += $"\n{result.ReadAt.Value:HH:mm:ss}";
        _copyToolTip.SetToolTip(label, detail);
        WriteLog($"IRQ.MODE: instance=\"{SanitizeLogValue(block.Device.InstanceId)}\" mode={result.Mode} simulated={block.Device.IsTestDevice} readAt={result.ReadAt:O} detail=\"{SanitizeLogValue(result.Detail)}\"");
    }

    private async Task CheckInterruptModeAsync()
    {
        if (_interruptModeCheckRunning || _devicesBusyDepth > 0 || _operations.IsBusy) return;
        DeviceBlock[] snapshot = _blocks.Where(block =>
            block.Device.InstanceId.StartsWith("PCI\\", StringComparison.OrdinalIgnoreCase)).ToArray();
        if (snapshot.Length == 0) return;
        DeviceInfo[] devices = snapshot.Select(block => block.Device).ToArray();
        string[] instances = devices.Select(device => device.InstanceId).ToArray();
        bool simulated = snapshot.All(block => block.Device.IsTestDevice);
        if (!simulated && QaRun.TryBlock("PCI MODE CHECK")) return;
        if (!simulated && !ConfirmDriverUse()) return;
        using IDisposable? operation = TryBeginDeviceOperation("READ IRQ MODE");
        if (operation is null) return;
        _interruptModeCheckRunning = true;
        // Disable actions individually. Disabling the form makes native RichEdit
        // blocks paint a system background and darkens otherwise readable labels.
        Button[] actions = InterruptModeActionButtons(this).Where(button => button.Enabled).ToArray();
        _devicesBusyDepth++;
        int generation = _irqRefreshGeneration;
        SetInterruptModeStatus(UiText.IrqReading, snapshot.Length);
        IProgress<string> progress = new Progress<string>(key =>
        {
            if (_interruptModeCheckRunning && !IsDisposed && !Disposing && generation == _irqRefreshGeneration)
                SetInterruptModeStatus(key, snapshot.Length);
        });
        try
        {
            foreach (Button button in actions) button.Enabled = false;
            UpdateInterruptModesStatus();
            // Keep the last result readable while its replacement is being queried.
            WriteLog($"IRQ.MODE.BEGIN: scope=all-pci devices={snapshot.Length} simulated={simulated}");
            var results = await Task.Run(() =>
            {
                PciInterruptReading[] Failed(string detail) => PciInterruptBatch.Unavailable(instances, detail);
                if (simulated)
                {
#if DEVELOPER_QA
                    return ReadTestInterruptModesForCheck(devices);
#else
                    return Failed("Device simulation requires the developer QA build.");
#endif
                }
                if (!IsAdministrator()) return Failed("administrator privileges required");
                progress.Report(UiText.IrqPreparing);
                // An explicit refresh prepares the current payload if it has not been staged yet.
                if (!EnsureImodDriverOnDisk(persistDriver: true, out string driverPath, out string? error)
                    || !ImodDriverContext.TryInitialize(driverPath, WriteLog, KduDriverLoader.TryLoadImodDriverWithKduFallback, out ImodDriverContext? context, out error))
                    return Failed(error ?? "DTIMOD load failed");
                progress.Report(UiText.IrqReading);
                using (context)
                    return PciInterruptBatch.Read(instances,
                        instance => PciInterruptReader.Query(context!.DriverHandle, instance),
                        (instance, ex) => WriteLog($"IRQ.MODE.ERROR: instance=\"{SanitizeLogValue(instance)}\" {ex}"));
            });
            if (!IsDisposed && generation == _irqRefreshGeneration)
            {
                for (int index = 0; index < results.Length; index++)
                {
                    DeviceBlock block = snapshot[index];
                    PciInterruptReading reading = results[index];
                    if (_blocks.Contains(block) && string.Equals(block.Device.InstanceId, reading.InstanceId, StringComparison.OrdinalIgnoreCase))
                        SetInterruptModeResult(block, reading.Result);
                }
                SetInterruptModeStatus(UiText.IrqCompleted, results.Length,
                    results.Count(item => item.Result.Mode == InterruptMode.Unavailable));
            }
            WriteLog($"IRQ.MODE.END: scope=all-pci devices={results.Length}");
        }
        catch (Exception ex)
        {
            WriteLog($"IRQ.MODE.ERROR: {ex}");
            if (!IsDisposed && generation == _irqRefreshGeneration)
            {
                foreach (DeviceBlock block in snapshot)
                    if (_blocks.Contains(block)) SetInterruptModeResult(block, new(InterruptMode.Unavailable, ex.Message, DateTimeOffset.Now));
                SetInterruptModeStatus(UiText.IrqCompleted, snapshot.Length, snapshot.Length);
            }
        }
        finally
        {
            _interruptModeCheckRunning = false;
            _devicesBusyDepth--;
            foreach (Button button in actions) if (!button.IsDisposed) button.Enabled = true;
            if (!IsDisposed) UpdateInterruptModesStatus();
        }
    }

    private static IEnumerable<Button> InterruptModeActionButtons(Control parent)
    {
        foreach (Control child in parent.Controls)
        {
            if (child is Button button) yield return button;
            foreach (Button nested in InterruptModeActionButtons(child)) yield return nested;
        }
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        if ((_startupConfigWorkPending || _interruptModeCheckRunning || (_operations.IsBusy && _activeOperationName != "REFRESH")) && e.CloseReason == CloseReason.UserClosing) e.Cancel = true;
        base.OnFormClosing(e);
    }

    private static InterruptModeResult CreateTestInterruptMode(DeviceInfo device)
    {
#if DEVELOPER_QA
        InterruptMode mode = device.TestInterruptMode switch
        {
            "Line" => InterruptMode.Line, "MSI" => InterruptMode.Msi, "MSI-X" => InterruptMode.MsiX,
            "MSI/MSI-X" => InterruptMode.Message, "No IRQ" => InterruptMode.None,
            "Read failed" => InterruptMode.Unavailable, "Not read" => InterruptMode.NotRead,
            _ => device.TestMsiStatus.Equals("Disabled", StringComparison.OrdinalIgnoreCase) ? InterruptMode.Line
                : device.Kind is DeviceKind.GPU or DeviceKind.NET_CX or DeviceKind.STOR ? InterruptMode.MsiX : InterruptMode.Msi,
        };
        return new(mode, "Simulated interrupt mode. No hardware was queried.");
#else
        return InterruptModeResult.Pending;
#endif
    }
}
