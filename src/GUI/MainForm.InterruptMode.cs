namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private bool _interruptModeCheckRunning;
    private Button _interruptModesReadButton = null!;

    private void UpdateInterruptModesStatus()
    {
        if (_interruptModesReadButton is null) return;
        if (!_interruptModeCheckRunning)
            _interruptModesReadButton.Enabled = _devicesBusyDepth == 0 && _blocks.Any(block =>
                block.Device.InstanceId.StartsWith("PCI\\", StringComparison.OrdinalIgnoreCase));
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
        if (_interruptModeCheckRunning || _devicesBusyDepth > 0) return;
        DeviceBlock[] snapshot = _blocks.Where(block =>
            block.Device.InstanceId.StartsWith("PCI\\", StringComparison.OrdinalIgnoreCase)).ToArray();
        if (snapshot.Length == 0) return;
        bool simulated = snapshot.All(block => block.Device.IsTestDevice);
        if (!simulated && QaRun.TryBlock("PCI MODE CHECK")) return;
        _interruptModeCheckRunning = true;
        // Disable actions individually. Disabling the form makes native RichEdit
        // blocks paint a system background and darkens otherwise readable labels.
        Button[] actions = InterruptModeActionButtons(this).Where(button => button.Enabled).ToArray();
        _devicesBusyDepth++;
        int generation = _irqRefreshGeneration;
        try
        {
            foreach (Button button in actions) button.Enabled = false;
            UpdateInterruptModesStatus();
            foreach (DeviceBlock block in snapshot)
                if (block.InterruptModeLabel is { IsDisposed: false } label)
                    label.Text = UiLanguage.Text("reading...");
            WriteLog($"IRQ.MODE.BEGIN: scope=all-pci devices={snapshot.Length} simulated={simulated}");
            var results = await Task.Run(() =>
            {
                (DeviceBlock Block, InterruptModeResult Result)[] Failed(string detail) => snapshot
                    .Select(block => (block, new InterruptModeResult(InterruptMode.Unavailable, detail, DateTimeOffset.Now))).ToArray();
                if (simulated)
                {
#if DEVELOPER_QA
                    return ReadTestInterruptModesForCheck(snapshot);
#else
                    return Failed("Device simulation requires the developer QA build.");
#endif
                }
                if (!IsAdministrator()) return Failed("administrator privileges required");
                // An explicit refresh prepares the current payload if it has not been staged yet.
                if (!EnsureImodDriverOnDisk(persistDriver: true, out string driverPath, out string? error)
                    || !ImodDriverContext.TryInitialize(driverPath, WriteLog, out ImodDriverContext? context, out error))
                    return Failed(error ?? "DTIMOD load failed");
                using (context)
                    return snapshot.Select(block =>
                    {
                        try { return (block, PciInterruptReader.Query(context!.DriverHandle, block.Device.InstanceId)); }
                        catch (Exception ex)
                        {
                            WriteLog($"IRQ.MODE.ERROR: instance=\"{SanitizeLogValue(block.Device.InstanceId)}\" {ex}");
                            return (block, new InterruptModeResult(InterruptMode.Unavailable, ex.Message, DateTimeOffset.Now));
                        }
                    }).ToArray();
            });
            if (!IsDisposed && generation == _irqRefreshGeneration)
                foreach (var (block, result) in results)
                    if (_blocks.Contains(block)) SetInterruptModeResult(block, result);
            WriteLog($"IRQ.MODE.END: scope=all-pci devices={results.Length}");
        }
        catch (Exception ex)
        {
            WriteLog($"IRQ.MODE.ERROR: {ex}");
            if (!IsDisposed && generation == _irqRefreshGeneration)
                foreach (DeviceBlock block in snapshot)
                    if (_blocks.Contains(block)) SetInterruptModeResult(block, new(InterruptMode.Unavailable, ex.Message, DateTimeOffset.Now));
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
        if (_interruptModeCheckRunning && e.CloseReason == CloseReason.UserClosing) e.Cancel = true;
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
