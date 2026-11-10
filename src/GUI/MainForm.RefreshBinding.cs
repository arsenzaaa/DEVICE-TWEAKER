namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private int _deviceTopologyGeneration;

    private Dictionary<string, DeviceBlock> SelectReusableCards(IReadOnlyList<DeviceInfo> devices)
    {
        Dictionary<string, DeviceBlock> reusable = new(StringComparer.OrdinalIgnoreCase);
        // Ambiguous IDs are never assigned an existing editor.
        Dictionary<string, DeviceBlock> existing = _blocks
            .GroupBy(block => block.Device.InstanceId, StringComparer.OrdinalIgnoreCase)
            .Where(group => group.Count() == 1).ToDictionary(group => group.Key, group => group.Single(), StringComparer.OrdinalIgnoreCase);
        foreach (var group in devices.GroupBy(device => device.InstanceId, StringComparer.OrdinalIgnoreCase))
        {
            if (string.IsNullOrWhiteSpace(group.Key) || group.Count() != 1 || !existing.TryGetValue(group.Key, out DeviceBlock? block)) continue;
            DeviceInfo device = group.Single();
            if (!block.Group.IsDisposed && block.Group.Parent == _devicesPanel
                && block.Group is DeviceCard card && card.LayoutScale == _uiScale
                && ReferenceEquals(block.CardCpuContext, _cpuInfo)
                && block.CardTopologyGeneration == _deviceTopologyGeneration
                && block.CardDefinition == DeviceCardDefinition.From(device))
                reusable.Add(group.Key, block);
        }
        return reusable;
    }

    private void ReloadRetainedCard(DeviceBlock block, DeviceInfo fresh)
    {
        // The definition matches, so captured immutable device data and handler
        // targets remain valid. Copy only the mutable discovery power values.
        block.Device.UsbSelectiveSuspend = fresh.UsbSelectiveSuspend;
        block.Device.NicPowerSaving = fresh.NicPowerSaving;
        block.NicItrOperationGeneration++;
        block.NicItrEditGeneration++;
        block.NicItrLastReadText = null;
        if (block.NicItrStatusLabel is not null) block.NicItrStatusLabel.ForeColor = _statusInactive;
        if (block.NicItrTimeLabel is not null) block.NicItrTimeLabel.ForeColor = _statusInactive;
        if (block.NicItrBox is not null) block.NicItrBox.Text = "0x0";
        if (block.NicItrStatusLabel is not null) block.NicItrStatusLabel.Text = "current: reading...";
        if (block.NicItrTimeLabel is not null) block.NicItrTimeLabel.Text = "time: reading...";
        block.ImodMapReadAt = null;
        block.ImodMapStale = false;
        block.ImodReadbackRequested = false;
        block.ImodDetailsExpanded = false;
        block.RestoreImodDetailsExpansion?.Invoke(false);
        if (block.ImodMapLabel is not null) block.ImodMapLabel.Text = string.Empty;
        block.IrqCount = null;
        block.InterruptCapabilities = null;
        block.IrqLabel.Text = "IRQ Count: reading...";
        SetInterruptModeResult(block, InterruptModeResult.Pending);
        block.SuppressCpuEvents++;
        block.SuppressImodEvents++;
        try { LoadBlockSettings(block); }
        finally { block.SuppressCpuEvents--; block.SuppressImodEvents--; }
        block.CaptureInitialState();
        if (block.Device.IsTestDevice && IsUsbImodTarget(block.Device))
            RefreshTestImodPreview(block, "test-block-refresh");
        block.RelayoutAction?.Invoke();
    }

    private void InvalidateDeviceTopology()
    {
        _deviceTopologyGeneration++;
        _imodReadbackGeneration++;
        _irqRefreshGeneration++;
        foreach (DeviceBlock block in _blocks) block.NicItrOperationGeneration++;
    }
}
