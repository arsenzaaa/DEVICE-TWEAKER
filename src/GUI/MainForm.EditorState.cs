namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private sealed record EditorDraft(DeviceEditorState Baseline, DeviceEditorState Values, bool DetailsExpanded, bool ReadbackRequested);
    private sealed record DeviceViewport(string? InstanceId, int WithinCard, int AbsoluteOffset);

    internal static DeviceEditorState CaptureEditorState(DeviceBlock block) => new()
    {
        AffinityMask = block.AffinityMask,
        MsiIndex = block.MsiCombo.SelectedIndex,
        LimitText = block.LimitBox.Text,
        PriorityIndex = block.PrioCombo.SelectedIndex,
        PolicyIndex = block.PolicyCombo.SelectedIndex,
        PowerSaving = block.PowerSavingCheck?.Checked,
        ImodText = block.ImodBox.Text,
        ImodEnabled = block.ImodAutoCheck.Checked,
        RawMouseEnabled = block.RawMouseThrottleCheck?.Checked,
        RawMouseIndex = block.RawMouseThrottleCombo?.SelectedIndex ?? -1,
        NdisModeIndex = block.NdisModeCombo?.SelectedIndex,
        RssQueues = block.RssQueueBox?.Value,
        NicItrText = block.NicItrBox?.Text,
    };

    private Dictionary<string, EditorDraft> CaptureEditorDrafts()
    {
        Dictionary<string, EditorDraft> drafts = new(StringComparer.OrdinalIgnoreCase);
        foreach (DeviceBlock block in _blocks)
        {
            if (block.InitialEditorState is not { } initial) continue;
            // A completed read is the NIC baseline, even when it finished after card creation.
            initial = initial with { NicItrText = block.NicItrLastReadText ?? initial.NicItrText };
            drafts[block.Device.InstanceId] = new(initial, CaptureEditorState(block), block.ImodDetailsExpanded, block.ImodReadbackRequested);
        }
        return drafts;
    }

    private DeviceViewport CaptureDeviceViewport()
    {
        int offset = Math.Max(0, -_devicesPanel.Top);
        DeviceBlock? anchor = _blocks.FirstOrDefault(block => block.Group.Visible && block.Group.Bottom > offset);
        return new(anchor?.Device.InstanceId, anchor is null ? 0 : offset - anchor.Group.Top, offset);
    }

    private void RestoreDeviceViewport(DeviceViewport viewport)
    {
        DeviceBlock? anchor = _blocks.FirstOrDefault(block => block.Group.Visible
            && string.Equals(block.Device.InstanceId, viewport.InstanceId, StringComparison.OrdinalIgnoreCase));
        int offset = anchor is null ? viewport.AbsoluteOffset : anchor.Group.Top + viewport.WithinCard;
        SyncDevicesScrollBar();
        if (_devicesScroll is not null) _devicesScroll.Value = Math.Max(0, offset);
        else SetDevicesScrollOffset(offset);
    }

    private void RestoreEditorDrafts(IReadOnlyDictionary<string, EditorDraft> drafts)
    {
        bool wasRawRefreshing = _rawMouseThrottleUiRefreshing;
        _rawMouseThrottleUiRefreshing = true;
        try
        {
            foreach (DeviceBlock block in _blocks)
            {
                if (!drafts.TryGetValue(block.Device.InstanceId, out EditorDraft? draft)) continue;
                DeviceEditorState current = CaptureEditorState(block);
                DeviceEditorState merged = DeviceEditorState.MergeDraft(current, draft.Baseline, draft.Values);
                block.SuppressCpuEvents++;
                block.SuppressImodEvents++;
                try
                {
                    static void Select(ThemedDropDownPicker picker, int index)
                    {
                        if (index >= -1 && index < picker.Items.Count) picker.SelectedIndex = index;
                    }
                    if (block.NdisModeCombo is not null && merged.NdisModeIndex is int mode) Select(block.NdisModeCombo, mode);
                    Select(block.MsiCombo, merged.MsiIndex);
                    block.LimitBox.Text = merged.LimitText;
                    Select(block.PrioCombo, merged.PriorityIndex);
                    Select(block.PolicyCombo, merged.PolicyIndex);
                    if (block.PowerSavingCheck is not null && merged.PowerSaving is bool power) block.PowerSavingCheck.Checked = power;
                    if (block.RssQueueBox is not null && merged.RssQueues is decimal queues)
                        block.RssQueueBox.Value = Math.Clamp(queues, block.RssQueueBox.Minimum, block.RssQueueBox.Maximum);
                    ApplyMaskToCpuBoxes(block, merged.AffinityMask);
                    if (block.Kind == DeviceKind.NET_NDIS)
                    {
                        block.RssBaseCore = GetFirstCheckedCore(block);
                        if (block.RssQueueBox is not null) block.RssQueueLast = (int)block.RssQueueBox.Value;
                    }
                    block.ImodBox.Text = merged.ImodText;
                    block.ImodAutoCheck.Checked = merged.ImodEnabled;
                    if (block.RawMouseThrottleCheck is not null && merged.RawMouseEnabled is bool mouse) block.RawMouseThrottleCheck.Checked = mouse;
                    if (block.RawMouseThrottleCombo is not null) Select(block.RawMouseThrottleCombo, merged.RawMouseIndex);
                    if (block.NicItrBox is not null && merged.NicItrText is string nic && nic != block.NicItrBox.Text)
                        block.NicItrBox.Text = nic;
                    block.ImodReadbackRequested = draft.ReadbackRequested;
                    block.ImodDetailsExpanded = draft.DetailsExpanded;
                    block.RestoreImodDetailsExpansion?.Invoke(draft.DetailsExpanded);
                }
                finally
                {
                    block.SuppressImodEvents--;
                    block.SuppressCpuEvents--;
                }
                RecalcAffinityMask(block);
                UpdateImodSelectorsFromText(block);
                UpdateBlockInfoText(block);
                block.RelayoutAction?.Invoke();
                block.CheckIsDirty();
            }
        }
        finally { _rawMouseThrottleUiRefreshing = wasRawRefreshing; }
        LayoutBlocks();
    }
}
