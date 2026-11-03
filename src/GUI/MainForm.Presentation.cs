namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private sealed record TextPresentation(string Text, string? Tag, Color Color, string? Tooltip);
    private sealed record PickerPresentation(object[] Items);
    private sealed record FocusedEditor(string InstanceId, int[] Path, int? Caret, int SelectionLength);
    private sealed record CardPresentation(
        DeviceInfo Device, DeviceEditorState Baseline, DeviceEditorState Values,
        Dictionary<string, TextPresentation> Text, Dictionary<string, PickerPresentation> Pickers,
        bool[] CpuAutoCheck, bool PolicyEnabled, decimal? RssMaximum, IReadOnlyList<int>? RssOptions,
        int? RssBaseCore, int RssQueueLast, NdisRssRuntimeState? RssRuntime,
        int? IrqCount, PciInterruptCapabilities? Capabilities, InterruptModeResult InterruptMode,
        bool DetailsExpanded, bool ReadbackRequested, DateTime? MapReadAt, bool MapStale, string? NicLastRead);

    private static IEnumerable<(string Key, Control? Control)> PresentationTexts(DeviceBlock block) =>
    [
        ("title", block.TitleLabel), ("policy", block.PolicyLabel), ("irq", block.IrqLabel),
        ("imod-default", block.ImodDefaultLabel), ("imod-current", block.ImodCurrentLabel),
        ("imod-map", block.ImodMapLabel), ("nic-status", block.NicItrStatusLabel),
        ("nic-time", block.NicItrTimeLabel), ("raw-status", block.RawMouseThrottleStatusLabel), ("info", block.InfoLabel),
    ];

    private static IEnumerable<(string Key, ThemedDropDownPicker? Picker)> PresentationPickers(DeviceBlock block) =>
    [
        ("msi", block.MsiCombo), ("priority", block.PrioCombo), ("policy", block.PolicyCombo),
        ("ndis", block.NdisModeCombo), ("imod", block.ImodModeCombo), ("raw", block.RawMouseThrottleCombo),
    ];

    private CardPresentation CaptureCardPresentation(DeviceBlock block) => new(
        block.Device.CopyForDiscovery(), block.InitialEditorState ?? CaptureEditorState(block), CaptureEditorState(block),
        PresentationTexts(block).Where(item => item.Control is not null).ToDictionary(item => item.Key,
            item => new TextPresentation(GetSourceControlText(item.Control!), item.Control!.Tag as string,
                item.Control.ForeColor, _copyToolTip.GetToolTip(item.Control))),
        PresentationPickers(block).Where(item => item.Picker is not null).ToDictionary(item => item.Key,
            item => new PickerPresentation(item.Picker!.Items.ToArray())),
        block.CpuBoxes.Select(box => box.AutoCheck).ToArray(), block.PolicyCombo.Enabled,
        block.RssQueueBox?.Maximum, block.RssQueueOptions?.ToArray(), block.RssBaseCore, block.RssQueueLast,
        block.NdisRssRuntime, block.IrqCount, block.InterruptCapabilities, block.InterruptMode,
        block.ImodDetailsExpanded, block.ImodReadbackRequested, block.ImodMapReadAt, block.ImodMapStale,
        block.NicItrLastReadText);

    private void RestoreCardPresentation(DeviceBlock block, CardPresentation saved)
    {
        bool wasRawRefreshing = _rawMouseThrottleUiRefreshing;
        _rawMouseThrottleUiRefreshing = true;
        try
        {
            block.InterruptCapabilities = saved.Capabilities;
            block.IrqCount = saved.IrqCount;
            block.RssQueueOptions = saved.RssOptions;
            block.NdisRssRuntime = saved.RssRuntime;
            if (block.RssQueueBox is not null && saved.RssMaximum is decimal maximum)
                block.RssQueueBox.Maximum = maximum;
            foreach (var (key, picker) in PresentationPickers(block))
            {
                if (picker is null || !saved.Pickers.TryGetValue(key, out PickerPresentation? state)) continue;
                picker.Items.Clear();
                picker.Items.AddRange(state.Items);
            }
            block.PolicyCombo.Enabled = saved.PolicyEnabled;
            block.PolicyLabel.Visible = block.Kind != DeviceKind.NET_NDIS;
            block.PolicyCombo.Visible = block.Kind != DeviceKind.NET_NDIS;
            ApplyEditorState(block, saved.Values);
            block.RssBaseCore = saved.RssBaseCore;
            block.RssQueueLast = saved.RssQueueLast;
            for (int i = 0; i < Math.Min(block.CpuBoxes.Count, saved.CpuAutoCheck.Length); i++)
                block.CpuBoxes[i].AutoCheck = saved.CpuAutoCheck[i];
            block.ImodReadbackRequested = saved.ReadbackRequested;
            block.ImodDetailsExpanded = saved.DetailsExpanded;
            block.ImodMapReadAt = saved.MapReadAt;
            block.ImodMapStale = saved.MapStale;
            block.NicItrLastReadText = saved.NicLastRead;
            foreach (var (key, control) in PresentationTexts(block))
            {
                if (control is null || !saved.Text.TryGetValue(key, out TextPresentation? state)) continue;
                control.Text = state.Text;
                control.Tag = state.Tag;
                control.ForeColor = state.Color;
                _copyToolTip.SetToolTip(control, state.Tooltip);
            }
            SetInterruptModeResult(block, saved.InterruptMode);
            block.RestoreImodDetailsExpansion?.Invoke(saved.DetailsExpanded);
        }
        finally { _rawMouseThrottleUiRefreshing = wasRawRefreshing; }
    }

    // Rebind existing presentation only. No discovery, startup/profile loading,
    // registry/WMI reads, polling refresh or native readback belongs here.
    private void RebuildDevicePresentation()
    {
        FocusedEditor? focus = CaptureFocusedDeviceEditor();
        CardPresentation[] cards = _blocks.Select(CaptureCardPresentation).ToArray();
        DeviceViewport viewport = CaptureDeviceViewport();
        float previousScale = cards.Length > 0 && _blocks[0].Group is DeviceCard card ? card.LayoutScale : _uiScale;
        bool hadReserved = _reservedCpuPanel is not null;
        bool[] reserved = _reservedCpuPanel?.Tag is ReservedCpuPanelTag tag
            ? tag.Meta.OrderBy(item => item.Index).Select(item => item.Control.Checked).ToArray() : [];
        Interlocked.Increment(ref _irqRefreshGeneration);
        Interlocked.Increment(ref _imodReadbackGeneration);
        _devicesPanel.SuspendLayout();
        try
        {
            DisposeDeviceBlocks();
            for (int i = 0; i < cards.Length; i++) NewDeviceBlock(cards[i].Device, i, presentation: cards[i]);
            _reservedCpuPanel = hadReserved ? NewReservedCpuSetsPanel(reserved) : null;
            if (_reservedCpuPanel is not null) _devicesPanel.Controls.Add(_reservedCpuPanel);
            RebuildFilterCategoryButtons();
            ApplyDeviceFilter();
            RestoreFocusedDeviceEditor(focus);
            float ratio = previousScale > 0 ? _uiScale / previousScale : 1f;
            RestoreDeviceViewport(viewport with
            {
                WithinCard = (int)Math.Round(viewport.WithinCard * ratio),
                AbsoluteOffset = (int)Math.Round(viewport.AbsoluteOffset * ratio),
            });
            UpdateApplyButtonDirtyCount();
            UpdateInterruptModesStatus();
        }
        finally { _devicesPanel.ResumeLayout(); }
        _devicesPanel.Invalidate(true);
        _devicesHost.Invalidate(true);
        WriteLog($"UI.LAYOUT.REBIND: cards={cards.Length} scale={_uiScale} discovery=false readback=false");
    }

    private FocusedEditor? CaptureFocusedDeviceEditor()
    {
        static Control? Find(Control root)
        {
            if (root.Focused) return root;
            if (!root.ContainsFocus) return null;
            foreach (Control child in root.Controls)
                if (Find(child) is { } focused) return focused;
            return null;
        }
        foreach (DeviceBlock block in _blocks)
        {
            if (Find(block.Group) is not { } target) continue;
            List<int> path = [];
            for (Control? current = target; current != block.Group && current?.Parent is { } parent; current = parent)
                path.Add(parent.Controls.GetChildIndex(current));
            path.Reverse();
            return new(block.Device.InstanceId, path.ToArray(), (target as TextBoxBase)?.SelectionStart,
                (target as TextBoxBase)?.SelectionLength ?? 0);
        }
        return null;
    }

    private void RestoreFocusedDeviceEditor(FocusedEditor? saved)
    {
        if (saved is null || _blocks.FirstOrDefault(block => block.Device.InstanceId == saved.InstanceId) is not { } block) return;
        Control target = block.Group;
        foreach (int index in saved.Path)
        {
            if (index < 0 || index >= target.Controls.Count) return;
            target = target.Controls[index];
        }
        if (!target.CanFocus || !target.Focus()) return;
        if (target is TextBoxBase editor && saved.Caret is int caret)
        {
            int start = Math.Clamp(caret, 0, editor.TextLength);
            editor.Select(start, Math.Clamp(saved.SelectionLength, 0, editor.TextLength - start));
        }
    }

    private void DisposeDeviceBlocks()
    {
        foreach (Control control in _devicesPanel.Controls.Cast<Control>().ToArray())
        {
            _copyToolTip.RemoveAll(control);
            control.Dispose();
        }
        _copyToolTip.RemoveDisposedRegistrations();
        _copyToolTip.RemoveDetachedRegistrations(this);
        _devicesPanel.Controls.Clear();
        _devicesPanel.Location = new Point(0, 0);
        if (_devicesScroll is not null) _devicesScroll.Value = 0;
        _blocks.Clear();
        _reservedCpuPanel = null;
    }
}
