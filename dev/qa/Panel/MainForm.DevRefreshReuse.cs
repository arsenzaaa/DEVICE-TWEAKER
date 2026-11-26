using System.Diagnostics;
using DeviceTweakerQa;

namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private bool _devDiscoveryFailure;

    private async Task DevAuditRefreshReuseAsync(string id)
    {
        if (!QaRun.IsReadOnly || !_testDevicesOnly || _blocks.Any(block => !block.Device.IsTestDevice))
            throw new InvalidOperationException("Refresh identity audit requires synthetic read-only devices.");
        Dictionary<string, Control> before = _blocks.ToDictionary(block => block.Device.InstanceId, block => block.Group,
            StringComparer.OrdinalIgnoreCase);
        long started = Stopwatch.GetTimestamp();
        OperationReport report = await RefreshBlocksAsync(includeImodReadback: false, preserveEdits: true);
        int retained = _blocks.Count(block => before.TryGetValue(block.Device.InstanceId, out Control? control)
            && ReferenceEquals(control, block.Group));
        string relative = "state/" + id + ".json";
        DevFiles.Write(DevArtifact(relative), new { before = before.Count, after = _blocks.Count,
            retained, errors = report.Errors.Count, elapsedMs = Stopwatch.GetElapsedTime(started).TotalMilliseconds });
        DevRecordState(relative);
        if (report.Errors.Count != 0 || _blocks.Count != before.Count || retained != before.Count)
            throw new InvalidOperationException($"Unchanged refresh recreated cards: retained={retained}/{before.Count}, errors={report.Errors.Count}.");
    }

    private async Task DevAuditRefreshLifecycleAsync(string id)
    {
        if (!QaRun.IsReadOnly || !_testDevicesOnly || _blocks.Any(block => !block.Device.IsTestDevice))
            throw new InvalidOperationException("Refresh lifecycle audit requires synthetic read-only devices.");
        List<object> checks = [];
        void Check(string name, bool passed)
        {
            checks.Add(new { name, passed });
            if (!passed) throw new InvalidOperationException("Refresh lifecycle: " + name);
        }
        Dictionary<string, Control> Cards() => _blocks.ToDictionary(block => block.Device.InstanceId,
            block => block.Group, StringComparer.OrdinalIgnoreCase);
        bool Same(Dictionary<string, Control> before, DeviceBlock block)
            => before.TryGetValue(block.Device.InstanceId, out Control? card) && ReferenceEquals(card, block.Group);
        static DeviceInfo Changed(DeviceInfo device, string property, object value)
        {
            DeviceInfo copy = device.CopyForDiscovery();
            typeof(DeviceInfo).GetProperty(property)!.SetValue(copy, value);
            return copy;
        }
        async Task<OperationReport> Refresh(bool edits = true)
        {
            OperationReport report = await RefreshBlocksAsync(includeImodReadback: false, preserveEdits: edits);
            Check("refresh-completed", report.Errors.Count == 0);
            return report;
        }
        try
        {
            DeviceBlock usb = _blocks.First(block => block.Kind == DeviceKind.USB && IsUsbImodTarget(block.Device));
            string targetId = usb.Device.InstanceId;
            DeviceInfo owner = _testDevices.Single(device => device.InstanceId == targetId);
            TestDeviceState state = EnsureTestDeviceState(owner);
            Dictionary<string, Control> before = Cards();
            int reads = _devDataReads;
            usb.LimitBox.Text = "1";
            state.MsiLimit = 2;
            state.Priority = usb.PrioCombo.SelectedItem?.ToString() == "High" ? 1 : 3;
            string expectedPriority = state.Priority == 1 ? "Low" : "High";
            DevReveal(usb.LimitBox);
            await DevWaitForSettledPresentation();
            Check("editor-focused-before-refresh", usb.LimitBox.Focus());
            usb.LimitBox.Select(0, 1);
            DeviceViewport viewportBefore = CaptureDeviceViewport();
            await Refresh();
            usb = _blocks.Single(block => block.Device.InstanceId == targetId);
            Check("unchanged-cards-retained", _blocks.All(block => Same(before, block)));
            Check("settings-read-for-every-card", _devDataReads >= reads + _blocks.Count);
            Check("draft-retained", usb.LimitBox.Text == "1");
            Check("baseline-fresh", usb.InitialLimitText == "2");
            Check("editor-focus-and-selection-retained", usb.LimitBox.Focused && usb.LimitBox.SelectionStart == 0 && usb.LimitBox.SelectionLength == 1);
            DeviceViewport viewportAfter = CaptureDeviceViewport();
            Check("viewport-anchor-retained", viewportBefore.InstanceId == viewportAfter.InstanceId && Math.Abs(viewportBefore.WithinCard - viewportAfter.WithinCard) <= 1);
            Check("unedited-priority-fresh", usb.PrioCombo.SelectedItem?.ToString() == expectedPriority);

            Dictionary<string, string> neighbors = _testDevices.Where(device => device.InstanceId != targetId)
                .ToDictionary(device => device.InstanceId, device => System.Text.Json.JsonSerializer.Serialize(EnsureTestDeviceState(device)));
            OperationReport applied = new();
            using (DevSandboxScope()) await SaveBlockSettings(usb, report: applied);
            Check("apply-report-successful", applied.Errors.Count == 0);
            Check("apply-target-after-retention", EnsureTestDeviceState(owner).MsiLimit == 1);
            Check("apply-neighbors-unchanged", _testDevices.Where(device => device.InstanceId != targetId)
                .All(device => neighbors[device.InstanceId] == System.Text.Json.JsonSerializer.Serialize(EnsureTestDeviceState(device))));

            before = Cards();
            Control removed = usb.Group;
            _testDevices.Remove(owner);
            await Refresh();
            Check("disconnected-card-removed", removed.IsDisposed && _blocks.All(block => block.Device.InstanceId != targetId));
            Check("other-cards-retained", _blocks.All(block => Same(before, block)));
            owner.TestState = state.Clone();
            owner.TestState.MsiLimit = 2;
            _testDevices.Add(owner);
            await Refresh();
            usb = _blocks.Single(block => block.Device.InstanceId == targetId);
            Check("same-id-reconnected-with-fresh-state", !ReferenceEquals(removed, usb.Group) && usb.LimitBox.Text == "2");

            DeviceInfo newPort = Changed(owner, nameof(DeviceInfo.InstanceId), targetId + "-PORT2");
            newPort.TestState = new TestDeviceState { MsiLimit = 2 };
            _testDevices.Add(newPort);
            await Refresh();
            DeviceBlock port = _blocks.Single(block => block.Device.InstanceId == newPort.InstanceId);
            Check("same-name-other-id-independent", port.LimitBox.Text == "2" && !ReferenceEquals(port.Group, usb.Group));
            port.LimitBox.Text = "1";
            using (DevSandboxScope()) await SaveBlockSettings(port);
            Check("apply-uses-id-not-name-or-order", EnsureTestDeviceState(newPort).MsiLimit == 1 && EnsureTestDeviceState(owner).MsiLimit == 2);

            before = Cards();
            DeviceInfo changedRoles = Changed(owner, nameof(DeviceInfo.UsbRoles), owner.UsbRoles + ", Gamepad 1K");
            _testDevices[_testDevices.IndexOf(owner)] = changedRoles;
            await Refresh();
            Check("same-id-new-roles-rebound", !Same(before, _blocks.Single(block => block.Device.InstanceId == targetId)));
            Check("unrelated-definitions-still-retained", _blocks.Where(block => block.Device.InstanceId != targetId).All(block => Same(before, block)));
            owner = changedRoles;

            before = Cards();
            InvalidateDeviceTopology();
            await Refresh();
            Check("pnp-notification-invalidates-old-bindings", _blocks.All(block => !Same(before, block)));

            _devDiscoveryDelayMs = 350;
            Task<OperationReport> pending = RefreshBlocksAsync(includeImodReadback: false);
            await Task.Delay(40);
            Check("discovery-really-pending", _devDiscoveryPending > 0);
            _testDevices.Remove(newPort);
            InvalidateDeviceTopology();
            OperationReport changedDuringScan = await pending;
            Check("changed-scan-retried-with-fresh-input", changedDuringScan.Errors.Count == 0
                && _blocks.All(block => block.Device.InstanceId != newPort.InstanceId));

            before = Cards();
            WriteLog("TEST.SCENARIO.BEGIN: fault=DiscoveryRead");
            _devDiscoveryFailure = true;
            OperationReport failure = await RefreshBlocksAsync(includeImodReadback: false);
            Check("failed-scan-is-not-empty-success", failure.Errors.Count > 0 && _blocks.Count == before.Count
                && _blocks.All(block => Same(before, block)));
            Check("discovery-fault-consumed", !_devDiscoveryFailure);
            WriteLog("TEST.SCENARIO.END: status=PASS assertions=2/0 faultConsumed=True");

            _devRefreshReadbackDelayMs = 350;
            _devRefreshReadbackTimeoutMs = 5000;
            pending = RefreshBlocksAsync(includeImodReadback: false);
            Stopwatch waiting = Stopwatch.StartNew();
            while (_devRefreshReadbackPending == 0 && !pending.IsCompleted && waiting.ElapsedMilliseconds < 3000) await Task.Delay(10);
            Check("readback-really-pending", _devRefreshReadbackPending > 0);
            InvalidateDeviceTopology();
            OperationReport interruptedRead = await pending;
            Check("late-readback-not-confirmed", interruptedRead.Warnings.Count > 0 && _blocks.All(block => block.IrqCount is null));

            await Refresh();
            usb = _blocks.Single(block => block.Device.InstanceId == targetId);
            usb.LimitBox.Text = "1";
            before = Cards();
            await Refresh(edits: false);
            usb = _blocks.Single(block => block.Device.InstanceId == targetId);
            Check("restore-style-refresh-discards-draft", Same(before, usb) && usb.LimitBox.Text == "2" && !usb.IsDirty);

            TestDeviceState usbState = EnsureTestDeviceState(usb.Device);
            usbState.Policy = 4;
            usbState.AffinityMask = 4;
            await Refresh(edits: false);
            usb = _blocks.Single(block => block.Device.InstanceId == targetId);
            ApplyMaskToCpuBoxes(usb, 16);
            usbState.Policy = 2;
            usbState.AffinityMask = 0;
            await Refresh();
            usb = _blocks.Single(block => block.Device.InstanceId == targetId);
            Check("edited-mask-retains-spec-cpu-policy", usb.AffinityMask == 16
                && usb.PolicyCombo.SelectedItem as string == "SpecCPU" && usb.IsDirty);
            await Refresh(edits: false);
            usb = _blocks.Single(block => block.Device.InstanceId == targetId);
            Check("fresh-policy-visible-without-draft", usb.AffinityMask == 0
                && usb.PolicyCombo.SelectedItem as string == "Single" && !usb.IsDirty);

            DeviceBlock? nic = _blocks.FirstOrDefault(block => block.Kind == DeviceKind.NET_NDIS && !block.Device.Wifi);
            if (nic is not null)
            {
                string nicId = nic.Device.InstanceId;
                TestDeviceState nicState = EnsureTestDeviceState(nic.Device);
                nicState.RssBaseCore = 2;
                nicState.RssQueues = 2;
                nicState.NdisMode = "RSS";
                await Refresh(edits: false);
                nic = _blocks.Single(block => block.Device.InstanceId == nicId);
                Check("rss-source-populates-selection", nic.AffinityMask != 0);
                before = Cards();
                nicState.RssBaseCore = null;
                nicState.AffinityMask = 0;
                nicState.Policy = 0;
                nicState.NdisMode = "IRQ";
                await Refresh(edits: false);
                nic = _blocks.Single(block => block.Device.InstanceId == nicId);
                Check("missing-rss-clears-retained-selection", Same(before, nic) && nic.AffinityMask == 0
                    && nic.CpuBoxes.All(box => !box.Checked) && nic.RssBaseCore is null);
            }

            DeviceInfo[] catalog = _testDevices.ToArray();
            before = Cards();
            _testDevices.Clear();
            await Refresh(edits: false);
            Check("empty-discovery-clears-cards", _blocks.Count == 0 && before.Values.All(card => card.IsDisposed));
            _testDevices.AddRange(catalog.Reverse());
            await Refresh(edits: false);
            Check("devices-reappear-after-empty-discovery", _blocks.Count == catalog.Length
                && _blocks.All(block => !Same(before, block)));
            Check("reappearance-binds-by-id", _blocks.All(block => catalog.Any(device => device.InstanceId == block.Device.InstanceId
                && ReferenceEquals(device.TestState, block.Device.TestState))));


        }
        finally
        {
            string relative = "state/" + id + ".json";
            DevFiles.Write(DevArtifact(relative), new { checks, scope = "Production refresh, model discovery changes and virtual settings writes; no physical hardware." });
            DevRecordState(relative);
        }
    }
}
