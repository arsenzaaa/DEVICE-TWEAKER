using System.Diagnostics;
using DeviceTweakerQa;

namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private async Task DevAuditToggleMotionAsync(string id)
    {
        if (!QaRun.IsReadOnly) throw new InvalidOperationException("Motion audit requires read-only QA.");
        using Form preview = new ThemedDialogForm()
        {
            Name = "TOGGLE_MOTION_AUDIT", Text = "Toggle component - synthetic input",
            ClientSize = new Size(440, 145), StartPosition = FormStartPosition.CenterParent,
            FormBorderStyle = FormBorderStyle.FixedDialog, ShowInTaskbar = false,
            BackColor = _bgForm, ForeColor = _fgMain, Font = _technicalFont,
        };
        using ThemedCheckBox toggle = new()
        {
            Name = "MOTION_TOGGLE", Location = new Point(26, 70), Size = new Size(230, 26),
            Text = "Disabled", SyncCheckedStateText = true, BackColor = _bgForm,
            ForeColor = _fgMain,
        };
        preview.Controls.Add(new Label { Text = "120 ms / immediate state", AutoSize = true, Location = new Point(26, 26) });
        preview.Controls.Add(toggle);
        LocalizeControlTree(preview);
        preview.Show(this);
        await Task.Delay(70);
        int events = 0;
        toggle.CheckedChanged += (_, _) => events++;
        List<object> samples = [];
        void Capture(string suffix)
        {
            float position = toggle.QaKnobPosition;
            DevCaptureWindow(id + "-" + suffix, "@active", false);
            using (Bitmap current = new(DevArtifact("images/" + id + "-" + suffix + ".png")))
            {
                int ink = 0;
                for (int y = 20; y < 48; y++) for (int x = 24; x < 320; x++)
                    if (current.GetPixel(x, y).GetBrightness() > .4f) ink++;
                if (ink < 30) throw new InvalidOperationException("Native capture omitted the component heading.");
            }
            if (suffix != "off")
            {
                using Bitmap reference = new(DevArtifact("images/" + id + "-off.png"));
                using Bitmap frame = new(DevArtifact("images/" + id + "-" + suffix + ".png"));
                for (int y = 20; y < 48; y++) for (int x = 24; x < 320; x++)
                    if (reference.GetPixel(x, y) != frame.GetPixel(x, y))
                        throw new InvalidOperationException("Static dialog text disappeared/changed during native capture: " + suffix);
            }
            samples.Add(new { suffix, timeMs = _devClock.Elapsed.TotalMilliseconds, position, toggle.Checked, events });
        }
        Capture("off");
        long started = Stopwatch.GetTimestamp();
        toggle.Checked = true;
        if (!toggle.Checked || events != 1 || toggle.Text != UiLanguage.Text("Enabled"))
            throw new InvalidOperationException("Logical state/caption waited for the animation.");
        Capture("start");
        await Task.Delay(30);
        float middle = toggle.QaKnobPosition;
        if (middle is <= 0 or >= 1) throw new InvalidOperationException("No intermediate native toggle position observed.");
        Capture("middle");
        // Reverse while moving. Events are synchronous and there is no queued click.
        float beforeReverse = toggle.QaKnobPosition;
        toggle.Checked = false;
        if (Math.Abs(toggle.QaKnobPosition - beforeReverse) > .15f || events != 2)
            throw new InvalidOperationException("Rapid reversal jumped or delayed the logical event.");
        Capture("reverse");
        await Task.Delay(200);
        if (toggle.Checked || toggle.QaKnobPosition != 0 || toggle.QaAnimationTimerRunning)
            throw new InvalidOperationException("Reversed animation did not stop at its final state.");
        Capture("settled");
        toggle.QaAnimationPreference = false;
        toggle.Checked = true;
        if (toggle.QaKnobPosition != 1 || toggle.QaAnimationTimerRunning || events != 3)
            throw new InvalidOperationException("Reduced-motion mode did not complete immediately.");
        Capture("reduced-motion");
        toggle.QaAnimationPreference = true;
        toggle.Checked = false;
        toggle.Enabled = false;
        if (toggle.QaAnimationTimerRunning || toggle.QaKnobPosition != 0)
            throw new InvalidOperationException("Disabling a control left an animation running.");
        Capture("disabled");
        toggle.Enabled = true;
        toggle.Checked = true;
        toggle.Visible = false;
        if (toggle.QaAnimationTimerRunning) throw new InvalidOperationException("Hidden control kept its timer running.");
        toggle.Visible = true;
        Capture("shown");
        toggle.Checked = false;
        toggle.Dispose();
        if (toggle.QaAnimationTimerRunning) throw new InvalidOperationException("Disposed control kept its timer running.");
        string relative = "state/" + id + ".json";
        DevFiles.Write(DevArtifact(relative), new { samples, events, elapsedMs = Stopwatch.GetElapsedTime(started).TotalMilliseconds,
            scope = "Native isolated component; default production motion policy plus QA-only local motion-off check; Windows settings unchanged",
            windowsAnimationPreference = NativeUser32.ClientAreaAnimationsEnabled(), systemHighContrast = SystemInformation.HighContrast });
        DevRecordState(relative);
        preview.Close();
    }

    private async Task DevAuditUiMotionAsync(string id)
    {
        if (!QaRun.IsReadOnly || !_testDevicesOnly || _blocks.Any(b => !b.Device.IsTestDevice))
            throw new InvalidOperationException("UI motion audit requires synthetic read-only devices.");
        List<object> samples = [];
        void Capture(string suffix)
        {
            DevCaptureWindow(id + "-" + suffix, "@main", false);
            samples.Add(new { suffix, timeMs = _devClock.Elapsed.TotalMilliseconds,
                offset = _devicesScroll.Value, target = _devicesScroll.TargetValue,
                scrolling = _devicesScroll.IsScrolling, panelTop = _devicesPanel.Top });
            if (-_devicesPanel.Top != _devicesScroll.Value)
                throw new InvalidOperationException("Content and scrollbar lost synchronization.");
        }
        async Task Settled()
        {
            Stopwatch timeout = Stopwatch.StartNew();
            while (_devicesScroll.IsScrolling && timeout.ElapsedMilliseconds < 1000) await Task.Delay(15);
            if (_devicesScroll.IsScrolling || _devicesScroll.Value != _devicesScroll.TargetValue)
                throw new InvalidOperationException("Scroll did not settle at its target.");
        }
        ResetDevicesScroll();
        Capture("top");
        DeviceBlock usb = _blocks.First(b => b.Kind == DeviceKind.USB && b.PowerSavingCheck is not null);
        ThemedCheckBox power = (ThemedCheckBox)usb.PowerSavingCheck!;
        DevReveal(power);
        bool originalPower = power.Checked;
        using (DevSandboxScope()) power.Checked = !originalPower;
        if (power.Checked == originalPower || !usb.IsDirty)
            throw new InvalidOperationException("Actual card toggle did not update its draft immediately.");
        await Task.Delay(30);
        if (power.QaKnobPosition is <= 0 or >= 1)
            throw new InvalidOperationException("Actual Power Saving control has no intermediate position.");
        Capture("power-middle");
        using (DevSandboxScope()) power.Checked = originalPower;
        await Task.Delay(180);
        if (power.QaAnimationTimerRunning || power.Checked != originalPower)
            throw new InvalidOperationException("Actual card toggle did not settle after reversal.");
        Capture("power-settled");
        ResetDevicesScroll();
        // Send a synthetic managed wheel event to the production scrollbar.
        DevEvent(_devicesScroll, "OnMouseWheel", new HandledMouseEventArgs(MouseButtons.None, 0, 0, 0, -120));
        int destination = _devicesScroll.TargetValue;
        if (destination < UiScale(120)) throw new InvalidOperationException("A wheel notch advances less than 120 scaled pixels.");
        if (!_devicesScroll.IsScrolling || destination <= 0 || _devicesScroll.Value != 0)
            throw new InvalidOperationException("Wheel movement jumped instead of animating.");
        await Task.Delay(30);
        if (_devicesScroll.Value <= 0 || _devicesScroll.Value >= destination)
            throw new InvalidOperationException("No intermediate scroll frame observed.");
        Capture("wheel-middle");
        _devicesScroll.ScrollWheel(-240);
        if (_devicesScroll.TargetValue != destination + 2 * _devicesScroll.SmallChange)
            throw new InvalidOperationException("Repeated wheel input was lost.");
        _devicesScroll.ScrollWheel(120);
        if (_devicesScroll.TargetValue != destination + _devicesScroll.SmallChange)
            throw new InvalidOperationException("Wheel reversal did not update the destination.");
        await Settled();
        Capture("wheel-settled");
        // Cancel in place on thumb capture, rather than completing a queued glide.
        ResetDevicesScroll();
        _devicesScroll.ScrollWheel(-120);
        DevEvent(_devicesScroll, "OnMouseDown", new MouseEventArgs(MouseButtons.Left, 1, _devicesScroll.Width / 2, 2, 0));
        if (_devicesScroll.IsScrolling || _devicesScroll.Value != 0)
            throw new InvalidOperationException("Thumb drag failed to cancel pending wheel motion.");
        DevEvent(_devicesScroll, "OnMouseMove", new MouseEventArgs(MouseButtons.Left, 0, _devicesScroll.Width / 2, 100, 0));
        int dragged = _devicesScroll.Value;
        if (_devicesScroll.IsScrolling || dragged <= 0 || _devicesScroll.TargetValue != dragged)
            throw new InvalidOperationException("Thumb drag did not update the viewport immediately.");
        DevEvent(_devicesScroll, "OnMouseUp", new MouseEventArgs(MouseButtons.Left, 1, 0, 100, 0));
        _devicesScroll.ScrollWheel(-120);
        if (_devicesScroll.TargetValue != dragged + _devicesScroll.SmallChange)
            throw new InvalidOperationException("Wheel after thumb dragging used a stale destination.");
        await Settled();
        ResetDevicesScroll();
        for (int i = 0; i < 12; i++) _devicesScroll.ScrollWheel(-1);
        if (_devicesScroll.TargetValue != _devicesScroll.SmallChange / 10)
            throw new InvalidOperationException("Partial wheel deltas were lost.");
        await Settled();
        SetCategoryFilter("GPU");
        if (!_devicesScroll.IsScrolling) throw new InvalidOperationException("Category navigation did not animate.");
        await Task.Delay(35);
        Capture("category-middle");
        // Retarget before arrival; no stale category animation may run afterward.
        SetCategoryFilter("NETWORK");
        await Settled();
        DeviceBlock network = _blocks.First(b => b.Kind is DeviceKind.NET_NDIS or DeviceKind.NET_CX);
        if (!DevInViewport(network.TitleLabel)) throw new InvalidOperationException("Category target did not enter the viewport.");
        Capture("category-settled");
        ThemedButton button = (ThemedButton)_filterCategoryButtons.First(b => (string?)b.Tag == "GPU");
        DevEvent(button, "OnMouseEnter", EventArgs.Empty);
        button.Refresh();
        await Task.Delay(25);
        if (!button.IsAnimating) throw new InvalidOperationException("Button hover did not animate.");
        Capture("hover-middle");
        DevEvent(button, "OnMouseLeave", EventArgs.Empty);
        button.Refresh();
        await Task.Delay(180);
        if (button.IsAnimating) throw new InvalidOperationException("Button feedback timer did not stop.");
        Capture("hover-settled");
        // Shrinking a scroll range must clamp a running animation and remain idle.
        using (Form component = new() { Size = new Size(120, 320), ShowInTaskbar = false })
        using (ThemedScrollBar bar = new() { Height = 200, Maximum = 1000, ViewportSize = 100 })
        {
            component.Controls.Add(bar); component.Show(this);
            bar.ScrollTo(900);
            bar.Maximum = 150;
            await Task.Delay(190);
            if (bar.IsScrolling || bar.Value > 50 || bar.TargetValue > 50)
                throw new InvalidOperationException("Range change left a stale scroll destination.");
            bar.ScrollTo(50); bar.Visible = false;
            if (bar.IsScrolling) throw new InvalidOperationException("Hidden scrollbar kept animating.");
            bar.Visible = true; bar.ScrollTo(0); bar.Enabled = false;
            if (bar.IsScrolling) throw new InvalidOperationException("Disabled scrollbar kept animating.");
            bar.Enabled = true; bar.ScrollTo(50); bar.Dispose();
            if (bar.IsScrolling) throw new InvalidOperationException("Disposed scrollbar kept animating.");
            component.Close();
        }
        SetCategoryFilter("ALL");
        await Settled();
        _devicesScroll.ScrollWheel(-120);
        ResetDevicesScroll();
        await Task.Delay(190);
        if (_devicesScroll.IsScrolling || _devicesScroll.Value != 0 || _devicesPanel.Top != 0)
            throw new InvalidOperationException("Reset left stale animation frames queued.");
        Capture("reset");
        string relative = "state/" + id + ".json";
        DevFiles.Write(DevArtifact(relative), new { samples,
            scope = "Production UI controls and category route with synthetic devices; no physical input or driver access",
            windowsAnimationPreference = NativeUser32.ClientAreaAnimationsEnabled() });
        DevRecordState(relative);
    }

    private async Task DevIrqClickBurstAsync(string id)
    {
        if (!QaRun.IsReadOnly || !_testDevicesOnly || _blocks.Any(b => !b.Device.IsTestDevice))
            throw new InvalidOperationException("IRQ burst requires synthetic devices.");
        _interruptModesReadButton.PerformClick();
        if (!_interruptModeCheckRunning) throw new InvalidOperationException("IRQ operation did not start synchronously.");
        string[] previous = _blocks.Select(b => b.InterruptModeLabel?.Text ?? "").ToArray();
        // Exercise the actual button plus a re-entrant caller while its lease is held.
        for (int i = 0; i < 20; i++)
        {
            _interruptModesReadButton.PerformClick();
            await CheckInterruptModeAsync();
        }
        if (!_blocks.Select(b => b.InterruptModeLabel?.Text ?? "").SequenceEqual(previous))
            throw new InvalidOperationException("IRQ fields changed to a transient busy caption.");
        int pulses = 0;
        using System.Windows.Forms.Timer heartbeat = new() { Interval = 20 };
        heartbeat.Tick += (_, _) => pulses++;
        heartbeat.Start();
        Stopwatch watch = Stopwatch.StartNew();
        while (_interruptModeCheckRunning && watch.ElapsedMilliseconds < 5000) await Task.Delay(15);
        heartbeat.Stop();
        if (_interruptModeCheckRunning || _operations.IsBusy || !_interruptModesReadButton.Enabled)
            throw new InvalidOperationException("IRQ operation did not release its controls/lease.");
        if (DevProperty(this, "irq-all-pci-queried-once") != "true")
            throw new InvalidOperationException("Rapid clicks queued or duplicated a PCI query.");
        if (_devIrqReadDelayMs >= 500 && pulses < 3)
            throw new InvalidOperationException("UI message loop did not respond during the simulated reader.");
        string relative = "state/" + id + ".json";
        DevFiles.Write(DevArtifact(relative), new { repeatButtonClicks = 20, reentrantCalls = 20,
            pciQueries = _devIrqReadInstances.Count, pulses, modeledReaderDelayMs = _devIrqReadDelayMs,
            outcome = _devIrqReadOutcome, scope = "Shared IRQ UI command with synthetic read backend; no driver loaded" });
        DevRecordState(relative);
    }
}
