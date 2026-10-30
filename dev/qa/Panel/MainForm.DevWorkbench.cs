using System.Diagnostics;
using System.Drawing.Imaging;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;
using DeviceTweakerQa;
using Microsoft.Win32;

namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private DevRun? _devRun;
    private readonly Stopwatch _devClock = new();
    private readonly List<Exception> _devQueuedErrors = [];
    private int _devPending;
    private int _devIrqReadDelayMs = 600;
    private string _devIrqReadOutcome = "success";
    private readonly List<string> _devIrqReadInstances = [];

    private (DeviceBlock Block, InterruptModeResult Result)[] ReadTestInterruptModesForCheck(DeviceBlock[] blocks)
    {
        if (!QaRun.IsReadOnly || blocks.Any(block => !block.Device.IsTestDevice)) throw new InvalidOperationException("IRQ model requires read-only synthetic devices");
        lock (_devIrqReadInstances) _devIrqReadInstances.AddRange(blocks.Select(block => block.Device.InstanceId));
        Thread.Sleep(_devIrqReadDelayMs);
        return blocks.Select(block => (block, _devIrqReadOutcome == "failure"
            ? new InterruptModeResult(InterruptMode.Unavailable, "Simulated PCI read failure. No hardware was queried.", DateTimeOffset.Now)
            : CreateTestInterruptMode(block.Device) with { ReadAt = DateTimeOffset.Now })).ToArray();
    }
    [DllImport("user32.dll", EntryPoint = "PrintWindow", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool DevPrintWindow(IntPtr window, IntPtr dc, uint flags);
    [DllImport("user32.dll", EntryPoint = "SendMessageW")]
    private static extern IntPtr DevSendMessage(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetKeyboardState(byte[] state);
    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetKeyboardState(byte[] state);

    private void StartDevWorkbenchIfRequested()
    {
        string? path = Environment.GetEnvironmentVariable("DEVICE_TWEAKER_DEV_SCENARIO");
        if (string.IsNullOrEmpty(path)) return;
        QaDesktop.DemandAssignedDesktop();
        if (!QaRun.IsReadOnly || !_testDevicesOnly || _blocks.Any(b => !b.Device.IsTestDevice))
            throw new InvalidOperationException("Workbench requires a private desktop, locked read-only mode and synthetic devices only");
        BeginInvoke(new Action(async () => await RunDevWorkbenchAsync(path)));
    }

    private async Task RunDevWorkbenchAsync(string path)
    {
        DevScenario scene = DevFiles.Read<DevScenario>(path);
        _devRun = new() { ScenarioId = scene.Id, ScenarioSha256 = DevFiles.Hash(path), ReadOnly = QaRun.IsReadOnly,
            Desktop = QaDesktop.CurrentName(), Environment = $"{Environment.OSVersion.Version}|dpi={DeviceDpi}|font={_blockFont.Name}/{_blockFont.Size}|{scene.Width}x{scene.Height}" };
        _devClock.Restart();
        if (Math.Abs(_uiScale - 1f) > .01f) _devRun.Environment += "|layoutScale=" + _uiScale.ToString(System.Globalization.CultureInfo.InvariantCulture);
        try
        {
            scene.Validate();
            DevResize(scene.Width, scene.Height);
            await Task.Delay(150);
            // Showcase presets configure controls; seed the sandbox backend from those exact initial values.
            using (DevSandboxScope()) foreach (DeviceBlock block in _blocks)
            {
                if (!ApplyTestBlockSettings(block, false, null)) throw new InvalidOperationException("Preset cannot seed virtual settings");
                EnsureTestDeviceState(block.Device).ImodValue = block.ImodBox.Text.Trim();
                EnsureTestDeviceState(block.Device).NicItrValue = block.NicItrBox?.Text.Trim() ?? "default";
            }
            foreach (DevStep step in scene.Steps)
            {
                try
                {
                    await DevExecute(step);
                    if (_devQueuedErrors.Count != 0) throw new AggregateException(_devQueuedErrors);
                    _devRun.Steps.Add(new(step.Id, step.Action, "PASS", _devClock.Elapsed.TotalMilliseconds, step.Value));
                }
                catch (Exception ex)
                {
                    _devRun.Steps.Add(new(step.Id, step.Action, "FAIL", _devClock.Elapsed.TotalMilliseconds, ex.ToString()));
                    try { DevCaptureWindow(step.Id + "-failure", "@active", false); } catch (Exception capture) { _devRun.Exceptions.Add(capture.ToString()); }
                    throw;
                }
                DevFiles.Write(DevArtifact("session.json"), _devRun);
            }
            if (_devPending != 0) throw new InvalidOperationException("Scenario ended with a pending modal action");
            _devRun.Status = "PASS";
            DevFiles.Verify(QaRun.OutputDirectory, _devRun);
        }
        catch (Exception ex) { _devRun.Status = "FAIL"; _devRun.Exceptions.Add(ex.ToString()); }
        finally
        {
            DevFiles.Write(DevArtifact("session.json"), _devRun);
            WorkbenchHtml.Session(QaRun.OutputDirectory, _devRun);
            Environment.ExitCode = _devRun.Status == "PASS" ? 0 : 1;
            foreach (Form dialog in OwnedForms.ToArray()) dialog.Close();
            Close();
        }
    }

    private string DevArtifact(string name)
    { string path = DevFiles.Inside(QaRun.OutputDirectory, name); QaRun.DemandArtifactPath(path); return path; }
    private void DevResize(int width, int height)
    {
        Rectangle screen = Screen.FromControl(this).WorkingArea;
        if (width < 700 || height < 500 || width > screen.Width - 32 || height > screen.Height - 64)
            throw new InvalidOperationException($"Requested client {width}x{height} does not fit private desktop {screen.Size}");
        WindowState = FormWindowState.Normal; ClientSize = new Size(width, height);
        Location = new Point(screen.Left + 8, screen.Top + 8); PerformLayout(); Refresh();
    }
    private Form DevActiveForm() => Application.OpenForms.Cast<Form>().LastOrDefault(f => f.Visible && (f == this || DevOwned(f))) ?? this;
    private bool DevOwned(Form form) { for (Form? f = form; f is not null; f = f.Owner) if (f == this) return true; return false; }
    private static IEnumerable<Control> DevDescendants(Control control)
    { yield return control; foreach (Control child in control.Controls) foreach (Control nested in DevDescendants(child)) yield return nested; }
    private DeviceBlock DevBlock(string device)
    {
        DeviceBlock[] blocks = _blocks.Where(b => b.Device.IsTestDevice && (b.Device.InstanceId.Contains(device, StringComparison.OrdinalIgnoreCase)
            || b.Device.Name.Contains(device, StringComparison.OrdinalIgnoreCase))).ToArray();
        if (string.IsNullOrEmpty(device) || blocks.Length != 1) throw new InvalidOperationException("Device selector must identify exactly one synthetic block: " + device);
        return blocks[0];
    }
    private Control DevTarget(DevStep step)
    {
        if (step.Target == "@main") return this;
        if (step.Target == "@active") return DevActiveForm();
        if (step.Target == "@search") return _searchFilterBox.Inner;
        if (step.Target == "@clear-filter") return _btnFilterClear;
        if (step.Target.StartsWith("active:id:", StringComparison.Ordinal))
        {
            Control[] active = DevDescendants(DevActiveForm()).Where(c => c.Name == step.Target[10..]).ToArray();
            return active.Length == 1 ? active[0] : throw new InvalidOperationException("Active dialog selector must identify one control: " + step.Target);
        }
        if (step.Target.StartsWith("role:", StringComparison.Ordinal))
        {
            DeviceBlock b = DevBlock(step.Device);
            return step.Target[5..] switch
            {
                "title" => b.TitleLabel,
                "msi" => b.MsiCombo, "limit" => b.LimitBox, "priority" => b.PrioCombo, "policy" => b.PolicyCombo,
                "imod" => b.ImodBox, "imod-toggle" => b.ImodAutoCheck, "imod-check" => b.ImodCheckButton!,
                "imod-current" => b.ImodCurrentLabel, "mode" => b.InterruptModeLabel!,
                "mode-refresh" => _interruptModesReadButton,
                _ => throw new InvalidDataException("Unknown device control role"),
            };
        }
        Control[] roots = string.IsNullOrEmpty(step.Device) ? Application.OpenForms.Cast<Form>().Where(f => f == this || DevOwned(f)).Cast<Control>().ToArray() : [DevBlock(step.Device).Group];
        Control[] matches = roots.SelectMany(DevDescendants).Where(c => step.Target.StartsWith("id:", StringComparison.Ordinal) ? c.Name == step.Target[3..]
            : step.Target.StartsWith("text:", StringComparison.Ordinal) && c.Text == step.Target[5..]).ToArray();
        return matches.Length == 1 ? matches[0] : throw new InvalidOperationException($"Selector {step.Target} matched {matches.Length} controls; specify device / unique ID");
    }
    private void DevQueue(Action action)
    {
        _devPending++;
        BeginInvoke(new Action(() => { try { action(); } catch (Exception ex) { _devQueuedErrors.Add(ex); } finally { _devPending--; } }));
    }
    private static void DevEvent(Control control, string method, EventArgs args)
        => typeof(Control).GetMethod(method, BindingFlags.Instance | BindingFlags.NonPublic)!.Invoke(control, [args]);
    private static bool DevInViewport(Control control)
    {
        if (!control.Visible || control.ClientSize.Width <= 0 || control.ClientSize.Height <= 0) return false;
        Rectangle bounds = control.RectangleToScreen(control.ClientRectangle);
        for (Control? parent = control.Parent; parent is not null; parent = parent.Parent)
            if (!parent.RectangleToScreen(parent.ClientRectangle).Contains(bounds)) return false;
        return true;
    }
    private void DevReveal(Control control)
    {
        for (Control? parent = control.Parent; parent is not null; parent = parent.Parent)
            if (parent is ScrollableControl scroll && scroll.AutoScroll) scroll.ScrollControlIntoView(control);
        if (_devicesPanel is not null && _devicesHost is not null && _devicesPanel.Contains(control))
        {
            int top = _devicesPanel.PointToClient(control.PointToScreen(Point.Empty)).Y;
            int offset = top - Math.Max(12, (_devicesHost.ClientSize.Height - control.Height) / 2);
            if (_devicesScroll is not null) _devicesScroll.Value = offset;
            else SetDevicesScrollOffset(offset);
        }
        if (!DevInViewport(control)) throw new InvalidOperationException("Target could not be fully revealed in its viewport");
    }
    private string DevProperty(Control control, string property) => property switch
    {
        "visibleDevices" when control == this => _blocks.Count(b => b.Group.Visible).ToString(),
        "irq-action-count" when control == this => DevDescendants(this).Count(item => item.Name == "IRQ_MODE_CHECK").ToString(),
        "top-action-count" when control == this => _btnApplyRef!.Parent!.Controls.OfType<Button>().Count().ToString(),
        "irq-action-top" when control == this => (DevDescendants(this).Count(item => item.Name == "IRQ_MODE_CHECK") == 1
            && _interruptModesReadButton.Parent == _btnApplyRef!.Parent
            && !_devicesPanel.Contains(_interruptModesReadButton)
            && !DevDescendants(this).Any(item => item.Name is "IRQ_MODES_BAR" or "IRQ_MODES_STATUS"))
            .ToString().ToLowerInvariant(),
        "irq-top-fits" when control == this => (_btnApplyRef!.Parent!.Controls.OfType<Button>() is var topButtons
            && topButtons.Count() == 5 && topButtons.All(DevInViewport)
            && topButtons.All(left => topButtons.All(right => left == right || !left.Bounds.IntersectsWith(right.Bounds)))
            && topButtons.Select(button => button.Top).Distinct().Count() == 1
            && topButtons.Select(button => button.Height).Distinct().Count() == 1
            && topButtons.All(button => TextRenderer.MeasureText(button.Text, button.Font).Width <= button.ClientSize.Width))
            .ToString().ToLowerInvariant(),
        "irq-all-pci-queried-once" when control == this => (_devIrqReadInstances.Count == _blocks.Count(block =>
            block.Device.InstanceId.StartsWith("PCI\\", StringComparison.OrdinalIgnoreCase)) && _blocks
            .Where(block => block.Device.InstanceId.StartsWith("PCI\\", StringComparison.OrdinalIgnoreCase))
            .All(block => _devIrqReadInstances.Count(id => id == block.Device.InstanceId) == 1))
            .ToString().ToLowerInvariant(),
        "filter-toolbar-fits" when control == this => (_filterCategoriesHost is not null && DevInViewport(_searchFilterBox)
            && !_searchFilterBox.Parent!.Bounds.IntersectsWith(_filterCategoriesHost.Bounds)
            && _filterCategoryButtons.All(DevInViewport)).ToString().ToLowerInvariant(),
        "irq-read-count" when control == this => _devIrqReadInstances.Count.ToString(),
        "irq-read-instance" when control == this => string.Join("|", _devIrqReadInstances),
        "readAt" => _blocks.Single(block => block.InterruptModeLabel == control).InterruptMode.ReadAt?.ToString("O") ?? "",
        "queried" => _devIrqReadInstances.Contains(_blocks.Single(block => block.InterruptModeLabel == control).Device.InstanceId).ToString().ToLowerInvariant(),
        "reading" when control == _interruptModesReadButton => _interruptModeCheckRunning.ToString().ToLowerInvariant(),
        "text" => control.Text, "visible" => control.Visible.ToString().ToLowerInvariant(), "enabled" => control.Enabled.ToString().ToLowerInvariant(),
        "focused" => control.Focused.ToString().ToLowerInvariant(), "viewport" => DevInViewport(control).ToString().ToLowerInvariant(),
        "checked" when control is CheckBox check => check.Checked.ToString().ToLowerInvariant(),
        "caret" when control is TextBoxBase editor => editor.SelectionStart.ToString(),
        "editorFits" when control is ThemedTextBox field => (field.Inner.Height >= field.Inner.PreferredHeight
            && field.Inner.Top >= 1 && field.Inner.Bottom < field.ClientSize.Height).ToString().ToLowerInvariant(),
        "value" when control is ThemedDropDownPicker picker => picker.SelectedItem?.ToString() ?? "",
        _ => throw new InvalidDataException("Unsupported assertion property: " + property),
    };
    private async Task DevExecute(DevStep step)
    {
        QaDesktop.DemandAssignedDesktop();
        if (!QaRun.IsReadOnly) throw new InvalidOperationException("Read-only policy lost");
        if (step.Action is "click" or "hover" or "leave" or "focus" or "text" or "select" or "toggle" or "key")
        {
            Control target = DevTarget(step);
            if (!target.Enabled || !DevInViewport(target)) throw new InvalidOperationException("Interaction target is disabled / clipped; use reveal before interacting");
        }
        if (step.Action == "wait")
        {
            Stopwatch timeout = Stopwatch.StartNew();
            while (timeout.ElapsedMilliseconds <= step.DurationMs)
            {
                try { if (DevProperty(DevTarget(step), step.Property) == step.Value) return; } catch (InvalidOperationException) { }
                if (_devQueuedErrors.Count != 0) throw new AggregateException(_devQueuedErrors);
                await Task.Delay(40);
            }
            throw new TimeoutException("Expected state was not observed: " + step.Target + "/" + step.Property + "=" + step.Value);
        }
        switch (step.Action)
        {
            case "capture": DevCaptureWindow(step.Id, step.Target, false, step.Device); break;
            case "reveal": DevReveal(DevTarget(step)); break;
            case "language": UiLanguage.Set(step.Value == "RU" ? UiLanguageCode.Russian : step.Value == "EN" ? UiLanguageCode.English : throw new InvalidDataException("Language must be RU / EN"), persist: false); break;
            case "resize": DevResize(step.Width, step.Height); break;
            case "gallery": DevQueue(ShowTestAdminDialog); break;
            case "dialog": DevQueue(() => DevPreviewDialog(step.Value)); break;
            case "audit-dialog": DevAuditDialog(step.Id); break;
            case "irq-check-model":
                if (_interruptModeCheckRunning || step.Value is not ("success" or "failure") || step.DurationMs < 100)
                    throw new InvalidDataException("IRQ model requires an idle reader, success/failure and a delay of at least 100 ms");
                _devIrqReadDelayMs = step.DurationMs; _devIrqReadOutcome = step.Value;
                lock (_devIrqReadInstances) _devIrqReadInstances.Clear();
                break;
            case "scroll":
                if (!Enum.TryParse(step.Value, out Keys key) || key is not (Keys.Home or Keys.End or Keys.PageUp or Keys.PageDown)) throw new InvalidDataException("Unsupported list key");
                ActiveControl = _btnScanRef; OnMainFormKeyDown(this, new KeyEventArgs(key)); break;
            case "key":
                if (!Enum.TryParse(step.Value, out Keys input)) throw new InvalidDataException("Unknown key");
                if ((input & Keys.Modifiers) != 0) throw new InvalidDataException("Key action supports one unmodified key");
                Control target = DevTarget(step);
                if (target == this) OnMainFormKeyDown(this, new KeyEventArgs(input));
                else if (target is Form && input is Keys.Escape or Keys.Enter)
                    typeof(Form).GetMethod("ProcessDialogKey", BindingFlags.Instance | BindingFlags.NonPublic)!.Invoke(target, [input]);
                else
                {
                    // Only this private UI thread's key table; never physical/system input.
                    byte[] savedKeys = new byte[256];
                    if (!GetKeyboardState(savedKeys)) throw new InvalidOperationException("Cannot read private thread key state.");
                    byte[] keys = (byte[])savedKeys.Clone();
                    foreach (int modifier in new[] { 0x10, 0x11, 0x12, 0xA0, 0xA1, 0xA2, 0xA3, 0xA4, 0xA5 }) keys[modifier] = 0;
                    try
                    {
                        if (!SetKeyboardState(keys)) throw new InvalidOperationException("Cannot set private thread key state.");
                        DevSendMessage(target.Handle, 0x0100, new IntPtr((int)input), new IntPtr(1));
                        DevSendMessage(target.Handle, 0x0101, new IntPtr((int)input), new IntPtr(unchecked((int)0xC0000001)));
                    }
                    finally { if (!SetKeyboardState(savedKeys)) throw new InvalidOperationException("Cannot restore private thread key state."); }
                }
                break;
            case "assert":
                string actual = DevProperty(DevTarget(step), step.Property);
                if (actual != step.Value) throw new InvalidOperationException($"Expected '{step.Value}', observed '{actual}' on {step.Target}/{step.Property}"); break;
            case "click":
                Control button = DevTarget(step);
                if (!button.Visible || !button.Enabled) throw new InvalidOperationException("Click target is unavailable");
                DevQueue(() => { if (button is IButtonControl b) b.PerformClick(); else DevEvent(button, "OnClick", EventArgs.Empty); }); break;
            case "hover": DevEvent(DevTarget(step), "OnMouseEnter", EventArgs.Empty); break;
            case "leave": DevEvent(DevTarget(step), "OnMouseLeave", EventArgs.Empty); break;
            case "focus": if (!DevTarget(step).Focus()) throw new InvalidOperationException("Focus rejected"); break;
            case "text":
                if (DevTarget(step) is not TextBoxBase editor || editor.ReadOnly) throw new InvalidOperationException("Text target is not an editable text box");
                editor.Text = step.Value; break;
            case "select":
                if (DevTarget(step) is ListBox list && int.TryParse(step.Value, out int row) && row >= 0 && row < list.Items.Count)
                { list.SelectedIndex = row; break; }
                if (DevTarget(step) is not ThemedDropDownPicker picker || !picker.Items.Contains(step.Value)) throw new InvalidOperationException("Picker item missing");
                picker.SelectedItem = step.Value; break;
            case "toggle":
                if (DevTarget(step) is not CheckBox check) throw new InvalidOperationException("Toggle target is not a checkbox");
                check.Checked = bool.Parse(step.Value); break;
            case "save-state":
                DevFiles.Id(step.Value);
                string saved = DevArtifact("state/" + step.Value + ".json");
                Directory.CreateDirectory(Path.GetDirectoryName(saved)!);
                File.WriteAllText(saved, JsonSerializer.Serialize(CaptureCurrentTestScenario(step.Value, TestSandboxOperation.Apply, TestSandboxFault.None, false), TestScenarioJsonOptions()));
                DevRecordState("state/" + step.Value + ".json"); break;
            case "restore-state":
                DevFiles.Id(step.Value);
                TestSandboxScenario snapshot = JsonSerializer.Deserialize<TestSandboxScenario>(File.ReadAllText(DevArtifact("state/" + step.Value + ".json")), TestScenarioJsonOptions())!;
                LoadTestSandboxScenario(snapshot);
                foreach (TestScenarioDevice savedDevice in snapshot.Devices)
                {
                    DeviceBlock restored = DevBlock(savedDevice.InstanceId);
                    TestDeviceState state = EnsureTestDeviceState(restored.Device);
                    restored.ImodBox.Text = state.ImodValue;
                    if (restored.NicItrBox is not null) restored.NicItrBox.Text = state.NicItrValue;
                    if (!StatesEqual(savedDevice.InitialState, state) || !TestStateMatchesBlock(state, restored, false))
                        throw new InvalidOperationException("Restored backend / controls differ from saved settings: " + savedDevice.InstanceId);
                }
                break;
            case "apply-model":
                DeviceBlock block = DevBlock(step.Device);
                using (DevSandboxScope()) if (!ApplyTestBlockSettings(block, false, null)) throw new InvalidOperationException("Virtual apply failed");
                DevPersistModel(block, step.Value); break;
            case "model-error": DevModelError(step.Value); break;
            case "audit-editors": await DevAuditEditors(step.Id); break;
            case "transition":
                void CheckObservedState()
                {
                    if (string.IsNullOrEmpty(step.Value)) return;
                    string observed = DevProperty(DevTarget(step), step.Property);
                    if (observed != step.Value)
                        throw new InvalidOperationException($"Transition {step.Id}: expected '{step.Value}', observed '{observed}' on {step.Target}/{step.Property}");
                }
                CheckObservedState();
                DevCaptureWindow(step.Id + "-frame-000", step.Target, true, step.Device);
                await DevExecute(step.Trigger!);
                Stopwatch interval = Stopwatch.StartNew(); int count = 1;
                // Native snapshots can take longer than the requested observation window on a busy desktop.
                // Always retain enough real samples to check change and settling; timestamps expose that overhead.
                while (count < 3 || interval.ElapsedMilliseconds < step.DurationMs)
                {
                    await Task.Delay(step.IntervalMs);
                    CheckObservedState();
                    DevCaptureWindow(step.Id + "-frame-" + count++.ToString("D3"), step.Target, true, step.Device);
                }
                DevCapture[] frames = _devRun!.Captures.Where(c => c.Timeline && c.Id.StartsWith(step.Id + "-frame-", StringComparison.Ordinal)).ToArray();
                if (frames.Length < 3) throw new InvalidOperationException("Transition has fewer than three observed frames");
                bool changed = frames.Select(f => f.Sha256).Distinct().Count() > 1;
                if (step.RequireChange && !changed) throw new InvalidOperationException("Expected transition did not change pixels");
                if (step.RequireSettled && frames[^1].Sha256 != frames[^2].Sha256) throw new InvalidOperationException("Transition did not settle in the observation window");
                TransitionEvidence.Write(QaRun.OutputDirectory, step.Id, frames, step.IntervalMs);
                DevCaptureWindow(step.Id + "-settled", step.Target, false, step.Device); break;
            default: throw new InvalidDataException("Unsupported action: " + step.Action);
        }
        await Task.Delay(70);
    }
    private IDisposable DevSandboxScope()
    { _testSandbox.Begin(TestSandboxFault.None); return new DevScope(() => _testSandbox.End()); }
    private sealed class DevScope(Action end) : IDisposable { public void Dispose() => end(); }

    private void DevPersistModel(DeviceBlock block, string name)
    {
        DevFiles.Id(name);
        DevRegistry values = new(DevArtifact("state/" + name + ".registry.json"));
        TestDeviceState state = EnsureTestDeviceState(block.Device);
        DeviceSettingsWriter.SetMsi(values, state.MsiEnabled);
        DeviceSettingsWriter.SetMsiLimit(values, state.MsiLimit);
        DeviceSettingsWriter.SetPriority(values, state.Priority);
        DeviceSettingsWriter.SetAffinity(values, state.Policy, state.AffinityMask, 8);
        var actual = DevFiles.Read<Dictionary<string, DevRegistryValue>>(values.Path);
        if (!actual.TryGetValue("MSISupported", out var msi) || msi.Kind != "DWord" || msi.Json != (state.MsiEnabled ? "1" : "0")
            || !actual.TryGetValue("DevicePolicy", out var policy) || policy.Json != (state.AffinityMask == 0 || state.Policy == 0 ? "0" : state.Policy.ToString())
            || (state.MsiLimit.HasValue ? !actual.TryGetValue("MessageNumberLimit", out var limit) || limit.Json != state.MsiLimit.Value.ToString() : actual.ContainsKey("MessageNumberLimit")))
            throw new InvalidOperationException("File-backed settings readback does not match the virtual UI state");
        DevRecordState("state/" + name + ".registry.json");
    }
    private void DevModelError(string name)
    {
        DevFiles.Id(name);
        if (name is "dtimod-incompatible" or "dtimod-pending")
        {
            bool pending = false;
            string expected = name == "dtimod-incompatible" ? DriverRuntime.RestartMessage
                : DriverLoadGate.CheckJournal("11111111-1111-1111-1111-111111111111|pending", "11111111-1111-1111-1111-111111111111");
            try
            {
                DriverReadyPolicy.Ensure(true,
                    () => { if (name == "dtimod-incompatible") throw new InvalidOperationException(expected); return false; },
                    () => { if (pending) throw new InvalidOperationException(expected); },
                    () => { pending = true; return false; },
                    () => throw new InvalidOperationException("Unexpected SCM call in pending-load model"));
                throw new InvalidOperationException("Expected DTIMOD model failure was not raised");
            }
            catch (InvalidOperationException ex) when (ex.Message == expected)
            {
                string relative = "state/" + name + ".exception.json";
                DevFiles.Write(DevArtifact(relative), new { expected = true, boundary = "production DriverReadyPolicy / fake loader", exception = ex.ToString(), driverLoaded = false });
                DevRecordState(relative);
                DevQueue(() => ShowThemedInfo(FormatImodUnavailableUserMessage(ex.Message), "USB IMOD"));
            }
            return;
        }
        DevRegistry values = new(DevArtifact("state/" + name + ".registry.json")) { FailWrites = true };
        try { DeviceSettingsWriter.SetMsi(values, true); throw new InvalidOperationException("Expected backend exception was not raised"); }
        catch (IOException ex)
        {
            string relative = "state/" + name + ".exception.json";
            DevFiles.Write(DevArtifact(relative), new { expected = true, boundary = "file-backed fake registry / production DeviceSettingsWriter", exception = ex.ToString() });
            DevRecordState(relative);
            DevQueue(() => ShowThemedInfo(ex.Message, "SANDBOX ERROR"));
        }
    }
    private sealed record DevRegistryValue(string Kind, string Json);
    private void DevRecordState(string relative)
    {
        _devRun!.StateFiles.RemoveAll(s => s.Path == relative);
        _devRun.StateFiles.Add(new(relative, DevFiles.Hash(DevArtifact(relative))));
    }
    private sealed class DevRegistry(string path) : IRegistryValues
    {
        public string Path => path;
        public bool FailWrites { get; init; }
        private readonly Dictionary<string, DevRegistryValue> _values = [];
        public void Set(string name, object value, RegistryValueKind kind)
        {
            if (FailWrites) throw new IOException("Sandbox settings write failed (injected I/O error).");
            _values[name] = new(kind.ToString(), JsonSerializer.Serialize(value)); DevFiles.Write(path, _values);
        }
        public void Delete(string name) { _values.Remove(name); DevFiles.Write(path, _values); }
    }
    private Bitmap DevCaptureClient(Form window)
    {
        Bitmap bitmap = new(window.ClientSize.Width, window.ClientSize.Height, PixelFormat.Format32bppArgb);
        using Graphics graphics = Graphics.FromImage(bitmap);
        IntPtr dc = graphics.GetHdc();
        try { if (!DevPrintWindow(window.Handle, dc, 3)) throw new InvalidOperationException("PrintWindow client capture failed"); }
        finally { graphics.ReleaseHdc(dc); }
        return bitmap;
    }

    private static bool DevIsBlankCapture(Bitmap bitmap)
    {
        int min = 765, max = 0, visible = 0, white = 0, samples = 0;
        for (int y = 0; y < bitmap.Height; y += 6) for (int x = 0; x < bitmap.Width; x += 6)
        {
            Color pixel = bitmap.GetPixel(x, y); int sum = pixel.R + pixel.G + pixel.B;
            min = Math.Min(min, sum); max = Math.Max(max, sum); samples++;
            if (sum > 90) visible++; if (sum > 720) white++;
        }
        return max - min < 45 || visible < 15 || white / (double)samples > .6;
    }

    private void DevCaptureWindow(string id, string selector, bool timeline, string device = "")
    {
        DevFiles.Id(id);
        Control control = DevTarget(new() { Target = selector, Device = device });
        Form window = control as Form ?? control.FindForm() ?? throw new InvalidOperationException("Capture target has no window");
        if (!window.Visible || (window != this && !DevOwned(window))) throw new InvalidOperationException("Capture window is not owned / visible");
        // Windows inherits keyboard cue visibility from prior input. Set the owned
        // window's real UI state explicitly; never hide focus differences in PNGs.
        // UIS_CLEAR | ((UISF_HIDEFOCUS | UISF_HIDEACCEL) << 16).
        DevSendMessage(window.Handle, 0x0127, new IntPtr(0x00030002), IntPtr.Zero);
        int uiState = DevSendMessage(window.Handle, 0x0129, IntPtr.Zero, IntPtr.Zero).ToInt32();
        if ((uiState & 3) != 0) throw new InvalidOperationException("Owned capture window did not expose keyboard focus cues");
        window.Refresh();
        double captureTime = _devClock.Elapsed.TotalMilliseconds;
        Bitmap captured = DevCaptureClient(window);
        // A dialog that has just opened can still be unpainted: two short retries, then
        // a blank image is a failure.
        for (int attempt = 1; DevIsBlankCapture(captured) && attempt < 3; attempt++)
        {
            captured.Dispose();
            Application.DoEvents();
            Thread.Sleep(150);
            window.Refresh();
            captureTime = _devClock.Elapsed.TotalMilliseconds;
            captured = DevCaptureClient(window);
        }
        using Bitmap bitmap = captured;
        string png = "images/" + id + ".png", tree = "trees/" + id + ".json";
        Directory.CreateDirectory(Path.GetDirectoryName(DevArtifact(png))!); bitmap.Save(DevArtifact(png), ImageFormat.Png);
        var nodes = DevDescendants(window).Select(c => new { id = c.Name, type = c.GetType().Name, text = c.Text,
            source = GetSourceControlText(c), c.Visible, c.Enabled, c.Focused, inViewport = DevInViewport(c), bounds = c.Bounds.ToString(),
            checkedState = c is CheckBox check ? check.Checked : (bool?)null,
            editorMetrics = c is ThemedTextBox field ? new { bounds=field.Inner.Bounds.ToString(),
                field.Inner.PreferredHeight, fontHeight=field.Inner.Font.Height } : null,
            selected = c is ThemedDropDownPicker picker ? picker.SelectedItem?.ToString() : null }).ToArray();
        DevFiles.Write(DevArtifact(tree), new { method = "WinForms control tree", uiState, client = window.ClientSize.ToString(), nodes });
        if (DevIsBlankCapture(bitmap))
            throw new InvalidDataException("Native capture is blank / mostly white; retained for diagnosis");
        _devRun!.Captures.Add(new(id, png, DevFiles.Hash(DevArtifact(png)), bitmap.Width, bitmap.Height, captureTime,
            tree, DevFiles.Hash(DevArtifact(tree)), timeline));
        // Check each visible technical block, not only the whole window's white fraction.
        foreach (InfoTextBox info in DevDescendants(window).OfType<InfoTextBox>().Where(c => c.Visible))
        {
            if (info.BackColor.R + info.BackColor.G + info.BackColor.B > 150) continue;
            Rectangle area = window.RectangleToClient(info.RectangleToScreen(info.ClientRectangle));
            for (Control? parent = info.Parent; parent is not null; parent = parent.Parent)
                area.Intersect(window.RectangleToClient(parent.RectangleToScreen(parent.ClientRectangle)));
            area.Intersect(new Rectangle(Point.Empty, bitmap.Size));
            int light = 0, count = 0;
            for (int y = area.Top + 2; y < area.Bottom - 2; y += 4)
                for (int x = area.Left + 2; x < area.Right - 2; x += 4)
                {
                    Color pixel = bitmap.GetPixel(x, y); count++;
                    if (pixel.R > 210 && pixel.G > 210 && pixel.B > 210) light++;
                }
            if (count >= 20 && light > count / 2)
                throw new InvalidDataException($"Technical block background became white: {id}, area={area}, white={light}/{count}. Native frame retained.");
        }
        if (_interruptModeCheckRunning && (!Enabled || _devicesBusyDepth <= 0
            || InterruptModeActionButtons(this).Any(button => button.Enabled)
            || _interruptModesReadButton.Enabled))
            throw new InvalidDataException("IRQ CHECK must keep the form readable and block all action buttons until the reader completes.");
    }
}
