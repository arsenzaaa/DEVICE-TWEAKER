using System.Diagnostics;
using DeviceTweakerQa;

namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private readonly List<object> _devUiStages = [];
    private void DevMarkUiStage(string stage) => _devUiStages.Add(new { stage, at = UiMotion.Now });
    private int _devLayoutCalls;
    private readonly List<double> _devCpuGridTimes = [];
    private readonly List<double> _devCardTimes = [];
    private async Task DevMeasureFramePacingAsync(string id, bool requireBudget)
    {
        if (!QaRun.IsReadOnly || !_testDevicesOnly || _blocks.Any(b => !b.Device.IsTestDevice))
            throw new InvalidOperationException("Pacing audit requires read-only synthetic devices.");
        UiFrameClock.QaRefreshRate = 300;
        using Form preview = new ThemedDialogForm
        {
            Text = "Native frame pacing - synthetic input", ClientSize = new Size(440, 140),
            ShowInTaskbar = false, BackColor = _bgForm, ForeColor = _fgMain,
            StartPosition = FormStartPosition.CenterParent,
        };
        using ThemedCheckBox toggle = new()
        {
            Text = "Animation pacing", Location = new Point(24, 60), Size = new Size(260, 26),
            BackColor = _bgForm, ForeColor = _fgMain,
        };
        preview.Controls.Add(toggle); preview.Show(this);
        await Task.Delay(80);
        List<double> intervals = [];
        List<object> frames = [];
        double previous = double.NaN;
        toggle.QaFramePainted += (_, _) =>
        {
            double now = UiMotion.Now;
            float position = toggle.QaKnobPosition;
            frames.Add(new { now, position });
            if (!double.IsNaN(previous)) intervals.Add(now - previous);
            previous = now;
        };
        for (int i = 0; i < 8; i++)
        {
            previous = double.NaN;
            toggle.Checked = !toggle.Checked;
            await Task.Delay(200);
        }
        preview.Close();
        UiFrameClock.QaRefreshRate = null;
        List<double> heartbeats = [];
        List<object> heartbeatSamples = [];
        _devUiStages.Clear();
        double heartbeatAt = UiMotion.Now;
        using System.Windows.Forms.Timer heartbeat = new() { Interval = 16 };
        heartbeat.Tick += (_, _) => { double now = UiMotion.Now; heartbeats.Add(now - heartbeatAt); heartbeatSamples.Add(new { now, gap = now - heartbeatAt }); heartbeatAt = now; };
        heartbeat.Start();
        await Task.Delay(60);
        List<object> refreshes = [];
        for (int i = 0; i < 2; i++)
        {
            if (i == 0) _cpuCheckboxWidthCache.Clear();
            _devCpuGridTimes.Clear(); _devCardTimes.Clear();
            int priorLayouts = _devLayoutCalls;
            long started = Stopwatch.GetTimestamp();
            OperationReport report = await RefreshBlocksAsync(includeImodReadback: false, preserveEdits: true);
            double elapsed = Stopwatch.GetElapsedTime(started).TotalMilliseconds;
            double trailingGap = UiMotion.Now - heartbeatAt;
            heartbeats.Add(trailingGap);
            refreshes.Add(new { elapsedMs = elapsed, report.Errors.Count, cards = _blocks.Count, layoutCalls = _devLayoutCalls - priorLayouts, cpuGridMs = _devCpuGridTimes.Sum(), cardMs = _devCardTimes.Sum(), measuredCpuLabels = _cpuCheckboxWidthCache.Count, createdCards = _devCardTimes.Count, retainedCards = _blocks.Count - _devCardTimes.Count, maxCardMs = _devCardTimes.DefaultIfEmpty(0).Max() });
            if (report.Errors.Count != 0) throw new InvalidOperationException("Synthetic refresh failed during timing audit.");
            await Task.Delay(80);
        }
        heartbeat.Stop();
        double[] sorted = intervals.Where(value => value > .1).Order().ToArray();
        if (sorted.Length < 8) throw new InvalidOperationException("Too few actual paint samples.");
        double median = sorted[sorted.Length / 2], p95 = sorted[(int)Math.Ceiling(sorted.Length * .95) - 1];
        string relative = "state/" + id + ".json";
        DevFiles.Write(DevArtifact(relative), new { frames, paintIntervalsMs = intervals, medianPaintMs = median,
            p95PaintMs = p95, refreshes, heartbeatGapsMs = heartbeats, heartbeatSamples, uiStages = _devUiStages, maxHeartbeatGapMs = heartbeats.Max(),
            scope = "Native off-screen paint callbacks and synthetic refresh UI occupancy; not physical display FPS", nativeRefreshHz = NativeAnimationClock.ReadRefreshRate(this), modeledRefreshHz = 300 });
        DevRecordState(relative);
        DevCaptureWindow(id + "-after-refresh", "@main", false);
        if (requireBudget && median > 8)
            throw new InvalidOperationException("Animation paint cadence exceeds the high-refresh budget: median=" + median.ToString("F2"));
    }

    // Measures real UI event handlers on synthetic devices. Delays below only
    // let the private desktop paint between samples; they are not app latency.
    private async Task DevMeasureUiLatencyAsync(string id)
    {
        if (!QaRun.IsReadOnly || !_testDevicesOnly || _blocks.Any(b => !b.Device.IsTestDevice))
            throw new InvalidOperationException("UI measurements require read-only synthetic devices.");
        int readsBefore = _devDataReads;
        List<object> groups = [];
        async Task Measure(string name, int count, Action<int> action)
        {
            List<double> elapsed = [], allocated = [];
            for (int i = 0; i < count; i++)
            {
                long memory = GC.GetAllocatedBytesForCurrentThread();
                long started = Stopwatch.GetTimestamp();
                action(i);
                elapsed.Add(Stopwatch.GetElapsedTime(started).TotalMilliseconds);
                allocated.Add(GC.GetAllocatedBytesForCurrentThread() - memory);
                await Task.Delay(30);
            }
            double[] sorted = elapsed.Order().ToArray();
            groups.Add(new { name, handlerMs = elapsed, medianMs = sorted[sorted.Length / 2],
                p95Ms = sorted[(int)Math.Ceiling(sorted.Length * .95) - 1], maxMs = sorted[^1], allocatedBytes = allocated });
        }

        string[] queries = ["USB", "", "NVMe", "", "no-such-device", ""];
        await Measure("search", 24, i => _searchFilterBox.Text = queries[i % queries.Length]);
        string[] categories = ["USB", "GPU", "NETWORK", "STORAGE", "ALL"];
        await Measure("navigation", 15, i => _filterCategoryButtons.Single(b => Equals(b.Tag, categories[i % categories.Length])).PerformClick());
        DeviceBlock Usb() => _blocks.First(b => b.Kind == DeviceKind.USB && b.PowerSavingCheck is not null);
        await Measure("power-saving-draft", 12, _ => Usb().PowerSavingCheck!.Checked = !Usb().PowerSavingCheck!.Checked);
        await Measure("mouse-throttle-draft", 12, _ => Usb().RawMouseThrottleCheck!.Checked = !Usb().RawMouseThrottleCheck!.Checked);
        await Measure("imod-details", 6, _ => Usb().ImodDetailsButton!.PerformClick());
        await Measure("language", 4, i => UiLanguage.Set(i % 2 == 0 ? UiLanguageCode.Russian : UiLanguageCode.English, persist: false));
        await Measure("presentation-rebind", 2, _ => RebuildDevicePresentation());
        await Measure("animation-preference", 12, i => { _ = SystemInformation.HighContrast; _ = NativeUser32.ClientAreaAnimationsEnabled(); });
        if (_devDataReads != readsBefore) throw new InvalidOperationException("A presentation interaction triggered data reads.");
        string path = DevArtifact("state/" + id + ".json");
        DevFiles.Write(path, new { scope = "Synchronous UI handler cost, not device latency or display FPS",
            devices = _blocks.Count, logicalCpus = _maxLogical, readsBefore, readsAfter = _devDataReads, groups });
        DevRecordState("state/" + id + ".json");
    }
}
