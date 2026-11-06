using System.Diagnostics;
using System.Runtime.InteropServices;
using DeviceTweakerQa;

namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    [DllImport("user32.dll")]
    private static extern IntPtr GetSystemMenu(IntPtr window, [MarshalAs(UnmanagedType.Bool)] bool revert);
    [DllImport("user32.dll")]
    private static extern uint GetMenuState(IntPtr menu, uint item, uint flags);

    private async Task DevAuditMotionLifecycleAsync(string id)
    {
        if (!QaRun.IsReadOnly) throw new InvalidOperationException("Motion audit requires read-only GUI.");
        IntPtr menu = GetSystemMenu(Handle, false);
        if (GetMenuState(menu, 0x1E00, 0) != uint.MaxValue || GetMenuState(menu, 0xF060, 0) == uint.MaxValue)
            throw new InvalidOperationException("Obsolete animation command remains, or Close is missing.");
        int workersBefore = UiFrameClock.QaWorkers;
        using Form preview = new ThemedDialogForm
        {
            Name = "MOTION_AUDIT", ClientSize = new Size(440, 180), ShowInTaskbar = false,
            BackColor = _bgForm, ForeColor = _fgMain, StartPosition = FormStartPosition.CenterParent,
        };
        using ThemedCheckBox toggle = new()
        {
            Text = "Motion lifecycle", Location = new Point(24, 40), Size = new Size(240, 26),
            BackColor = _bgForm, ForeColor = _fgMain,
        };
        using ThemedScrollBar bar = new()
        { Location = new Point(400, 10), Height = 160, Maximum = 1000, ViewportSize = 100 };
        preview.Controls.Add(toggle); preview.Controls.Add(bar);
        preview.Show(this);
        await Task.Delay(50);
        if (!UiMotion.CanAnimate(toggle) || !UiMotion.CanAnimate(bar))
            throw new InvalidOperationException("Visible enabled controls do not animate by default.");
        toggle.Checked = true; bar.ScrollTo(800);
        if (!toggle.QaAnimationTimerRunning || !bar.IsScrolling)
            throw new InvalidOperationException("Default motion did not start.");
        preview.Hide();
        // Native parent hiding can suppress child VisibleChanged until the next frame.
        await Task.Delay(70);
        if (toggle.QaAnimationTimerRunning || bar.IsScrolling || toggle.QaKnobPosition != 1 || bar.Value != 800)
            throw new InvalidOperationException($"Hiding transitions: toggle visible={toggle.Visible}, running={toggle.QaAnimationTimerRunning}, position={toggle.QaKnobPosition}; bar visible={bar.Visible}, running={bar.IsScrolling}, value={bar.Value}, target={bar.TargetValue}.");
        long posted = UiFrameClock.QaPostedFrames;
        await Task.Delay(100);
        if (posted != UiFrameClock.QaPostedFrames || UiFrameClock.QaQueuedFrames != 0)
            throw new InvalidOperationException("Idle motion continued posting frames.");
        DevCaptureWindow(id + "-idle", "@main", false);
        preview.Show(this);
        toggle.Checked = false; bar.ScrollTo(0);
        await Task.Delay(30);
        if (toggle.QaKnobPosition is <= 0 or >= 1 || !bar.IsScrolling)
            throw new InvalidOperationException("Visible controls did not resume motion.");
        Thread.Sleep(45); // Block only the isolated QA UI to test queue coalescing.
        if (UiFrameClock.QaQueuedFrames > 1)
            throw new InvalidOperationException("Frame callbacks accumulated during UI congestion.");
        preview.Dispose();
        Stopwatch timeout = Stopwatch.StartNew();
        while (UiFrameClock.QaWorkers > workersBefore && timeout.ElapsedMilliseconds < 1000) await Task.Delay(10);
        if (UiFrameClock.QaWorkers > workersBefore || UiFrameClock.QaQueuedFrames != 0)
            throw new InvalidOperationException("Owner disposal left an animation worker or callback.");
        DevCaptureWindow(id + "-closed", "@main", false);
        string relative = "state/" + id + ".json";
        DevFiles.Write(DevArtifact(relative), new { workersBefore, workersAfter = UiFrameClock.QaWorkers,
            queuedFrames = UiFrameClock.QaQueuedFrames, removedCommand = true,
            scope = "Always-on default policy, native controls and real frame clock; no system settings modified" });
        DevRecordState(relative);
    }
}
