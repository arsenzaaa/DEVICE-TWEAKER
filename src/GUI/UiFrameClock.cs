using System.Runtime.CompilerServices;

namespace DeviceTweakerCS;

// One clock per form. A single coalesced callback renders only the latest state.
// Sleeping is off the UI thread; native timer objects are disarmed while idle.
internal sealed class UiFrameClock
{
    private static readonly ConditionalWeakTable<Form, UiFrameClock> Clocks = new();
    private readonly Form _form;
    private readonly NativeAnimationClock.FrameTimer _timer = new();
    private readonly AutoResetEvent _wake = new(false);
    private readonly HashSet<Subscription> _active = [];
    private readonly List<Subscription> _frame = [];
    private volatile bool _disposed;
    private volatile bool _failed;
    private volatile bool _hasWork;
    private double _periodMs = 1000.0 / 60;
    private int _pending;
    internal bool HighResolution => _timer.HighResolution;
#if DEVELOPER_QA
    internal static bool QaHasActiveAnimations(Form form)
        => Clocks.TryGetValue(form, out UiFrameClock? clock) && clock._active.Count != 0;
    internal static int? QaRefreshRate { get; set; }
    internal static int QaWorkers;
    internal static int QaQueuedFrames;
    internal static long QaPostedFrames;
#endif

    private UiFrameClock(Form form)
    {
        _form = form;
        form.Disposed += OnDisposed;
        new Thread(WaitForFrames) { IsBackground = true, Name = "DEVICE TWEAKER animation clock" }.Start();
    }

    internal static Subscription? Start(Control owner, Func<bool, bool> draw)
    {
        if (owner.FindForm() is not { IsDisposed: false } form) return null;
        try
        {
            UiFrameClock clock = Clocks.GetValue(form, root => new(root));
            if (clock._disposed) return null;
            int rate = NativeAnimationClock.ReadRefreshRate(owner);
#if DEVELOPER_QA
            rate = QaRefreshRate ?? rate;
#endif
            Volatile.Write(ref clock._periodMs, 1000.0 / Math.Clamp(rate, 30, 1000));
            Subscription item = new(clock, draw);
            clock._active.Add(item);
            clock._hasWork = true;
            clock.Wake();
            return item;
        }
        catch (Exception error) when (error is System.ComponentModel.Win32Exception or InvalidOperationException)
        {
            AppDiagnostics.Warn("UI.ANIMATION.CLOCK", error);
            return null; // Settings still change immediately if a timer cannot be created.
        }
    }

    private void WaitForFrames()
    {
#if DEVELOPER_QA
        Interlocked.Increment(ref QaWorkers);
#endif
        try
        {
            WaitHandle[] waits = [_wake, _timer];
            while (!_disposed)
            {
                if (!_hasWork) { _wake.WaitOne(); continue; }
                double deadline = UiMotion.Now + Volatile.Read(ref _periodMs);
                while (_hasWork && !_disposed)
                {
                    _timer.Arm(Math.Max(.1, deadline - UiMotion.Now));
                    if (WaitHandle.WaitAny(waits) == 0) continue;
                    PostFrame();
                    double period = Volatile.Read(ref _periodMs);
                    deadline += period;
                    if (deadline <= UiMotion.Now) deadline = UiMotion.Now + period;
                }
                NativeAnimationClock.CancelWaitableTimer(_timer.SafeWaitHandle);
            }
        }
        catch (Exception error) when (error is System.ComponentModel.Win32Exception or ObjectDisposedException)
        {
            AppDiagnostics.Warn("UI.ANIMATION.WAIT", error);
            // Finish active transitions on the UI thread, without a retry loop.
            _failed = true;
            _disposed = true;
            PostFrame(finish: true);
        }
        finally
        {
            _timer.Dispose(); _wake.Dispose();
#if DEVELOPER_QA
            Interlocked.Decrement(ref QaWorkers);
#endif
        }
    }

    private void PostFrame(bool finish = false)
    {
        if ((!finish && _disposed) || Interlocked.CompareExchange(ref _pending, 1, 0) != 0) return;
#if DEVELOPER_QA
        Interlocked.Increment(ref QaQueuedFrames);
        Interlocked.Increment(ref QaPostedFrames);
#endif
        try
        {
            _form.BeginInvoke(new Action(() =>
            {
                try { DrawFrame(finish); }
                finally { ReleasePending(); }
            }));
        }
        catch (InvalidOperationException) { ReleasePending(); }
    }

    private void ReleasePending()
    {
        if (Interlocked.Exchange(ref _pending, 0) == 0) return;
#if DEVELOPER_QA
        Interlocked.Decrement(ref QaQueuedFrames);
#endif
    }

    private void DrawFrame(bool finish = false)
    {
        _frame.Clear(); _frame.AddRange(_active);
        foreach (Subscription item in _frame)
        {
            if (!item.IsActive) continue;
            if (_form.IsDisposed || (_disposed && !_failed && !finish)) { item.Dispose(); continue; }
            if (!item.Draw(finish || _failed) || finish || _failed) item.Dispose();
        }
        _frame.Clear();
    }

    private void Remove(Subscription item)
    {
        _active.Remove(item);
        _hasWork = _active.Count != 0;
        if (!_disposed) Wake();
    }

    private void OnDisposed(object? sender, EventArgs e)
    {
        _form.Disposed -= OnDisposed;
        _disposed = true;
        foreach (Subscription item in _active.ToArray()) item.Dispose();
        ReleasePending();
        Wake();
    }

    private void Wake()
    {
        try { _wake.Set(); }
        catch (ObjectDisposedException) when (_disposed) { }
    }

    internal sealed class Subscription(UiFrameClock clock, Func<bool, bool> draw) : IDisposable
    {
        internal bool IsActive { get; private set; } = true;
        internal bool HighResolution => clock.HighResolution;
        internal bool Draw(bool finish) => draw(finish);
        public void Dispose()
        {
            if (!IsActive) return;
            IsActive = false;
            clock.Remove(this);
        }
    }
}
