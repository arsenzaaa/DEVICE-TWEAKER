using System.Diagnostics;

namespace DeviceTweakerCS;

// Local feedback is an application choice. Do not change Windows preferences.
internal static class UiMotion
{
    internal static double Now => Stopwatch.GetTimestamp() * 1000.0 / Stopwatch.Frequency;
    internal static bool CanAnimate(Control owner) => owner.IsHandleCreated && owner.Visible
        && owner.Enabled && !owner.IsDisposed && !SystemInformation.HighContrast;

    internal static Color Mix(Color from, Color to, float amount)
    {
        int Lerp(int a, int b) => (int)Math.Round(a + (b - a) * amount);
        return Color.FromArgb(Lerp(from.A, to.A), Lerp(from.R, to.R),
            Lerp(from.G, to.G), Lerp(from.B, to.B));
    }
}

// Time-based, retargetable motion. A delayed frame catches up without replaying ticks.
internal sealed class ValueTransition(double durationMs)
{
    private float _from;
    private double _started;
    internal float Target { get; private set; }
    internal bool IsRunning { get; private set; }

    internal float Position(double now)
    {
        if (!IsRunning) return Target;
        float t = (float)Math.Clamp((now - _started) / durationMs, 0, 1);
        if (t >= 1) { Complete(); return Target; }
        // Ease out promptly, with no overshoot at either end.
        return _from + (Target - _from) * (1 - (1 - t) * (1 - t) * (1 - t));
    }

    internal void MoveTo(float target, double now, bool animate)
    {
        if (animate && IsRunning && target == Target) return;
        float current = Position(now);
        Target = target;
        _from = current;
        _started = now;
        IsRunning = animate && current != target;
        if (!IsRunning) Complete();
    }

    internal void Complete() { _from = Target; IsRunning = false; }
}

// Owns only presentation. There is no background work or timer while idle.
internal sealed class ControlAnimation : IDisposable
{
    private readonly Control _owner;
    private readonly ValueTransition _motion;
    private readonly Action<float> _frame;
    private UiFrameClock.Subscription? _frames;
    private bool _disposed;
    internal float Position => _motion.Position(UiMotion.Now);
    internal float Target => _motion.Target;
    internal bool IsRunning => _frames?.IsActive == true;

    internal ControlAnimation(Control owner, double durationMs, Action<float> frame)
    {
        _owner = owner; _motion = new(durationMs); _frame = frame;
        owner.VisibleChanged += OnAvailabilityChanged;
        owner.EnabledChanged += OnAvailabilityChanged;
        owner.HandleDestroyed += OnHandleDestroyed;
        owner.Disposed += OnDisposed;
    }

    internal void MoveTo(float target, bool animate = true)
    {
        if (_disposed) return;
        _motion.MoveTo(target, UiMotion.Now, animate && UiMotion.CanAnimate(_owner));
        _frame(Position);
        if (!_motion.IsRunning) { _frames?.Dispose(); _frames = null; return; }
        if (_frames?.IsActive != true)
        {
            _frames = UiFrameClock.Start(_owner, finish =>
            {
                if (finish || !UiMotion.CanAnimate(_owner)) _motion.Complete();
                _frame(Position);
                _owner.Update();
                return _motion.IsRunning;
            });
            if (_frames is null) Complete();
        }
    }

    internal void Complete()
    {
        _frames?.Dispose(); _frames = null; _motion.Complete();
        if (!_disposed && !_owner.IsDisposed) _frame(Target);
    }

    private void OnAvailabilityChanged(object? sender, EventArgs e)
    { if (!_owner.Visible || !_owner.Enabled) Complete(); }
    private void OnHandleDestroyed(object? sender, EventArgs e) => Complete();
    private void OnDisposed(object? sender, EventArgs e) => Dispose();

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        _frames?.Dispose(); _frames = null;
        _motion.Complete();
        _owner.VisibleChanged -= OnAvailabilityChanged;
        _owner.EnabledChanged -= OnAvailabilityChanged;
        _owner.HandleDestroyed -= OnHandleDestroyed;
        _owner.Disposed -= OnDisposed;
    }
}
