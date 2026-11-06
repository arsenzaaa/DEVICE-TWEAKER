namespace DeviceTweakerCS;

// A visual transition only. Checked and its events never wait for this clock.
internal sealed class ToggleTransition
{
    internal const double DurationMs = 120;
    private float _from;
    private float _target;
    private double _startedMs;
    internal bool IsRunning { get; private set; }

    internal float Position(double nowMs)
    {
        if (!IsRunning) return _target;
        float t = (float)Math.Clamp((nowMs - _startedMs) / DurationMs, 0, 1);
        if (t >= 1) { Complete(); return _target; }
        float eased = t * t * (3 - 2 * t);
        return _from + (_target - _from) * eased;
    }

    internal void SetTarget(bool value, double nowMs, bool animate)
    {
        float target = value ? 1 : 0;
        if (animate && IsRunning && target == _target) return;
        float current = Position(nowMs);
        _target = target;
        if (!animate || current == target) { Complete(); return; }
        _from = current;
        _startedMs = nowMs;
        IsRunning = true;
    }

    internal void Complete() { _from = _target; IsRunning = false; }
}
