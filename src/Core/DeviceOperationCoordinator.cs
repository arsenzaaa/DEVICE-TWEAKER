namespace DeviceTweakerCS;

// Reject a second UI command rather than queue stale editor inputs. The lease
// belongs to the complete operation, including backup, native work and readback.
internal sealed class DeviceOperationCoordinator
{
    private int _busy;
    internal bool IsBusy => Volatile.Read(ref _busy) != 0;

    internal IDisposable? TryEnter()
        => Interlocked.CompareExchange(ref _busy, 1, 0) == 0 ? new Lease(this) : null;

    private sealed class Lease(DeviceOperationCoordinator owner) : IDisposable
    {
        private DeviceOperationCoordinator? _owner = owner;
        public void Dispose()
        {
            DeviceOperationCoordinator? owner = Interlocked.Exchange(ref _owner, null);
            if (owner is not null) Volatile.Write(ref owner._busy, 0);
        }
    }
}
