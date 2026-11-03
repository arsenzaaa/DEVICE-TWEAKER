namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private sealed record DisplayReadTarget(string InstanceId, DeviceKind Kind, bool Wifi);
    private sealed record DeviceDisplayState(string InstanceId, bool? PowerSaving, NdisRssRuntimeState? Rss);
    private sealed record DisplaySnapshot(DeviceDisplayState[] Devices,
        Dictionary<string, SignedDriverInfo>? Drivers, string? DriverError);

    private readonly BackgroundSnapshotReader<DisplaySnapshot> _displayReader = new();
    private readonly CancellationTokenSource _displayReadLifetime = new();
    private readonly Dictionary<string, bool> _displayPowerCache = new(StringComparer.OrdinalIgnoreCase);
    private string? _signedDriverInfoError;

    private static bool HasDisplayPower(DeviceKind kind) => kind is DeviceKind.USB or DeviceKind.NET_NDIS or DeviceKind.NET_CX;

    private Task<SnapshotReadResult<DisplaySnapshot>> ReadDisplaySnapshotAsync(IEnumerable<DeviceInfo> devices)
    {
        DisplayReadTarget[] targets = devices.Where(device => !device.IsTestDevice)
            .Select(device => new DisplayReadTarget(device.InstanceId, device.Kind, device.Wifi)).ToArray();
        bool readDrivers = _detailedLogEnabled && _signedDriverInfoCache is null;
        // Pure virtual fixtures must never query host WMI or network providers.
        if (_testDevicesEnabled && _testDevicesOnly)
        {
#if DEVELOPER_QA
            if (_devDisplayReadModel is { } model)
                return ReadDevDisplaySnapshotAsync(model);
#endif
            return Task.FromResult(new SnapshotReadResult<DisplaySnapshot>(new([], null, null), null));
        }
        CancellationToken lifetime = _displayReadLifetime.Token;
        return _displayReader.ReadAsync(() => ReadDisplaySnapshot(targets, readDrivers,
            target => { lifetime.ThrowIfCancellationRequested(); return ReadDisplayPower(target); },
            id => ReadNdisRssRuntimeState(id, lifetime),
            () => { lifetime.ThrowIfCancellationRequested(); return ReadSignedDriverSnapshot(); }),
            TimeSpan.FromSeconds(30), error => AppDiagnostics.Warn("DISPLAY.LATE", error));
    }

    private static DisplaySnapshot ReadDisplaySnapshot(DisplayReadTarget[] targets, bool readDrivers,
        Func<DisplayReadTarget, bool> readPower, Func<string, NdisRssRuntimeState> readRss,
        Func<(Dictionary<string, SignedDriverInfo> Drivers, string? Error)> readDriverMap)
    {
        List<DeviceDisplayState> states = [];
        foreach (DisplayReadTarget target in targets)
        {
            bool? power = null;
            if (!target.Wifi && HasDisplayPower(target.Kind))
                power = readPower(target);
            NdisRssRuntimeState? rss = target.Kind == DeviceKind.NET_NDIS && !target.Wifi
                ? readRss(target.InstanceId) : null;
            states.Add(new(target.InstanceId, power, rss));
        }
        var driverMap = readDrivers ? readDriverMap() : (null, null);
        return new(states.ToArray(), driverMap.Item1, driverMap.Item2);
    }

    private bool ReadDisplayPower(DisplayReadTarget target)
    {
        if (target.Kind == DeviceKind.USB)
        {
            _ = UsbSelectiveSuspendPolicy.TryReadEnabled(target.InstanceId, out bool enabled);
            return enabled;
        }
        bool? live = DevicePowerPolicy.TryReadDevicePowerEnable(target.InstanceId);
        if (live is bool value) return value;
        int? caps = DevicePowerPolicy.TryReadNicPnPCapabilities(GetClassKeyForDevice(target.InstanceId));
        return caps is null || DevicePowerPolicy.IsNicTurnOffAllowed(caps.Value);
    }

    private static (Dictionary<string, SignedDriverInfo>, string?) ReadSignedDriverSnapshot()
    {
        var map = ReadSignedDriverInfoMap(out string? error);
        return (map, error);
    }

    private void PublishDisplaySnapshot(DisplaySnapshot snapshot, bool replaceDevices)
    {
        if (replaceDevices)
        {
            _ndisRssRuntimeCache.Clear();
            _displayPowerCache.Clear();
        }
        foreach (DeviceDisplayState state in snapshot.Devices)
        {
            string id = NormalizeInstanceId(state.InstanceId);
            if (state.PowerSaving is bool power) _displayPowerCache[id] = power;
            if (state.Rss is not null) _ndisRssRuntimeCache[id] = state.Rss;
        }
        if (snapshot.Drivers is not null)
        {
            _signedDriverInfoError = snapshot.DriverError;
            if (snapshot.DriverError is null && snapshot.Drivers.Count > 0)
                _signedDriverInfoCache = snapshot.Drivers;
        }
    }

    private bool GetDisplayPowerState(string instanceId)
        => _displayPowerCache.TryGetValue(NormalizeInstanceId(instanceId), out bool enabled)
            ? enabled : throw new InvalidOperationException("Device power state must be read before binding its card.");

    private async Task ReloadBlockSettingsAsync(DeviceBlock block, OperationReport report)
    {
        SnapshotReadResult<DisplaySnapshot> result = await ReadDisplaySnapshotAsync([block.Device]);
        if (_uiDisposing || IsDisposed || Disposing || !_blocks.Contains(block)) return;
        if (result.Snapshot is null)
        {
            report.AddWarning("Device readback", "Some device data could not be refreshed within the time limit.", result.Error);
            return;
        }
        PublishDisplaySnapshot(result.Snapshot, replaceDevices: false);
        LoadBlockSettings(block);
    }
}
