namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private sealed record DisplayReadTarget(string InstanceId, DeviceKind Kind, bool Wifi);
    private sealed record DeviceDisplayState(string InstanceId, bool? PowerSaving, NdisRssRuntimeState? Rss);
    private sealed record DisplaySnapshot(DeviceDisplayState[] Devices);

    private sealed record DriverMetadata(Dictionary<string, SignedDriverInfo> Drivers, string? Error);
    private readonly BackgroundSnapshotReader<DriverMetadata> _driverMetadataReader = new();
    private Task? _driverMetadataRead;
    private readonly BackgroundSnapshotReader<DisplaySnapshot> _displayReader = new();
    private readonly CancellationTokenSource _displayReadLifetime = new();
    private readonly Dictionary<string, bool> _displayPowerCache = new(StringComparer.OrdinalIgnoreCase);
    private string? _signedDriverInfoError;

    private static bool HasDisplayPower(DeviceKind kind) => kind is DeviceKind.USB or DeviceKind.NET_NDIS or DeviceKind.NET_CX;

    private Task<SnapshotReadResult<DisplaySnapshot>> ReadDisplaySnapshotAsync(IEnumerable<DeviceInfo> devices)
    {
        DisplayReadTarget[] targets = devices.Where(device => !device.IsTestDevice)
            .Select(device => new DisplayReadTarget(device.InstanceId, device.Kind, device.Wifi)).ToArray();
        // Pure virtual fixtures must never query host WMI or network providers.
        if (_testDevicesEnabled && _testDevicesOnly)
        {
#if DEVELOPER_QA
            if (_devDisplayReadModel is { } model)
                return ReadDevDisplaySnapshotAsync(model);
#endif
            return Task.FromResult(new SnapshotReadResult<DisplaySnapshot>(new([]), null));
        }
        CancellationToken lifetime = _displayReadLifetime.Token;
        return _displayReader.ReadAsync(() => ReadDisplaySnapshot(targets,
            target => { lifetime.ThrowIfCancellationRequested(); return ReadDisplayPower(target); },
            id => ReadNdisRssRuntimeState(id, lifetime)),
            TimeSpan.FromSeconds(30), error => AppDiagnostics.Warn("DISPLAY.LATE", error));
    }

    // Driver metadata is diagnostic data independent of device discovery and RSS.
    // Start it once per uncached refresh; the reader owns even a timed-out worker.
    private void BeginDriverMetadataRead()
    {
        if (!_detailedLogEnabled || _signedDriverInfoCache is not null || _driverMetadataRead is { IsCompleted: false }
            || (_testDevicesEnabled && _testDevicesOnly))
            return;
        CancellationToken lifetime = _displayReadLifetime.Token;
        _driverMetadataRead = PublishDriverMetadataAsync(_driverMetadataReader.ReadAsync(() =>
        {
            lifetime.ThrowIfCancellationRequested();
            var drivers = ReadSignedDriverInfoMap(out string? error);
            return new DriverMetadata(drivers, error);
        }, TimeSpan.FromSeconds(30), error => AppDiagnostics.Warn("DISPLAY.DRIVERS.LATE", error)));
    }

    private async Task PublishDriverMetadataAsync(Task<SnapshotReadResult<DriverMetadata>> read)
    {
        SnapshotReadResult<DriverMetadata> result = await read;
        if (_uiDisposing || IsDisposed || Disposing) return;
        _signedDriverInfoError = result.Error ?? result.Snapshot?.Error;
        if (_signedDriverInfoError is not null)
            WriteLog($"DISPLAY.DRIVERS.WARN: {FlattenLogText(_signedDriverInfoError)}");
        else if (result.Snapshot is { } metadata)
        {
            if (metadata.Drivers.Count > 0) _signedDriverInfoCache = metadata.Drivers;
            WriteLog($"DISPLAY.DRIVERS: signedDrivers={metadata.Drivers.Count}");
        }
    }

    private static DisplaySnapshot ReadDisplaySnapshot(DisplayReadTarget[] targets,
        Func<DisplayReadTarget, bool> readPower, Func<string, NdisRssRuntimeState> readRss)
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
        return new(states.ToArray());
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
    }

    private bool GetDisplayPowerState(string instanceId)
        => _displayPowerCache.TryGetValue(NormalizeInstanceId(instanceId), out bool enabled)
            ? enabled : throw new InvalidOperationException("Device power state must be read before binding its card.");

    private async Task ReloadBlockSettingsAsync(DeviceBlock block, OperationReport report)
    {
        BeginDriverMetadataRead();
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
