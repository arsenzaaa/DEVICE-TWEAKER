namespace DeviceTweakerCS;

// Captured on the UI thread before entering the backend. No control or mutable
// DeviceBlock reference crosses the worker boundary.
internal sealed record DeviceSettingsWritePlan
{
    internal required string InstanceId { get; init; }
    internal required string DeviceName { get; init; }
    internal required string RegBase { get; init; }
    internal required DeviceKind Kind { get; init; }
    internal required string MsiMode { get; init; }
    internal required string LimitText { get; init; }
    internal PciInterruptCapabilities? InterruptCapabilities { get; init; }
    internal required string Priority { get; init; }
    internal required string Policy { get; init; }
    internal ulong AffinityMask { get; init; }
    internal bool? PowerSaving { get; init; }
    internal bool Wifi { get; init; }
    internal NdisAffinityMode NdisMode { get; init; }
    internal NdisRssLayout.Layout? RssLayout { get; init; }
    internal bool AutoMsiOnly { get; init; }
    internal string AutoReason { get; init; } = "auto-policy";
}

// Written only by the worker, consumed only after its task completes.
internal sealed class DeviceSettingsWriteResult
{
    internal OperationReport Report { get; } = new();
    internal string? NormalizedLimit { get; set; }
    internal string? UsbPower { get; set; }
    internal string? NicPower { get; set; }
    internal NdisRssRuntimeState? RssRuntime { get; set; }
}
