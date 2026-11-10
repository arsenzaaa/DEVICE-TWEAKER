namespace DeviceTweakerCS;

// Values captured by the card's layout and event handlers. Mutable settings are
// deliberately absent: LoadBlockSettings reads those again for every refresh.
internal sealed record DeviceCardDefinition(
    string Name, string InstanceId, string Class, string RegBase, DeviceKind Kind,
    string UsbRoles, string UsbPollingRates, string AudioEndpoints, string StorageTag,
    bool IntegratedGpu, bool Wifi, bool Xhci, bool HasUsbDevices, UsbChipPathInfo? UsbChipPath,
    bool TestDevice, int? TestIrqCount, int TestIrqMatchCount, string TestMsiStatus,
    string TestMsiCapabilities, uint? TestMsiMaximum, string TestInterruptMode, TestDeviceState? TestState)
{
    internal static DeviceCardDefinition From(DeviceInfo device) => new(
        device.Name, device.InstanceId, device.Class, device.RegBase, device.Kind,
        device.UsbRoles, device.UsbPollingRates, device.AudioEndpoints, device.StorageTag,
        device.IsIntegratedGpu, device.Wifi, device.UsbIsXhci, device.UsbHasDevices, device.UsbChipPath is { } path ? path with { } : null,
        device.IsTestDevice, device.TestIrqCount, device.TestIrqMatchCount, device.TestMsiStatus,
        device.TestMsiCapabilities, device.TestMsiMessageMaximum,
#if DEVELOPER_QA
        device.TestInterruptMode,
#else
        string.Empty,
#endif
        device.TestState);
}
