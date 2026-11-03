namespace DeviceTweakerCS;

internal sealed record DeviceSettingsResetPlan(
    string InstanceId, string DeviceName, string RegBase, DeviceKind Kind, bool Wifi, bool HasPowerEditor);
