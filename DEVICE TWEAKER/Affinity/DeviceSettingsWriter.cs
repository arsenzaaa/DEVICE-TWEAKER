using Microsoft.Win32;
namespace DeviceTweakerCS;

internal interface IRegistryValues
{
    void Set(string name, object value, RegistryValueKind kind);
    void Delete(string name);
}

internal sealed class WindowsRegistryValues(RegistryKey key) : IRegistryValues
{
    public void Set(string name, object value, RegistryValueKind kind)
    { QaRun.DemandWrite("DeviceSettings.Set"); key.SetValue(name, value, kind); }
    public void Delete(string name)
    { QaRun.DemandWrite("DeviceSettings.Delete"); key.DeleteValue(name, throwOnMissingValue: false); }
}

internal static class DeviceSettingsWriter
{
    internal static void SetMsi(IRegistryValues values, bool enabled)
        => values.Set("MSISupported", enabled ? 1 : 0, RegistryValueKind.DWord);
    internal static void SetMsiLimit(IRegistryValues values, int? limit)
    {
        if (limit.HasValue) values.Set("MessageNumberLimit", limit.Value, RegistryValueKind.DWord);
        else values.Delete("MessageNumberLimit");
    }
    internal static void SetPriority(IRegistryValues values, int? priority)
    {
        if (priority.HasValue) values.Set("DevicePriority", priority.Value, RegistryValueKind.DWord);
        else values.Delete("DevicePriority");
    }
    internal static void SetAffinity(IRegistryValues values, int policy, ulong mask, int pointerSize)
    {
        if (mask == 0 || policy == 0)
        {
            values.Set("DevicePolicy", 0, RegistryValueKind.DWord);
            values.Delete("AssignmentSetOverride");
            return;
        }
        values.Set("DevicePolicy", policy, RegistryValueKind.DWord);
        values.Set("AssignmentSetOverride", pointerSize >= 8 ? BitConverter.GetBytes(mask) : BitConverter.GetBytes((uint)mask), RegistryValueKind.Binary);
    }
    internal static void ClearStorageAffinity(IRegistryValues values)
    {
        values.Delete("AssignmentSetOverride");
        values.Delete("DevicePolicy");
    }
}
