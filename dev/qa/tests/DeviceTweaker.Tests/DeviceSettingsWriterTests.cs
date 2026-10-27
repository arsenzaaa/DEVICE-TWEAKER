using DeviceTweakerCS;
using Microsoft.Win32;
using Xunit;
namespace DeviceTweaker.Tests;

public sealed class DeviceSettingsWriterTests
{
    private sealed class Values : IRegistryValues
    {
        internal readonly Dictionary<string, (object Value, RegistryValueKind Kind)> Data = [];
        internal readonly List<string> Trace = [];
        internal string? FailName;
        public void Set(string name, object value, RegistryValueKind kind)
        { Trace.Add("set:" + name); if (name == FailName) throw new IOException("denied"); Data[name] = (value, kind); }
        public void Delete(string name) { Trace.Add("delete:" + name); Data.Remove(name); }
    }
    [Theory] [InlineData(0, 16UL)] [InlineData(4, 0UL)]
    public void EmptyOrDefaultAffinityClearsOverride(int policy, ulong mask)
    {
        Values v = new(); v.Set("AssignmentSetOverride", new byte[] { 16 }, RegistryValueKind.Binary); v.Trace.Clear();
        DeviceSettingsWriter.SetAffinity(v, policy, mask, 8);
        Assert.Equal((0, RegistryValueKind.DWord), v.Data["DevicePolicy"]); Assert.False(v.Data.ContainsKey("AssignmentSetOverride"));
        Assert.Equal(new[] { "set:DevicePolicy", "delete:AssignmentSetOverride" }, v.Trace);
    }
    [Fact] public void ExplicitAffinityWritesFull64BitLittleEndianMask()
    {
        Values v = new(); DeviceSettingsWriter.SetAffinity(v, 4, 0x8000000000000001, 8);
        Assert.Equal((4, RegistryValueKind.DWord), v.Data["DevicePolicy"]);
        Assert.Equal(new byte[] { 1, 0, 0, 0, 0, 0, 0, 128 }, (byte[])v.Data["AssignmentSetOverride"].Value);
        Assert.Equal(RegistryValueKind.Binary, v.Data["AssignmentSetOverride"].Kind);
    }
    [Fact] public void MsiOnlyWritePreservesExistingLimitAndPriority()
    {
        Values v = new(); v.Set("MessageNumberLimit", 8, RegistryValueKind.DWord); v.Set("DevicePriority", 3, RegistryValueKind.DWord); v.Trace.Clear();
        DeviceSettingsWriter.SetMsi(v, true);
        Assert.Equal((1, RegistryValueKind.DWord), v.Data["MSISupported"]);
        Assert.Equal((8, RegistryValueKind.DWord), v.Data["MessageNumberLimit"]); Assert.Equal((3, RegistryValueKind.DWord), v.Data["DevicePriority"]);
        Assert.Equal(new[] { "set:MSISupported" }, v.Trace);
    }
    [Fact] public void UnlimitedLimitAndUndefinedPriorityRemoveValues()
    {
        Values v = new(); DeviceSettingsWriter.SetMsiLimit(v, 8); DeviceSettingsWriter.SetPriority(v, 3);
        DeviceSettingsWriter.SetMsiLimit(v, null); DeviceSettingsWriter.SetPriority(v, null); Assert.Empty(v.Data);
    }
    [Fact] public void StorageUsesWindowsDefaultByRemovingBothOverrides()
    {
        Values v = new(); DeviceSettingsWriter.SetAffinity(v, 4, 16, 8); v.Trace.Clear(); DeviceSettingsWriter.ClearStorageAffinity(v);
        Assert.Empty(v.Data); Assert.Equal(new[] { "delete:AssignmentSetOverride", "delete:DevicePolicy" }, v.Trace);
    }
    [Fact] public void PolicyWriteFailureStopsBeforeWritingMask()
    {
        Values v = new() { FailName = "DevicePolicy" };
        Assert.Throws<IOException>(() => DeviceSettingsWriter.SetAffinity(v, 4, 16, 8));
        Assert.Equal(new[] { "set:DevicePolicy" }, v.Trace); Assert.Empty(v.Data);
    }
}
