using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public sealed class DeviceCardDefinitionTests
{
    private static DeviceInfo Device(UsbChipPathInfo? path = null) => new()
    {
        Name = "USB controller", InstanceId = "PCI\\VEN_8086&DEV_1234\\ONE",
        Class = "USB", RegBase = "SYSTEM\\CurrentControlSet\\Enum\\PCI\\ONE", Kind = DeviceKind.USB, UsbChipPath = path,
    };

    private static readonly string[] DefinitionProperties =
    [
        "Name", "InstanceId", "Class", "RegBase", "Kind", "UsbRoles", "UsbPollingRates", "AudioEndpoints", "StorageTag",
        "IsIntegratedGpu", "Wifi", "UsbIsXhci", "UsbHasDevices", "UsbChipPath", "IsTestDevice",
        "TestIrqCount", "TestIrqMatchCount", "TestMsiStatus", "TestMsiCapabilities", "TestMsiMessageMaximum", "TestState",
    ];

    [Fact]
    public void EveryDevicePropertyMustBeClassifiedForReuse()
    {
        // A newly added device property must not silently be omitted from the key.
        string[] known = [.. DefinitionProperties, "UsbSelectiveSuspend", "NicPowerSaving"];
        Assert.Equal(known.Order(), typeof(DeviceInfo).GetProperties().Select(property => property.Name).Order());
    }

    public static IEnumerable<object[]> CapturedChanges()
    {
        foreach (string name in DefinitionProperties.Where(name => name != "IsTestDevice")) yield return [name];
    }

    [Theory]
    [MemberData(nameof(CapturedChanges))]
    public void CapturedDeviceChangesInvalidateDefinition(string name)
    {
        DeviceInfo before = Device();
        DeviceCardDefinition key = DeviceCardDefinition.From(before);
        DeviceInfo after = before.CopyForDiscovery();
        var property = typeof(DeviceInfo).GetProperty(name)!;
        object changed = name switch
        {
            "Kind" => DeviceKind.GPU,
            "UsbChipPath" => new UsbChipPathInfo(UsbChipOrigin.CpuDirect, 0, "CHIP 0", "Intel", "USB 3", "Controller", "8086", "1234"),
            "TestState" => new TestDeviceState(),
            _ when property.PropertyType == typeof(string) => "changed",
            _ when property.PropertyType == typeof(bool) => true,
            _ when property.PropertyType == typeof(uint?) => (uint)2,
            _ => 2,
        };
        property.SetValue(after, changed);
        Assert.NotEqual(key, DeviceCardDefinition.From(after));
        Assert.Equal(key, DeviceCardDefinition.From(before));
    }

    [Fact]
    public void MutablePowerAndRegisterValuesDoNotDefineTheCard()
    {
        DeviceInfo device = Device();
        device.TestState = new TestDeviceState();
        DeviceCardDefinition before = DeviceCardDefinition.From(device);
        device.UsbSelectiveSuspend = "off";
        device.NicPowerSaving = "on";
        device.TestState.MsiLimit = 2;
        device.TestState.AffinityMask = 4;
        Assert.Equal(before, DeviceCardDefinition.From(device));
    }

    [Fact]
    public void AReplacementDeviceStateIsNotThePreviousBinding()
    {
        DeviceInfo device = Device();
        device.TestState = new TestDeviceState();
        DeviceCardDefinition before = DeviceCardDefinition.From(device);
        device.TestState = device.TestState.Clone();
        Assert.NotEqual(before, DeviceCardDefinition.From(device));
    }

    [Fact]
    public void ChangedControllerIndexInvalidatesPreviouslyCapturedDefinition()
    {
        DeviceInfo device = Device(new UsbChipPathInfo(UsbChipOrigin.CpuDirect, 0, "CHIP 0", "Intel", "USB 3", "Controller", "8086", "1234"));
        DeviceCardDefinition before = DeviceCardDefinition.From(device);
        device.UsbChipPath!.InstanceIndex++;
        Assert.NotEqual(before, DeviceCardDefinition.From(device));
    }
}
