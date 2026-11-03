using System.Reflection;
using System.Runtime.CompilerServices;
using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public class BackupWorkTests
{
    [Fact]
    public void BackupTargetsSurviveReplacementAndRemovalOfUiCards()
    {
        // No form constructor, handles, discovery, registry or WMI calls.
        var form = (MainForm)RuntimeHelpers.GetUninitializedObject(typeof(MainForm));
        var block = (DeviceBlock)RuntimeHelpers.GetUninitializedObject(typeof(DeviceBlock));
        var device = new DeviceInfo
        {
            Name = "Adapter", InstanceId = "PCI\\NIC-A", Class = "Net", RegBase = "Enum\\NIC-A",
            Kind = DeviceKind.NET_CX, Wifi = true,
        };
        typeof(DeviceBlock).GetProperty("Device")!.SetValue(block, device);
        typeof(DeviceBlock).GetProperty("Kind")!.SetValue(block, DeviceKind.NET_CX);
        List<DeviceBlock> cards = [block];
        typeof(MainForm).GetField("_blocks", BindingFlags.Instance | BindingFlags.NonPublic)!.SetValue(form, cards);

        var targets = (MainForm.BackupDeviceTarget[])typeof(MainForm)
            .GetMethod("CaptureBackupTargets", BindingFlags.Instance | BindingFlags.NonPublic)!.Invoke(form, null)!;
        typeof(DeviceBlock).GetProperty("Device")!.SetValue(block, new DeviceInfo
        {
            Name = "Replacement", InstanceId = "USB\\B", Class = "USB", RegBase = "Enum\\B", Kind = DeviceKind.USB,
        });
        cards.Clear();

        var target = Assert.Single(targets);
        Assert.Equal("PCI\\NIC-A", target.InstanceId);
        Assert.Equal("Enum\\NIC-A", target.RegBase);
        Assert.Equal(DeviceKind.NET_CX, target.Kind);
        Assert.True(target.Wifi);
    }
}
