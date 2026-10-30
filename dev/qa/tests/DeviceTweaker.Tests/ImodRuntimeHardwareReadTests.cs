using Xunit;

namespace DeviceTweaker.Tests;

// Read-only checks of the shared runtime on the machine running the tests:
// SetupAPI / Configuration Manager only, no driver and no register access.
public sealed class ImodRuntimeHardwareReadTests
{
    [Fact]
    public void ControllerEnumerationAndSingleDeviceLookupAgree()
    {
        DeviceTweakerImodController[] controllers = DeviceTweakerImodRuntime.EnumerateXhciControllers();

        foreach (DeviceTweakerImodController controller in controllers)
        {
            Assert.StartsWith(@"PCI\", controller.DeviceId, StringComparison.OrdinalIgnoreCase);
            if (!controller.HasBase)
            {
                continue;
            }

            Assert.NotEqual(0UL, controller.BaseAddress);
            Assert.True(controller.MemoryLength >= 0x20, $"{controller.DeviceId} length=0x{controller.MemoryLength:X}");
            Assert.True(DeviceTweakerImodRuntime.TryGetPciMemoryRange(controller.DeviceId, out ulong baseAddress, out ulong length, out string error), error);
            Assert.Equal(controller.BaseAddress, baseAddress);
            Assert.Equal(controller.MemoryLength, length);
        }
    }

    [Fact]
    public void PciDeviceEnumerationIncludesEveryXhciController()
    {
        HashSet<string> pci = DeviceTweakerImodRuntime.EnumeratePciDevices()
            .Select(device => device.DeviceId)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        Assert.All(DeviceTweakerImodRuntime.EnumerateXhciControllers(), controller => Assert.Contains(controller.DeviceId, pci));
    }

    [Fact]
    public void UnknownDeviceIsReportedNotThrown()
    {
        Assert.False(DeviceTweakerImodRuntime.TryGetPciMemoryRange(@"PCI\VEN_FFFF&DEV_FFFF\0&0", out _, out _, out string error));
        Assert.Contains("not found", error);
    }
}
