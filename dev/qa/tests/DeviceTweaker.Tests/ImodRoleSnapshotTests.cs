using Xunit;

namespace DeviceTweakerCS.Tests;

public sealed class ImodRoleSnapshotTests
{
    private const string Controller = @"PCI\VEN_8086&DEV_1234\A";

    [Fact]
    public void LateRoleResolutionUsesCapturedDevicesAndPollingOverrides()
    {
        List<ImodRoleDevice> devices = [new(Controller, "Mouse (1000Hz), Gamepad", "Audio")];
        Dictionary<string, string> overrides = new(StringComparer.OrdinalIgnoreCase) { [Controller] = "Mouse (8000Hz)" };
        ImodRoleSnapshot captured = new(devices, overrides);
        devices.Clear();
        overrides.Clear();

        HashSet<string> roles = MainForm.BuildAdaptiveControllerRoleFallback(captured, Controller);
        Assert.Equal(new[] { "Audio", "Gamepad", "Mouse" }, roles.OrderBy(role => role).ToArray());
        Dictionary<string, List<string>> labels = MainForm.BuildAdaptiveRoleDisplayLabels(captured, Controller);
        Assert.Contains(labels["Mouse"], label => label.Contains("8000", StringComparison.Ordinal)
            || label.Contains("8k", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(labels["Mouse"], label => label.Contains("1000", StringComparison.Ordinal));
    }

    [Fact]
    public void OtherControllersNeverBorrowCapturedRoles()
    {
        ImodRoleSnapshot captured = new([new(Controller, "Mouse", "")], new Dictionary<string, string>());
        Assert.Empty(MainForm.BuildAdaptiveControllerRoleFallback(captured, @"PCI\VEN_1022&DEV_5678\B"));
        Assert.Empty(MainForm.BuildAdaptiveRoleDisplayLabels(captured, @"PCI\VEN_1022&DEV_5678\B"));
    }
}
