using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public sealed class SingleInstanceTests
{
    private const string WinFormsClass = "WindowsForms10.Window.8.app.0.141b42a_r6_ad1";

    [Fact]
    public void RunningInstanceWindowIsActivated()
    {
        Assert.True(Program.IsInstanceWindow("DEVICE TWEAKER", WinFormsClass, 4321, 1234));
    }

    [Fact]
    public void ExplorerFolderWithTheSameTitleIsIgnored()
    {
        Assert.False(Program.IsInstanceWindow("DEVICE TWEAKER", "CabinetWClass", 4321, 1234));
    }

    [Fact]
    public void OwnWindowAndOtherTitlesAreIgnored()
    {
        Assert.False(Program.IsInstanceWindow("DEVICE TWEAKER", WinFormsClass, 1234, 1234));
        Assert.False(Program.IsInstanceWindow("DEVICE TWEAKER", WinFormsClass, 0, 1234));
        Assert.False(Program.IsInstanceWindow("DEVICE TWEAKER - Notepad", WinFormsClass, 4321, 1234));
        Assert.False(Program.IsInstanceWindow("device tweaker", WinFormsClass, 4321, 1234));
    }
}
