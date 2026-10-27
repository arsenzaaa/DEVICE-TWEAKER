using System.Reflection;
using DeviceTweakerCS;
using DeviceTweakerQa;
using Xunit;

namespace DeviceTweaker.Tests;

public sealed class DialogTextTests
{
    [Theory]
    [InlineData("Ошибка и\nследующая строка")]
    [InlineData("Параметр не\n\nзадан")]
    [InlineData("Путь в\r\nследующей строке")]
    public void NonBreakingPrepositionsNeverConsumeExplicitLineOrParagraphBreaks(string text)
    {
        MethodInfo method = typeof(MainForm).GetMethod("NormalizeDialogMessage", BindingFlags.Static | BindingFlags.NonPublic)!;
        string result = (string)method.Invoke(null, [text])!;
        Assert.Equal(text.Replace("\r\n", "\n").Count(c => c == '\n'), result.Count(c => c == '\n'));
    }

    [Theory]
    [InlineData(0)] [InlineData(97)] [InlineData(192)]
    public void UnsupportedLayoutModelCannotSilentlyBecomeAnUntestedScale(int dpi)
    {
        DevScenario scene = new() { Id = "invalid-scale", LayoutDpi = dpi, Steps = [new() { Id = "screen", Action = "capture" }] };
        Assert.Throws<InvalidDataException>(scene.Validate);
    }
}
