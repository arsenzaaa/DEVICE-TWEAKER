using System.Reflection;
using System.Runtime.CompilerServices;
using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public sealed class QaBackupDirectoryTests
{
    [Theory]
    [InlineData("Local")]
    [InlineData("Roaming")]
    public void ReadOnlyRestoreLooksOnlyInsideTheRunDirectory(string location)
    {
        Assert.True(QaRun.IsReadOnly);
        var form = (MainForm)RuntimeHelpers.GetUninitializedObject(typeof(MainForm));
        Type locationType = typeof(MainForm).GetNestedType("BackupLocation", BindingFlags.NonPublic)!;
        MethodInfo method = typeof(MainForm).GetMethod("GetBackupDirectory", BindingFlags.NonPublic | BindingFlags.Instance,
            null, [locationType], null)!;
        string directory = (string)method.Invoke(form, [Enum.Parse(locationType, location)])!;
        string relative = Path.GetRelativePath(QaRun.OutputDirectory, directory);
        Assert.False(Path.IsPathRooted(relative));
        Assert.NotEqual("..", relative);
        Assert.False(relative.StartsWith(".." + Path.DirectorySeparatorChar, StringComparison.Ordinal));
        QaRun.DemandArtifactPath(directory);
    }
}
