using System.Reflection;
using System.Text.Json;
using DeviceTweakerCS;
using Xunit;
namespace DeviceTweaker.Tests;

public sealed class QaArtifactTests
{
    [Fact] public async Task ConcurrentReportReadsNeverSeeTruncatedJson()
    {
        QaRun.WriteReport(); string path = Path.Combine(QaRun.OutputDirectory, $"process-{Environment.ProcessId}.json");
        var writer = Task.Run(() => { for (int i = 0; i < 1000; i++) QaRun.WriteReport(); });
        int reads = 0;
        try
        {
            do
            {
                for (int attempt = 0; ; attempt++)
                {
                    try
                    {
                        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
                        using var json = JsonDocument.Parse(stream);
                        Assert.True(json.RootElement.GetProperty("readOnly").GetBoolean()); reads++;
                        break;
                    }
                    catch (IOException) when (attempt < 10) { await Task.Delay(5); }
                }
            } while (!writer.IsCompleted);
        }
        finally { await writer; }
        Assert.True(reads > 0);
    }
    [Fact] public void ReleaseAssemblyContainsNoTestPanelOrScenarioRuntime()
    {
        Assert.Null(typeof(MainForm).GetMethod("ShowTestAdminDialog", BindingFlags.NonPublic | BindingFlags.Instance));
        Assert.Null(typeof(MainForm).GetMethod("TryHandleVisualQa", BindingFlags.NonPublic | BindingFlags.Instance));
        Assert.Null(typeof(MainForm).GetNestedType("TestCpuConfig", BindingFlags.NonPublic));
        Assert.Null(typeof(MainForm).Assembly.GetType("DeviceTweakerCS.TestSandboxRuntime"));
        Assert.Null(typeof(MainForm).Assembly.GetType("DeviceTweakerCS.TestSandboxScenario"));
        Assert.False(typeof(MainForm).GetProperty("_testAutoDryRun", BindingFlags.NonPublic | BindingFlags.Instance)!.CanWrite);
    }
}
