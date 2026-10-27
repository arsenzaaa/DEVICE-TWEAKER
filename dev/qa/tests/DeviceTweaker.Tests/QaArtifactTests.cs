using System.Diagnostics;
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
                using var json = await ReadReportWithHandoffRetryAsync(path, TimeSpan.FromSeconds(5));
                Assert.True(json.RootElement.GetProperty("readOnly").GetBoolean()); reads++;
            } while (!writer.IsCompleted);
        }
        finally { await writer; }
        Assert.True(reads > 0);
    }

    [Fact] public async Task ReportReaderWaitsForTransientSharingViolation()
    {
        string path = CreateReaderFixture("{\"readOnly\":true}");
        try
        {
            using var held = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None);
            Task<JsonDocument> reading = ReadReportWithHandoffRetryAsync(path, TimeSpan.FromSeconds(5));
            await Task.Delay(100);
            Assert.False(reading.IsCompleted);
            held.Dispose();
            using var json = await reading;
            Assert.True(json.RootElement.GetProperty("readOnly").GetBoolean());
        }
        finally { File.Delete(path); }
    }

    [Fact] public async Task ReportReaderFailsWhenSharingViolationPersists()
    {
        string path = CreateReaderFixture("{\"readOnly\":true}");
        try
        {
            using var held = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None);
            await Assert.ThrowsAsync<IOException>(() => ReadReportWithHandoffRetryAsync(path, TimeSpan.FromMilliseconds(100)));
        }
        finally { File.Delete(path); }
    }

    [Fact] public async Task ReportReaderDoesNotRetryMalformedJson()
    {
        string path = CreateReaderFixture("{\"readOnly\":");
        try
        {
            Task<JsonDocument> reading = ReadReportWithHandoffRetryAsync(path, TimeSpan.FromSeconds(5));
            Assert.True(reading.IsFaulted);
            await Assert.ThrowsAnyAsync<JsonException>(() => reading);
        }
        finally { File.Delete(path); }
    }

    [Fact] public async Task ReportReaderWaitsForReportNameHandoff()
    {
        Directory.CreateDirectory(QaRun.OutputDirectory);
        string path = Path.Combine(QaRun.OutputDirectory, $"reader-{Guid.NewGuid():N}.json");
        try
        {
            Task<JsonDocument> reading = ReadReportWithHandoffRetryAsync(path, TimeSpan.FromSeconds(5));
            await Task.Delay(100);
            Assert.False(reading.IsCompleted);
            File.WriteAllText(path, "{\"readOnly\":true}");
            using var json = await reading;
            Assert.True(json.RootElement.GetProperty("readOnly").GetBoolean());
        }
        finally { File.Delete(path); }
    }

    [Fact] public async Task ReportReaderFailsWhenReportRemainsMissing()
    {
        Directory.CreateDirectory(QaRun.OutputDirectory);
        string path = Path.Combine(QaRun.OutputDirectory, $"reader-{Guid.NewGuid():N}.json");
        await Assert.ThrowsAsync<FileNotFoundException>(() => ReadReportWithHandoffRetryAsync(path, TimeSpan.FromMilliseconds(100)));
    }

    private static string CreateReaderFixture(string text)
    {
        Directory.CreateDirectory(QaRun.OutputDirectory);
        string path = Path.Combine(QaRun.OutputDirectory, $"reader-{Guid.NewGuid():N}.json");
        File.WriteAllText(path, text);
        return path;
    }

    private static async Task<JsonDocument> ReadReportWithHandoffRetryAsync(string path, TimeSpan timeout)
    {
        var elapsed = Stopwatch.StartNew();
        while (true)
        {
            try
            {
                using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
                return JsonDocument.Parse(stream);
            }
            catch (IOException ex) when ((ex.HResult & 0xffff) is 2 or 32 or 33 && elapsed.Elapsed < timeout)
            {
                await Task.Delay(10);
            }
        }
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
