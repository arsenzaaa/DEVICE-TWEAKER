using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public sealed class DriverStagingTests
{
    [Fact]
    public void ParallelStagesDoNotCompeteForTheSameFile()
    {
        string folder = Path.Combine(QaRun.OutputDirectory, "driver-stage-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        string path = Path.Combine(folder, "fixture.sys");
        byte[] payload = Enumerable.Range(0, 4096).Select(i => (byte)i).ToArray();
        Exception?[] failures = new Exception?[8];
        using ManualResetEventSlim start = new(false);
        Thread[] threads = Enumerable.Range(0, 8).Select(index => new Thread(() =>
        {
            start.Wait();
            try
            {
                Assert.True(DriverStagingGate.Run(() =>
                {
                    using FileStream stream = new(path, FileMode.Create, FileAccess.Write, FileShare.None);
                    Thread.Sleep(25); // Keep the real exclusive file handle open while other callers contend.
                    stream.Write(payload); stream.Flush(true);
                    return true;
                }));
            }
            catch (Exception ex) { failures[index] = ex; }
        })).ToArray();
        foreach (Thread thread in threads) thread.Start();
        start.Set();
        foreach (Thread thread in threads) Assert.True(thread.Join(10000));
        Assert.All(failures, failure => Assert.Null(failure));
        Assert.Equal(payload, File.ReadAllBytes(path));
    }
}
