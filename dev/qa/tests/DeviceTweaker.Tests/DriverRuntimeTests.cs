using System.Reflection;
using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public sealed class DriverRuntimeTests
{
    [Fact]
    public void IncompleteLoadBlocksBothLoadersUntilReboot()
    {
        const string boot = "11111111-1111-1111-1111-111111111111";
        Assert.NotNull(DriverLoadGate.CheckJournal(boot + "|pending", boot));
        Assert.Null(DriverLoadGate.CheckJournal(boot + "|pending", "22222222-2222-2222-2222-222222222222"));
        Assert.NotNull(DriverLoadGate.CheckJournal("broken", boot));
        Assert.Null(DriverLoadGate.CheckJournal("", boot));
    }

    [Fact]
    public void AutonomousQaCannotEnterNativeCoordinator()
    {
        Assert.True(QaRun.IsReadOnly);
        Assert.Throws<QaWriteBlockedException>(() => DriverLoadGate.Enter());
        Assert.Throws<QaWriteBlockedException>(() => DriverRuntime.DescribeLoadedDriver());
        Assert.Throws<QaWriteBlockedException>(() => DriverLoadGate.BeginLoader());
    }

    [Fact]
    public void ReadOnlyExistingDeviceProbeStopsBeforeNativeAccess()
    {
        Assert.False(ImodDriverContext.TryOpenExisting(_ => throw new Exception("native diagnostics reached"), out var context, out string? error));
        Assert.Null(context);
        Assert.NotNull(error);
        Assert.Contains("QA read-only", error);
    }

    [Theory] [InlineData("existing")] [InlineData("kdu")] [InlineData("service")] [InlineData("failed")]
    public void ReadyPolicyUsesServiceOnlyAfterCompletedKdu(string scenario)
    {
        List<string> calls = [];
        bool available = scenario == "existing";
        var result = DriverReadyPolicy.Ensure(
            () => { calls.Add("open"); return available; },
            () => { calls.Add("verify"); return true; },
            () => { calls.Add("kdu"); return scenario == "kdu"; }, () => calls.Add("idle"),
            () => { calls.Add("service"); return available = scenario == "service"; });
        string[] expected = scenario switch {
            "existing" => ["open"], "kdu" => ["open", "verify", "kdu"],
            _ => ["open", "verify", "kdu", "idle", "verify", "service"] };
        Assert.Equal(expected, calls);
        Assert.Equal(scenario switch { "existing" => DriverReadyResult.Existing, "kdu" => DriverReadyResult.Kdu,
            "service" => DriverReadyResult.Service, _ => DriverReadyResult.Unavailable }, result);
    }

    [Fact]
    public void PendingKduStopsServiceEvenIfDeviceAppeared()
    {
        bool pending = false, appeared = false;
        Assert.Throws<TimeoutException>(() => DriverReadyPolicy.Ensure(() => appeared, () => true,
            () => { pending = appeared = true; return false; },
            () => { if (pending) throw new TimeoutException(); },
            () => throw new Exception("SCM must never run after timeout")));
    }

    [Fact]
    public void ExistingDeviceProbeFailureStopsAllLoading()
        => Assert.Throws<IOException>(() => DriverReadyPolicy.Ensure(
            () => throw new IOException("device probe failed"), () => true,
            () => throw new Exception("must not map"), () => { }, () => throw new Exception("must not start")));

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RefusedFilePreventsTheNextLoader(bool validAtFirst)
    {
        int attempts = 0;
        DriverReadyResult result = DriverReadyPolicy.Ensure(
            () => false, () => validAtFirst && attempts == 0,
            () => { attempts++; return false; }, () => { },
            () => throw new InvalidOperationException("SCM received an unverified file"));
        Assert.Equal(DriverReadyResult.Unavailable, result);
        Assert.Equal(validAtFirst ? 1 : 0, attempts);
    }

    [Fact]
    public async Task ProductionMutexLeaseSerializesCompetingOperationsAndReleasesAfterFailure()
    {
        string name = @"Local\DT_QA_" + Guid.NewGuid().ToString("N");
        int active = 0, completed = 0;
        using ManualResetEventSlim start = new();
        Task[] workers = Enumerable.Range(0, 4).Select(_ => Task.Run(() =>
        {
            start.Wait();
            for (int i = 0; i < 8; i++)
            {
                using DriverOperationLease lease = DriverOperationLease.Acquire(name, 5000);
                Assert.Equal(1, Interlocked.Increment(ref active));
                Thread.Sleep(2);
                Assert.Equal(0, Interlocked.Decrement(ref active));
                Interlocked.Increment(ref completed);
            }
        })).ToArray();
        start.Set(); await Task.WhenAll(workers);
        Assert.Equal(32, completed);
        Assert.Throws<IOException>((Action)(() => { using var lease = DriverOperationLease.Acquire(name, 1000); throw new IOException(); }));
        using var final = DriverOperationLease.Acquire(name, 1000);
    }

    [Fact]
    public void ProductionMutexLeaseHasBoundedWaitAndRecoversAbandonedOwnership()
    {
        string name = @"Local\DT_QA_" + Guid.NewGuid().ToString("N");
        Mutex? abandoned = null;
        using (DriverOperationLease owner = DriverOperationLease.Acquire(name, 1000))
        {
            Exception? failure = null;
            Thread contender = new(() => { try { using var ignored = DriverOperationLease.Acquire(name, 10); } catch (Exception ex) { failure = ex; } });
            contender.Start(); contender.Join();Assert.IsType<TimeoutException>(failure);
        }
        Thread thread = new(() => { abandoned = new Mutex(false, name); abandoned.WaitOne(); });
        thread.Start(); thread.Join();
        using (DriverOperationLease recovered = DriverOperationLease.Acquire(name, 1000)) { }
        abandoned!.Dispose();
    }

    [Fact]
    public void PublishedContextWithoutOwnershipDisposesWithoutNativeServiceAccess()
    {
        Type type = typeof(ImodDriverContext);
        var context = (IDisposable)System.Runtime.CompilerServices.RuntimeHelpers.GetUninitializedObject(type);
        type.GetField("<DriverHandle>k__BackingField", BindingFlags.Instance | BindingFlags.NonPublic)!.SetValue(context, new IntPtr(-1));
        Assert.Null(type.GetField("_operationGate", BindingFlags.Instance | BindingFlags.NonPublic));
        context.Dispose();
        context.Dispose();
    }
}
