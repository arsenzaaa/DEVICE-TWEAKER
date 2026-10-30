using System.Reflection;
using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public sealed class DriverRuntimeTests
{
    private static byte[] Info(uint mode = 2, ulong build = 2026102301)
    {
        byte[] data = new byte[32];
        BitConverter.GetBytes(32u).CopyTo(data, 0);
        BitConverter.GetBytes(1u).CopyTo(data, 4);
        BitConverter.GetBytes(build).CopyTo(data, 8);
        BitConverter.GetBytes(7u).CopyTo(data, 16);
        BitConverter.GetBytes(mode).CopyTo(data, 20);
        return data;
    }

    [Theory] [InlineData(1u)] [InlineData(2u)]
    public void CompatibleProtocolAcceptsDifferentBuildsAndReportsActualMode(uint mode)
    {
        Assert.True(DriverRuntime.ValidateInfo(Info(mode, 123), 32, out string result));
        Assert.Contains("build=123", result);
        Assert.Contains(mode == 1 ? "mode=service" : "mode=kdu", result);
    }

    [Theory] [InlineData(0, 31u)] [InlineData(4, 2u)] [InlineData(16, 3u)] [InlineData(20, 0u)] [InlineData(24, 1u)]
    public void MalformedOrIncompatibleReplyIsRejected(int offset, uint value)
    {
        byte[] data = Info(); BitConverter.GetBytes(value).CopyTo(data, offset);
        Assert.False(DriverRuntime.ValidateInfo(data, data.Length, out _));
    }

    [Theory] [InlineData(0)] [InlineData(31)] [InlineData(33)]
    public void TruncatedAndOversizedReplyIsRejected(int returned)
        => Assert.False(DriverRuntime.ValidateInfo(Info(), returned, out _));

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
    public void AutonomousQaCannotEnterNativeCoordinatorOrQueryDriver()
    {
        Assert.True(QaRun.IsReadOnly);
        Assert.Throws<QaWriteBlockedException>(() => DriverLoadGate.Enter());
        Assert.Throws<QaWriteBlockedException>(() => DriverRuntime.QueryInfo(IntPtr.Zero, out _));
        Assert.Throws<QaWriteBlockedException>(() => DriverLoadGate.BeginLoader());
    }

    [Theory] [InlineData(false)] [InlineData(true)]
    public void ReadbackNeverChecksJournalOrInvokesEitherLoader(bool available)
    {
        var result = DriverReadyPolicy.Ensure(false, () => available,
            () => throw new Exception("journal/staging touched by readback"),
            () => throw new Exception("KDU called by readback"), () => throw new Exception("SCM called by readback"));
        Assert.Equal(available ? DriverReadyResult.Existing : DriverReadyResult.Unavailable, result);
    }

    [Theory] [InlineData("existing")] [InlineData("kdu")] [InlineData("service")] [InlineData("failed")]
    public void ReadyPolicyRechecksDeviceAndUsesServiceOnlyAfterCompletedKdu(string scenario)
    {
        List<string> calls = [];
        bool available = scenario == "existing";
        var result = DriverReadyPolicy.Ensure(true,
            () => { calls.Add("open"); return available; }, () => calls.Add("idle"),
            () => { calls.Add("kdu"); return available = scenario == "kdu"; },
            () => { calls.Add("service"); return available = scenario == "service"; });
        string[] expected = scenario switch {
            "existing" => ["idle", "open"], "kdu" => ["idle", "open", "kdu", "idle", "open"],
            "service" => ["idle", "open", "kdu", "idle", "open", "service", "open"],
            _ => ["idle", "open", "kdu", "idle", "open", "service"] };
        Assert.Equal(expected, calls);
        Assert.Equal(scenario switch { "existing" => DriverReadyResult.Existing, "kdu" => DriverReadyResult.Kdu,
            "service" => DriverReadyResult.Service, _ => DriverReadyResult.Unavailable }, result);
    }

    [Fact]
    public void PendingKduStopsServiceEvenIfDeviceAppeared()
    {
        bool pending = false, appeared = false;
        Assert.Throws<TimeoutException>(() => DriverReadyPolicy.Ensure(true, () => appeared,
            () => { if (pending) throw new TimeoutException(); },
            () => { pending = appeared = true; return false; },
            () => throw new Exception("SCM must never run after timeout")));
    }

    [Fact]
    public void IncompatibleLoadedDriverStopsAllLoading()
        => Assert.Throws<InvalidDataException>(() => DriverReadyPolicy.Ensure(true,
            () => throw new InvalidDataException("incompatible protocol"), () => { },
            () => throw new Exception("must not map"), () => throw new Exception("must not start")));

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
        Type type = typeof(MainForm).GetNestedType("ImodDriverContext", BindingFlags.NonPublic)!;
        var context = (IDisposable)System.Runtime.CompilerServices.RuntimeHelpers.GetUninitializedObject(type);
        type.GetField("<DriverHandle>k__BackingField", BindingFlags.Instance | BindingFlags.NonPublic)!.SetValue(context, new IntPtr(-1));
        Assert.Null(type.GetField("_operationGate", BindingFlags.Instance | BindingFlags.NonPublic));
        context.Dispose();
        context.Dispose();
    }
}
