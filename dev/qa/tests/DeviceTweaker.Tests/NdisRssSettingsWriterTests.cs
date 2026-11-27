using Microsoft.Win32;
using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public sealed class NdisRssSettingsWriterTests
{
    private const string Instance = @"PCI\TEST\NIC";
    private const string Path = @"SYSTEM\CurrentControlSet\Control\Class\{TEST}\0001";

    private sealed class Key : INdisRssKey
    {
        internal readonly Dictionary<string, object> Values = [];
        internal readonly List<(string Action, string Name, object? Value, RegistryValueKind Kind)> Calls = [];
        internal string? FailAt;
        internal bool FailDispose;
        internal int Disposed;

        public void Set(string name, object value, RegistryValueKind kind)
        {
            Calls.Add(("set", name, value, kind));
            if (name == FailAt) throw new UnauthorizedAccessException("denied\nwith \"details\"");
            Values[name] = value;
        }

        public void Delete(string name)
        {
            Calls.Add(("delete", name, null, RegistryValueKind.None));
            if (name == FailAt) throw new UnauthorizedAccessException("delete denied");
            Values.Remove(name);
        }

        public void Dispose()
        {
            Disposed++;
            if (FailDispose) throw new IOException("close failed");
        }
    }

    private static NdisRssSettingsWriter Writer(Key? key, List<string> log, Action<bool>? opened = null)
        => new(_ => Path, (path, create) => { Assert.Equal(Path, path); opened?.Invoke(create); return key; }, log.Add);

    [Fact]
    public void ExtendedSettings_KeepPhysicalCoreSpanAndAutomaticNuma()
    {
        Key key = new(); key.Values["*NumaNodeId"] = 4;
        List<string> log = []; OperationReport report = new();
        Writer(key, log, create => Assert.True(create)).SetExtra(Instance,
            new(4, 8, new[] { 4, 6, 8 }), report, "NIC");
        Assert.Equal(new[] { "*RssBaseProcGroup", "*MaxRssProcessors", "*RSSMaxProcGroup", "*RssMaxProcNumber", "*RSSProfile", "*NumaNodeId" }, key.Calls.Select(x => x.Name));
        Assert.All(key.Calls.Take(5), x => Assert.Equal(RegistryValueKind.DWord, x.Kind));
        Assert.Equal(0, key.Values["*RssBaseProcGroup"]);
        Assert.Equal(3, key.Values["*MaxRssProcessors"]);
        Assert.Equal(0, key.Values["*RSSMaxProcGroup"]);
        Assert.Equal(8, key.Values["*RssMaxProcNumber"]);
        Assert.Equal(4, key.Values["*RSSProfile"]);
        Assert.False(key.Values.ContainsKey("*NumaNodeId"));
        Assert.Equal(1, key.Disposed); Assert.True(report.Succeeded);
        Assert.Equal(5, log.Count(x => x.StartsWith("RSS.REG.SET:")));
        Assert.Contains(log, x => x.Contains("profile=NUMAStatic(4)") && x.Contains("numa=automatic") && x.Contains("cpus=[4,6,8]"));
    }

    [Theory]
    [InlineData(true, -4, 0)]
    [InlineData(true, 12, 12)]
    [InlineData(false, 0, 1)]
    [InlineData(false, 4, 4)]
    public void ScalarSettings_PreserveClampingAndDwordType(bool baseCore, int requested, int expected)
    {
        Key key = new(); List<string> log = []; var writer = Writer(key, log);
        if (baseCore) writer.SetBaseCore(Instance, requested, null, null);
        else writer.SetQueues(Instance, requested, null, null);
        var call = Assert.Single(key.Calls);
        Assert.Equal(baseCore ? "*RssBaseProcNumber" : "*NumRssQueues", call.Name);
        Assert.Equal(expected, call.Value); Assert.Equal(RegistryValueKind.DWord, call.Kind);
        Assert.Equal(1, key.Disposed);
        Assert.Contains(log, x => x.Contains("status=written") && x.Contains($"requested={expected}"));
        Assert.DoesNotContain(log, x => x.Contains("status=active"));
    }

    [Fact]
    public void FailedExtendedValue_PreservesPriorWritesAndLogsExactPartialResult()
    {
        Key key = new() { FailAt = "*RSSMaxProcGroup" }; key.Values["*NumaNodeId"] = 7;
        List<string> log = []; OperationReport report = new();
        Writer(key, log).SetExtra(Instance, new(2, 4, new[] { 2, 4 }), report, "Network card");
        Assert.Equal(3, key.Calls.Count); Assert.Equal(1, key.Disposed);
        Assert.Equal(2, key.Values["*MaxRssProcessors"]);
        Assert.Equal(7, key.Values["*NumaNodeId"]);
        Assert.False(key.Values.ContainsKey("*RssMaxProcNumber"));
        string error = Assert.Single(log.Where(x => x.StartsWith("RSS.EXTRA.SET.ERROR:")));
        Assert.Contains("completed=2 status=partial", error);
        Assert.Contains("value=*RSSMaxProcGroup action=set requested=0", error);
        Assert.Contains("UnauthorizedAccessException", error);
        Assert.DoesNotContain('\n', error); Assert.DoesNotContain('\r', error);
        Assert.DoesNotContain(log, x => x.StartsWith("RSS.EXTRA.SET:"));
        Assert.Contains("Network card: set extended RSS values", Assert.Single(report.Issues).Component);
    }

    [Fact]
    public void MissingClassKey_IsAnErrorForSetAndANoopForClear()
    {
        List<string> log = []; int opened = 0;
        var writer = new NdisRssSettingsWriter(_ => null, (_, _) => { opened++; return null; }, log.Add);
        OperationReport set = new(), clear = new();
        writer.SetBaseCore(Instance, 4, set, null);
        writer.ClearBaseCore(Instance, clear, null);
        writer.ClearQueues(Instance, clear, null);
        writer.ClearExtra(Instance, clear, null);
        Assert.Equal(0, opened); Assert.Single(set.Errors); Assert.True(clear.Succeeded);
        Assert.Contains(log, x => x.Contains("status=not-written"));
        Assert.Equal(3, log.Count(x => x.Contains("status=not-present")));
    }

    [Fact]
    public void NullWritableKey_LogsOpenFailureWithoutReportingSuccess()
    {
        List<string> log = []; OperationReport report = new();
        Writer(null, log).SetQueues(Instance, 2, report, null);
        string error = Assert.Single(log);
        Assert.Contains("RSS.SET.ERROR:", error); Assert.Contains("stage=open-key", error);
        Assert.Contains("completed=0 status=failed", error); Assert.Single(report.Errors);
    }

    [Fact]
    public void Clear_OpensExistingKeyOnlyAndPreservesUnrelatedValues()
    {
        Key key = new(); key.Values["unrelated"] = 42; key.Values["*RssMaxProcNumber"] = 8; key.Values["*RSSProfile"] = 4;
        List<string> log = []; OperationReport report = new();
        Writer(key, log, create => Assert.False(create)).ClearExtra(Instance, report, null);
        Assert.Equal(6, key.Calls.Count); Assert.Single(key.Values); Assert.Equal(42, key.Values["unrelated"]);
        Assert.All(key.Calls, x => Assert.Equal("delete", x.Action));
        Assert.True(report.Succeeded); Assert.Equal(1, key.Disposed);
        Assert.Equal(6, log.Count(x => x.Contains("status=cleared-or-absent")));
    }

    [Fact]
    public void MissingWritableKeyDuringClear_IsLoggedWithoutCreatingOrClaimingDeletion()
    {
        List<string> log = []; OperationReport report = new();
        Writer(null, log, create => Assert.False(create)).ClearQueues(Instance, report, null);
        Assert.True(report.Succeeded); Assert.Contains("status=not-present", Assert.Single(log));
        Assert.DoesNotContain(log, x => x.StartsWith("RSS.REG.CLEAR:"));
    }

    [Theory]
    [InlineData(true, "*RssBaseProcNumber", "*NumRssQueues")]
    [InlineData(false, "*NumRssQueues", "*RssBaseProcNumber")]
    public void ScalarClear_RemovesOnlyTheSelectedOverride(bool baseCore, string removed, string retained)
    {
        Key key = new(); key.Values[removed] = 2; key.Values[retained] = 4;
        List<string> log = []; var writer = Writer(key, log, create => Assert.False(create));
        if (baseCore) writer.ClearBaseCore(Instance, null, null);
        else writer.ClearQueues(Instance, null, null);
        Assert.Equal(removed, Assert.Single(key.Calls).Name);
        Assert.False(key.Values.ContainsKey(removed)); Assert.Equal(4, key.Values[retained]);
        Assert.Equal(1, key.Disposed);
    }

    [Fact]
    public void FailedClear_StopsAtFailedValueAndDisposesKey()
    {
        Key key = new() { FailAt = "*MaxRssProcessors" };
        key.Values["*RssBaseProcGroup"] = 0;
        key.Values["*MaxRssProcessors"] = 4;
        key.Values["*RSSMaxProcGroup"] = 0;
        List<string> log = []; OperationReport report = new();
        Writer(key, log).ClearExtra(Instance, report, null);
        Assert.Equal(2, key.Calls.Count); Assert.Equal(1, key.Disposed); Assert.Single(report.Errors);
        Assert.False(key.Values.ContainsKey("*RssBaseProcGroup"));
        Assert.Equal(4, key.Values["*MaxRssProcessors"]); Assert.Equal(0, key.Values["*RSSMaxProcGroup"]);
        Assert.Contains(log, x => x.Contains("value=*MaxRssProcessors action=delete requested=absent completed=1 status=partial"));
        Assert.DoesNotContain(log, x => x.StartsWith("RSS.CLEAR:"));
    }

    [Fact]
    public void CloseFailure_IsDistinctFromFailedRegistryValue()
    {
        Key key = new() { FailDispose = true }; List<string> log = []; OperationReport report = new();
        Writer(key, log).SetBaseCore(Instance, 8, report, null);
        Assert.Equal(8, key.Values["*RssBaseProcNumber"]); Assert.Equal(1, key.Disposed);
        Assert.Contains(log, x => x.Contains("stage=close-key") && x.Contains("completed=1 status=partial"));
        Assert.DoesNotContain(log, x => x.StartsWith("RSS.SET:")); Assert.Single(report.Errors);
    }

    [Fact]
    public void ResolveFailure_IsLoggedBeforeAnyKeyIsOpened()
    {
        List<string> log = []; OperationReport report = new(); int opened = 0;
        var writer = new NdisRssSettingsWriter(_ => throw new IOException("lookup failed"),
            (_, _) => { opened++; return null; }, log.Add);
        writer.SetQueues(Instance, 2, report, null);
        Assert.Equal(0, opened); Assert.Contains("stage=resolve-key", Assert.Single(log)); Assert.Single(report.Errors);
    }

    [Fact]
    public void WindowsKeyFactory_BlocksBeforeTouchingRegistryInReadOnlyProcess()
    {
        Assert.True(QaRun.IsReadOnly);
        Assert.Throws<QaWriteBlockedException>(() => NdisRssSettingsWriter.OpenWindowsClassKey("", create: true));
        Assert.Throws<QaWriteBlockedException>(() => NdisRssSettingsWriter.OpenWindowsClassKey("", create: false));
    }

    [Fact]
    public void ProductionWriter_DoesNotConvertReadOnlyViolationIntoAnOrdinaryWriteError()
    {
        Assert.True(QaRun.IsReadOnly);
        List<string> log = []; OperationReport report = new();
        var writer = new NdisRssSettingsWriter(_ => Path, log.Add);
        Assert.Throws<QaWriteBlockedException>(() => writer.SetBaseCore(Instance, 4, report, null));
        Assert.Empty(log); Assert.Empty(report.Issues);
    }
}
