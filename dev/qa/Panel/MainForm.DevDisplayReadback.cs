namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private sealed record DevDisplayReadModel(string Stage, string Outcome, int DelayMs);
    private DevDisplayReadModel? _devDisplayReadModel;
    private int _devDisplayReadPending;
    private int _devDisplayReadUiThread;
    private readonly System.Collections.Concurrent.ConcurrentQueue<string> _devDisplayStages = new();
    private Task? _devDisplayRefresh;

    private Task<SnapshotReadResult<DisplaySnapshot>> ReadDevDisplaySnapshotAsync(DevDisplayReadModel model)
    {
        _devDisplayReadModel = null;
        int uiThread = Environment.CurrentManagedThreadId;
        void Stage(string name)
        {
            if (Environment.CurrentManagedThreadId == uiThread)
                Interlocked.Increment(ref _devDisplayReadUiThread);
            _devDisplayStages.Enqueue(name);
            if (name != model.Stage) return;
            Thread.Sleep(model.DelayMs); // Only a fake provider on the QA worker.
            if (model.Outcome == "failure") throw new IOException("Injected display provider failure.");
        }
        return _displayReader.ReadAsync(() =>
        {
            Interlocked.Increment(ref _devDisplayReadPending);
            try
            {
                return ReadDisplaySnapshot([new("QA_DISPLAY", DeviceKind.NET_NDIS, false)], true,
                    _ => { Stage("power"); return false; },
                    _ => { Stage("rss"); return EmptyNdisRssRuntimeState("model-only"); },
                    () => { Stage("drivers"); return (new() { ["QA_DISPLAY"] = new() { DeviceId = "QA_DISPLAY" } }, null); });
            }
            finally { Interlocked.Decrement(ref _devDisplayReadPending); }
        }, TimeSpan.FromMilliseconds(model.Outcome == "timeout" ? 200 : 5000),
            error => AppDiagnostics.Warn("DISPLAY.LATE", error));
    }

    private async Task DevRefreshDisplayFailureAsync()
    {
        WriteLog("TEST.SCENARIO.BEGIN: name=\"Display read failure\" fault=DisplayRead");
        try
        {
            OperationReport report = await RefreshBlocksAsync();
            if (report.Errors.Count != 1 || _displayPowerCache.ContainsKey("QA_DISPLAY"))
                throw new InvalidOperationException("Expected one failed read and no published display snapshot.");
            WriteLog("TEST.SCENARIO.END: status=PASS assertions=2/0 faultConsumed=True");
        }
        catch (Exception error)
        {
            WriteLog("TEST.SCENARIO.END: status=FAIL assertions=0/1 faultConsumed=False");
            _devQueuedErrors.Add(error);
        }
    }
}
