namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private NdisRssSettingsWriter CreateNdisRssWriter() => new(GetClassKeyForDevice, WriteLog);

    private void SetNdisBaseCore(string instanceId, int baseCore, OperationReport? report = null, string? deviceName = null)
    {
        DemandRealWrite("SetNdisBaseCore");
        CreateNdisRssWriter().SetBaseCore(instanceId, baseCore, report, deviceName);
    }

    private void SetNdisRssQueues(string instanceId, int queues, OperationReport? report = null, string? deviceName = null)
    {
        DemandRealWrite("SetNdisRssQueues");
        CreateNdisRssWriter().SetQueues(instanceId, queues, report, deviceName);
    }

    private void SetNdisRssExtraValues(string instanceId, NdisRssLayout.Layout layout, OperationReport? report = null, string? deviceName = null)
    {
        DemandRealWrite("SetNdisRssExtraValues");
        CreateNdisRssWriter().SetExtra(instanceId, layout, report, deviceName);
    }

    private void ClearNdisBaseCore(string instanceId, OperationReport? report = null, string? deviceName = null)
    {
        DemandRealWrite("ClearNdisBaseCore");
        CreateNdisRssWriter().ClearBaseCore(instanceId, report, deviceName);
    }

    private void ClearNdisRssQueues(string instanceId, OperationReport? report = null, string? deviceName = null)
    {
        DemandRealWrite("ClearNdisRssQueues");
        CreateNdisRssWriter().ClearQueues(instanceId, report, deviceName);
    }

    private void ClearNdisRssExtraValues(string instanceId, OperationReport? report = null, string? deviceName = null)
    {
        DemandRealWrite("ClearNdisRssExtraValues");
        CreateNdisRssWriter().ClearExtra(instanceId, report, deviceName);
    }
}
