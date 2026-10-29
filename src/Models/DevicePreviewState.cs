namespace DeviceTweakerCS;

internal sealed class TestDeviceState
{
    public bool MsiEnabled { get; set; } = true;
    public int? MsiLimit { get; set; }
    public int? Priority { get; set; }
    public int Policy { get; set; }
    public ulong AffinityMask { get; set; }
    public int? RssBaseCore { get; set; }
    public int RssQueues { get; set; } = 1;
    public string NdisMode { get; set; } = "RSS";
    public bool? PowerSavingEnabled { get; set; }
    public string ImodValue { get; set; } = "0x0";
    public string NicItrValue { get; set; } = "default";

    public TestDeviceState Clone() => new()
    {
        MsiEnabled = MsiEnabled,
        MsiLimit = MsiLimit,
        Priority = Priority,
        Policy = Policy,
        AffinityMask = AffinityMask,
        RssBaseCore = RssBaseCore,
        RssQueues = RssQueues,
        NdisMode = NdisMode,
        PowerSavingEnabled = PowerSavingEnabled,
        ImodValue = ImodValue,
        NicItrValue = NicItrValue,
    };
}

