using System.Text.Json.Serialization;

namespace DeviceTweakerCS;

internal enum TestSandboxOperation
{
    Auto,
    Apply,
    SafeReset,
    Backup,
    Restore,
}

internal enum TestSandboxFault
{
    None,
    CpuTopologyUnavailable,
    CpuMapUnavailable,
    BackupAccessDenied,
    RegistryWriteDenied,
    UsbPowerWriteFailed,
    WmiTimeout,
    KduMissing,
    KduTimeout,
    KduExitCode,
    KduDeviceUnavailable577,
    ImodReadFailed,
    ImodWriteFailed,
    ImodVerificationMismatch,
    NicItrReadFailed,
    NicItrWriteFailed,
    CorruptBackup,
    RestoreWriteFailed,
    RollbackFailed,
    RefreshFailed,
}

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

internal sealed class TestScenarioCpu
{
    public string Name { get; set; } = string.Empty;
    public int LogicalCount { get; set; }
    public bool SmtEnabled { get; set; }
    public bool UseHyperThreadingLabel { get; set; }
    public List<int> ECoreLps { get; set; } = [];
    public Dictionary<int, int> CoreMap { get; set; } = [];
    public Dictionary<int, int> CcdMap { get; set; } = [];
    public Dictionary<int, int> CcxMap { get; set; } = [];
    public Dictionary<int, int> CppcRatings { get; set; } = [];
}

internal sealed class TestScenarioDevice
{
    public DeviceKind Kind { get; set; }
    public string Name { get; set; } = string.Empty;
    public string InstanceId { get; set; } = string.Empty;
    public string UsbRoles { get; set; } = string.Empty;
    public string AudioEndpoints { get; set; } = string.Empty;
    public string StorageTag { get; set; } = string.Empty;
    public bool Wifi { get; set; }
    public bool UsbIsXhci { get; set; }
    public bool UsbHasDevices { get; set; }
    public bool IntegratedGpu { get; set; }
    public int? IrqCount { get; set; }
    public string MsiStatus { get; set; } = "Auto";
    public string? UsbSelectiveSuspend { get; set; }
    public string? NicPowerSaving { get; set; }
    public TestDeviceState InitialState { get; set; } = new();
}

internal sealed class TestSandboxScenario
{
    public int Version { get; set; } = 1;
    public string Name { get; set; } = "Untitled scenario";
    [JsonConverter(typeof(JsonStringEnumConverter))]
    public TestSandboxOperation Operation { get; set; } = TestSandboxOperation.Auto;
    [JsonConverter(typeof(JsonStringEnumConverter))]
    public TestSandboxFault Fault { get; set; }
    public bool OptimizeUsbImod { get; set; } = true;
    public TestScenarioCpu Cpu { get; set; } = new();
    public List<TestScenarioDevice> Devices { get; set; } = [];
}

internal sealed class TestSandboxRunResult
{
    public required OperationReport Report { get; init; }
    public required string ScenarioName { get; init; }
    public required TestSandboxOperation Operation { get; init; }
    public required TestSandboxFault Fault { get; init; }
    public int AssertionsPassed { get; set; }
    public int AssertionsFailed { get; set; }
    public int VirtualWrites { get; set; }
    public int BlockedRealWrites { get; set; }
    public bool Passed => AssertionsFailed == 0;
}

internal sealed class TestSandboxRuntime
{
    public bool Active { get; set; }
    public TestSandboxFault Fault { get; set; }
    public bool FaultConsumed { get; set; }
    public int VirtualWrites { get; set; }
    public int BlockedRealWrites { get; set; }
    public Dictionary<string, TestDeviceState>? Backup { get; set; }

    public void Begin(TestSandboxFault fault)
    {
        Active = true;
        Fault = fault;
        FaultConsumed = false;
        VirtualWrites = 0;
        BlockedRealWrites = 0;
    }

    public void End()
    {
        Active = false;
        Fault = TestSandboxFault.None;
        FaultConsumed = false;
    }
}
