using System.Windows.Forms;

namespace DeviceTweakerCS;

internal enum DeviceKind
{
    USB,
    GPU,
    AUDIO,
    NET_NDIS,
    NET_CX,
    STOR,
    OTHER,
}

internal sealed record CpuVendorInfo(string Name, string Vendor);

internal sealed record CpuLpInfo(
    int Group,
    int LP,
    int Core,
    int LLC,
    int NUMA,
    int EffClass,
    int LocalIndex = -1,
    int CpuSetId = -1);

internal sealed class CpuTopology
{
    public CpuTopology(List<CpuLpInfo> lps)
    {
        LPs = lps;
        ByCore = lps
            .GroupBy(lp => MakeCoreKey(lp.Group, lp.Core))
            .ToDictionary(g => g.Key, g => g.OrderBy(x => x.LP).ToList());

        ByLLC = lps
            .GroupBy(lp => MakeLlcKey(lp.Group, lp.LLC))
            .ToDictionary(g => g.Key, g => g.OrderBy(x => x.LP).ToList());
    }

    public static int MakeGroupKey(int group, int id)
    {
        int g = group & 0xFFFF;
        int v = id & 0xFFFF;
        return (g << 16) | v;
    }

    public static int MakeCoreKey(int group, int core)
    {
        return MakeGroupKey(group, core);
    }

    public static int MakeLlcKey(int group, int llc)
    {
        if (llc < 0)
        {
            return -1;
        }

        return MakeGroupKey(group, llc);
    }

    public List<CpuLpInfo> LPs { get; }
    public Dictionary<int, List<CpuLpInfo>> ByCore { get; }
    public Dictionary<int, List<CpuLpInfo>> ByLLC { get; }
    public int Logical => LPs.Count;
    public int PhysicalCores => ByCore.Count;
}

internal sealed class CpuInfo
{
    public required CpuTopology Topology { get; init; }
    public required Dictionary<int, int> CcdMap { get; init; }
}

internal sealed class DeviceInfo
{
    public required string Name { get; init; }
    public required string InstanceId { get; init; }
    public required string Class { get; init; }
    public required string RegBase { get; init; }
    public required DeviceKind Kind { get; init; }
    public string UsbRoles { get; init; } = string.Empty;
    public string AudioEndpoints { get; init; } = string.Empty;
    public string StorageTag { get; init; } = string.Empty;
    public bool Wifi { get; init; }
    public bool UsbIsXhci { get; init; }
    public bool UsbHasDevices { get; init; }
    public bool IsTestDevice { get; init; }
}

internal sealed class DeviceBlock
{
    public required DeviceInfo Device { get; init; }
    public required DeviceKind Kind { get; init; }
    public required Panel Group { get; init; }
    public required List<CheckBox> CpuBoxes { get; init; }
    public required Label AffinityLabel { get; init; }
    public required Label IrqLabel { get; init; }
    public required ComboBox MsiCombo { get; init; }
    public required TextBox LimitBox { get; init; }
    public required ComboBox PrioCombo { get; init; }
    public required ComboBox PolicyCombo { get; init; }
    public required Label PolicyLabel { get; init; }
    public required CheckBox ImodAutoCheck { get; init; }
    public required TextBox ImodBox { get; init; }
    public required Label ImodDefaultLabel { get; init; }
    public required Label InfoLabel { get; init; }

    public ulong AffinityMask { get; set; }
    public int? IrqCount { get; set; }
    public int SuppressCpuEvents { get; set; }
}

internal sealed record UsbControllerInfo(string ControllerPNPID, string ControllerName);

internal sealed class HidDeviceInfo
{
    public required string ProductString { get; init; }
    public required string DevicePath { get; init; }
    public string? DeviceType { get; init; }
    public string? DeviceInstanceId { get; init; }
    public List<UsbControllerInfo> UsbControllers { get; init; } = [];
    public int? UsagePage { get; init; }
    public int? UsageId { get; init; }
}

internal sealed record WmiPnPDevice(
    string InstanceId,
    string? Name,
    string? Class,
    string? Service,
    string? Status,
    int? ConfigManagerErrorCode);

internal sealed record WmiPhysicalDisk(
    string Name,
    ushort BusType,
    ushort MediaType);

internal sealed class ReservedCpuEntry
{
    public required CheckBox Control { get; init; }
    public required int Ccd { get; init; }
    public required int Eff { get; init; }
    public required int Index { get; init; }
}

internal sealed class ReservedCpuPanelTag
{
    public required Panel InnerPanel { get; init; }
    public required Label Title { get; init; }
    public required Label Description { get; init; }
    public required List<ReservedCpuEntry> Meta { get; init; }
    public required Label PathLabel { get; init; }
}
