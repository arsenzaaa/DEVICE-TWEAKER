using System.Globalization;
using System.Text.RegularExpressions;

namespace DeviceTweakerCS;

/// <summary>
/// Classifies USB host controllers by PCI VID/DID into CPU-direct vs chipset vs add-in.
/// Device ID tables adapted (MIT) from MariusHeier/cpu-direct-usb (Linux xhci-pci.c + pci.ids).
/// CHIP N here is an estimated controller-class hop, not a measured physical path
/// or latency value; downstream USB hubs are outside this classification.
/// </summary>
internal enum UsbChipOrigin
{
    Unknown = 0,
    CpuDirect = 1,
    Thunderbolt = 2,
    Chipset = 3,
    Addon = 4,
}

internal sealed record UsbChipPathInfo(
    UsbChipOrigin Origin,
    int BaseChipCount,
    string ShortLabel,
    string Platform,
    string UsbSpec,
    string ControllerName,
    string Vid,
    string Did)
{
    public string CompactTag => BaseChipCount < 0
        ? "CHIP ?"
        : BaseChipCount == 0
            ? "CHIP 0"
            : BaseChipCount == 1
                ? "CHIP 1"
                : $"CHIP {BaseChipCount}";

    public string OriginWord => Origin switch
    {
        UsbChipOrigin.CpuDirect => "CPU-direct",
        UsbChipOrigin.Thunderbolt => "Thunderbolt/USB4",
        UsbChipOrigin.Chipset => "chipset/PCH",
        UsbChipOrigin.Addon => "PCIe add-in",
        _ => "unclassified",
    };

    /// <summary>Info-box lines distinguish inferred topology from controller capability.</summary>
    public string DetailLine
    {
        get
        {
            string platform = string.IsNullOrWhiteSpace(Platform) ? string.Empty : $" | {Platform}";
            string topology = $"Topology: estimated {CompactTag} | {OriginWord}{platform}";
            string capability = string.IsNullOrWhiteSpace(UsbSpec)
                ? string.Empty
                : $"\nController capability: {UsbSpec}";
            return topology + capability;
        }
    }

    public string TooltipText =>
        $"Estimated topology: {CompactTag} ({OriginWord}).\n" +
        "This estimate is based on the controller PCI ID.\n" +
        "It does not represent measured latency or the negotiated device link.\n" +
        "USB hubs connected downstream are not included.\n\n" +
        $"Controller: {ControllerName}." +
        (string.IsNullOrWhiteSpace(Platform) ? string.Empty : $"\nPlatform: {Platform}.") +
        (string.IsNullOrWhiteSpace(UsbSpec) ? string.Empty : $"\nController capability: {UsbSpec}.") +
        (string.IsNullOrWhiteSpace(Vid) || string.IsNullOrWhiteSpace(Did) ? string.Empty : $"\nPCI ID: {Vid}:{Did}.");
}

internal static class UsbChipPath
{
    private static readonly Regex PciVid = new(@"VEN_([0-9A-Fa-f]{4})", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex PciDid = new(@"DEV_([0-9A-Fa-f]{4})", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    // Intel CPU-integrated / SoC USB (CHIP 0)
    private static readonly Dictionary<string, (string Name, string Platform, string Usb)> IntelCpu = new(StringComparer.OrdinalIgnoreCase)
    {
        ["8a13"] = ("Ice Lake Thunderbolt 3 USB", "Ice Lake (10th Gen)", "USB 3.2/TB3"),
        ["9a13"] = ("Tiger Lake-LP Thunderbolt 4 USB", "Tiger Lake (11th Gen)", "USB4/TB4"),
        ["9a17"] = ("Tiger Lake-H Thunderbolt 4 USB", "Tiger Lake-H (11th Gen)", "USB4/TB4"),
        ["461e"] = ("Alder Lake-P Thunderbolt 4 USB", "Alder Lake (12th Gen)", "USB4/TB4"),
        ["464e"] = ("Alder Lake-N Processor USB 3.2 xHCI", "Alder Lake-N", "USB 3.2"),
        ["a71e"] = ("Raptor Lake-P Thunderbolt 4 USB", "Raptor Lake (13th Gen)", "USB4/TB4"),
        ["7ec0"] = ("Meteor Lake-P Thunderbolt 4 USB", "Meteor Lake (Core Ultra)", "USB4/TB4"),
        ["a831"] = ("Lunar Lake-M Thunderbolt 4 USB", "Lunar Lake", "USB4/TB4"),
    };

    // Intel discrete Thunderbolt host (CHIP 0 - CPU-attached path)
    private static readonly Dictionary<string, (string Name, string Platform, string Usb)> IntelTb = new(StringComparer.OrdinalIgnoreCase)
    {
        ["5782"] = ("JHL9580 Thunderbolt 5 USB", "Barlow Ridge Host 80G", "USB4/TB5"),
        ["5785"] = ("JHL9540 Thunderbolt 4 USB", "Barlow Ridge Host 40G", "USB4/TB4"),
        ["1138"] = ("Thunderbolt 4 USB [Maple Ridge 4C]", "Maple Ridge 4C", "USB4/TB4"),
        ["1135"] = ("Thunderbolt 4 USB [Maple Ridge 2C]", "Maple Ridge 2C", "USB4/TB4"),
        ["0b27"] = ("Thunderbolt 4 USB [Goshen Ridge]", "Goshen Ridge", "USB4/TB4"),
        ["15e9"] = ("JHL7540 Thunderbolt 3 USB", "Titan Ridge 2C", "USB 3.1/TB3"),
        ["15ec"] = ("JHL7540 Thunderbolt 3 USB", "Titan Ridge 4C", "USB 3.1/TB3"),
        ["15b6"] = ("DSL6540 USB 3.1 [Alpine Ridge 4C]", "Alpine Ridge 4C", "USB 3.1/TB3"),
        ["15b5"] = ("DSL6340 USB 3.1 [Alpine Ridge 2C]", "Alpine Ridge 2C", "USB 3.1/TB3"),
    };

    // Intel PCH / chipset USB (CHIP 1)
    private static readonly Dictionary<string, (string Name, string Platform, string Usb)> IntelPch = new(StringComparer.OrdinalIgnoreCase)
    {
        ["7f6e"] = ("800 Series PCH USB 3.1 xHCI", "800 Series PCH", "USB 3.1"),
        ["7a60"] = ("Raptor Lake USB 3.2 Gen 2x2 xHCI", "700 Series PCH", "USB 3.2 Gen 2x2"),
        ["7ae0"] = ("Alder Lake-S PCH USB 3.2 Gen 2x2 xHCI", "600 Series PCH (Desktop)", "USB 3.2 Gen 2x2"),
        ["7ae1"] = ("Alder Lake-S PCH USB 3.2 xHCI", "600 Series PCH (Desktop)", "USB 3.2"),
        ["51ed"] = ("Alder Lake PCH USB 3.2 xHCI", "600 Series PCH", "USB 3.2"),
        ["54ed"] = ("Alder Lake-N PCH USB 3.2 xHCI", "Alder Lake-N PCH", "USB 3.2 Gen 2"),
        ["7e7d"] = ("Meteor Lake-P USB 3.2 xHCI", "Meteor Lake PCH", "USB 3.2 Gen 2"),
        ["777d"] = ("Arrow Lake USB 3.2 xHCI", "Arrow Lake", "USB 3.2"),
        ["a87d"] = ("Lunar Lake-M USB 3.2 xHCI", "Lunar Lake PCH", "USB 3.2 Gen 2"),
        ["a0ed"] = ("Tiger Lake-LP USB 3.2 xHCI", "500 Series PCH", "USB 3.2 Gen 2"),
        ["43ed"] = ("Tiger Lake-H USB 3.2 xHCI", "500 Series PCH-H", "USB 3.2 Gen 2"),
        ["a3af"] = ("Comet Lake PCH-V USB", "400 Series PCH", "USB 3.1"),
        ["02ed"] = ("Comet Lake PCH-LP USB 3.1 xHCI", "400 Series PCH-LP", "USB 3.1"),
        ["06ed"] = ("Comet Lake USB 3.1 xHCI", "400 Series PCH", "USB 3.1"),
        ["a36d"] = ("Cannon Lake PCH USB 3.1 xHCI", "300 Series PCH", "USB 3.1"),
        ["9ded"] = ("Cannon Point-LP USB 3.1 xHCI", "300 Series PCH-LP", "USB 3.1"),
        ["a2af"] = ("200 Series/Z370 USB 3.0 xHCI", "200 Series PCH", "USB 3.0"),
        ["a12f"] = ("100 Series/C230 USB 3.0 xHCI", "100 Series PCH", "USB 3.0"),
        ["9d2f"] = ("Sunrise Point-LP USB 3.0 xHCI", "100 Series PCH-LP", "USB 3.0"),
        ["8cb1"] = ("9 Series Chipset USB xHCI", "9 Series PCH", "USB 3.0"),
        ["8c31"] = ("8 Series/C220 USB xHCI", "8 Series PCH", "USB 3.0"),
        ["1e31"] = ("7 Series/C210 USB xHCI", "7 Series PCH", "USB 3.0"),
        ["8d31"] = ("C610/X99 USB xHCI", "X99/C610", "USB 3.0"),
    };

    // AMD CPU-integrated (CHIP 0)
    private static readonly Dictionary<string, (string Name, string Platform, string Usb)> AmdCpu = new(StringComparer.OrdinalIgnoreCase)
    {
        ["15b6"] = ("Raphael/Granite Ridge USB 3.1 xHCI", "Ryzen 7000/9000 (AM5)", "USB 3.1"),
        ["15b7"] = ("Raphael/Granite Ridge USB 3.1 xHCI", "Ryzen 7000/9000 (AM5)", "USB 3.1"),
        ["15b8"] = ("Raphael/Granite Ridge USB 2.0 xHCI", "Ryzen 7000/9000 (AM5)", "USB 2.0"),
        ["1587"] = ("Strix Halo USB 3.1 xHCI", "Strix Halo (Zen 5)", "USB 3.1"),
        ["1588"] = ("Strix Halo USB 3.1 xHCI", "Strix Halo (Zen 5)", "USB 3.1"),
        ["161d"] = ("Rembrandt USB4 xHCI", "Ryzen 6000 Mobile", "USB4"),
        ["15c4"] = ("Phoenix USB4/Thunderbolt NHI", "Ryzen 7040 Mobile", "USB4/TB"),
        ["1639"] = ("Renoir/Cezanne USB 3.1", "Ryzen 4000/5000 APU", "USB 3.1"),
        ["149c"] = ("Matisse USB 3.0 Host Controller", "Ryzen 3000/5000 Desktop", "USB 3.0"),
        ["148c"] = ("Starship USB 3.0 Host Controller", "EPYC Rome / TR 3rd Gen", "USB 3.0"),
        ["145f"] = ("Zeppelin USB 3.0 xHCI", "Ryzen 1000", "USB 3.0"),
        ["145c"] = ("Family 17h USB 3.0 Host Controller", "Ryzen 1000", "USB 3.0"),
    };

    // AMD chipset (CHIP 1)
    private static readonly Dictionary<string, (string Name, string Platform, string Usb)> AmdChipset = new(StringComparer.OrdinalIgnoreCase)
    {
        ["43fc"] = ("800 Series Chipset USB 3.x xHCI", "X870/B850 (AM5)", "USB 3.2"),
        ["43fd"] = ("800 Series Chipset USB 3.x xHCI", "X870/B850 (AM5)", "USB 3.2"),
        ["43f7"] = ("600 Series Chipset USB 3.2", "X670/B650 (AM5)", "USB 3.2"),
        ["43ee"] = ("500 Series Chipset USB 3.1 xHCI", "X570/B550 (AM4)", "USB 3.1"),
        ["43ec"] = ("A520 Series Chipset USB 3.1 xHCI", "A520 (AM4)", "USB 3.1"),
        ["43d5"] = ("400 Series Chipset USB 3.1 xHCI", "X470/B450 (AM4)", "USB 3.1"),
        ["43b9"] = ("X370 Series Chipset USB 3.1 xHCI", "X370 (AM4)", "USB 3.1"),
        ["43bb"] = ("300 Series Chipset USB 3.1 xHCI", "B350 (AM4)", "USB 3.1"),
        ["43bc"] = ("A320 USB 3.1 xHCI", "A320 (AM4)", "USB 3.1"),
        ["7814"] = ("FCH USB xHCI", "Legacy FCH", "USB 3.0"),
    };

    // Third-party add-in (CHIP 1+)
    private static readonly Dictionary<string, (string Name, string Platform, string Usb)> Addon = new(StringComparer.OrdinalIgnoreCase)
    {
        ["1b21:1042"] = ("ASM1042 USB 3.0", "ASMedia add-in", "USB 3.0"),
        ["1b21:1142"] = ("ASM1042A USB 3.0", "ASMedia add-in", "USB 3.0"),
        ["1b21:1242"] = ("ASM1142 USB 3.1", "ASMedia add-in", "USB 3.1 Gen 2"),
        ["1b21:2142"] = ("ASM2142/3142 USB 3.1", "ASMedia add-in", "USB 3.1 Gen 2"),
        ["1b21:3242"] = ("ASM3242 USB 3.2", "ASMedia add-in", "USB 3.2 Gen 2x2"),
        ["1b21:2426"] = ("ASM4242 USB 3.2 xHCI", "ASMedia add-in", "USB 3.2"),
        ["1106:3483"] = ("VL805/806 USB 3.0 xHCI", "VIA add-in", "USB 3.0"),
        ["1b73:1100"] = ("FL1100 USB 3.0", "Fresco Logic add-in", "USB 3.0"),
        ["1912:0015"] = ("uPD720202 USB 3.0", "Renesas add-in", "USB 3.0"),
    };

    private static readonly HashSet<string> KnownAddonVendors = new(StringComparer.OrdinalIgnoreCase)
    {
        "1b21", "1106", "1b73", "1912", "1b6f", "104c",
    };

    public static bool TryParsePciIds(string instanceId, out string vid, out string did)
    {
        vid = string.Empty;
        did = string.Empty;
        if (string.IsNullOrWhiteSpace(instanceId))
        {
            return false;
        }

        Match vidMatch = PciVid.Match(instanceId);
        Match didMatch = PciDid.Match(instanceId);
        if (!vidMatch.Success || !didMatch.Success)
        {
            return false;
        }

        vid = vidMatch.Groups[1].Value.ToLowerInvariant();
        did = didMatch.Groups[1].Value.ToLowerInvariant();
        return true;
    }

    public static UsbChipPathInfo Classify(string instanceId)
    {
        if (!TryParsePciIds(instanceId, out string vid, out string did))
        {
            return new UsbChipPathInfo(
                UsbChipOrigin.Unknown,
                -1,
                "CHIP ?",
                string.Empty,
                string.Empty,
                "Unknown USB controller",
                string.Empty,
                string.Empty);
        }

        string key = $"{vid}:{did}";

        if (vid is "8086" && IntelCpu.TryGetValue(did, out var cpu))
        {
            return Make(UsbChipOrigin.CpuDirect, 0, cpu, vid, did);
        }

        if (vid is "8086" && IntelTb.TryGetValue(did, out var tb))
        {
            return Make(UsbChipOrigin.Thunderbolt, 0, tb, vid, did);
        }

        if (vid is "8086" && IntelPch.TryGetValue(did, out var pch))
        {
            return Make(UsbChipOrigin.Chipset, 1, pch, vid, did);
        }

        if (vid is "1022" && AmdCpu.TryGetValue(did, out var amdCpu))
        {
            return Make(UsbChipOrigin.CpuDirect, 0, amdCpu, vid, did);
        }

        if (vid is "1022" && AmdChipset.TryGetValue(did, out var amdCs))
        {
            return Make(UsbChipOrigin.Chipset, 1, amdCs, vid, did);
        }

        if (Addon.TryGetValue(key, out var addon))
        {
            return Make(UsbChipOrigin.Addon, 1, addon, vid, did);
        }

        if (KnownAddonVendors.Contains(vid))
        {
            string vendor = vid switch
            {
                "1b21" => "ASMedia",
                "1106" => "VIA",
                "1b73" => "Fresco Logic",
                "1912" => "Renesas",
                "1b6f" => "Etron",
                "104c" => "Texas Instruments",
                _ => "Add-in",
            };
            return Make(
                UsbChipOrigin.Addon,
                1,
                ($"{vendor} USB controller", "PCIe add-in", "USB 3.x"),
                vid,
                did);
        }

        if (vid is "8086")
        {
            return Make(
                UsbChipOrigin.Unknown,
                -1,
                ("Intel USB controller (unclassified PCI ID)", $"Intel vendor, DID:{did}", "unknown"),
                vid,
                did);
        }

        if (vid is "1022")
        {
            return Make(
                UsbChipOrigin.Unknown,
                -1,
                ("AMD USB controller (unclassified PCI ID)", $"AMD vendor, DID:{did}", "unknown"),
                vid,
                did);
        }

        return Make(
            UsbChipOrigin.Unknown,
            1,
            ("Unknown USB controller", $"VID:{vid} DID:{did}", "?"),
            vid,
            did);
    }

    public static string DeviceParametersPath(string instanceId)
        => $@"SYSTEM\CurrentControlSet\Enum\{instanceId}\Device Parameters";

    public static string? TryReadSelectiveSuspendLabel(string instanceId)
    {
        if (!TryReadSelectiveSuspendEnabled(instanceId, out int? enabled) || enabled is null)
        {
            return null;
        }

        return enabled.Value == 0 ? "off" : "on";
    }

    /// <summary>
    /// Reads PCI USB controller Device Parameters\SelectiveSuspendEnabled.
    /// Returns false on access failure; enabled is null when the value is missing/unknown.
    /// </summary>
    public static bool TryReadSelectiveSuspendEnabled(string instanceId, out int? enabled)
    {
        enabled = null;
        if (string.IsNullOrWhiteSpace(instanceId) || !instanceId.StartsWith("PCI\\", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        try
        {
            using Microsoft.Win32.RegistryKey? key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(
                DeviceParametersPath(instanceId),
                writable: false);
            object? value = key?.GetValue("SelectiveSuspendEnabled");
            if (value is null)
            {
                return true;
            }

            int numeric = value switch
            {
                int i => i,
                uint u => unchecked((int)u),
                string s when int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) => parsed,
                _ => -1,
            };

            if (numeric is 0 or 1)
            {
                enabled = numeric;
            }

            return true;
        }
        catch
        {
            return false;
        }
    }

    public static void SetSelectiveSuspendEnabled(string instanceId, bool enabled)
    {
        if (string.IsNullOrWhiteSpace(instanceId) || !instanceId.StartsWith("PCI\\", StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException("Selective Suspend applies only to PCI USB controllers.");
        }

        string path = DeviceParametersPath(instanceId);
        using Microsoft.Win32.RegistryKey? key = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(path);
        if (key is null)
        {
            throw new InvalidOperationException($"Could not open HKLM\\{path} for writing.");
        }

        key.SetValue("SelectiveSuspendEnabled", enabled ? 1 : 0, Microsoft.Win32.RegistryValueKind.DWord);
    }

    private static UsbChipPathInfo Make(
        UsbChipOrigin origin,
        int chips,
        (string Name, string Platform, string Usb) data,
        string vid,
        string did)
    {
        string shortLabel = chips <= 0 ? "CHIP 0" : chips == 1 ? "CHIP 1" : $"CHIP {chips}";
        return new UsbChipPathInfo(origin, chips, shortLabel, data.Platform, data.Usb, data.Name, vid, did);
    }
}
