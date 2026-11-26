using System.Drawing;
using System.Text.RegularExpressions;

namespace DeviceTweakerCS;

// Known USB/GPU vendors retain their colours. Network and storage use distinct
// category accents; operation feedback keeps its own status palette.
internal static class DeviceBadgePalette
{
    internal static readonly Color Neutral = Color.FromArgb(240, 240, 240);
    internal static readonly Color Network = Color.FromArgb(52, 201, 189);
    internal static readonly Color Storage = Color.FromArgb(204, 209, 216);
    internal static readonly Color Nvidia = Color.FromArgb(118, 185, 0);
    internal static readonly Color Amd = Color.FromArgb(235, 60, 60);
    internal static readonly Color Intel = Color.FromArgb(88, 192, 255);

    internal static Color GetUsbColor(string instanceId)
    {
        // The controller may be an add-in device from another vendor. Never infer
        // its colour from the CPU, the connected peripheral or a localized name.
        Match vendor = Regex.Match(instanceId ?? string.Empty,
            @"^PCI\\VEN_([0-9A-F]{4})(?:&|\\|$)",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        return vendor.Groups[1].Value.ToUpperInvariant() switch
        {
            "1022" or "1002" => Amd,
            "8086" => Intel,
            _ => Neutral,
        };
    }
}
